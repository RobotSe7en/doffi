import { captureWorkbookModel, DocumentError } from "@doffi/core";
import type { DocumentSelection, TransactionPrepareOptions, TransactionPreparationExecution, TransactionReceipt, WorkbookCandidateChange, WorkbookSession } from "@doffi/core";
import { createWorkbookInputTransaction, readWorkbookInputText } from "./input-source.js";
import type { WorkbookTextInputMode } from "./text-input.js";
import { copyWorkbookInputSelection } from "./input-selection.js";

export interface WorkbookInputState {
  readonly selection?: DocumentSelection;
  readonly text: string;
  readonly mode: WorkbookTextInputMode;
  readonly dirty: boolean;
  readonly composing: boolean;
}
export interface WorkbookInputController {
  state(): WorkbookInputState;
  select(selection: DocumentSelection): void;
  setText(text: string): void;
  setMode(mode: WorkbookTextInputMode): void;
  compositionStart(): void;
  compositionEnd(text: string): void;
  apply(options?: TransactionPrepareOptions): TransactionReceipt;
  applyAsync(options?: TransactionPrepareOptions, execution?: TransactionPreparationExecution): Promise<TransactionReceipt>;
  cancel(): void;
  dispose(): void;
}
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
/** Transient human input owns neither source cells nor history. Every apply
 * prepares one revision-bound core candidate; failures retain the draft. */
export function createWorkbookInputController(session: WorkbookSession): WorkbookInputController {
  if (session.kind !== "workbook") invalid("Workbook input requires its owning workbook session.");
  let selection: DocumentSelection | undefined, text = "", original = "", mode: WorkbookTextInputMode = "auto", composing = false, disposed = false;
  let version = 0, pending: AbortController | undefined, committing = false;
  const check = (): void => { if (disposed) invalid("This workbook input controller has been disposed."); };
  const dirty = (): boolean => text !== original || mode !== "auto";
  const mutate = (): void => { version++; if (!committing) pending?.abort(); };
  function read(): void {
    if (!selection || selection.anchor.kind !== "cells") invalid("Workbook input requires a cell selection.");
    const at = selection.anchor.ranges[0]!.from, model = captureWorkbookModel(session), cell = model.getCell(selection.anchor.sheetId, at.row, at.column);
    text = readWorkbookInputText(cell);
    original = text;
    mode = "auto"; composing = false;
  }
  function setText(value: string): void {
    check(); if (!selection || typeof value !== "string") invalid("Workbook text input requires a selected cell and string value.");
    if (value.length > 32767 && !(mode === "auto" && value.startsWith("'") && value.length === 32768)) throw new DocumentError("ResourceLimitExceeded", "Workbook input exceeds its text budget.");
    if (text !== value) { mutate(); text = value; }
  }
  function inputTransaction() {
    check(); if (!selection || selection.anchor.kind !== "cells") invalid("Workbook input requires a selected cell.");
    if (composing) invalid("An active composition must finish before a document transaction.");
    const selected = selection, at = selection.anchor.ranges[0]!.from, sheetId = selection.anchor.sheetId, model = captureWorkbookModel(session), manifest = model.snapshot().manifest;
    if (selected.documentId !== manifest.documentId || selected.baseRevision !== manifest.revision) throw new DocumentError("RevisionConflict", "Input selection belongs to another workbook revision.");
    return { selected, at, sheetId, transaction: createWorkbookInputTransaction(model, sheetId, at, text, mode) };
  }
  function commitInput(candidate: WorkbookCandidateChange, input: ReturnType<typeof inputTransaction>, ticket: number): TransactionReceipt {
    const canonical = readWorkbookInputText(candidate.model.getCell(input.sheetId, input.at.row, input.at.column));
    committing = true;
    try {
      const receipt = session.commit(candidate);
      if (!disposed) {
        selection = Object.freeze({ ...input.selected, baseRevision: receipt.revision }); original = canonical;
        if (version === ticket) { text = canonical; mode = "auto"; composing = false; }
      }
      return receipt;
    } finally { committing = false; }
  }
  return Object.freeze({
    state() { check(); return Object.freeze({ ...(selection ? { selection } : {}), text, mode, dirty: dirty(), composing }); },
    select(input: DocumentSelection) {
      check();
      const model = captureWorkbookModel(session), manifest = model.snapshot().manifest;
      const next = copyWorkbookInputSelection(input, manifest);
      if (!model.snapshot().metadata.workbook.sheets.some(sheet => sheet.sheetId === next.anchor.sheetId)) invalid("Input selection requires an existing worksheet.");
      if (JSON.stringify(selection) === JSON.stringify(next)) return;
      if (pending || composing || dirty()) invalid("Finish or cancel the current cell input before changing its selection.");
      mutate(); selection = next; read();
    },
    setText,
    setMode(value: WorkbookTextInputMode) {
      check(); if (!selection || !["auto", "text", "number"].includes(value)) invalid("Invalid workbook input mode.");
      if (mode === value) return; mutate();
      if (mode === "auto" && value !== "auto" && text.startsWith("'")) text = text.slice(1);
      mode = value;
    },
    compositionStart() { check(); if (!selection) invalid("Composition requires a selected cell."); mutate(); composing = true; },
    compositionEnd(value: string) { check(); mutate(); composing = false; setText(value); },
    apply(options?: TransactionPrepareOptions) {
      if (pending && !pending.signal.aborted) invalid("Another workbook input is preparing.");
      const input = inputTransaction(), ticket = version, candidate = session.prepare(input.transaction, options);
      if (disposed || version !== ticket) { session.discard(candidate); throw new DocumentError("OperationCancelled", "Workbook input changed while preparing."); }
      return commitInput(candidate, input, ticket);
    },
    async applyAsync(options?: TransactionPrepareOptions, execution: TransactionPreparationExecution = {}) {
      if (pending) invalid("Another workbook input is preparing.");
      if (!execution || typeof execution !== "object" || Array.isArray(execution) || Object.keys(execution).some(key => !["signal", "yield", "onProgress"].includes(key))
        || execution.signal !== undefined && (!execution.signal || typeof execution.signal.aborted !== "boolean")
        || execution.yield !== undefined && typeof execution.yield !== "function" || execution.onProgress !== undefined && typeof execution.onProgress !== "function") invalid("Invalid workbook input execution options.");
      const input = inputTransaction(), ticket = version, abort = new AbortController(), config = { ...execution }, external = config.signal as AbortSignal | undefined;
      const stop = (): void => abort.abort(); pending = abort;
      if (external?.aborted) stop(); else external?.addEventListener?.("abort", stop, { once: true });
      let candidate: WorkbookCandidateChange | undefined, consumed = false;
      const guard = (): void => { if (disposed || abort.signal.aborted || config.signal?.aborted || version !== ticket) throw new DocumentError("OperationCancelled", "Workbook input changed while preparing."); };
      try {
        guard(); candidate = await session.prepareAsync(input.transaction, options, { ...config, signal: abort.signal,
          yield: async () => { guard(); if (config.yield) await config.yield(); else await new Promise<void>(resolve => setTimeout(resolve, 0)); guard(); } });
        guard(); const receipt = commitInput(candidate, input, ticket); consumed = true; return receipt;
      } finally {
        if (candidate && !consumed) try { session.discard(candidate); } catch { /* Source changes or disposal can release this candidate. */ }
        external?.removeEventListener?.("abort", stop); if (pending === abort) pending = undefined;
      }
    },
    cancel() { check(); if (committing) invalid("The workbook input is awaiting its commit outcome."); mutate(); text = original; mode = "auto"; composing = false; },
    dispose() { if (disposed) return; disposed = true; mutate(); selection = undefined; text = original = ""; composing = false; },
  });
}
