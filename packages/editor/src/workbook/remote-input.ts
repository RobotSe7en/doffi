import { DocumentError } from "@doffi/core";
import type { DocumentSelection, TransactionPrepareOptions } from "@doffi/core";
import type { WorkbookInputState } from "./input-controller.js";
import { copyWorkbookInputSelection } from "./input-selection.js";
import type { WorkbookInputSelection } from "./input-selection.js";
import type { WorkbookTextInputMode } from "./text-input.js";
import type { WorkbookSessionCandidate, WorkbookSessionConnection, WorkbookSessionReceipt } from "./session-transport.js";

export interface WorkbookRemoteInputState extends WorkbookInputState { readonly applying: boolean; readonly selecting: boolean }
export interface WorkbookRemoteInputController {
  state(): WorkbookRemoteInputState;
  select(selection: DocumentSelection): Promise<void>;
  setText(text: string): void;
  setMode(mode: WorkbookTextInputMode): void;
  compositionStart(): void;
  compositionEnd(text: string): void;
  apply(options?: TransactionPrepareOptions): Promise<WorkbookSessionReceipt>;
  cancel(): void;
  dispose(): void;
}
/** The main thread owns only a transient draft. Reads, input conversion,
 * rich source and every history step belong to the host's Worker session. */
export function createWorkbookRemoteInputController(client: WorkbookSessionConnection, options: { readonly onChange?: () => void } = {}): WorkbookRemoteInputController {
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => key !== "onChange") || options.onChange !== undefined && typeof options.onChange !== "function") throw new DocumentError("InvalidOperation", "Invalid remote workbook input options.");
  let selection: WorkbookInputSelection | undefined, text = "", original = "", mode: WorkbookTextInputMode = "auto", composing = false;
  let disposed = false, applying = false, committing = false, version = 0, selectionAbort: AbortController | undefined, preparationAbort: AbortController | undefined;
  const invalid = (message: string): never => { throw new DocumentError("InvalidOperation", message); };
  const check = (): void => { if (disposed) invalid("This remote workbook input controller has been disposed."); };
  const notify = (): void => { if (!disposed) try { options.onChange?.(); } catch { /* Reporting never owns a source mutation. */ } };
  const dirty = (): boolean => text !== original || mode !== "auto";
  const mutate = (): void => { check(); version++; selectionAbort?.abort(); if (!committing) preparationAbort?.abort(); };
  function setText(value: string): void {
    check(); if (!selection || typeof value !== "string") invalid("Workbook text input requires a selected cell and string value.");
    if (value.length > 32767 && !(mode === "auto" && value.startsWith("'") && value.length === 32768)) throw new DocumentError("ResourceLimitExceeded", "Workbook input exceeds its text budget.");
    if (text !== value) { mutate(); text = value; notify(); }
  }
  return Object.freeze({
    state() { check(); return Object.freeze({ ...(selection ? { selection } : {}), text, mode, dirty: dirty(), composing, applying, selecting: !!selectionAbort && !selectionAbort.signal.aborted }); },
    async select(input: DocumentSelection) {
      check(); const next = copyWorkbookInputSelection(input, client.state());
      if (JSON.stringify(selection) === JSON.stringify(next)) return;
      if (applying || composing || dirty()) invalid("Finish or cancel the current input before changing its selection.");
      selectionAbort?.abort(); const abort = new AbortController(), ticket = ++version; selectionAbort = abort; notify();
      try {
        const value = await client.readInput(next.anchor.sheetId, next.anchor.ranges[0]!.from, undefined, { signal: abort.signal });
        check(); if (abort.signal.aborted || ticket !== version) throw new DocumentError("OperationCancelled", "Workbook input selection was superseded.");
        if (value.revision !== next.baseRevision || value.documentId !== next.documentId || client.state().revision !== next.baseRevision) throw new DocumentError("RevisionConflict", "Workbook input source changed while selecting.");
        selection = next; text = original = value.text; mode = "auto"; composing = false;
      } finally { if (selectionAbort === abort) selectionAbort = undefined; notify(); }
    },
    setText,
    setMode(value: WorkbookTextInputMode) {
      check(); if (!selection || !["auto", "text", "number"].includes(value)) invalid("Invalid workbook input mode.");
      if (mode === value) return; mutate(); if (mode === "auto" && value !== "auto" && text.startsWith("'")) text = text.slice(1); mode = value; notify();
    },
    compositionStart() { check(); if (!selection) invalid("Composition requires a selected cell."); mutate(); composing = true; notify(); },
    compositionEnd(value: string) { check(); composing = false; setText(value); notify(); },
    async apply(preparation?: TransactionPrepareOptions) {
      check(); const selected = selection;
      if (!selected) throw new DocumentError("InvalidOperation", "Workbook input requires a selected cell.");
      if (composing || applying || selectionAbort && !selectionAbort.signal.aborted) invalid("Finish the active input or selection before applying.");
      const submitted = text, submittedMode = mode, ticket = version, abort = new AbortController();
      const guard = (): void => {
        check(); if (abort.signal.aborted || ticket !== version) throw new DocumentError("OperationCancelled", "The workbook input draft changed before application.");
        const current = client.state(); if (selected.documentId !== current.documentId || selected.baseRevision !== current.revision) throw new DocumentError("RevisionConflict", "Workbook input belongs to another source revision.");
      };
      let candidate: WorkbookSessionCandidate | undefined, consumed = false; applying = true; preparationAbort = abort; notify();
      try {
        guard(); const at = selected.anchor.ranges[0]!.from;
        candidate = await client.prepare({ kind: "input", documentId: selected.documentId, revision: selected.baseRevision, sheetId: selected.anchor.sheetId, at, text: submitted, mode: submittedMode }, preparation, { signal: abort.signal });
        guard(); const canonical = await client.readInput(selected.anchor.sheetId, at, candidate, { signal: abort.signal }); guard();
        committing = true; const receipt = await client.commit(candidate); consumed = true;
        if (!disposed) {
          selection = Object.freeze({ ...selected, baseRevision: receipt.revision }); original = canonical.text;
          if (ticket === version) { text = canonical.text; mode = "auto"; composing = false; }
          // Continued typing after this controller sent its own commit is a
          // new transient draft. Preserve it against that known receipt;
          // any later external revision still conflicts at the next apply.
        }
        return receipt;
      } finally {
        if (candidate && !consumed) try { await client.discard(candidate); } catch { /* Host disposal or a changed source can consume this lease. */ }
        if (preparationAbort === abort) preparationAbort = undefined;
        committing = applying = false; notify();
      }
    },
    cancel() { check(); if (committing) invalid("The workbook input is awaiting its commit outcome."); mutate(); text = original; mode = "auto"; composing = false; notify(); },
    dispose() { if (disposed) return; disposed = true; version++; selectionAbort?.abort(); preparationAbort?.abort(); selection = undefined; text = original = ""; composing = false; },
  });
}
