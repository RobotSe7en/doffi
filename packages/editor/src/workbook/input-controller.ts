import { captureWorkbookModel, DocumentError } from "@doffi/core";
import type { DocumentSelection, TransactionPrepareOptions, TransactionReceipt, WorkbookSession } from "@doffi/core";
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
  cancel(): void;
  dispose(): void;
}
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
/** Transient human input owns neither source cells nor history. Every apply
 * prepares one revision-bound core candidate; failures retain the draft. */
export function createWorkbookInputController(session: WorkbookSession): WorkbookInputController {
  if (session.kind !== "workbook") invalid("Workbook input requires its owning workbook session.");
  let selection: DocumentSelection | undefined, text = "", original = "", mode: WorkbookTextInputMode = "auto", composing = false, disposed = false;
  const check = (): void => { if (disposed) invalid("This workbook input controller has been disposed."); };
  const dirty = (): boolean => text !== original || mode !== "auto";
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
    text = value;
  }
  return Object.freeze({
    state() { check(); return Object.freeze({ ...(selection ? { selection } : {}), text, mode, dirty: dirty(), composing }); },
    select(input: DocumentSelection) {
      check();
      const model = captureWorkbookModel(session), manifest = model.snapshot().manifest;
      const next = copyWorkbookInputSelection(input, manifest);
      if (!model.snapshot().metadata.workbook.sheets.some(sheet => sheet.sheetId === next.anchor.sheetId)) invalid("Input selection requires an existing worksheet.");
      if (JSON.stringify(selection) === JSON.stringify(next)) return;
      if (composing || dirty()) invalid("Finish or cancel the current cell input before changing its selection.");
      selection = next; read();
    },
    setText,
    setMode(value: WorkbookTextInputMode) {
      check(); if (!selection || !["auto", "text", "number"].includes(value)) invalid("Invalid workbook input mode.");
      if (mode === "auto" && value !== "auto" && text.startsWith("'")) text = text.slice(1);
      mode = value;
    },
    compositionStart() { check(); if (!selection) invalid("Composition requires a selected cell."); composing = true; },
    compositionEnd(value: string) { check(); composing = false; setText(value); },
    apply(options?: TransactionPrepareOptions) {
      check(); if (!selection || selection.anchor.kind !== "cells") invalid("Workbook input requires a selected cell.");
      if (composing) invalid("An active composition must finish before a document transaction.");
      const at = selection.anchor.ranges[0]!.from, model = captureWorkbookModel(session), manifest = model.snapshot().manifest;
      if (selection.documentId !== manifest.documentId || selection.baseRevision !== manifest.revision) throw new DocumentError("RevisionConflict", "Input selection belongs to another workbook revision.");
      const candidate = session.prepare(createWorkbookInputTransaction(model, selection.anchor.sheetId, at, text, mode), options);
      const receipt = session.commit(candidate);
      selection = Object.freeze({ ...selection, baseRevision: receipt.revision }); read(); return receipt;
    },
    cancel() { check(); text = original; mode = "auto"; composing = false; },
    dispose() { if (disposed) return; disposed = true; selection = undefined; text = original = ""; composing = false; },
  });
}
