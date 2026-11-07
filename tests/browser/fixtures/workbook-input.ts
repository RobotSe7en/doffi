import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentSelection, DocumentTransaction, WorkbookSession } from "@doffi/core";
import { bindWorkbookTextInput, createWorkbookInputController, createWorkbookPasteTransaction } from "@doffi/editor";
import type { WorkbookInputController, WorkbookTextBinding } from "@doffi/editor";

let session: WorkbookSession, controller: WorkbookInputController, binding: WorkbookTextBinding, input: HTMLTextAreaElement;
const errors: { code: string; text: string }[] = [], navigation: { direction: string; extend: boolean }[] = [];
function selection(row = 0, column = 0, toRow = row, toColumn = column): DocumentSelection {
  const manifest = session.snapshot().manifest;
  return { documentId: manifest.documentId, baseRevision: manifest.revision, anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row, column }, to: { row: toRow, column: toColumn } }] } };
}
function renew(): void {
  const old = controller.state().selection!; controller.cancel(); controller.select({ ...old, baseRevision: session.snapshot().manifest.revision }); binding.refresh(); binding.focus();
}
function perform(transaction: DocumentTransaction): void { session.commit(session.prepare(transaction)); renew(); }
export function mount(row = 0, column = 0, toRow = row, toColumn = column): void {
  binding?.dispose(); controller?.dispose(); session?.dispose(); input?.remove(); errors.length = navigation.length = 0;
  session = createDocumentSession(openDocument(createWorkbook({ documentId: "native-workbook-input" }), { kind: "workbook" }));
  controller = createWorkbookInputController(session); controller.select(selection(row, column, toRow, toColumn));
  input = document.createElement("textarea"); input.setAttribute("aria-label", "工作簿输入"); document.body.append(input);
  binding = bindWorkbookTextInput(input, controller, {
    onNavigate(direction, extend) { navigation.push({ direction, extend }); binding.focus(); },
    onHistory(direction) { session[direction](); renew(); },
    onClear(selected) {
      if (selected.anchor.kind !== "cells") throw Error("Expected cells");
      perform({ protocolVersion: 1, documentId: selected.documentId, baseRevision: selected.baseRevision, operations: [{ op: "clearWorkbookRanges", sheetId: selected.anchor.sheetId, ranges: selected.anchor.ranges.map(range => ({ from: { ...range.from }, to: { ...range.to } })), mode: "contents" }] });
    },
    onPaste(text, selected) {
      if (selected.anchor.kind !== "cells") throw Error("Expected cells");
      const transaction = createWorkbookPasteTransaction(session.snapshot(), selected.anchor.sheetId, selected.anchor.ranges[0]!.from, text);
      perform({ ...transaction, documentId: selected.documentId, baseRevision: selected.baseRevision });
    },
    onError(error, text) { errors.push({ code: error.code, text }); },
  });
  binding.focus();
}
export function inspect() {
  const snapshot = session.snapshot();
  return { revision: snapshot.manifest.revision, history: session.history(), state: controller.state(), errors: [...errors], navigation: [...navigation],
    cells: [[0, 0], [255, 31], [255, 32], [256, 31], [256, 32]].map(([row, column]) => ({ row, column, cell: session.getCell("sheet-1", row!, column!) })), value: input.value };
}
export function select(row: number, column: number): void { controller.cancel(); controller.select(selection(row, column)); binding.refresh(); binding.focus(); }
export function external(text: string): void {
  const snapshot = session.snapshot(); session.commit(session.prepare({ protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision,
    operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: text } } }] }] }));
}
export function seedText(text: string): void { external(text); renew(); session.clearHistory(); }
export function paste(text: string): boolean {
  const data = new DataTransfer(); data.setData("text/plain", text);
  const event = new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", { value: data });
  return !input.dispatchEvent(event);
}
export function releaseBinding(): void { binding.dispose(); }
export function dispose(): void { binding.dispose(); controller.dispose(); input.remove(); }
export function hostStillUsable(): boolean { external("host remains usable"); return session.getCell("sheet-1", 0, 0)?.value.type === "text"; }
(window as unknown as { workbookInput: unknown }).workbookInput = { mount, inspect, select, external, seedText, paste, releaseBinding, dispose, hostStillUsable };
