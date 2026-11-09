import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, TransactionPrepareOptions, WorkbookSession } from "@doffi/core";
import { createWorkbookLayoutService } from "@doffi/render";
import { mountWorkbookEditor } from "@doffi/editor";
import type { WorkbookEditorView } from "@doffi/editor";
import { decodeTestPackage } from "./package-transport.js";
import { hostOwnershipProbeLayout } from "./workbook-host-probe.js";
import type { TestPackage } from "./package-transport.js";
import { measureWorkbookCellAppearance, measureWorkbookCellBounds, measureWorkbookCellDecorations, measureWorkbookChart, measureWorkbookDrawingText, measureWorkbookDrawingBounds, observeWorkbookViews } from "./workbook-view-observer.js";
let session: WorkbookSession, editor: WorkbookEditorView, container: HTMLDivElement;
let observed: ReturnType<typeof observeWorkbookViews> | undefined;
const errors: string[] = [];
let pendingAction: (() => void) | undefined;
export async function mount(entries: TestPackage, observeText = false, prepareOptions?: TransactionPrepareOptions, confirmWarnings = false): Promise<void> {
  observed?.dispose(); observed = undefined;
  editor?.dispose(); session?.dispose(); container?.remove(); errors.length = 0; pendingAction = undefined;
  session = createDocumentSession(openDocument(decodeTestPackage(entries), { kind: "workbook" }));
  container = document.createElement("div"); document.body.append(container);
  // The fixture opens only its own parent-realm UI shadows for browser
  // automation. Production source and render surfaces remain closed.
  const attach = Element.prototype.attachShadow;
  Element.prototype.attachShadow = function(init) { return attach.call(this, { ...init, mode: "open" }); };
  try { editor = mountWorkbookEditor(container, session, { widthPt: 420, heightPt: 240, ...(prepareOptions ? { prepareOptions: () => prepareOptions } : {}), ...(confirmWarnings ? { confirmValidationWarnings: () => window.confirm("应用验证警告？") } : {}),
    onPendingChange(value) { if (value && pendingAction) { const action = pendingAction; pendingAction = undefined; action(); } }, onError: error => errors.push((error as { code?: string }).code ?? error.message) }); }
  finally { Element.prototype.attachShadow = attach; }
  if (observeText) observed = observeWorkbookViews(editor.element);
  await editor.ready;
}
export async function mountObserved(entries: TestPackage, prepareOptions?: TransactionPrepareOptions): Promise<void> { await mount(entries, true, prepareOptions); }
export function displayedCellAppearance(row: number, column: number) { return measureWorkbookCellAppearance(observed?.current(), row, column); }
export function displayedCellDecorations(row: number, column: number) { return measureWorkbookCellDecorations(observed?.current(), row, column); }
export function displayedChart() { return measureWorkbookChart(observed?.current()); }
export function displayedDrawingText(nodeId: string, includeBounds = false) { return measureWorkbookDrawingText(observed?.current(), nodeId, includeBounds); }
export function displayedCell(row: number, column: number): string | undefined {
  const frame = editor.element.shadowRoot!.querySelector(".preview")!.firstElementChild!.shadowRoot!.querySelector<HTMLIFrameElement>("iframe")!;
  return observed?.root(frame)?.querySelector(`[data-office-view="workbook-cell"][data-row="${row}"][data-column="${column}"]`)?.textContent ?? undefined;
}
export function displayedTextStyles(row: number, column: number) {
  const frame = editor.element.shadowRoot!.querySelector(".preview")!.firstElementChild!.shadowRoot!.querySelector<HTMLIFrameElement>("iframe")!;
  const cell = observed?.root(frame)?.querySelector<HTMLElement>(`[data-office-view="workbook-cell"][data-row="${row}"][data-column="${column}"]`);
  if (!cell) return [];
  return [cell, ...cell.querySelectorAll<HTMLElement>("[data-office-text-style]")].map(element => {
    const style = element.ownerDocument.defaultView!.getComputedStyle(element);
    return { text: element.textContent, color: style.color, weight: style.fontWeight, italic: style.fontStyle, size: style.fontSize, family: style.fontFamily };
  });
}
export async function settled() { return editor.settled(); }
export async function perform(operations: readonly DocumentOperation[]) { const receipt = await editor.perform(operations); await editor.settled(); return receipt; }
export function mutateInputOnPending(text: string): void { pendingAction = () => {
  const field = editor.element.shadowRoot!.querySelector("textarea")!; field.value = text; field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: text }));
}; }
export async function performWithPendingMutation() {
  const cell = { row: 0, column: 0, source: { kind: "value" as const, value: { type: "text" as const, value: "captured" } } };
  pendingAction = () => { cell.source.value.value = "changed-by-host"; };
  const receipt = await editor.perform([{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [cell] }]); await editor.settled(); return { receipt, callerValue: cell.source.value.value };
}
export async function measuredApply(operations: readonly DocumentOperation[]) {
  const layout = createWorkbookLayoutService(container, { widthPt: 420, heightPt: 240 });
  try {
    const m = session.snapshot().manifest;
    const candidate = await session.prepareAsync({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [...operations] }, { measureWorkbookLayout: layout.measure });
    const receipt = session.commit(candidate); await editor.settled(); return receipt;
  } finally { layout.dispose(); }
}
export function displayedDrawingBounds() { return measureWorkbookDrawingBounds(observed?.current()); }
export function displayedCellBounds(row: number, column: number) { return measureWorkbookCellBounds(observed?.current(), row, column); }
export async function select(fromRow: number, fromColumn = 0, toRow = fromRow, toColumn = fromColumn): Promise<void> {
  await editor.selectRange("sheet-1", { from: { row: fromRow, column: fromColumn }, to: { row: toRow, column: toColumn } }); editor.focusInput();
}
export function inspect() {
  return { input: editor.inputState(), selection: editor.selection(), revision: session.snapshot().manifest.revision, history: session.history(), errors: [...errors], pending: editor.hasPendingInput(),
    merges: session.snapshot().metadata.workbook.sheets[0]!.merges,
    tables: session.snapshot().metadata.workbook.sheets.map(sheet => ({ sheetId: sheet.sheetId, tables: sheet.tables ?? [] })),
    names: session.snapshot().metadata.workbook.names ?? [],
    sheets: session.snapshot().metadata.workbook.sheets,
    frozen: session.snapshot().metadata.workbook.sheets[0]!.frozen,
    drawings: session.snapshot().metadata.workbook.objects ?? [], charts: session.snapshot().metadata.workbook.charts ?? [], content: session.snapshot().content?.html,
    annotations: session.snapshot().annotations, comments: session.snapshot().metadata.workbook.comments ?? [], stories: session.snapshot().metadata.workbook.stories ?? [],
    cells: [session.getCell("sheet-1", 0, 0), session.getCell("sheet-1", 0, 1)], calculated: session.getCalculatedCell("sheet-1", 0, 1),
    cleanSource: [...session.serialize()].filter(([path]) => path.endsWith(".html")).every(([, bytes]) => !/contenteditable|data-office-view|textarea/.test(new TextDecoder().decode(bytes))) };
}
export function external(value: string, at = { row: 0, column: 0 }, preparation?: TransactionPrepareOptions): void {
  const manifest = session.snapshot().manifest;
  session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision,
    operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ ...at, source: { kind: "value", value: { type: "text", value } } }] }] }, preparation));
}
export function tableStyleSource() { const workbook = session.snapshot().metadata.workbook; return { tableStyles: workbook.tableStyles, defaultTableStyle: workbook.defaultTableStyle }; }
export function cell(sheetId: string, row: number, column: number) { return session.getCell(sheetId, row, column); }
export function paste(text: string): boolean {
  const input = editor.element.shadowRoot!.querySelector("textarea")!, data = new DataTransfer(); data.setData("text/plain", text);
  const event = new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: data });
  return !input.dispatchEvent(event);
}
export function point(row: number, column: string): { x: number; y: number } {
  const root = editor.element.shadowRoot!.querySelector(".preview")!.firstElementChild!.shadowRoot!, headers = [...root.querySelectorAll(".header>div")];
  const x = headers.find(node => node.textContent === column)!.getBoundingClientRect(), y = headers.find(node => node.textContent === String(row + 1))!.getBoundingClientRect();
  return { x: x.x + x.width / 2, y: y.y + y.height / 2 };
}
export function cancelFillAtProgress(): void {
  const root = editor.element.shadowRoot!, status = root.querySelector('[role="status"]')!;
  const observer = new MutationObserver(() => {
    if (!/正在填充 256 /.test(status.textContent!)) return; observer.disconnect();
    [...root.querySelectorAll<HTMLButtonElement>(".tools button")].find(button => button.textContent === "取消操作")!.click();
  }); observer.observe(status, { childList: true, subtree: true });
}
export function cancelSearchAtProgress(): void {
  const root = editor.element.shadowRoot!, status = root.querySelector('[role="status"]')!;
  const observer = new MutationObserver(() => {
    if (!/正在(?:查找|替换) · 已检查 /.test(status.textContent!)) return; observer.disconnect();
    [...root.querySelectorAll<HTMLButtonElement>(".tools button")].find(button => button.textContent === "取消操作")!.click();
  }); observer.observe(status, { childList: true, subtree: true });
}
export function cancelPreparationAtProgress(): void {
  const root = editor.element.shadowRoot!, status = root.querySelector('[role="status"]')!;
  const observer = new MutationObserver(() => {
    if (!/正在准备更改 · 256/.test(status.textContent!)) return; observer.disconnect();
    [...root.querySelectorAll<HTMLButtonElement>(".tools button")].find(button => button.textContent === "取消操作")!.click();
  }); observer.observe(status, { childList: true, subtree: true });
}
export function dispose(): { revision: number; usable: boolean } {
  observed?.dispose(); observed = undefined;
  editor.dispose(); const revision = session.snapshot().manifest.revision;
  external("host remains usable", undefined, { measureWorkbookLayout: hostOwnershipProbeLayout }); container.remove();
  return { revision, usable: session.getCell("sheet-1", 0, 0)?.value.type === "text" };
}
(window as unknown as { workbookEditor: unknown }).workbookEditor = { tableStyleSource, mount, mountObserved, displayedCellAppearance, displayedCellDecorations, displayedChart, displayedDrawingText, displayedDrawingBounds, displayedCellBounds, displayedCell, displayedTextStyles, settled, perform, mutateInputOnPending, performWithPendingMutation, measuredApply, select, inspect, external, cell, paste, point, cancelFillAtProgress, cancelSearchAtProgress, cancelPreparationAtProgress, dispose };
