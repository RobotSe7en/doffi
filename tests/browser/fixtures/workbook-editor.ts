import { createDocumentSession, openDocument } from "@doffi/core";
import type { WorkbookSession } from "@doffi/core";
import { mountWorkbookEditor } from "@doffi/editor";
import type { WorkbookEditorView } from "@doffi/editor";
import { decodeTestPackage } from "./package-transport.js";
import type { TestPackage } from "./package-transport.js";
let session: WorkbookSession, editor: WorkbookEditorView, container: HTMLDivElement;
let observer: MutationObserver | undefined;
const textViews = new WeakMap<HTMLIFrameElement, ShadowRoot>(), restores: (() => void)[] = [];
const errors: string[] = [];
export async function mount(entries: TestPackage, observeText = false): Promise<void> {
  observer?.disconnect(); for (const restore of restores.splice(0)) restore();
  editor?.dispose(); session?.dispose(); container?.remove(); errors.length = 0;
  session = createDocumentSession(openDocument(decodeTestPackage(entries), { kind: "workbook" }));
  container = document.createElement("div"); document.body.append(container);
  // The fixture opens only its own parent-realm UI shadows for browser
  // automation. Production source and render surfaces remain closed.
  const attach = Element.prototype.attachShadow;
  Element.prototype.attachShadow = function(init) { return attach.call(this, { ...init, mode: "open" }); };
  try { editor = mountWorkbookEditor(container, session, { widthPt: 420, heightPt: 240, onError: error => errors.push((error as { code?: string }).code ?? error.message) }); }
  finally { Element.prototype.attachShadow = attach; }
  if (observeText) {
    const root = editor.element.shadowRoot!.querySelector(".preview")!.firstElementChild!.shadowRoot!, watched = new WeakSet<HTMLIFrameElement>();
    const watch = () => { for (const frame of root.querySelectorAll<HTMLIFrameElement>("iframe")) {
      if (watched.has(frame)) continue; watched.add(frame);
      frame.addEventListener("load", () => {
        const prototype = (frame.contentWindow as Window & typeof globalThis).Element.prototype, attach = prototype.attachShadow;
        prototype.attachShadow = function(init) { const value = attach.call(this, init); if (this.localName === "office-workbook-view-surface") textViews.set(frame, value); return value; };
        restores.push(() => { prototype.attachShadow = attach; });
      }, { capture: true, once: true });
    } };
    observer = new MutationObserver(watch); observer.observe(root, { childList: true, subtree: true }); watch();
  }
  await editor.ready;
}
export async function mountObserved(entries: TestPackage): Promise<void> { await mount(entries, true); }
export function displayedCell(row: number, column: number): string | undefined {
  const frame = editor.element.shadowRoot!.querySelector(".preview")!.firstElementChild!.shadowRoot!.querySelector<HTMLIFrameElement>("iframe")!;
  return textViews.get(frame)?.querySelector(`[data-office-view="workbook-cell"][data-row="${row}"][data-column="${column}"]`)?.textContent ?? undefined;
}
export async function settled() { return editor.settled(); }
export async function select(fromRow: number, fromColumn = 0, toRow = fromRow, toColumn = fromColumn): Promise<void> {
  await editor.selectRange("sheet-1", { from: { row: fromRow, column: fromColumn }, to: { row: toRow, column: toColumn } }); editor.focusInput();
}
export function inspect() {
  return { input: editor.inputState(), selection: editor.selection(), revision: session.snapshot().manifest.revision, history: session.history(), errors: [...errors],
    cells: [session.getCell("sheet-1", 0, 0), session.getCell("sheet-1", 0, 1)], calculated: session.getCalculatedCell("sheet-1", 0, 1),
    cleanSource: [...session.serialize()].filter(([path]) => path.endsWith(".html")).every(([, bytes]) => !/contenteditable|data-office-view|textarea/.test(new TextDecoder().decode(bytes))) };
}
export function external(value: string): void {
  const manifest = session.snapshot().manifest;
  session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision,
    operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value } } }] }] }));
}
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
  observer?.disconnect(); for (const restore of restores.splice(0)) restore();
  editor.dispose(); const revision = session.snapshot().manifest.revision; external("host remains usable"); container.remove();
  return { revision, usable: session.getCell("sheet-1", 0, 0)?.value.type === "text" };
}
(window as unknown as { workbookEditor: unknown }).workbookEditor = { mount, mountObserved, displayedCell, settled, select, inspect, external, cell, paste, point, cancelFillAtProgress, cancelSearchAtProgress, cancelPreparationAtProgress, dispose };
