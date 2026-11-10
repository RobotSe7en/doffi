import { DocumentError } from "@doffi/core";
import type { WorkbookSnapshot } from "@doffi/core";
import { WorkbookScrollAxis } from "./scroll.js";
import type { WorkbookRenderedWindow } from "./window-controller.js";

export function workbookColumnLabel(column: number): string {
  let label = "";
  for (let value = column + 1; value; value = Math.floor((value - 1) / 26)) label = String.fromCharCode(65 + (value - 1) % 26) + label;
  return label;
}
export interface WorkbookPreviewChrome {
  readonly element: HTMLDivElement;
  readonly canvas: HTMLElement;
  prepare(window: WorkbookRenderedWindow, snapshot: WorkbookSnapshot): () => void;
  cell(address: string, value: string): void;
  validation(message: string): void;
  status(message: string, error?: boolean): void;
  busy(value: boolean): void;
  dispose(): void;
}
export function createWorkbookPreviewChrome(container: HTMLElement, title: string,
  actions: { sheet(id: string): void; scroll(left: number, top: number): void; comments(): void }): WorkbookPreviewChrome {
  const document = container.ownerDocument, view = document.defaultView;
  if (!container.isConnected || !view) throw new DocumentError("InvalidOperation", "Workbook preview requires a connected host container.");
  const element = document.createElement("div"), shadow = element.attachShadow({ mode: "closed" }), controls = new view.AbortController();
  element.style.cssText = "all:initial!important;display:block!important;position:relative!important;direction:ltr!important";
  element.dataset.officeView = "workbook-preview"; element.setAttribute("role", "region"); element.setAttribute("aria-label", title);
  const style = document.createElement("style");
  style.textContent = ':host{color:#20242a}*{box-sizing:border-box}button,input{font:inherit;color:inherit}button{cursor:pointer;background:#f4f6fa;border:1px solid #dce0e6;padding:4px 8px}button:disabled{cursor:default}button[aria-selected="true"]{background:#dfeaff;color:#173d7a}input{min-width:0;border:1px solid #dce0e6;padding:4px;background:white}.root{font:12px sans-serif;color:#20242a;background:#fff;width:max-content;border:1px solid #dce0e6}.toolbar{display:var(--doffi-workbook-preview-toolbar,flex);gap:4px;align-items:center;padding:4px}.board{position:relative;background:#f4f6fa}.canvas,.rows,.columns,.x,.y{position:absolute;overflow:hidden}.rows{left:0;top:18pt;width:36pt}.columns{left:36pt;top:0;height:18pt}.canvas{left:36pt;top:18pt}.x{left:36pt;height:12pt;overflow-x:scroll}.y{top:18pt;width:12pt;overflow-y:scroll}.tabs{display:flex;overflow:auto;gap:2px;padding:4px;max-width:100%}.status[data-mode-label=true]{display:var(--doffi-workbook-mode-label,block)}.status{padding:4px 8px;min-height:24px;white-space:pre-wrap;overflow-wrap:anywhere}.header{position:absolute;overflow:hidden;background:#f4f6fa;border:1px solid #dce0e6}.header>div{position:absolute;display:flex;align-items:center;justify-content:center;font:10px sans-serif;color:#39414c}';
  const root = document.createElement("div"); root.className = "root";
  const toolbar = document.createElement("div"); toolbar.className = "toolbar";
  const address = document.createElement("input"), value = document.createElement("input"), comments = document.createElement("button");
  address.readOnly = value.readOnly = true; address.setAttribute("aria-label", "单元格地址"); address.style.width = "88px";
  value.setAttribute("aria-label", "单元格值或公式"); value.style.flex = "1";
  comments.type = "button"; comments.textContent = "批注"; comments.addEventListener("click", actions.comments, { signal: controls.signal }); toolbar.append(address, value, comments);
  const board = document.createElement("div"); board.className = "board";
  const canvas = document.createElement("div"), rows = document.createElement("div"), columns = document.createElement("div"), x = document.createElement("div"), y = document.createElement("div");
  canvas.className = "canvas"; rows.className = "rows"; columns.className = "columns"; x.className = "x"; y.className = "y";
  rows.setAttribute("aria-label", "行标题"); columns.setAttribute("aria-label", "列标题"); x.setAttribute("aria-label", "水平滚动"); y.setAttribute("aria-label", "垂直滚动");
  x.tabIndex = y.tabIndex = 0; const xTrack = document.createElement("div"), yTrack = document.createElement("div"); xTrack.style.height = "1px"; yTrack.style.width = "1px"; x.append(xTrack); y.append(yTrack);
  board.append(rows, columns, canvas, x, y);
  const tabs = document.createElement("div"), status = document.createElement("div"); tabs.className = "tabs"; tabs.setAttribute("role", "tablist"); tabs.setAttribute("aria-label", "工作表");
  const validation = document.createElement("div"); validation.className = "status"; validation.setAttribute("role", "note"); validation.setAttribute("aria-live", "polite"); validation.setAttribute("aria-label", "数据验证"); validation.hidden = true;
  status.className = "status"; status.setAttribute("role", "status"); root.append(toolbar, board, tabs, validation, status); shadow.append(style, root); container.append(element);
  let horizontal: WorkbookScrollAxis | undefined, vertical: WorkbookScrollAxis | undefined, left = 0, top = 0, trackLeft = 0, trackTop = 0, busy = false;
  let columnSignature: string | undefined, tabSignature: string | undefined;
  x.addEventListener("scroll", () => { if (x.scrollLeft === trackLeft) return; trackLeft = x.scrollLeft; if (!busy && horizontal) actions.scroll(horizontal.fromTrack(trackLeft), top); }, { signal: controls.signal });
  y.addEventListener("scroll", () => { if (y.scrollTop === trackTop) return; trackTop = y.scrollTop; if (!busy && vertical) actions.scroll(left, vertical.fromTrack(trackTop)); }, { signal: controls.signal });
  tabs.addEventListener("click", event => { const button = (event.target as Element).closest<HTMLButtonElement>("button[data-sheet-id]"); if (!busy && button && tabs.contains(button)) actions.sheet(button.dataset.sheetId!); }, { signal: controls.signal });
  return Object.freeze({ element, canvas,
    prepare(window: WorkbookRenderedWindow, snapshot: WorkbookSnapshot): () => void {
      const layout = window.frame.layout, sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === window.frame.sheetId)!;
      const rowNodes = document.createDocumentFragment();
      const columnPlans: [number, number, number, number, number][] = [];
      const seenRows = new Set<string>(), seenColumns = new Set<string>();
      const header = (parent: DocumentFragment, index: number, start: number, size: number, clipStart: number, clipSize: number, column: boolean): void => {
        const node = document.createElement("div"), inner = document.createElement("div"); node.className = "header"; node.setAttribute("role", column ? "columnheader" : "rowheader");
        inner.textContent = column ? workbookColumnLabel(index) : String(index + 1);
        const before = column && sheet.rightToLeft ? layout.widthPt - clipStart - clipSize : clipStart;
        const offset = column && sheet.rightToLeft ? clipStart + clipSize - start - size : start - clipStart;
        node.style.cssText = column ? `left:${before}pt;top:0;width:${clipSize}pt;height:18pt` : `left:0;top:${clipStart}pt;width:36pt;height:${clipSize}pt`;
        inner.style.cssText = column ? `left:${offset}pt;top:0;width:${size}pt;height:18pt` : `left:0;top:${offset}pt;width:36pt;height:${size}pt`;
        node.append(inner); parent.append(node);
      };
      for (const pane of layout.panes) for (const column of [false, true]) {
        const start = column ? pane.inlineStartPt : pane.blockStartPt, end = start + (column ? pane.widthPt : pane.heightPt), source = column ? pane.sourceLeftPt : pane.sourceTopPt;
        for (const cell of column ? pane.columns : pane.rows) {
          const position = cell.startPt - source + start, clipped = Math.max(start, position), size = Math.min(end, position + cell.sizePt) - clipped;
          const key = `${cell.index}:${clipped}:${size}`, seen = column ? seenColumns : seenRows; if (seen.has(key)) continue;
          if (seenRows.size + seenColumns.size >= 10_000) throw new DocumentError("ResourceLimitExceeded", "Workbook headers exceed their visible-label budget.");
          seen.add(key);
          if (column) columnPlans.push([cell.index, position, cell.sizePt, clipped, size]);
          else header(rowNodes, cell.index, position, cell.sizePt, clipped, size, false);
        }
      }
      const shown = snapshot.metadata.workbook.sheets.filter(item => item.visibility === "visible" || item.sheetId === sheet.sheetId);
      if (shown.length > 1000) throw new DocumentError("ResourceLimitExceeded", "Workbook navigation exceeds its worksheet-tab budget.");
      const nextColumns = JSON.stringify([layout.widthPt, sheet.rightToLeft, columnPlans]);
      const nextTabs = JSON.stringify([sheet.sheetId, shown.map(item => [item.sheetId, item.name, item.visibility])]);
      const makeColumns = (): DocumentFragment => { const fragment = document.createDocumentFragment(); for (const plan of columnPlans) header(fragment, ...plan, true); return fragment; };
      const makeTabs = (): DocumentFragment => {
        const fragment = document.createDocumentFragment();
        for (const item of shown) {
          const button = document.createElement("button"); button.type = "button"; button.dataset.sheetId = item.sheetId; button.textContent = item.name + (item.visibility !== "visible" ? "（隐藏）" : "");
          button.setAttribute("role", "tab"); button.setAttribute("aria-selected", String(item.sheetId === sheet.sheetId)); fragment.append(button);
        }
        return fragment;
      };
      const columnNodes = columnSignature === nextColumns ? undefined : makeColumns(), tabNodes = tabSignature === nextTabs ? undefined : makeTabs();
      const nextX = new WorkbookScrollAxis(layout.maxScrollLeftPt + layout.widthPt, layout.widthPt), nextY = new WorkbookScrollAxis(layout.maxScrollTopPt + layout.heightPt, layout.heightPt);
      return () => {
        // The board excludes the root's two one-pixel borders (1.5 pt).
        root.style.width = layout.widthPt + 49.5 + "pt";
        board.style.width = layout.widthPt + 48 + "pt"; board.style.height = layout.heightPt + 30 + "pt";
        canvas.style.width = columns.style.width = x.style.width = layout.widthPt + "pt"; canvas.style.height = rows.style.height = y.style.height = layout.heightPt + "pt";
        x.style.top = layout.heightPt + 18 + "pt"; y.style.left = layout.widthPt + 36 + "pt";
        rows.replaceChildren(rowNodes);
        // Vertical scrolling does not replace identical column labels or tabs.
        // Compare again at commit so prepared candidates may be discarded or
        // committed after another candidate without retaining stale chrome.
        if (columnSignature !== nextColumns) { columns.replaceChildren(columnNodes ?? makeColumns()); columnSignature = nextColumns; }
        if (tabSignature !== nextTabs) { tabs.replaceChildren(tabNodes ?? makeTabs()); tabSignature = nextTabs; }
        horizontal = nextX; vertical = nextY; left = layout.scrollLeftPt; top = layout.scrollTopPt;
        xTrack.style.width = nextX.trackSizePx + "px"; yTrack.style.height = nextY.trackSizePx + "px";
        x.scrollLeft = nextX.toTrack(left); y.scrollTop = nextY.toTrack(top); trackLeft = x.scrollLeft; trackTop = y.scrollTop;
        for (const button of tabs.querySelectorAll("button")) button.disabled = busy;
      };
    },
    cell(nextAddress: string, nextValue: string) { if (address.value !== nextAddress) address.value = nextAddress; if (value.value !== nextValue) value.value = nextValue; },
    validation(message: string) { if (validation.textContent !== message) validation.textContent = message; validation.hidden = !message; },
    status(message: string, error = false) { status.dataset.modeLabel = String(message === "只读预览"); if (status.textContent !== message) status.textContent = message; status.style.color = error ? "#a12d22" : "#596273"; },
    busy(value: boolean) { busy = value; element.setAttribute("aria-busy", String(value)); comments.disabled = value; for (const button of tabs.querySelectorAll("button")) button.disabled = value; },
    dispose() {
      controls.abort(); element.remove(); element.removeAttribute("aria-label");
      address.value = value.value = ""; status.textContent = validation.textContent = "";
      rows.replaceChildren(); columns.replaceChildren(); tabs.replaceChildren(); canvas.replaceChildren(); root.replaceChildren(); shadow.replaceChildren();
      columnSignature = tabSignature = undefined;
    },
  });
}
