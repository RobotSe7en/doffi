import{wordTableRepeatHeader}from'./word-table-layout.js';
import { DocumentError } from "@doffi/core";
import type { RenderedPage, RenderLimits } from "./contracts.js";
import { markViewNode, viewNodeRole } from "./view-nodes.js";
import type { PreviewSource } from "./source-index.js";

export interface RepeatedHeaders { clear(): void; project(pages: readonly RenderedPage[]): void; dispose(): void }
/** Source headers remain editable once; owned copies occupy the space reserved
 * by the measured table flow on subsequent pages. */
export function repeatTableHeaders(document: Document, root: HTMLElement, source: PreviewSource, limits: Readonly<RenderLimits>): RepeatedHeaders {
  const tables = [...root.querySelectorAll<HTMLTableElement>("table")], clones: HTMLElement[] = [];
  const clear = (): void => { for (const clone of clones.splice(0)) clone.remove(); source.clearInstances(); };
  return { clear, project(pages) {
    clear(); let nodes = root.querySelectorAll("*").length;
    const surface = root.getBoundingClientRect(), px = (value: number): number => value * 96 / 72;
    for (const table of tables) {
      if(table.closest<HTMLElement>('.sdk-page-flow')?.hasAttribute('data-sdk-column-count'))continue;
      const rows: HTMLTableRowElement[] = [];
      for (const row of table.rows) {
        if (viewNodeRole(row) === "remove") continue;
        if (!wordTableRepeatHeader(row)) break;
        rows.push(row);
      }
      if (!rows.length) continue;
      const bounds = table.getBoundingClientRect();
      const containing = pages.filter((page) => bounds.bottom > surface.top + px(page.rect.y + page.setup.margins.top) + 0.5 && bounds.top < surface.top + px(page.rect.y + page.rect.height - page.setup.margins.bottom) - 0.5);
      for (const page of containing.slice(1)) {
        const header = markViewNode(table.cloneNode(false) as HTMLTableElement);
        header.classList.add("sdk-repeat-header"); header.contentEditable = "false"; header.setAttribute("aria-hidden", "true"); header.dataset.sdkRepeatPageId = page.pageId;
        header.style.cssText += `;position:absolute;left:${bounds.left - surface.left}px;top:${px(page.rect.y + page.setup.margins.top)}px;width:${bounds.width}px;margin:0;pointer-events:none;user-select:none;z-index:2`;
        const columns = table.querySelector(":scope > colgroup"); if (columns) header.append(columns.cloneNode(true));
        const body = document.createElement("tbody");
        for (const row of rows) {
          const clone = row.cloneNode(true) as HTMLTableRowElement;
          for (const [index, cell] of [...clone.cells].entries()) cell.style.width = row.cells[index]!.getBoundingClientRect().width + "px";
          body.append(clone);
        }
        header.append(body); nodes += header.querySelectorAll("*").length + 1;
        if (nodes > limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Repeated table headers exceed the render node budget.");
        root.append(header); clones.push(header); source.registerInstances(header);
      }
    }
  }, dispose: clear };
}
