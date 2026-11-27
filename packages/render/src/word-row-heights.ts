import { DEFAULT_DOCUMENT_READ_LIMITS, resolveSourceCssLength, SourceStyleResolver } from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { ownViewStyle } from "./view-nodes.js";
const rowInsets = new WeakMap<HTMLTableRowElement, number>();
/** Physical points added to the editable native height on this composed row. */
export function getWordRowHeightInset(row: HTMLTableRowElement): number {
  return rowInsets.get(row) ?? 0;
}

/** Native trHeight is not a CSS border-box height. Word adds the largest
 * bottom cell margin to exact heights; minimum heights reserve the cell's
 * margins and horizontal strokes as well. Keep these projections view-owned. */
export function composeWordRowHeights(document: Document, model: ContentDocumentModel): Map<HTMLTableRowElement, number> {
  const snapshot = model.snapshot(), exact = new Map<HTMLTableRowElement, number>();
  if (!snapshot.metadata.word) return exact;
  const view = document.defaultView!,
    resolver = new SourceStyleResolver(snapshot.styles.map(sheet => sheet.css), DEFAULT_DOCUMENT_READ_LIMITS),
    cache = new WeakMap<HTMLElement, Readonly<Record<string, string>>>(),
    project = (id: string): SourceStyleNode => {
      const node = model.getNode(id)!;
      return { tag: node.tagName, attrs: node.attributes, parent: node.parentNodeId ? project(node.parentNodeId) : undefined };
    },
    css = (node: HTMLElement) => {
      const known = cache.get(node);
      if (known) return known;
      const id = node.dataset.nodeId;
      const values = id && model.getNode(id) ? resolver.get(project(id)) : {};
      cache.set(node, values); return values;
    },
    length = (node: HTMLElement, values: Readonly<Record<string, string>>, property: string): number => {
      const computed = view.getComputedStyle(node), raw = values[property];
      if (raw && !/^(?:initial|auto|thin|medium|thick)$/.test(raw))
        return resolveSourceCssLength(raw, { font: parseFloat(computed.fontSize) * 0.75, percent: 0 }) * 4 / 3;
      return parseFloat(computed.getPropertyValue(property)) || 0;
    },
    border = (node: HTMLElement, side: "top" | "bottom"): number => {
      const values = css(node), property = "border-" + side;
      const style = values[property + "-style"] ?? view.getComputedStyle(node).getPropertyValue(property + "-style");
      return style === "none" || style === "hidden" ? 0 : length(node, values, property + "-width");
    };
  for (const table of document.querySelectorAll<HTMLTableElement>("table[data-node-id]")) {
    const rows = [...table.rows];
    const top = rows.map(() => 0), bottom = rows.map(() => 0),
      strokes = Array<number>(rows.length + 1).fill(0),
      collapsed = view.getComputedStyle(table).borderCollapse === "collapse",
      ownStrokes = rows.map(() => ({ top: 0, bottom: 0 }));
    if (collapsed && rows.length) {
      strokes[0] = border(table, "top"); strokes[rows.length] = border(table, "bottom");
    }
    rows.forEach((row, index) => {
      for (const cell of row.cells) {
        if (cell.dataset.wordGridGap) continue;
        const last = Math.min(rows.length, index + (cell.rowSpan || rows.length - index)), values = css(cell),
          before = border(cell, "top"), after = border(cell, "bottom");
        top[index] = Math.max(top[index]!, length(cell, values, "padding-top"));
        bottom[last - 1] = Math.max(bottom[last - 1]!, length(cell, values, "padding-bottom"));
        if (collapsed) {
          strokes[index] = Math.max(strokes[index]!, before);
          strokes[last] = Math.max(strokes[last]!, after);
        } else {
          ownStrokes[index]!.top = Math.max(ownStrokes[index]!.top, before);
          ownStrokes[last - 1]!.bottom = Math.max(ownStrokes[last - 1]!.bottom, after);
        }
      }
    });
    rows.forEach((row, index) => {
      for (const cell of row.cells) {
        if (cell.dataset.wordGridGap) continue;
        const last = Math.min(rows.length, index + (cell.rowSpan || rows.length - index)),
          target = cell.querySelector<HTMLElement>(":scope > .sdk-word-row-exception-content") ?? cell;
        // Word shares the largest vertical margins across a physical row,
        // even when individual native cells declare different values.
        for (const [property, value] of [["padding-top", top[index]!], ["padding-bottom", bottom[last - 1]!]] as const) {
          if (Math.abs((parseFloat(view.getComputedStyle(target).getPropertyValue(property)) || 0) - value) < 0.01) continue;
          ownViewStyle(target, property); target.style.setProperty(property, value + "px");
        }
      }
      const values = css(row), rule = values["--word-row-height-rule"];
      const stroke = collapsed ? (strokes[index]! + strokes[index + 1]!) / 2
        : ownStrokes[index]!.top + ownStrokes[index]!.bottom;
      const inset = rule === "exact" ? bottom[index]! : top[index]! + bottom[index]! + stroke;
      rowInsets.set(row, inset * 0.75);
      if (rule === "auto") {
        ownViewStyle(row, "height"); row.style.height = "auto";
      } else if ((rule === "atLeast" || rule === "exact") && values.height && !/^(?:auto|initial)$/.test(values.height)) {
        const height = length(row, values, "height");
        if (rule === "exact") {
          const target = height + bottom[index]!;
          ownViewStyle(row, "height"); row.style.height = target + "px";
          exact.set(row, target);
        }
        else {
          ownViewStyle(row, "height");
          row.style.height = height + inset + "px";
        }
      }
    });
  }
  return exact;
}
