import {
  DEFAULT_DOCUMENT_READ_LIMITS,
  DocumentError,
  SourceStyleResolver,
  wordRowExceptionProperties,
} from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { markViewNode, ownViewStyle } from "./view-nodes.js";
const leadingInsets = new WeakMap<HTMLTableElement, number>();
export function wordRowGeometryProperties(css: Readonly<Record<string, string>>): Record<string, string> {
  return { ...wordRowExceptionProperties(css),
    ...(css["--word-row-alignment"] ? { alignment: css["--word-row-alignment"] } : {}) };
}

/** Independent row geometry uses reversible content boxes. The native table,
 * grid, rowspans and source selection remain one tree; text is never scaled. */
export function composeWordRowExceptions(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): void {
  const snapshot = model.snapshot();
  if (!snapshot.metadata.word?.tables.length) return;
  if ((snapshot.metadata.word.compatibilityMode ?? 12) >= 15) return;
  const resolver = new SourceStyleResolver(
    snapshot.styles.map((sheet) => sheet.css),
    DEFAULT_DOCUMENT_READ_LIMITS,
  );
  const project = (id: string): SourceStyleNode => {
    const node = model.getNode(id)!;
    return {
      tag: node.tagName,
      attrs: node.attributes,
      parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
    };
  };
  let count = document.querySelectorAll("*").length;
  for (const table of document.querySelectorAll<HTMLTableElement>("table[data-node-id]")) {
    const info = snapshot.metadata.word?.tables.find(
      (info) => info.nodeId === table.dataset.nodeId,
    );
    if (!info) continue;
    const rows = [...table.rows],
      properties = rows.map((row) =>
        wordRowGeometryProperties(resolver.specified(project(row.dataset.nodeId!))),
      );
    if (!properties.some((css) => css.alignment)) continue;
    const bounds = table.getBoundingClientRect(),
      total = info.grid.reduce((a, b) => a + b, 0),
      occupied: number[] = [];
    const leading = rows[0]?.cells[0];
    leadingInsets.set(table, leading ? parseFloat(document.defaultView!.getComputedStyle(leading)
      .getPropertyValue(table.dataset.sdkWordTableBidi === "true" ? "padding-right" : "padding-left")) || 0 : 0);
    const rowBounds = properties.map(css => wordRowExceptionBounds(table, css));
    ownViewStyle(table, "table-layout", "width");
    table.style.tableLayout = "fixed";
    table.style.width = bounds.width + "px";
    rows.forEach((row, index) => {
      const { left, width } = rowBounds[index]!;
      let column = 0;
      for (const cell of row.cells) {
        while ((occupied[column] ?? 0) > index) column++;
        const start = column,
          span = cell.colSpan;
        for (let c = start; c < start + span; c++)
          occupied[c] = index + (cell.rowSpan || rows.length - index);
        column += span;
        if (cell.dataset.wordGridGap) continue;
        if (++count > limits.maxNodes)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Row exception geometry exceeds the render node budget.",
          );
        const box = markViewNode(document.createElement("div"), "unwrap"),
          original = cell.getBoundingClientRect(),
          computed = document.defaultView!.getComputedStyle(cell);
        const offset = info.grid.slice(0, start).reduce((a, b) => a + b, 0) / total,
          size = info.grid.slice(start, column).reduce((a, b) => a + b, 0) / total,
          bidi = table.dataset.sdkWordTableBidi === "true";
        box.className = "sdk-word-row-exception-content";
        box.style.cssText = `box-sizing:border-box;position:relative;display:flow-root;min-width:0;pointer-events:auto;z-index:1;width:${width * size}px;margin-left:${left + width * (bidi ? 1 - offset - size : offset) - original.left}px`;
        for (const key of [
          "background-color",
          ...["top", "right", "bottom", "left"].flatMap((side) => [
            "padding-" + side,
            ...["width", "style", "color"].map((part) => "border-" + side + "-" + part),
          ]),
        ])
          box.style.setProperty(key, computed.getPropertyValue(key));
        ownViewStyle(cell, "padding", "border", "background-color", "pointer-events");
        cell.style.pointerEvents = "none";
        cell.style.padding = "0";
        cell.style.border = "0";
        cell.style.backgroundColor = "transparent";
        box.append(...cell.childNodes);
        cell.append(box);
      }
    });
  }
}

/** Physical merged-cell segments use the containing row's own geometry. */
export function wordRowExceptionBounds(
  table: HTMLTableElement,
  css: Readonly<Record<string, string>>,
): { left: number; width: number } {
  const bounds = table.getBoundingClientRect(),
    parent = table.parentElement!,
    rect = parent.getBoundingClientRect(),
    computed = table.ownerDocument.defaultView!.getComputedStyle(parent);
  const left = rect.left + (parseFloat(computed.paddingLeft) || 0),
    available =
      rect.width -
      (parseFloat(computed.paddingLeft) || 0) -
      (parseFloat(computed.paddingRight) || 0);
  const width = table.rows[0]?.getBoundingClientRect().width ?? bounds.width,
    alignment = css.alignment,
    bidi = table.dataset.sdkWordTableBidi === "true",
    leading = table.rows[0]?.cells[0],
    inset = leadingInsets.get(table) ?? (leading ? parseFloat(table.ownerDocument.defaultView!.getComputedStyle(leading)
      .getPropertyValue(bidi ? "padding-right" : "padding-left")) || 0 : 0);
  if (!alignment) return { left: bounds.left, width };
  const physical = bidi ? alignment === "left" || alignment === "start" ? "right"
    : alignment === "right" ? "left" : alignment : alignment;
  return { left: physical === "center" ? left + (available - width) / 2
    : physical === "right" ? left + available - width + inset
    : bounds.left, width };
}
