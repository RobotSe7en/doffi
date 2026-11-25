import {
  DEFAULT_DOCUMENT_READ_LIMITS,
  DocumentError,
  SourceStyleResolver,
  resolveSourceCssLength,
  wordRowExceptionProperties,
} from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { markViewNode, ownViewStyle } from "./view-nodes.js";

/** Independent row geometry uses reversible content boxes. The native table,
 * grid, rowspans and source selection remain one tree; text is never scaled. */
export function composeWordRowExceptions(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): void {
  const snapshot = model.snapshot();
  if (!snapshot.metadata.word?.tables.length) return;
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
        wordRowExceptionProperties(resolver.get(project(row.dataset.nodeId!))),
      );
    if (!properties.some((css) => css.width || css["margin-left"] || css["margin-right"])) continue;
    const bounds = table.getBoundingClientRect(),
      total = info.grid.reduce((a, b) => a + b, 0),
      occupied: number[] = [];
    ownViewStyle(table, "table-layout", "width");
    table.style.tableLayout = "fixed";
    table.style.width = bounds.width + "px";
    rows.forEach((row, index) => {
      const { left, width } = wordRowExceptionBounds(table, properties[index]!);
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
        box.className = "sdk-word-row-exception-content";
        box.style.cssText = `box-sizing:border-box;position:relative;display:flow-root;min-width:0;pointer-events:auto;z-index:1;width:${(width * info.grid.slice(start, column).reduce((a, b) => a + b, 0)) / total}px;margin-left:${left + (width * info.grid.slice(0, start).reduce((a, b) => a + b, 0)) / total - original.left}px`;
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
  const length = (raw: string) =>
    resolveSourceCssLength(raw, { font: 12, percent: available * 0.75 }) / 0.75;
  const width = css.width && css.width !== "auto" ? length(css.width) : bounds.width;
  const x =
    css["margin-left"] === "auto"
      ? css["margin-right"] === "auto"
        ? (available - width) / 2
        : available - width
      : css["margin-left"]
        ? length(css["margin-left"])
        : bounds.left - left;
  return { left: left + x, width };
}
