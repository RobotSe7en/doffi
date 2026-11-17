import type { WordDocumentMetadata } from "../contracts/document.js";
import type { ParsedHtml } from "./html.js";
import type { SourceStyleNode, SourceStyleResolver } from "./source-css.js";
import type { DocumentReadLimits } from "../runtime.js";
import { length } from "./css-values.js";
import { checkLimit } from "./validation.js";
import { wordDrawingDimensions } from "../word-drawing-size.js";
import { evaluatePresentationPreset } from "../presentation-geometry.js";
/** Derive an insertion's containing width from authoritative section, table and
 * drawing definitions. A host can supply an explicit width for an unequal
 * column; without layout coordinates use the narrowest column that can hold it. */
export function wordContainerWidth(
  metadata: WordDocumentMetadata,
  source: ParsedHtml,
  nodeId: string,
  resolver: SourceStyleResolver,
  limits: DocumentReadLimits,
): number {
  const ancestors: string[] = [];
  for (let id: string | undefined = nodeId; id; id = source.peekNode(id)!.parentNodeId ?? undefined)
    ancestors.unshift(id);
  const word = metadata.word,
    story = word?.stories.find((story) => ancestors.includes(story.rootNodeId)),
    section =
      metadata.sections.find((section) => ancestors.includes(section.rootNodeId)) ??
      metadata.sections.find(
        (section) =>
          story &&
          [
            ...Object.values(section.headers ?? {}),
            ...Object.values(section.footers ?? {}),
          ].includes(story.storyId),
      ) ??
      metadata.sections[0],
    setup = section?.pageSetup ?? metadata.pageSetup;
  let width =
    setup.width -
    setup.margins.left -
    setup.margins.right -
    (section?.gutter && section.gutterPosition !== "top" ? section.gutter : 0);
  if (!story || story.kind === "main") {
    const columns = section?.columns;
    if (columns)
      width = columns.widths
        ? Math.min(...columns.widths)
        : (width - columns.gap * (columns.count - 1)) / columns.count;
  }
  const drawing = word?.drawings?.find(
    (drawing) =>
      (drawing.kind === "textBox" || drawing.kind === "shape") &&
      drawing.textStoryId === story?.storyId,
  );
  if (drawing?.kind === "textBox" || drawing?.kind === "shape") {
    const side =
        section?.gutterPosition === "top"
          ? "top"
          : section?.gutterPosition === "right"
            ? "right"
            : "left",
      effective = {
        ...setup,
        margins: { ...setup.margins, [side]: setup.margins[side] + (section?.gutter ?? 0) },
      },
      size = wordDrawingDimensions(drawing, { index: 0, setup: effective }, section),
      textWidth =
        drawing.kind === "shape"
          ? (evaluatePresentationPreset(
              drawing.preset,
              size.width,
              size.height,
              drawing.adjustments,
            ).textRect?.width ?? size.width)
          : size.width;
    width = textWidth - (drawing.text?.insets?.left ?? 7.2) - (drawing.text?.insets?.right ?? 7.2);
  }
  const projected = new Map<string, SourceStyleNode>(),
    project = (id: string): SourceStyleNode => {
      let result = projected.get(id);
      if (!result) {
        const node = source.peekNode(id)!;
        result = {
          tag: node.tagName,
          attrs: node.attributes,
          parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
        };
        projected.set(id, result);
      }
      return result;
    };
  const size = (value: string | undefined, font: number, percent: number): number =>
    value && value !== "auto" && value !== "normal" ? length(value, { font, percent }) : 0;
  let tableId: string | undefined,
    tableWidth = width;
  for (const id of ancestors.slice(0, -1)) {
    const node = source.peekNode(id)!,
      css = resolver.get(project(id)),
      font = size(css["font-size"], 12, 12) || 12;
    if (node.tagName === "table") {
      tableId = id;
      tableWidth =
        word?.tables
          .find((table) => table.nodeId === id)
          ?.grid.reduce((sum, item) => sum + item, 0) ??
        (size(css.width, font, width) || width);
      width = tableWidth;
    } else if (node.tagName === "td" || node.tagName === "th") {
      const info = word?.tables.find((table) => table.nodeId === tableId),
        table = tableId ? source.peekNode(tableId) : undefined,
        rows =
          table?.childNodeIds.flatMap((id) => {
            const row = source.peekNode(id)!;
            return row.tagName === "tr"
              ? [id]
              : row.childNodeIds.filter((id) => source.peekNode(id)!.tagName === "tr");
          }) ?? [],
        occupied: number[] = [],
        rowCount = rows.length;
      let cellWidth: number | undefined,
        columns = 0;
      for (const [r, rowId] of rows.entries()) {
        let column = 0;
        for (const cellId of source.peekNode(rowId)!.childNodeIds) {
          const cell = source.peekNode(cellId)!;
          if (!["td", "th"].includes(cell.tagName)) continue;
          while ((occupied[column] ?? 0) > r) column++;
          const span = Number(cell.attributes.colspan ?? 1),
            height = Number(cell.attributes.rowspan ?? 1) || rowCount - r;
          checkLimit((column + span) * rowCount, limits.maxNodes);
          if (cellId === id && info)
            cellWidth = info.grid.slice(column, column + span).reduce((sum, item) => sum + item, 0);
          for (let c = column; c < column + span; c++) occupied[c] = r + height;
          column += span;
          columns = Math.max(columns, column);
        }
      }
      width =
        cellWidth ??
        (size(css.width, font, tableWidth) ||
          (tableWidth * Math.max(1, Number(node.attributes.colspan ?? 1))) / Math.max(1, columns));
      width -=
        size(css["padding-left"], font, width) +
        size(css["padding-right"], font, width) +
        size(css["border-left-width"], font, width) +
        size(css["border-right-width"], font, width);
    } else if (
      !node.attributes["data-office-kind"] &&
      !word?.stories.some((story) => story.rootNodeId === id) &&
      !metadata.sections.some((section) => section.rootNodeId === id)
    )
      width = size(css.width, font, width) || width;
  }
  return Math.max(1, width);
}
