import type { WordMetadata } from "../contracts/word.js";
import type { WordTableEmptyParagraph } from "../contracts/word-table-continuation.js";
import type { ParsedHtml } from "../document/html.js";
import type { DocumentEdit } from "../document/edit.js";
import type { DocumentReadLimits } from "../runtime.js";
import { checkLimit } from "../document/validation.js";
import { DocumentError } from "../errors.js";
import type { SourceStyleEdit } from "./source-styles.js";

type Cell = { id: string; row: number; column: number; rows: number; columns: number };
type Continuation = NonNullable<WordMetadata["tables"][number]["continuationFormats"]>[number];
const escape = (text: string): string =>
  text.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
const attributes = (values: Record<string, string | undefined>): string =>
  Object.entries(values)
    .flatMap(([key, value]) =>
      value === undefined ? [] : [" " + key + '="' + escape(value) + '"'],
    )
    .join("");
const rows = (source: ParsedHtml, id: string): string[] =>
  source.peekNode(id)!.childNodeIds.flatMap((id) => {
    const node = source.peekNode(id)!;
    return node.tagName === "tr"
      ? [id]
      : node.childNodeIds.filter((id) => source.peekNode(id)!.tagName === "tr");
  });

function materialize(
  word: WordMetadata,
  edit: DocumentEdit,
  styles: SourceStyleEdit,
  entry: Continuation,
  cellId: string,
): void {
  const paragraphs = entry.emptyParagraphs ??
    entry.paragraphs?.map((format) => ({ ...(format ? { format } : {}), runs: [] })) ?? [
      { runs: [] },
    ];
  edit.removeSource(cellId, true);
  if (entry.cellCssClass)
    edit.attributes(cellId, {
      class: entry.cellCssClass,
      style: Object.entries(styles.resolver().specifiedClass(entry.cellCssClass))
        .map(([key, value]) => key + ":" + value)
        .join(";"),
      "data-word-text-direction": entry.textDirection,
    });
  if (entry.format) {
    word.conditionalFormats ??= [];
    word.conditionalFormats = word.conditionalFormats.filter((item) => item.nodeId !== cellId);
    word.conditionalFormats.push({
      nodeId: cellId,
      format: JSON.parse(JSON.stringify(entry.format)),
    });
  }
  for (const paragraph of paragraphs) {
    const definition: WordTableEmptyParagraph = paragraph,
      style = word.styles.find((style) => style.styleId === definition.styleId),
      level = definition.outlineLevel ?? style?.outlineLevel,
      tag = level !== undefined && level < 6 ? "h" + (level + 1) : "p";
    const attrs = attributes({
      class: [style?.cssClass, definition.paragraphCssClass].filter(Boolean).join(" "),
      "data-word-style": definition.styleId,
      "data-word-outline-level":
        definition.outlineLevel === undefined ? undefined : String(definition.outlineLevel),
      "data-word-numbering": definition.numbering?.numberingId,
      "data-word-level": definition.numbering ? String(definition.numbering.level) : undefined,
      "data-word-numbering-off": definition.numberingDisabled ? "true" : undefined,
    });
    const mark = definition.runCssClass
      ? '<span data-word-paragraph-mark="true" class="' +
        escape(definition.runCssClass) +
        '"></span>'
      : "";
    const runs = definition.runs
      .map(
        (run) =>
          "<span" +
          attributes({
            "data-word-run": "true",
            "data-word-style": run.styleId,
            class: [
              word.styles.find((style) => style.styleId === run.styleId)?.cssClass,
              run.cssClass,
            ]
              .filter(Boolean)
              .join(" "),
          }) +
          "></span>",
      )
      .join("");
    const id = edit.insertSource(cellId, "<" + tag + attrs + ">" + mark + runs + "</" + tag + ">")
      .rootNodeIds[0]!;
    if (definition.format) {
      word.conditionalFormats ??= [];
      word.conditionalFormats.push({
        nodeId: id,
        format: JSON.parse(JSON.stringify(definition.format)),
      });
    }
    {
      const clears = (definition.tabs ?? []).filter((tab) => tab.alignment === "clear");
      word.tabs.push({
        paragraphNodeId: id,
        defaultInterval: definition.defaultTabInterval ?? word.tabs[0]?.defaultInterval ?? 36,
        stops: [
          ...(definition.resolvedTabs ??
            definition.tabs?.filter((tab) => tab.alignment !== "clear") ??
            []),
          ...clears,
        ].map((tab) => ({ ...tab })),
      });
    }
  }
}

/** Refined row/column boundaries keep old physical continuation slots. Slots
 * that become real cells receive their empty paragraphs; other merged slots
 * retain metadata, including outside the selected split. */
export function restoreWordSplitContinuations(
  word: WordMetadata,
  table: WordMetadata["tables"][number],
  original: ParsedHtml,
  originalRows: readonly string[],
  originalGrid: readonly number[],
  previous: readonly Continuation[],
  cells: readonly Cell[],
  edit: DocumentEdit,
  source: ParsedHtml,
  limits: DocumentReadLimits,
  styles: SourceStyleEdit,
): void {
  if (!previous.length) return;
  const currentRows = rows(source, table.nodeId),
    width = table.grid.length,
    oldWidth = originalGrid.length;
  checkLimit(originalRows.length * oldWidth, limits.maxNodes);
  checkLimit(currentRows.length * width, limits.maxNodes);
  const old: Array<{ column: number; columns: number } | undefined> = new Array(
      originalRows.length * oldWidth,
    ),
    occupied: Array<Cell | undefined> = new Array(currentRows.length * width);
  for (const [row, rowId] of originalRows.entries())
    for (const id of original.peekNode(rowId)!.childNodeIds) {
      const node = original.peekNode(id)!;
      if (!["td", "th"].includes(node.tagName)) continue;
      let column = 0;
      while (old[row * oldWidth + column]) column++;
      const columns = Number(node.attributes.colspan ?? 1),
        height = Number(node.attributes.rowspan ?? 1) || originalRows.length - row;
      for (let r = row; r < row + height; r++)
        for (let c = column; c < column + columns; c++) old[r * oldWidth + c] = { column, columns };
    }
  for (const cell of cells)
    for (let r = cell.row; r < cell.row + cell.rows; r++)
      for (let c = cell.column; c < cell.column + cell.columns; c++) occupied[r * width + c] = cell;
  const rowIndex = new Map(currentRows.map((id, index) => [id, index])),
    positions = [0];
  for (const size of table.grid) positions.push(positions.at(-1)! + size);
  const columns: number[] = [];
  let offset = 0,
    index = 0;
  for (const size of [...originalGrid, 0]) {
    while (index < positions.length - 1 && positions[index]! < offset - 1e-5) index++;
    if (Math.abs(positions[index]! - offset) > 1e-5)
      throw new DocumentError("InvalidOperation", "A split lost an original column boundary.");
    columns.push(index);
    offset += size;
  }
  const retained: Continuation[] = [],
    materializations = new Map<string, Continuation>();
  for (const entry of previous) {
    const row = rowIndex.get(originalRows[entry.row]!)!,
      oldCell = old[entry.row * oldWidth + entry.column];
    if (!oldCell)
      throw new DocumentError("InvalidOperation", "A continuation has no original merged cell.");
    const start = columns[entry.column]!,
      end = columns[oldCell.column + oldCell.columns]!;
    for (let column = start; column < end; ) {
      const cell = occupied[row * width + column];
      if (!cell)
        throw new DocumentError("InvalidOperation", "A split continuation has no current cell.");
      const copy: Continuation = {
        ...(JSON.parse(JSON.stringify(entry)) as Continuation),
        row,
        column: cell.column,
      };
      if (cell.row === row && !original.hasNode(cell.id)) {
        if (materializations.has(cell.id))
          throw new DocumentError(
            "InvalidOperation",
            "Several continuation formats target the same split cell.",
          );
        materializations.set(cell.id, copy);
      } else if (cell.row < row) retained.push(copy);
      column = cell.column + cell.columns;
    }
  }
  // Bound the expanded source before copying a large number of empty
  // paragraphs into many subdivisions. Candidate failure stays atomic.
  let nodes = source.nodeIds().length;
  for (const [id, entry] of materializations) {
    const pending = [...source.peekNode(id)!.childNodeIds];
    for (let index = 0; index < pending.length; index++)
      pending.push(...source.peekNode(pending[index]!)!.childNodeIds);
    nodes -= pending.length;
    const paragraphs = entry.emptyParagraphs ??
      entry.paragraphs?.map((format) => ({ ...(format ? { format } : {}), runs: [] })) ?? [
        { runs: [] },
      ];
    for (const paragraph of paragraphs)
      nodes +=
        1 + paragraph.runs.length + ("runCssClass" in paragraph && paragraph.runCssClass ? 1 : 0);
    if (nodes > limits.maxNodes)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Split continuation paragraphs exceed the source node budget.",
      );
  }
  for (const [id, entry] of materializations) materialize(word, edit, styles, entry, id);
  table.continuationFormats = retained;
}
