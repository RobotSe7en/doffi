import { DocumentError, resolveWordGridInsertion } from "@doffi/core";
import type { WordMetadata, WordRevision } from "@doffi/core";
import { unsupported } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
interface Node {
  tag: string;
  attrs: Record<string, string>;
  children: (Node | string)[];
}
/** A column-count history is editable only when the native marks and old cell
 * spans describe the complete inverse. Ambiguous native structures stay native. */
export function bindWordGridInsertion(
  word: WordMetadata,
  table: Node,
  revision: WordRevision,
  source: XmlNode,
  maxCells: number,
): void {
  const tableId = table.attrs["data-node-id"]!,
    info = word.tables.find((info) => info.nodeId === tableId)!,
    previous = revision.before!.grid!;
  const rows = table.children.filter(
      (node): node is Node => typeof node !== "string" && node.tag === "tr",
    ),
    cells: { node: Node; column: number; columns: number; rows: number; row: number }[] = [],
    occupied: (typeof cells)[] = Array.from({ length: rows.length }, () => []);
  if (rows.length * info.grid.length > maxCells)
    throw new DocumentError(
      "ResourceLimitExceeded",
      "Native grid revision mapping exceeds its cell budget.",
    );
  const own = new Map(
    word.revisions.map((item) => [item.nodeId + "\n" + (item.property ?? item.kind), item]),
  );
  for (const [row, node] of rows.entries())
    for (const child of node.children) {
      if (typeof child === "string") continue;
      let column = 0;
      while (occupied[row]![column]) column++;
      const cell = {
        node: child,
        column,
        columns: Number(child.attrs.colspan ?? 1),
        rows: Number(child.attrs.rowspan ?? 1),
        row,
      };
      cells.push(cell);
      for (let r = row; r < row + cell.rows; r++)
        for (let c = column; c < column + cell.columns; c++) occupied[r]![c] = cell;
    }
  const historyRows = occupied.map((row, index) => {
    const result: import("@doffi/core").WordGridHistoryCell[] = [];
    for (let column = 0; column < row.length; ) {
      const cell = row[column];
      if (!cell) return [];
      const id = cell.node.attrs["data-node-id"]!,
        gap = cell.node.attrs["data-word-grid-gap"],
        format = own.get(id + "\ncell"),
        rowHistory = own.get(rows[index]!.attrs["data-node-id"] + "\nrow");
      const old =
        gap === "before" || gap === "after"
          ? (rowHistory?.before?.rowOmissions?.[gap] ?? cell.columns)
          : own.has(id + "\ncellInsert")
            ? 0
            : format?.before?.attributes && "colspan" in format.before.attributes
              ? Number(format.before.attributes.colspan ?? 1)
              : cell.columns;
      result.push({ column, columns: cell.columns, beforeColumns: old });
      column += cell.columns;
    }
    return result;
  });
  const resolved = resolveWordGridInsertion(
    info.grid.length,
    previous.length,
    historyRows,
    maxCells,
  );
  if (!resolved)
    unsupported(
      source,
      "The changed native column count has no complete cell/span/omission inverse.",
    );
  const groupId = tableId + "-column-insertion",
    ids = new Set([
      tableId,
      ...rows.map((row) => row.attrs["data-node-id"]!),
      ...cells.map((cell) => cell.node.attrs["data-node-id"]!),
    ]);
  revision.groupId = groupId;
  for (const item of word.revisions)
    if (
      ids.has(item.nodeId) &&
      (item.kind === "cellInsert" ||
        (item.kind === "format" && ["cell", "row"].includes(item.property ?? "")))
    ) {
      item.groupId = groupId;
      if (item.kind === "cellInsert" || item.kind === "format") {
        if (item.author !== undefined) revision.author = item.author;
        if (item.date !== undefined) revision.date = item.date;
      }
    }
}
