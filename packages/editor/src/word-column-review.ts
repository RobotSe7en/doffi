import { DocumentError, resolveWordGridInsertion } from "@doffi/core";
import type { DeepReadonly, WordMetadata } from "@doffi/core";
import type { SourceProjection } from "@doffi/render";
/** Keep pending cells in the live DOM for source reconciliation. Only their
 * displayed spans and widths use the original/accepted column projection. */
export function projectWordColumnReview(
  root: HTMLElement,
  word: DeepReadonly<WordMetadata> | undefined,
  display: "all" | "original" | "final",
  projection: SourceProjection,
  maxNodes: number,
): void {
  if (!word || display === "all") return;
  const hiddenKind = display === "original" ? "cellInsert" : "cellDelete",
    hidden = new Set(
      word.revisions
        .filter((revision) => revision.kind === hiddenKind)
        .map((revision) => revision.nodeId),
    );
  if (!hidden.size && !word.revisions.some((item) => item.property === "tableGrid")) return;
  const histories = new Map(
    word.revisions.map((item) => [item.nodeId + "\n" + (item.property ?? item.kind), item]),
  );
  const tables = new Map(word.tables.map((table) => [table.nodeId, table])),
    oldGrids = new Map(
      word.revisions
        .filter((revision) => revision.property === "tableGrid")
        .map((revision) => [revision.nodeId, revision.before!.grid!]),
    );
  let work = 0;
  for (const table of root.querySelectorAll<HTMLTableElement>("table[data-node-id]")) {
    if (projection.ignored(table)) continue;
    const info = tables.get(table.dataset.nodeId!);
    if (!info) continue;
    const rows = [...table.rows].filter(
        (row) => row.closest("table") === table && !projection.ignored(row),
      ),
      cells: { node: HTMLTableCellElement; column: number; columns: number; rows: number }[] = [],
      occupied: (typeof cells)[] = Array.from({ length: rows.length }, () => []),
      candidates = new Set<number>();
    work += rows.length * info.grid.length;
    if (work > maxNodes)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Column review exceeds its table projection budget.",
      );
    for (const [row, node] of rows.entries())
      for (const child of [...node.cells].filter((cell) => !projection.ignored(cell))) {
        let column = 0;
        while (occupied[row]![column]) column++;
        const cell = {
          node: child,
          column,
          columns: child.colSpan,
          rows: child.rowSpan || rows.length - row,
        };
        cells.push(cell);
        for (let r = row; r < Math.min(rows.length, row + cell.rows); r++)
          for (let c = column; c < column + cell.columns; c++) occupied[r]![c] = cell;
        if (hidden.has(child.dataset.nodeId!))
          for (let c = column; c < column + cell.columns; c++) candidates.add(c);
      }
    const previous = oldGrids.get(info.nodeId);
    if (display === "original" && previous && previous.length < info.grid.length) {
      const segments = occupied.map((row, index) => {
        const out: import("@doffi/core").WordGridHistoryCell[] = [];
        for (let column = 0; column < row.length; ) {
          const cell = row[column];
          if (!cell) return [];
          const id = cell.node.dataset.nodeId!,
            gap = cell.node.dataset.wordGridGap,
            format = histories.get(id + "\ncell"),
            rowHistory = histories.get(rows[index]!.dataset.nodeId + "\nrow");
          const old =
            gap === "before" || gap === "after"
              ? (rowHistory?.before?.rowOmissions?.[gap] ?? cell.columns)
              : hidden.has(id)
                ? 0
                : format?.before?.attributes && "colspan" in format.before.attributes
                  ? Number(format.before.attributes.colspan ?? 1)
                  : cell.columns;
          out.push({ column, columns: cell.columns, beforeColumns: old });
          column += cell.columns;
        }
        return out;
      });
      const resolved = resolveWordGridInsertion(
        info.grid.length,
        previous.length,
        segments,
        maxNodes,
      );
      if (resolved) {
        candidates.clear();
        for (const column of resolved) candidates.add(column);
      }
    }
    if (!candidates.size) continue;
    const inserted =
      display === "original" && previous && previous.length + candidates.size === info.grid.length;
    const removed = [...candidates]
      .filter(
        (column) =>
          inserted ||
          occupied.every((row) => {
            const cell = row[column];
            return (
              cell &&
              (hidden.has(cell.node.dataset.nodeId!) ||
                cell.node.hasAttribute("data-word-grid-gap"))
            );
          }),
      )
      .sort((a, b) => a - b);
    if (!removed.length) continue;
    const removedSet = new Set(removed),
      counts = [0];
    for (let column = 0; column < info.grid.length; column++)
      counts.push(counts.at(-1)! + (removedSet.has(column) ? 1 : 0));
    const before = (column: number) => counts[column] ?? 0,
      total = info.grid.reduce((a, b) => a + b, 0),
      remaining = info.grid.filter((_, index) => !removedSet.has(index)),
      remainingWidth = remaining.reduce((a, b) => a + b, 0),
      grid = inserted ? [...previous] : remaining.map((width) => (width * total) / remainingWidth);
    const group = table.querySelector(":scope > colgroup");
    if (group && projection.ignored(group))
      group.replaceChildren(
        ...grid.map((width) => {
          const column = table.ownerDocument.createElement("col");
          column.style.width = width + "pt";
          return column;
        }),
      );
    for (const cell of cells) {
      if (hidden.has(cell.node.dataset.nodeId!)) continue;
      const column = cell.column - before(cell.column),
        columns = cell.columns - (before(cell.column + cell.columns) - before(cell.column));
      if (!columns) {
        projection.setViewStyles(cell.node, { display: "none" });
        continue;
      }
      if (columns !== cell.columns)
        projection.setViewAttribute(cell.node, "colspan", columns === 1 ? null : String(columns));
      projection.setViewStyles(cell.node, {
        width: grid.slice(column, column + columns).reduce((a, b) => a + b, 0) + "pt",
      });
    }
  }
}
