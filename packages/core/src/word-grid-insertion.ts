import { DocumentError } from "./errors.js";
export interface WordGridHistoryCell {
  readonly column: number;
  readonly columns: number;
  /** Zero belongs to an inserted cell; omissions retain their previous count. */
  readonly beforeColumns: number;
}
/** All physical row boundaries constrain one monotone insertion map. Within a
 * region spanned identically by every row, column identities are equivalent:
 * the inverse restores the exact old grid, spans and content in either case. */
export function resolveWordGridInsertion(
  current: number,
  previous: number,
  rows: readonly (readonly WordGridHistoryCell[])[],
  maxCells: number,
): readonly number[] | undefined {
  if (
    !Number.isSafeInteger(current) ||
    !Number.isSafeInteger(previous) ||
    previous < 1 ||
    current <= previous
  )
    return;
  if (current * Math.max(1, rows.length) > maxCells)
    throw new DocumentError(
      "ResourceLimitExceeded",
      "Column insertion mapping exceeds its grid budget.",
    );
  const delta = current - previous,
    boundaries = new Map<number, number>([
      [0, 0],
      [current, delta],
    ]);
  const bind = (column: number, count: number): boolean => {
    if (
      column < 0 ||
      column > current ||
      count < 0 ||
      count > delta ||
      (boundaries.has(column) && boundaries.get(column) !== count)
    )
      return false;
    boundaries.set(column, count);
    return true;
  };
  for (const row of rows) {
    let old = 0,
      column = 0;
    for (const cell of row) {
      if (
        cell.column !== column ||
        !Number.isSafeInteger(cell.columns) ||
        cell.columns < 1 ||
        !Number.isSafeInteger(cell.beforeColumns) ||
        cell.beforeColumns < 0 ||
        cell.beforeColumns > cell.columns
      )
        return;
      if (!bind(column, column - old)) return;
      column += cell.columns;
      old += cell.beforeColumns;
      if (!bind(column, column - old)) return;
    }
    if (column !== current || old !== previous) return;
  }
  const points = [...boundaries].sort(([a], [b]) => a - b),
    inserted: number[] = [];
  for (let index = 1; index < points.length; index++) {
    const [from, before] = points[index - 1]!,
      [to, after] = points[index]!,
      count = after - before;
    if (count < 0 || count > to - from) return;
    for (let column = to - count; column < to; column++) inserted.push(column);
  }
  return inserted.length === delta ? inserted : undefined;
}
