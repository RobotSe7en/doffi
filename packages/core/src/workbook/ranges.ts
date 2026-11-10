import type { WorkbookRange } from "../contracts/workbook.js";
import { invalid } from "../document/validation.js";
import { MAX_SHEET_COLUMNS } from "./chunk-index.js";

export function checkRange(range: WorkbookRange): void {
  if (range.from.row > range.to.row || range.from.column > range.to.column) invalid("Workbook ranges must be ordered.");
}
/** Sweep rows with a column segment tree; no quadratic pairwise merge scan. */
export function checkNonOverlappingRanges(ranges: readonly WorkbookRange[]): void {
  const events = ranges.flatMap(range => { checkRange(range); return [{ row: range.from.row, delta: 1, range }, { row: range.to.row + 1, delta: -1, range }]; }).sort((a, b) => a.row - b.row || a.delta - b.delta);
  const maximum = new Int32Array(MAX_SHEET_COLUMNS * 4), lazy = new Int32Array(MAX_SHEET_COLUMNS * 4);
  function update(node: number, left: number, right: number, from: number, to: number, delta: number): void {
    if (from <= left && right <= to) { maximum[node]! += delta; lazy[node]! += delta; return; }
    const middle = (left + right) >>> 1;
    if (from <= middle) update(node * 2, left, middle, from, to, delta);
    if (to > middle) update(node * 2 + 1, middle + 1, right, from, to, delta);
    maximum[node] = lazy[node]! + Math.max(maximum[node * 2]!, maximum[node * 2 + 1]!);
  }
  for (const event of events) {
    update(1, 0, MAX_SHEET_COLUMNS - 1, event.range.from.column, event.range.to.column, event.delta);
    if (maximum[1]! > 1) invalid("Workbook regions must not overlap.");
  }
}

/** Compare unions of validated nonoverlapping rectangles without expanding sparse axes or
 * treating a different fragmentation as a source change. */
export function sameWorkbookRangeCoverage(left: readonly WorkbookRange[], right: readonly WorkbookRange[]): boolean {
  if (left.length === right.length && left.every((range, index) => {
    const other = right[index]!; return range.from.row === other.from.row && range.from.column === other.from.column && range.to.row === other.to.row && range.to.column === other.to.column;
  })) return true;
  const events = [...left.flatMap(range => [{ row: range.from.row, delta: 1, range }, { row: range.to.row + 1, delta: -1, range }]),
    ...right.flatMap(range => [{ row: range.from.row, delta: -1, range }, { row: range.to.row + 1, delta: 1, range }])].sort((a, b) => a.row - b.row);
  const maximum = new Int32Array(MAX_SHEET_COLUMNS * 4), minimum = new Int32Array(MAX_SHEET_COLUMNS * 4), lazy = new Int32Array(MAX_SHEET_COLUMNS * 4);
  function update(node: number, from: number, to: number, first: number, last: number, delta: number): void {
    if (first <= from && to <= last) { maximum[node]! += delta; minimum[node]! += delta; lazy[node]! += delta; return; }
    const middle = (from + to) >>> 1;
    if (first <= middle) update(node * 2, from, middle, first, last, delta);
    if (last > middle) update(node * 2 + 1, middle + 1, to, first, last, delta);
    maximum[node] = lazy[node]! + Math.max(maximum[node * 2]!, maximum[node * 2 + 1]!);
    minimum[node] = lazy[node]! + Math.min(minimum[node * 2]!, minimum[node * 2 + 1]!);
  }
  for (let index = 0; index < events.length;) {
    const row = events[index]!.row;
    do { const event = events[index++]!; update(1, 0, MAX_SHEET_COLUMNS - 1, event.range.from.column, event.range.to.column, event.delta); } while (index < events.length && events[index]!.row === row);
    if (maximum[1] !== 0 || minimum[1] !== 0) return false;
  }
  return true;
}
