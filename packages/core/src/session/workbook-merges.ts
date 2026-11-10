import type { WorkbookRange, WorkbookSheet } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { DeepReadonly } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { checkRange } from "../workbook/ranges.js";
import { subtractWorkbookRange } from "../workbook/range-mapping.js";
import { rangesIntersect } from "./workbook-cells.js";
import type { WorkbookCellEdit } from "./workbook-cells.js";

type MergeOperation = Extract<WorkbookOperation, { op: "mergeWorkbookCells" | "unmergeWorkbookCells" }>;
const contains = (outer: WorkbookRange, inner: WorkbookRange): boolean => outer.from.row <= inner.from.row && outer.from.column <= inner.from.column && outer.to.row >= inner.to.row && outer.to.column >= inner.to.column;
const same = (a: DeepReadonly<WorkbookRange>, b: DeepReadonly<WorkbookRange>): boolean => a.from.row === b.from.row && a.from.column === b.from.column && a.to.row === b.to.row && a.to.column === b.to.column;
export function sameWorkbookMerges(left: readonly DeepReadonly<WorkbookRange>[], right: readonly DeepReadonly<WorkbookRange>[]): boolean {
  const key = (range: DeepReadonly<WorkbookRange>): string => [range.from.row, range.from.column, range.to.row, range.to.column].join(":");
  const keys = new Set(left.map(key)); return keys.size === left.length && left.length === right.length && new Set(right.map(key)).size === right.length && right.every(range => keys.has(key(range)));
}
/** Only explicit clearing can discard covered source. Reads/clears visit stored
 * cells in intersecting chunks; enormous sparse merges never expand an area. */
export function* editWorkbookMerge(sheet: WorkbookSheet, cells: WorkbookCellEdit, operation: MergeOperation): Generator<{ completed: number; total: number }, void> {
  const { range } = operation; checkRange(range);
  const intersected = sheet.merges.filter(merge => rangesIntersect(merge, range));
  if (operation.op === "unmergeWorkbookCells") { sheet.merges = sheet.merges.filter(merge => !rangesIntersect(merge, range)); return; }
  if (range.from.row === range.to.row && range.from.column === range.to.column) throw new DocumentError("InvalidOperation", "Merging requires at least two cells.");
  if (intersected.some(merge => same(merge, range))) return;
  if (intersected.some(merge => !contains(range, merge))) throw new DocumentError("InvalidOperation", "A merge cannot split an existing merged region.");
  if (sheet.tables?.some(table => rangesIntersect(table.range, range))) throw new DocumentError("InvalidOperation", "Table cells cannot be merged.");
  const covered = subtractWorkbookRange(range, { from: range.from, to: range.from });
  if (operation.coveredContent === "clear") yield* cells.clearSteps(sheet.sheetId, covered, "contents");
  else {
    let completed = 0;
    for (const cell of cells.entries(sheet.sheetId, false, covered)) {
      if (cell.formula || cell.value.type !== "blank") throw new DocumentError("InvalidOperation", "The merge covers other cell content; choose explicit clearing or cancel.", [{ code: "merge-covered-content", path: [sheet.sheetId, cell.row, cell.column], message: "Covered cell content requires an explicit clearing choice." }]);
      if (++completed % 256 === 0) yield { completed, total: 0 };
    }
    if (completed % 256 !== 0) yield { completed, total: completed };
  }
  const replaced = new Set(intersected);
  sheet.merges = sheet.merges.filter(merge => !replaced.has(merge));
  sheet.merges.push({ from: { ...range.from }, to: { ...range.to } });
}
