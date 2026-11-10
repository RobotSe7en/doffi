import type { WorkbookRange, WorkbookSheet } from "../contracts/workbook.js";
import { WorkbookRangeSchema } from "../contracts/workbook.js";
import type { DeepReadonly } from "../runtime.js";
import { DocumentError } from "../errors.js";

export interface WorkbookSortLayout { readonly rowsPerRecord: number; readonly columnsPerCell: number }
/** Native merged sort cells must tile the selected rectangle at one size.
 * Inspect only stored merges, never expand the worksheet or selected cells. */
export function resolveWorkbookSortLayout(sheet: Pick<DeepReadonly<WorkbookSheet>, "merges">, range: DeepReadonly<WorkbookRange>): WorkbookSortLayout {
  if (!WorkbookRangeSchema.safeParse(range).success || range.from.row > range.to.row || range.from.column > range.to.column)
    throw new DocumentError("InvalidOperation", "Sorting requires an ordered worksheet range.");
  const merges = sheet.merges.filter(merge => merge.from.row <= range.to.row && range.from.row <= merge.to.row && merge.from.column <= range.to.column && range.from.column <= merge.to.column);
  const first = merges[0]; if (!first) return Object.freeze({ rowsPerRecord: 1, columnsPerCell: 1 });
  const rowsPerRecord = first.to.row - first.from.row + 1, columnsPerCell = first.to.column - first.from.column + 1;
  const height = range.to.row - range.from.row + 1, width = range.to.column - range.from.column + 1;
  if (height % rowsPerRecord || width % columnsPerCell || merges.length * rowsPerRecord * columnsPerCell !== height * width
    || merges.some(merge => merge.from.row < range.from.row || merge.to.row > range.to.row || merge.from.column < range.from.column || merge.to.column > range.to.column
      || merge.to.row - merge.from.row + 1 !== rowsPerRecord || merge.to.column - merge.from.column + 1 !== columnsPerCell
      || (merge.from.row - range.from.row) % rowsPerRecord || (merge.from.column - range.from.column) % columnsPerCell))
    throw new DocumentError("InvalidOperation", "Sorted merged cells must occupy a complete grid of equal row and column sizes.");
  return Object.freeze({ rowsPerRecord, columnsPerCell });
}
