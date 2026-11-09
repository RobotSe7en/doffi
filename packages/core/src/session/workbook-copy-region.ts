import type { WorkbookCoordinate, WorkbookRange } from "../contracts/workbook.js";
import { WorkbookCoordinateSchema } from "../contracts/workbook.js";
import type { DeepReadonly } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { checkNonOverlappingRanges, checkRange } from "../workbook/ranges.js";

export interface WorkbookRegionCopy {
  readonly sheetId: string; readonly range: DeepReadonly<WorkbookRange>;
  readonly targetSheetId: string; readonly target: DeepReadonly<WorkbookCoordinate>;
  readonly ranges?: readonly DeepReadonly<WorkbookRange>[] | undefined;
}
/** A mask describes actual destination cells, not its bounding rectangle.
 * All consumers use the same native bounds and source/destination translation. */
export function planWorkbookRegionCopy(operation: WorkbookRegionCopy) {
  checkRange(operation.range);
  const delta = { row: operation.target.row - operation.range.from.row, column: operation.target.column - operation.range.from.column };
  const translate = (range: DeepReadonly<WorkbookRange>, sign = 1): WorkbookRange => ({
    from: { row: range.from.row + delta.row * sign, column: range.from.column + delta.column * sign },
    to: { row: range.to.row + delta.row * sign, column: range.to.column + delta.column * sign },
  });
  const target = translate(operation.range);
  if (!WorkbookCoordinateSchema.safeParse(target.to).success) throw new DocumentError("InvalidOperation", "Range copying exceeds the destination worksheet bounds.");
  const footprint = (operation.ranges ?? [target]).map(range => translate(range, 0));
  for (const range of footprint) {
    checkRange(range);
    if (range.from.row < target.from.row || range.to.row > target.to.row || range.from.column < target.from.column || range.to.column > target.to.column)
      throw new DocumentError("InvalidOperation", "The copy footprint must lie inside its destination rectangle.");
  }
  checkNonOverlappingRanges(footprint);
  return { delta, target, footprint, source: footprint.map(range => translate(range, -1)), translate,
    unchanged: operation.sheetId === operation.targetSheetId && delta.row === 0 && delta.column === 0 };
}
