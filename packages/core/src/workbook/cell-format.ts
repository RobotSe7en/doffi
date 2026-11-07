import { DocumentError } from "../errors.js";
import type { DeepReadonly, WorkbookCell } from "../runtime.js";
import type { WorkbookCoordinate, WorkbookSheet } from "../contracts/workbook.js";

export interface WorkbookCellFormat { readonly cssClass: string; readonly numberFormatId?: string }
type Dimension = Pick<DeepReadonly<WorkbookSheet>["rowDimensions"][number], "from" | "to" | "cssClass" | "numberFormatId">;
function at(dimensions: readonly Dimension[], coordinate: number): Dimension | undefined {
  let low = 0, high = dimensions.length;
  while (low < high) { const mid = (low + high) >>> 1; if (dimensions[mid]!.to < coordinate) low = mid + 1; else high = mid; }
  const item = dimensions[low]; return item && item.from <= coordinate ? item : undefined;
}
/** Resolve source format identities without evaluating CSS or expanding sparse
 * dimensions. Class order remains governed by the stylesheet cascade. */
export function resolveWorkbookCellFormat(sheet: DeepReadonly<WorkbookSheet>, coordinate: WorkbookCoordinate, cell?: WorkbookCell): WorkbookCellFormat {
  const { row, column } = coordinate;
  if (!Number.isSafeInteger(row) || row < 0 || row >= 1_048_576 || !Number.isSafeInteger(column) || column < 0 || column >= 16_384
    || cell && (cell.sheetId !== sheet.sheetId || cell.row !== row || cell.column !== column)) throw new DocumentError("InvalidOperation", "Cell format resolution requires a matching worksheet coordinate.");
  const rowStyle = at(sheet.rowDimensions, row), columnStyle = at(sheet.columnDimensions, column);
  const cssClass = cell?.cssClass ?? rowStyle?.cssClass ?? columnStyle?.cssClass ?? sheet.defaultCssClass;
  const numberFormatId = cell?.formatId ?? rowStyle?.numberFormatId ?? columnStyle?.numberFormatId ?? sheet.defaultNumberFormatId;
  return Object.freeze({ cssClass, ...(numberFormatId === undefined ? {} : { numberFormatId }) });
}
