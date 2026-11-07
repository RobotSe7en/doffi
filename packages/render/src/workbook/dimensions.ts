import { DocumentError, WorkbookAxisDimensions as AxisDimensions } from "@doffi/core";
import type { DeepReadonly, WorkbookSheet, WorkbookAxisOverride as AxisOverride } from "@doffi/core";
export { WorkbookAxisDimensions as AxisDimensions } from "@doffi/core";
export type { WorkbookAxisOverride as AxisOverride, WorkbookAxisCell as AxisCell } from "@doffi/core";
const invalid = (message: string): never => { throw new DocumentError("InvalidOperation", message); };

/** Stored OOXML width uses the Normal font's maximum digit advance at 96 dpi.
 * The caller measures the resolved font; substitutions remain render diagnostics.
 * Source character widths never become authoritative screen-pixel widths. */
export function columnWidthPt(width: number, maximumDigitWidthPx: number, unit: "characters" | "baseCharacters" = "characters"): number {
  if (!Number.isFinite(width) || width < 0 || width > 255 || !Number.isFinite(maximumDigitWidthPx) || maximumDigitWidthPx <= 0 || maximumDigitWidthPx > 16384) invalid("Invalid column width or measured digit advance.");
  const stored = unit === "baseCharacters" ? Math.floor((width * maximumDigitWidthPx + 5) / maximumDigitWidthPx * 256) / 256 : width;
  return Math.floor((256 * stored + Math.floor(128 / maximumDigitWidthPx)) / 256 * maximumDigitWidthPx) * 0.75;
}

/** Automatic row height uses the stored pre-hide width. Display geometry
 * keeps hidden columns at zero extent; measurement never changes that axis. */
export function worksheetColumnDimensions(sheet: DeepReadonly<WorkbookSheet>, maximumDigitWidthPx: number, includeHidden = false): AxisDimensions {
  const column = columnWidthPt(sheet.defaultColumnWidth, maximumDigitWidthPx, sheet.defaultColumnWidthUnit);
  return new AxisDimensions(16_384, column, sheet.columnDimensions.map(run => ({ from: run.from, to: run.to,
    sizePt: run.hidden && !includeHidden ? 0 : run.width === undefined ? column : columnWidthPt(run.width, maximumDigitWidthPx) })));
}

export function worksheetDimensions(sheet: DeepReadonly<WorkbookSheet>, maximumDigitWidthPx: number, measuredRows: ReadonlyMap<number, number> = new Map()): { readonly rows: AxisDimensions; readonly columns: AxisDimensions } {
  const row = sheet.defaultRowsHidden ? 0 : sheet.defaultRowHeightPt;
  const rows: AxisOverride[] = [], measured = [...measuredRows].sort((a, b) => a[0] - b[0]); let index = 0;
  const appendMeasured = (before: number): void => {
    while (index < measured.length && measured[index]![0] < before) {
      const [from, sizePt] = measured[index++]!;
      if (!Number.isSafeInteger(from) || from < 0 || from >= 1_048_576 || !Number.isFinite(sizePt) || sizePt <= 0 || sizePt > 409)
        invalid("Measured worksheet row heights must lie within native bounds.");
      if (!sheet.defaultRowsHidden && !sheet.defaultRowHeightCustom) rows.push({ from, to: from, sizePt });
    }
  };
  for (const run of sheet.rowDimensions) {
    appendMeasured(run.from); let next = run.from;
    while (index < measured.length && measured[index]![0] <= run.to) {
      const [from, sizePt] = measured[index++]!;
      if (!Number.isFinite(sizePt) || sizePt <= 0 || sizePt > 409) invalid("Measured worksheet row heights must lie within native bounds.");
      const sourceSize = (run.hidden ?? sheet.defaultRowsHidden) ? 0 : run.heightPt ?? sheet.defaultRowHeightPt;
      if (next < from) rows.push({ from: next, to: from - 1, sizePt: sourceSize });
      rows.push({ from, to: from, sizePt: sourceSize === 0 || run.heightPt !== undefined || run.autoHeight === false || sheet.defaultRowHeightCustom && run.autoHeight !== true ? sourceSize : sizePt }); next = from + 1;
    }
    if (next <= run.to) rows.push({ from: next, to: run.to, sizePt: (run.hidden ?? sheet.defaultRowsHidden) ? 0 : run.heightPt ?? sheet.defaultRowHeightPt });
  }
  appendMeasured(1_048_576);
  if (index !== measured.length) invalid("A measured worksheet row is outside its source axis.");
  return {
    rows: new AxisDimensions(1_048_576, row, rows),
    columns: worksheetColumnDimensions(sheet, maximumDigitWidthPx),
  };
}
