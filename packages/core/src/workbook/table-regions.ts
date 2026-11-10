import { DocumentError } from "../errors.js";
import type { DeepReadonly } from "../runtime.js";
import type { WorkbookCoordinate, WorkbookRange, WorkbookTable, WorkbookTableStyle } from "../contracts/workbook.js";
export type WorkbookTableRegion = WorkbookTableStyle["elements"][number]["region"];
import type { AxisDimensions } from "./axis-dimensions.js";

/** ECMA-376 Part 1 §18.8.41, low to high. Source element/CSS order does
 * not determine regional precedence. This order is different from Word. */
export const workbookTableRegionOrder: readonly WorkbookTableRegion[] = Object.freeze([
  "wholeTable", "firstColumnStripe", "secondColumnStripe", "firstRowStripe", "secondRowStripe", "lastColumn", "firstColumn",
  "headerRow", "totalRow", "firstHeaderCell", "lastHeaderCell", "firstTotalCell", "lastTotalCell",
]);
export interface WorkbookTableRegionElement<Style> {
  readonly region: WorkbookTableRegion;
  readonly size: number;
  readonly style?: Style | undefined;
}
export interface WorkbookTableRegionMatch<Style> {
  readonly region: WorkbookTableRegion;
  readonly range: DeepReadonly<WorkbookRange>;
  readonly style: Style;
}
type Geometry = Pick<DeepReadonly<WorkbookTable>, "range" | "header" | "totals" | "rowStripes" | "columnStripes" | "showFirstColumn" | "showLastColumn">;
type Range = DeepReadonly<WorkbookRange>;
function range(row: number, column: number, endRow: number, endColumn: number): Range {
  return Object.freeze({ from: Object.freeze({ row, column }), to: Object.freeze({ row: endRow, column: endColumn }) });
}

/** Compile one validated table's structural regions without enumerating rows,
 * columns or cells. Coordinates here are source coordinates: viewport clipping
 * and RTL never restart a stripe or reverse the table's logical columns.
 * Visibility/filter-dependent banding is composed separately from this layer. */
export class WorkbookTableRegions<Style> {
  private readonly table: Geometry;
  private readonly elements: readonly WorkbookTableRegionElement<Style>[];
  private readonly sizes: ReadonlyMap<WorkbookTableRegion, number>;
  constructor(table: Geometry, elements: readonly WorkbookTableRegionElement<Style>[], private readonly visibleRows: Pick<AxisDimensions, "position" | "size" | "hit"> | undefined = undefined) {
    const { from, to } = table.range;
    if (![from.row, to.row, from.column, to.column].every(Number.isSafeInteger) || from.row < 0 || from.column < 0
      || to.row >= 1_048_576 || to.column >= 16_384 || from.row > to.row || from.column > to.column
      || to.row - from.row + 1 < Number(table.header) + Number(table.totals)) throw new DocumentError("InvalidDocument", "A table style has invalid source geometry.");
    this.table = Object.freeze({ ...table, range: range(from.row, from.column, to.row, to.column) });
    const byRegion = new Map<WorkbookTableRegion, WorkbookTableRegionElement<Style>>();
    for (const element of elements) {
      if (!workbookTableRegionOrder.includes(element.region) || byRegion.has(element.region) || !Number.isSafeInteger(element.size) || element.size < 1 || element.size > 9
        || !element.region.endsWith("Stripe") && element.size !== 1)
        throw new DocumentError("InvalidDocument", "A table style has an invalid or repeated regional definition.");
      byRegion.set(element.region, Object.freeze({ ...element }));
    }
    this.elements = Object.freeze(workbookTableRegionOrder.flatMap(region => byRegion.has(region) ? [byRegion.get(region)!] : []));
    this.sizes = new Map(this.elements.map(element => [element.region, element.size]));
  }
  at(coordinate: WorkbookCoordinate): readonly WorkbookTableRegionMatch<Style>[] {
    const { row, column } = coordinate, t = this.table, { from, to } = t.range;
    if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || row >= 1_048_576 || column < 0 || column >= 16_384)
      throw new DocumentError("InvalidOperation", "Table region resolution requires a worksheet coordinate.");
    if (row < from.row || row > to.row || column < from.column || column > to.column) return Object.freeze([]);
    const header = t.header && row === from.row, total = t.totals && row === to.row;
    const dataFrom = from.row + Number(t.header), dataTo = to.row - Number(t.totals);
    const band = (axis: "row" | "column", second: boolean): Range | undefined => {
      if (axis === "row" && (!t.rowStripes || header || total) || axis === "column" && !t.columnStripes) return undefined;
      const firstName = axis === "row" ? "firstRowStripe" : "firstColumnStripe", secondName = axis === "row" ? "secondRowStripe" : "secondColumnStripe";
      const first = this.sizes.get(firstName) ?? 1, next = this.sizes.get(secondName) ?? 1, period = first + next;
      const start = axis === "row" ? dataFrom : from.column, end = axis === "row" ? dataTo : to.column;
      const value = axis === "row" ? row : column, visible = axis === "row" ? this.visibleRows : undefined;
      if (visible && !visible.size(row)) return undefined;
      const rankStart = visible ? visible.position(start) : start, rankEnd = visible ? visible.position(end + 1) - 1 : end;
      const offset = (visible ? visible.position(value) : value) - rankStart, position = offset % period;
      if ((position >= first) !== second) return undefined;
      const rankFrom = rankStart + Math.floor(offset / period) * period + (second ? first : 0);
      const rankTo = Math.min(rankEnd, rankFrom + (second ? next : first) - 1);
      const bandFrom = visible ? visible.hit(rankFrom + .5)! : rankFrom;
      const bandTo = visible ? visible.hit(rankTo + .5)! : rankTo;
      return axis === "row" ? range(bandFrom, from.column, bandTo, to.column) : range(from.row, bandFrom, to.row, bandTo);
    };
    const regionRange = (region: WorkbookTableRegion): Range | undefined => {
      switch (region) {
        case "wholeTable": return t.range;
        case "firstColumnStripe": return band("column", false);
        case "secondColumnStripe": return band("column", true);
        case "firstRowStripe": return band("row", false);
        case "secondRowStripe": return band("row", true);
        case "firstColumn": return t.showFirstColumn && column === from.column ? range(from.row, column, to.row, column) : undefined;
        case "lastColumn": return t.showLastColumn && column === to.column ? range(from.row, column, to.row, column) : undefined;
        case "headerRow": return header ? range(row, from.column, row, to.column) : undefined;
        case "totalRow": return total ? range(row, from.column, row, to.column) : undefined;
        case "firstHeaderCell": return header && column === from.column ? range(row, column, row, column) : undefined;
        case "lastHeaderCell": return header && column === to.column ? range(row, column, row, column) : undefined;
        case "firstTotalCell": return total && column === from.column ? range(row, column, row, column) : undefined;
        case "lastTotalCell": return total && column === to.column ? range(row, column, row, column) : undefined;
      }
    };
    return Object.freeze(this.elements.flatMap(element => {
      if (element.style === undefined) return [];
      const area = regionRange(element.region);
      return area ? [Object.freeze({ region: element.region, range: area, style: element.style })] : [];
    }));
  }
}

/** Regional outer and inner border slots map to logical cell sides. Undefined
 * leaves a lower layer intact; an explicit `none` paint can clear that layer. */
export function workbookTableRegionBorders<Border>(area: Range, coordinate: WorkbookCoordinate,
  borders: Readonly<Partial<Record<"left" | "right" | "top" | "bottom" | "vertical" | "horizontal", Border>>>
): Readonly<Partial<Record<"before" | "after" | "top" | "bottom", Border>>> {
  const { row, column } = coordinate;
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < area.from.row || row > area.to.row || column < area.from.column || column > area.to.column)
    throw new DocumentError("InvalidOperation", "Table borders require a coordinate inside their region.");
  const before = column === area.from.column ? borders.left : borders.vertical, after = column === area.to.column ? borders.right : borders.vertical;
  const top = row === area.from.row ? borders.top : borders.horizontal, bottom = row === area.to.row ? borders.bottom : borders.horizontal;
  return Object.freeze({ ...(before === undefined ? {} : { before }), ...(after === undefined ? {} : { after }),
    ...(top === undefined ? {} : { top }), ...(bottom === undefined ? {} : { bottom }) });
}
