import { DocumentError } from "@doffi/core";
import type { DeepReadonly, PageSetup, WorkbookRange, WorkbookSheet } from "@doffi/core";
import type { AxisDimensions } from "./dimensions.js";
import type { WorkbookPane, WorkbookViewportLayout } from "./viewport.js";

interface Interval { readonly from: number; readonly to: number }
interface Slice {
  readonly parts: readonly Interval[];
  readonly body: Interval;
  readonly extent: number;
  readonly clipped: boolean;
}
export interface WorkbookPrintTile {
  readonly rows: Slice;
  readonly columns: Slice;
  readonly areaIndex: number;
}
export type PrintDimensions = { readonly rows: AxisDimensions; readonly columns: AxisDimensions };
type PrintSettings = DeepReadonly<WorkbookSheet>["print"];
export const PRINT_HEADING_WIDTH = 28;
export const PRINT_HEADING_HEIGHT = 15;

export function workbookPaper(print: PrintSettings): PageSetup {
  const { widthPt = 612, heightPt = 792 } = print?.paper ?? {};
  const portrait = print?.orientation === "portrait", landscape = print?.orientation === "landscape";
  const width = landscape ? Math.max(widthPt, heightPt) : portrait ? Math.min(widthPt, heightPt) : widthPt;
  const height = landscape ? Math.min(widthPt, heightPt) : portrait ? Math.max(widthPt, heightPt) : heightPt;
  const { left = 54, right = 54, top = 54, bottom = 54 } = print?.marginsPt ?? {};
  if (left + right >= width || top + bottom >= height)
    throw new DocumentError("InvalidOperation", "Workbook print margins leave no usable space on the paper.");
  return { unit: "pt", width, height, margins: { left, right, top, bottom } };
}

export function workbookPrintLimit(detail: string): never {
  throw new DocumentError("ResourceLimitExceeded", detail + " Select a smaller workbook print area.");
}

function parts(body: Interval, titles?: Interval): readonly Interval[] {
  if (!titles) return [body];
  if (body.to + 1 < titles.from || titles.to + 1 < body.from) return [titles, body];
  return [{ from: Math.min(body.from, titles.from), to: Math.max(body.to, titles.to) }];
}
const extent = (axis: AxisDimensions, interval: Interval): number => axis.position(interval.to + 1) - axis.position(interval.from);

/** Split at whole source rows/columns. Prefix sums let us jump across long
 * hidden runs and reject excessive page counts without expanding the grid. */
function splitAxis(axis: AxisDimensions, bounds: Interval, titles: Interval | undefined, size: number, breaks: readonly number[], max: number): Slice[] | undefined {
  if (size <= 0) return undefined;
  const pages: Slice[] = [];
  if (titles && extent(axis, titles) > size + 1e-7) return undefined;
  let from = bounds.from, breakIndex = 0;
  while (from <= bounds.to) {
    const visible = axis.size(from) > 0 ? from : axis.neighbor(from, 1);
    if (visible === undefined || visible > bounds.to) break;
    from = visible;
    while (breakIndex < breaks.length && breaks[breakIndex]! <= from) breakIndex++;
    const last = Math.min(bounds.to, (breaks[breakIndex] ?? axis.count) - 1);
    const length = (to: number): number => parts({ from, to }, titles).reduce((sum, range) => sum + extent(axis, range), 0);
    let low = from, high = last + 1;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (length(middle) <= size + 1e-7) low = middle + 1;
      else high = middle;
    }
    const to = Math.max(from, low - 1), ranges = parts({ from, to }, titles);
    const clipped = length(to) > size + 1e-7;
    // An oversized single cell can be clipped. Repeated titles must leave
    // room for at least part of that cell on the same page.
    if (clipped && titles && (from < titles.from || from > titles.to) && extent(axis, titles) >= size - 1e-7) return undefined;
    pages.push({ parts: ranges, body: { from, to }, extent: Math.min(size, length(to)), clipped });
    if (pages.length > max) return undefined;
    from = to + 1;
  }
  if (!pages.length) pages.push({ parts: titles ? [titles] : [], body: bounds, extent: titles ? extent(axis, titles) : 0, clipped: false });
  return pages;
}

export function planWorkbookPrintSheet(dimensions: PrintDimensions, areas: readonly DeepReadonly<WorkbookRange>[], print: PrintSettings,
  setup: PageSetup, maxPages: number): { readonly scale: number; readonly tiles: readonly WorkbookPrintTile[] } {
  const width = setup.width - setup.margins.left - setup.margins.right, height = setup.height - setup.margins.top - setup.margins.bottom;
  const headings = print?.headings;
  const fit = print?.scale.kind === "fit" ? print.scale : undefined;
  // Excel applies manual breaks across the entirety of the orthogonal axis.
  const rowBreaks = fit ? [] : (print?.rowBreaks ?? []).map(item => item.before);
  const columnBreaks = fit ? [] : (print?.columnBreaks ?? []).map(item => item.before);
  const split = (range: DeepReadonly<WorkbookRange>, scale: number, maxRows: number, maxColumns: number) => {
    const rows = splitAxis(dimensions.rows, { from: range.from.row, to: range.to.row }, print?.repeatRows,
      height / scale - (headings ? PRINT_HEADING_HEIGHT : 0), rowBreaks, maxRows);
    if (!rows) return;
    const columns = splitAxis(dimensions.columns, { from: range.from.column, to: range.to.column }, print?.repeatColumns,
      width / scale - (headings ? PRINT_HEADING_WIDTH : 0), columnBreaks, maxColumns);
    return columns ? { rows, columns } : undefined;
  };
  let scale = print?.scale.kind === "percent" ? print.scale.percent / 100 : 1;
  if (fit && (fit.width || fit.height)) {
    // Choose the largest whole percent that satisfies all selected areas. A
    // zero target leaves its axis unconstrained. Test only the requested axes.
    const fits = (percent: number) => areas.every(range =>
      (!fit.height || splitAxis(dimensions.rows, { from: range.from.row, to: range.to.row }, print?.repeatRows,
        height * 100 / percent - (headings ? PRINT_HEADING_HEIGHT : 0), [], fit.height)?.every(slice => !slice.clipped))
      && (!fit.width || splitAxis(dimensions.columns, { from: range.from.column, to: range.to.column }, print?.repeatColumns,
        width * 100 / percent - (headings ? PRINT_HEADING_WIDTH : 0), [], fit.width)?.every(slice => !slice.clipped)));
    let low = 10, high = 100;
    while (low < high) {
      const middle = Math.ceil((low + high) / 2);
      if (fits(middle)) low = middle;
      else high = middle - 1;
    }
    scale = low / 100;
  }
  const tiles: WorkbookPrintTile[] = [];
  for (const [areaIndex, range] of areas.entries()) {
    const remaining = maxPages - tiles.length;
    const grid = split(range, scale, remaining, remaining);
    if (!grid || grid.rows.length * grid.columns.length > remaining) workbookPrintLimit("Workbook print layout exceeds its page budget or printable space.");
    const append = (rows: Slice, columns: Slice): void => { tiles.push({ rows, columns, areaIndex }); };
    if (print?.pageOrder === "overThenDown") {
      for (const rows of grid.rows) for (const columns of grid.columns) append(rows, columns);
    } else {
      for (const columns of grid.columns) for (const rows of grid.rows) append(rows, columns);
    }
  }
  return { scale, tiles };
}

/** Each repeated-title strip is a source window. The paint pipeline uses the
 * same pane projection as the grid, independent of saved freeze/scroll state. */
export function workbookPrintViewport(dimensions: PrintDimensions, tile: WorkbookPrintTile, maxCells: number): WorkbookViewportLayout {
  const slices = (axis: AxisDimensions, slice: Slice) => {
    let start = 0;
    return slice.parts.flatMap(part => {
      const origin = axis.position(part.from), size = Math.min(slice.extent - start, extent(axis, part));
      if (size <= 0) return [];
      const result = { start, origin, size, cells: axis.visible(origin, origin + size, maxCells) };
      start += size;
      return [result];
    });
  };
  const rows = slices(dimensions.rows, tile.rows), columns = slices(dimensions.columns, tile.columns);
  const panes: WorkbookPane[] = [];
  let cellCount = 0;
  for (const [r, row] of rows.entries()) for (const [c, column] of columns.entries()) {
    cellCount += row.cells.length * column.cells.length;
    if (cellCount > maxCells) workbookPrintLimit("Workbook print page exceeds its cell budget.");
    panes.push({
      kind: rows.length === 2 && r === 0 ? columns.length === 2 && c === 0 ? "corner" : "frozenRows"
        : columns.length === 2 && c === 0 ? "frozenColumns" : "body",
      inlineStartPt: column.start, blockStartPt: row.start, widthPt: column.size, heightPt: row.size,
      sourceTopPt: row.origin, sourceLeftPt: column.origin, rows: row.cells, columns: column.cells,
    });
  }
  return { widthPt: tile.columns.extent, heightPt: tile.rows.extent, panes, cellCount,
    scrollLeftPt: 0, scrollTopPt: 0, maxScrollLeftPt: 0, maxScrollTopPt: 0 };
}
