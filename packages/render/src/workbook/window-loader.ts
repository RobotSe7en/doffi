import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookCell, WorkbookCoordinate, WorkbookSheet, WorkbookSnapshot } from "@doffi/core";
import type { AxisDimensions } from "./dimensions.js";
import { projectWorkbookCells, projectionWindows, WorkbookMergeIndex } from "./projection.js";
import type { WorkbookGridCell } from "./projection.js";
import { workbookViewport } from "./viewport.js";
import type { WorkbookViewport, WorkbookViewportLayout } from "./viewport.js";
import { cellSourceUnits, checkWindowIdentity, windowLimits } from "./window-source.js";
import type { WorkbookWindowLimits, WorkbookWindowSource } from "./window-source.js";
import { workbookAppearanceCoordinates } from "./appearance-coordinates.js";
import { workbookRowSourceCoordinates, workbookWindowBatches } from "./row-sources.js";
import { workbookSkewColumns } from "./skew-sources.js";
import type { WorkbookSkewColumns } from "./skew-sources.js";

export interface WorkbookWindowFrame {
  readonly sheetId: string;
  readonly documentId: string; readonly revision: number;
  readonly layout: WorkbookViewportLayout;
  readonly projections: readonly WorkbookGridCell[];
  readonly sourceCoordinates: readonly Readonly<WorkbookCoordinate>[];
  readonly storedCellCount: number;
  /** These displayed rows have been read across every actual source chunk. */
  readonly completeRows: readonly number[];
  readonly skewColumns: readonly WorkbookSkewColumns[];
  getCell(row: number, column: number): WorkbookCell | undefined;
}
export interface WorkbookViewportReadLimits extends WorkbookWindowLimits { readonly maxSourceCells: number; readonly maxRowReadCells: number }

/** At most one accepted viewport is retained. A complete frame replaces the old
 * frame atomically; cancellation, errors and late responses cannot splice old
 * data into new geometry. Every loader is bound to one immutable source. */
export class WorkbookWindowLoader {
  private controller: AbortController | undefined;
  private accepted: WorkbookWindowFrame | undefined;
  private readonly limits: WorkbookWindowLimits;
  private readonly maxSourceCells: number;
  private readonly maxRowReadCells: number;
  private readonly snapshot: WorkbookSnapshot;
  private readonly sheets: ReadonlyMap<string, { readonly sheet: DeepReadonly<WorkbookSheet>; readonly merges: WorkbookMergeIndex }>;
  private disposed = false;
  constructor(private readonly source: WorkbookWindowSource, limits: Partial<WorkbookViewportReadLimits> = {}, private readonly includeRowSources = false, private readonly skewCoefficient = 0) {
    const { maxSourceCells = 25_000, maxRowReadCells = 250_000, ...window } = limits;
    if (![maxSourceCells, maxRowReadCells].every(value => Number.isSafeInteger(value) && value > 0) || typeof includeRowSources !== "boolean" || !Number.isFinite(skewCoefficient) || skewCoefficient < 0 || skewCoefficient > 58) throw new DocumentError("InvalidOperation", "Worksheet source-cell limits must be positive safe integers with a bounded skew coefficient.");
    this.maxSourceCells = maxSourceCells; this.maxRowReadCells = maxRowReadCells; this.limits = windowLimits(window); this.snapshot = source.snapshot;
    this.sheets = new Map(this.snapshot.metadata.workbook.sheets.map(sheet => [sheet.sheetId, { sheet, merges: new WorkbookMergeIndex(sheet.merges) }]));
  }
  get frame(): WorkbookWindowFrame | undefined { return this.accepted; }
  async load(sheetId: string, dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, viewport: WorkbookViewport, printLayout?: WorkbookViewportLayout): Promise<boolean> {
    if (this.disposed) throw new DocumentError("InvalidOperation", "This workbook window loader has been disposed.");
    this.controller?.abort(); const controller = new AbortController(); this.controller = controller;
    const entry = this.sheets.get(sheetId);
    if (!entry) throw new DocumentError("InvalidOperation", "The requested worksheet is not part of this workbook source.");
    const layout = printLayout ?? workbookViewport(dimensions, entry.sheet.frozen, viewport, this.limits.maxCells);
    const projections = projectWorkbookCells(dimensions, layout, entry.merges, this.limits.maxCells);
    const skewColumns = workbookSkewColumns(this.snapshot, sheetId, layout, dimensions, this.skewCoefficient, this.maxSourceCells);
    const appearance = workbookAppearanceCoordinates(dimensions, layout, projections, entry.merges, this.maxSourceCells, new Map(skewColumns.map(item => [item.pane, item.columns])));
    const rowSources = this.includeRowSources ? workbookRowSourceCoordinates(this.snapshot, sheetId, layout, appearance, this.maxRowReadCells, this.maxSourceCells) : undefined;
    const batches = workbookWindowBatches(projectionWindows(sheetId, rowSources?.coordinates ?? appearance), this.maxSourceCells);
    const { documentId, revision } = this.snapshot.manifest;
    try {
      const cells = new Map<number, WorkbookCell>(); let units = 0;
      for (const [index, windows] of batches.entries()) {
        if (controller.signal.aborted || this.disposed || this.controller !== controller) return false;
        const response = await this.source.read(Object.freeze({ documentId, revision, windows, maxCells: this.maxSourceCells, maxCodeUnits: this.limits.maxCodeUnits }), controller.signal);
        if (controller.signal.aborted || this.disposed || this.controller !== controller) return false;
        checkWindowIdentity(response, this.snapshot);
        if (response.cells.length + cells.size > this.maxSourceCells) throw new DocumentError("ResourceLimitExceeded", "Workbook window response exceeds its retained source-cell budget.");
        const expected = new Set<number>();
        for (const window of windows) for (let row = window.row; row < window.row + window.rows; row++) for (let column = window.column; column < window.column + window.columns; column++) expected.add(row * 16_384 + column);
        for (const cell of response.cells) {
          const key = cell.row * 16_384 + cell.column;
          if (cell.sheetId !== sheetId || !Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.column) || cell.row < 0 || cell.row >= 1_048_576 || cell.column < 0 || cell.column >= 16_384 || !expected.has(key) || cells.has(key)) throw new DocumentError("InvalidDocument", "Workbook window response contains duplicate or unrequested coordinates.");
          units += cellSourceUnits(cell);
          if (units > this.limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook window response exceeds its cumulative source-text budget.");
          cells.set(key, Object.freeze({ ...cell, value: Object.freeze({ ...cell.value }), ...(cell.formula ? { formula: Object.freeze({ ...cell.formula }) } : {}) }));
        }
        if (index + 1 < batches.length) await new Promise<void>(resolve => setTimeout(resolve, 0));
      }
      const retained = new Map((rowSources?.required ?? appearance).map(point => [point.row * 16_384 + point.column, point]));
      for (const [key, cell] of cells) if (!retained.has(key)) {
        if (retained.size >= this.maxSourceCells) throw new DocumentError("ResourceLimitExceeded", "Retained worksheet style sources exceed their cell budget.");
        retained.set(key, Object.freeze({ row: cell.row, column: cell.column }));
      }
      const sourceCoordinates = Object.freeze([...retained.values()].sort((a, b) => a.row - b.row || a.column - b.column));
      const completeRows = Object.freeze(this.includeRowSources ? [...new Set(layout.panes.flatMap(pane => pane.rows.map(row => row.index)))].sort((a, b) => a - b) : []);
      for (const pane of layout.panes) {
        for (const cell of [...pane.rows, ...pane.columns]) Object.freeze(cell);
        Object.freeze(pane.rows); Object.freeze(pane.columns); Object.freeze(pane);
      }
      Object.freeze(layout.panes); Object.freeze(layout);
      for (const cell of projections) { Object.freeze(cell.range.from); Object.freeze(cell.range.to); Object.freeze(cell.range); Object.freeze(cell.rect); Object.freeze(cell.clip); Object.freeze(cell); }
      this.accepted = Object.freeze({ sheetId, documentId, revision, layout, projections: Object.freeze(projections), sourceCoordinates, completeRows, skewColumns, storedCellCount: cells.size,
        getCell: (row: number, column: number) => Number.isSafeInteger(row) && Number.isSafeInteger(column) && row >= 0 && row < 1_048_576 && column >= 0 && column < 16_384 ? cells.get(row * 16_384 + column) : undefined });
      return true;
    } catch (error) {
      if (controller.signal.aborted || this.disposed || this.controller !== controller) return false;
      throw error;
    }
  }
  dispose(): void { this.disposed = true; this.controller?.abort(); this.controller = undefined; this.accepted = undefined; }
}
