import { captureWorkbookModel, DocumentError } from "@doffi/core";
import type { WorkbookCell, WorkbookModel, WorkbookSnapshot, WorkbookWindow } from "@doffi/core";

export interface WorkbookWindowLimits { readonly maxCells: number; readonly maxCodeUnits: number }
export interface WorkbookWindowRequest extends WorkbookWindowLimits {
  readonly documentId: string; readonly revision: number; readonly windows: readonly WorkbookWindow[];
}
export interface WorkbookWindowResponse {
  readonly documentId: string; readonly revision: number; readonly cells: readonly WorkbookCell[];
}
/** The provider owns an already validated immutable model. A Worker adapter may
 * transport these queries, but IDs/revisions never confer access to a model. */
export interface WorkbookWindowSource {
  readonly snapshot: WorkbookSnapshot;
  /** Opt in to bounded neighboring queries while idle. This is a traffic
   * preference; the provider must still authorize and validate every read. */
  readonly allowReadAhead?: boolean;
  read(request: WorkbookWindowRequest, signal: AbortSignal): Promise<WorkbookWindowResponse>;
}
export const DEFAULT_WINDOW_LIMITS: WorkbookWindowLimits = Object.freeze({ maxCells: 5000, maxCodeUnits: 8 * 1024 * 1024 });
export function windowLimits(input: Partial<WorkbookWindowLimits> = {}): WorkbookWindowLimits {
  const result = { ...DEFAULT_WINDOW_LIMITS, ...input };
  if (Object.keys(input).some(key => !Object.hasOwn(DEFAULT_WINDOW_LIMITS, key)) || !Object.values(result).every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Workbook window limits must be supported positive safe integers.");
  return Object.freeze(result);
}
export function checkWindowIdentity(source: { readonly documentId: string; readonly revision: number }, snapshot: WorkbookSnapshot): void {
  if (source.documentId !== snapshot.manifest.documentId || source.revision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "Workbook window data belongs to another document or revision.");
}
export function checkWindowCancellation(signal: AbortSignal): void {
  if (signal.aborted) throw new DocumentError("OperationCancelled", "The workbook window read was cancelled.");
}
export function cellSourceUnits(cell: WorkbookCell): number {
  return cell.html.length + cell.text.length + (cell.value.type === "text" || cell.value.type === "error" ? cell.value.value.length : 0)
    + (cell.formula?.expression.length ?? 0) + (cell.cssClass?.length ?? 0) + (cell.formatId?.length ?? 0) + cell.sheetId.length;
}
/** Check the whole batch before reading its first cell. Coordinate expansion is
 * bounded by the viewport budget, including sparse gaps and repeated queries. */
function validateRequest(request: WorkbookWindowRequest, snapshot: WorkbookSnapshot): void {
  checkWindowIdentity(request, snapshot); windowLimits({ maxCells: request.maxCells, maxCodeUnits: request.maxCodeUnits });
  const sheets = new Set(snapshot.metadata.workbook.sheets.map(sheet => sheet.sheetId)); let total = 0;
  if (request.windows.length > request.maxCells) throw new DocumentError("ResourceLimitExceeded", "Too many workbook window queries.");
  for (const window of request.windows) {
    if (!sheets.has(window.sheetId) || ![window.row, window.column, window.rows, window.columns].every(Number.isSafeInteger)
      || window.row < 0 || window.column < 0 || window.rows <= 0 || window.columns <= 0 || window.row + window.rows > 1_048_576 || window.column + window.columns > 16_384) throw new DocumentError("InvalidOperation", "A workbook window lies outside its source worksheet.");
    total += window.rows * window.columns;
    if (total > request.maxCells) throw new DocumentError("ResourceLimitExceeded", "The workbook window batch exceeds its total cell budget.");
  }
  const seen = new Map<string, Set<number>>();
  for (const window of request.windows) {
    let cells = seen.get(window.sheetId); if (!cells) { cells = new Set(); seen.set(window.sheetId, cells); }
    for (let row = window.row; row < window.row + window.rows; row++) for (let column = window.column; column < window.column + window.columns; column++) {
      const key = row * 16_384 + column;
      if (cells.has(key)) throw new DocumentError("InvalidOperation", "Workbook window queries must not overlap.");
      cells.add(key);
    }
  }
}
/** Direct small-document adapter; the same bounded batch can run in a host-owned
 * Worker without transferring the complete package or allocating implicit cells. */
export function modelWindowSource(model: WorkbookModel): WorkbookWindowSource {
  model = captureWorkbookModel(model);
  const snapshot = model.snapshot();
  return Object.freeze({ snapshot, allowReadAhead: true, async read(input: WorkbookWindowRequest, signal: AbortSignal): Promise<WorkbookWindowResponse> {
    checkWindowCancellation(signal);
    const limits = windowLimits({ maxCells: input.maxCells, maxCodeUnits: input.maxCodeUnits });
    if (input.windows.length > limits.maxCells) throw new DocumentError("ResourceLimitExceeded", "Too many workbook window queries.");
    const request = { documentId: input.documentId, revision: input.revision, ...limits,
      windows: input.windows.map(window => ({ sheetId: window.sheetId, row: window.row, column: window.column, rows: window.rows, columns: window.columns })) };
    validateRequest(request, snapshot);
    const cells: WorkbookCell[] = []; let units = 0, visited = 0, checkpoint = performance.now(), firstYield = true;
    const total = request.windows.reduce((sum, window) => sum + window.rows * window.columns, 0);
    for (const window of request.windows) {
      const columns = Math.min(256, window.columns), rows = Math.floor(256 / columns);
      for (let row = window.row; row < window.row + window.rows; row += rows) for (let column = window.column; column < window.column + window.columns; column += columns) {
        checkWindowCancellation(signal);
        const query = { sheetId: window.sheetId, row, column, rows: Math.min(rows, window.row + window.rows - row), columns: Math.min(columns, window.column + window.columns - column) };
        for (const cell of model.readWindow(query)) {
          units += cellSourceUnits(cell);
          if (units > request.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook window content exceeds its source-text budget.");
          cells.push(cell);
        }
        visited += query.rows * query.columns;
        // Yield a task, not only a microtask, so a host/Worker cancellation
        // message can interrupt a large or highly fragmented viewport read.
        if (visited < total && (firstYield || performance.now() - checkpoint >= 8)) { firstYield = false; await new Promise<void>(resolve => setTimeout(resolve, 0)); checkpoint = performance.now(); }
      }
    }
    checkWindowCancellation(signal);
    return Object.freeze({ documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, cells: Object.freeze(cells) });
  } });
}
