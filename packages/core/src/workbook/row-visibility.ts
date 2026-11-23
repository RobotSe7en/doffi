import type { WorkbookModel } from "../runtime.js";
import type { WorkbookCalculationResult } from "./calculation.js";
import type { WorkbookScalarContext, WorkbookScalarUnavailable } from "./scalar.js";
import { copyWorkbookScalarContext, createWorkbookScalarEvaluator } from "./scalar.js";
import { createWorkbookCachedEvaluator, createWorkbookCalculatedEvaluator } from "./calculated-evaluator.js";
import { workbookCalculationSource } from "./calculation-source.js";
import { createWorkbookRowFilter } from "./filters.js";
import { createWorkbookFilterDisplay } from "./filter-display.js";
import { currentWorkbookModel } from "./source.js";
import { checkLimit, freeze } from "../document/validation.js";
import { DocumentError } from "../errors.js";

export interface WorkbookRowVisibilityLimits { readonly maxRows: number; readonly maxSteps: number; readonly maxReadCells: number; readonly maxCodeUnits: number }
export const DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS: WorkbookRowVisibilityLimits = Object.freeze({ maxRows: 250_000, maxSteps: 5_000_000, maxReadCells: 250_000, maxCodeUnits: 8 * 1024 * 1024 });
export interface WorkbookRowVisibilityRequest extends WorkbookRowVisibilityLimits {
  readonly documentId: string; readonly revision: number; readonly sheetId: string; readonly from: number; readonly to: number;
}
export interface WorkbookRowVisibilityOptions {
  readonly calculation?: WorkbookCalculationResult; readonly context?: WorkbookScalarContext; readonly cachedValues?: boolean;
  readonly signal?: { readonly aborted: boolean }; readonly yield?: () => Promise<void>;
}
export interface WorkbookRowVisibility {
  readonly row: number; readonly hidden: boolean; readonly filtered: boolean | WorkbookScalarUnavailable;
}
export interface WorkbookRowVisibilityResult {
  readonly documentId: string; readonly revision: number; readonly sheetId: string;
  readonly rows: readonly WorkbookRowVisibility[]; readonly steps: number; readonly readCells: number; readonly codeUnits: number;
}
/** Manual dimensions and current filtering are separate sources. Consumers
 * share SUBTOTAL's first-party predicate and the exact completed batch context;
 * imported hidden rows or cached filter statistics cannot impersonate it. */
export async function readWorkbookRowVisibility(input: WorkbookModel, supplied: WorkbookRowVisibilityRequest, settings: WorkbookRowVisibilityOptions = {}): Promise<WorkbookRowVisibilityResult> {
  if (!supplied || typeof supplied !== "object" || Array.isArray(supplied) || !settings || typeof settings !== "object" || Array.isArray(settings)) throw new DocumentError("InvalidOperation", "Invalid row visibility input.");
  const request = { ...supplied }, options = { ...settings };
  const model = currentWorkbookModel(input), snapshot = model.snapshot(), sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === request?.sheetId);
  if (!request || typeof request !== "object" || Object.keys(request).some(key => !["documentId", "revision", "sheetId", "from", "to", ...Object.keys(DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS)].includes(key))
    || !sheet || !Number.isSafeInteger(request.from) || !Number.isSafeInteger(request.to) || request.from < 0 || request.to < request.from || request.to >= 1_048_576)
    throw new DocumentError("InvalidOperation", "Invalid workbook row visibility query.");
  if (request.documentId !== snapshot.manifest.documentId || request.revision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "Row visibility belongs to another workbook revision.");
  if (!options || typeof options !== "object" || Object.keys(options).some(key => !["calculation", "context", "cachedValues", "signal", "yield"].includes(key))
    || options.yield !== undefined && typeof options.yield !== "function" || options.signal !== undefined && (!options.signal || typeof options.signal.aborted !== "boolean")
    || options.cachedValues !== undefined && typeof options.cachedValues !== "boolean"
    || options.calculation !== undefined && (options.context !== undefined || options.cachedValues)) throw new DocumentError("InvalidOperation", "Invalid row visibility options.");
  const limits = { maxRows: request.maxRows, maxSteps: request.maxSteps, maxReadCells: request.maxReadCells, maxCodeUnits: request.maxCodeUnits };
  if (!Object.values(limits).every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Row visibility budgets must be positive safe integers.");
  checkLimit(request.to - request.from + 1, limits.maxRows);
  const context = options.calculation ? workbookCalculationSource(model, options.calculation).context : copyWorkbookScalarContext(options.context);
  const evaluatorLimits = { maxSteps: limits.maxSteps, maxCells: limits.maxReadCells, maxCodeUnits: limits.maxCodeUnits };
  const evaluator = options.calculation ? createWorkbookCalculatedEvaluator(model, options.calculation, evaluatorLimits)
    : options.cachedValues ? createWorkbookCachedEvaluator(model, evaluatorLimits, context) : createWorkbookScalarEvaluator(model, evaluatorLimits, context);
  let steps = 0, readCells = 0, codeUnits = 0;
  function guard(): void {
    if (options.signal?.aborted) throw new DocumentError("OperationCancelled", "Row visibility was cancelled.");
    if (currentWorkbookModel(input) !== model) throw new DocumentError("RevisionConflict", "Workbook changed while reading row visibility.");
  }
  const step = () => { checkLimit(++steps, limits.maxSteps); if (steps % 256 === 0) guard(); };
  let display: ReturnType<typeof createWorkbookFilterDisplay> | undefined;
  const filter = createWorkbookRowFilter(sheet, { ...context, dateSystem: snapshot.metadata.workbook.dateSystem, step,
    display: (at, value) => (display ??= createWorkbookFilterDisplay(model, context, units => { codeUnits += units; checkLimit(codeUnits, limits.maxCodeUnits); }, step)).text(at, value),
    read(location) {
      step(); guard();
      const result = evaluator.cell(location, { maxSteps: Math.max(1, limits.maxSteps - steps), maxCodeUnits: Math.max(0, limits.maxCodeUnits - codeUnits), ...(options.signal ? { signal: options.signal } : {}) });
      steps += result.steps; readCells += result.readCells; codeUnits += result.codeUnits;
      checkLimit(steps, limits.maxSteps); checkLimit(readCells, limits.maxReadCells); checkLimit(codeUnits, limits.maxCodeUnits);
      return result.kind === "value" ? result.value : { kind: "unavailable", reason: result.reason, ...(result.functionName ? { functionName: result.functionName } : {}) };
    },
    *stored(area) {
      for (const chunk of snapshot.manifest.chunks) {
        step(); if (chunk.sheetId !== area.sheetId || chunk.row > area.to.row || chunk.row + 255 < area.from.row || chunk.column > area.to.column || chunk.column + 31 < area.from.column) continue;
        for (const offset of [0, 128]) for (const cell of model.readWindow({ sheetId: chunk.sheetId, row: chunk.row + offset, column: chunk.column, rows: 128, columns: 32 })) {
          step(); if (cell.row >= area.from.row && cell.row <= area.to.row && cell.column >= area.from.column && cell.column <= area.to.column)
            yield { sheetId: chunk.sheetId, row: cell.row, column: cell.column };
        }
      }
    },
  });
  try {
    guard(); const rows: WorkbookRowVisibility[] = []; let dimension = 0;
    for (let row = request.from; row <= request.to; row++) {
      while (sheet.rowDimensions[dimension] && sheet.rowDimensions[dimension]!.to < row) dimension++;
      const span = sheet.rowDimensions[dimension], hidden = span && span.from <= row ? (span.hidden ?? sheet.defaultRowsHidden ?? false) || span.heightPt === 0 : sheet.defaultRowsHidden ?? false;
      step(); rows.push({ row, hidden, filtered: filter(row) });
      if (rows.length % 256 === 0 && options.yield) { await options.yield(); guard(); }
    }
    guard(); return freeze({ documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, sheetId: sheet.sheetId, rows, steps, readCells, codeUnits });
  } finally { display?.dispose(); evaluator.dispose(); }
}

/** Read the bounded union of active sheet/table filter ranges for one output.
 * The returned rows are additional to any explicitly hidden source rows. */
export async function collectWorkbookFilteredRows(input: WorkbookModel, sheetId: string, options: WorkbookRowVisibilityOptions = {}): Promise<{
  readonly rows?: ReadonlySet<number>; readonly unresolved: boolean;
}> {
  const model = currentWorkbookModel(input), { metadata, manifest } = model.snapshot();
  const sheet = metadata.workbook.sheets.find(item => item.sheetId === sheetId);
  if (!sheet) throw new DocumentError("InvalidOperation", "Workbook filters require an existing worksheet.");
  const intervals = [sheet.autoFilter, ...(sheet.tables ?? []).map(table => table.autoFilter)].flatMap(filter =>
    filter?.columns.some(column => column.criterion) && filter.range.from.row < filter.range.to.row
      ? [{ from: filter.range.from.row + 1, to: filter.range.to.row }] : []);
  if (!intervals.length) return { unresolved: false };
  intervals.sort((left, right) => left.from - right.from);
  const spans: typeof intervals = [];
  for (const interval of intervals) {
    const last = spans.at(-1);
    if (last && interval.from <= last.to + 1) last.to = Math.max(last.to, interval.to);
    else spans.push(interval);
  }
  const limits = { ...DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS };
  checkLimit(spans.reduce((sum, { from, to }) => sum + to - from + 1, 0), limits.maxRows);
  const rows = new Set<number>();
  let unresolved = false;
  for (const span of spans) {
    const result = await readWorkbookRowVisibility(input, { ...limits, documentId: manifest.documentId, revision: manifest.revision, sheetId, ...span }, options);
    for (const row of result.rows) {
      if (row.filtered === true) rows.add(row.row);
      else if (row.filtered !== false) unresolved = true;
    }
    limits.maxSteps -= result.steps;
    limits.maxReadCells -= result.readCells;
    limits.maxCodeUnits -= result.codeUnits;
    if (limits.maxSteps < 1 || limits.maxReadCells < 1 || limits.maxCodeUnits < 1)
      throw new DocumentError("ResourceLimitExceeded", "Workbook filter evaluation exceeds its cumulative budget.");
  }
  return { rows, unresolved };
}
