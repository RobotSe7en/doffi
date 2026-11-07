import { captureWorkbookModel, createWorkbookCalculatedEvaluator, createWorkbookScalarEvaluator, DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookCalculationResult, WorkbookChart, WorkbookChartLiteral, WorkbookChartVector, WorkbookFormulaSource, WorkbookModel, WorkbookReferenceArea, WorkbookReferenceResult, WorkbookScalarEvaluator, WorkbookScalarResult, WorkbookScalarUnavailable, WorkbookSheet, WorkbookSnapshot, WorkbookValue } from "@doffi/core";
import { workbookCellFormat } from "./cell-format.js";

export interface WorkbookChartReadLimits {
  readonly maxCharts: number; readonly maxReferences: number; readonly maxPoints: number; readonly maxCodeUnits: number;
  readonly maxFormulaCells: number; readonly maxFormulaSteps: number;
}
export const DEFAULT_CHART_READ_LIMITS: WorkbookChartReadLimits = Object.freeze({ maxCharts: 128, maxReferences: 5000, maxPoints: 50_000, maxCodeUnits: 8 * 1024 * 1024,
  maxFormulaCells: 250_000, maxFormulaSteps: 5_000_000 });
export interface WorkbookChartReadRequest extends WorkbookChartReadLimits {
  readonly documentId: string; readonly revision: number; readonly chartIds: readonly string[];
}
export interface WorkbookChartCell {
  readonly row: number; readonly column: number; readonly value: DeepReadonly<WorkbookValue>;
  readonly state: "constant" | WorkbookFormulaSource["cacheState"];
  readonly issue?: Pick<WorkbookScalarUnavailable, "reason" | "functionName">;
  readonly numberFormatId?: string; readonly rowHidden: boolean; readonly columnHidden: boolean;
}
/** Areas retain source order and duplicates. Cells within each area are in row
 * order, including implicit blanks. This is source data, not a flattened plot:
 * hiding, numeric coercion and display formatting belong to the chart adapter. */
export type WorkbookChartReferenceData =
  | { readonly kind: "cells"; readonly areas: readonly { readonly range: DeepReadonly<WorkbookReferenceArea>; readonly cells: readonly WorkbookChartCell[] }[] }
  | DeepReadonly<WorkbookChartLiteral>
  | Exclude<WorkbookReferenceResult, { kind: "ranges" }>;
export type WorkbookChartReferenceSlot =
  | { readonly role: "categories" | "title" | "category-axis-title" | "value-axis-title" }
  | { readonly role: "series-name" | "series-values"; readonly seriesId: string };
export interface WorkbookChartData {
  readonly chartId: string; readonly sheetId: string;
  readonly references: readonly (WorkbookChartReferenceSlot & { readonly data: WorkbookChartReferenceData })[];
}
export interface WorkbookChartReadResponse {
  readonly documentId: string; readonly revision: number; readonly charts: readonly WorkbookChartData[];
}
export interface WorkbookChartSource {
  readonly snapshot: WorkbookSnapshot;
  read(request: WorkbookChartReadRequest, signal: AbortSignal): Promise<WorkbookChartReadResponse>;
}

function cancelled(signal: AbortSignal): void {
  if (signal.aborted) throw new DocumentError("OperationCancelled", "The workbook chart read was cancelled.");
}
function limits(input: WorkbookChartReadLimits): WorkbookChartReadLimits {
  const result = { maxCharts: input.maxCharts, maxReferences: input.maxReferences, maxPoints: input.maxPoints, maxCodeUnits: input.maxCodeUnits,
    maxFormulaCells: input.maxFormulaCells, maxFormulaSteps: input.maxFormulaSteps };
  if (!Object.values(result).every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Chart read budgets must be positive safe integers.");
  return result;
}
function hidden(runs: readonly { readonly from: number; readonly to: number; readonly hidden?: boolean | undefined }[], index: number, fallback = false): boolean {
  let low = 0, high = runs.length;
  while (low < high) { const mid = (low + high) >>> 1; if (runs[mid]!.to < index) low = mid + 1; else high = mid; }
  const run = runs[low]; return run && run.from <= index ? run.hidden ?? fallback : fallback;
}
function references(chart: DeepReadonly<WorkbookChart>, add: (slot: WorkbookChartReferenceSlot, reference: DeepReadonly<WorkbookChartVector>) => void): void {
  if (chart.categories) add({ role: "categories" }, chart.categories);
  if (chart.titleReference) add({ role: "title" }, chart.titleReference);
  if (chart.categoryAxisTitle?.reference) add({ role: "category-axis-title" }, chart.categoryAxisTitle.reference);
  if (chart.valueAxisTitle?.reference) add({ role: "value-axis-title" }, chart.valueAxisTitle.reference);
  for (const series of chart.series) {
    if (series.name?.kind === "reference") add({ role: "series-name", seriesId: series.seriesId }, series.name.reference);
    add({ role: "series-values", seriesId: series.seriesId }, series.values);
  }
}

/** Read an immutable revision through the public model boundary. Resolve and
 * admit the entire batch before its first cell read. Imported chart caches stay
 * in the snapshot; they never replace unresolved references or current cells.
 * The response is structured-cloneable for a host-owned Worker transport. */
export function modelChartSource(model: WorkbookModel, calculation?: WorkbookCalculationResult): WorkbookChartSource {
  model = captureWorkbookModel(model);
  const snapshot = model.snapshot(), workbook = snapshot.metadata.workbook;
  const charts = new Map((workbook.charts ?? []).map(chart => [chart.chartId, chart]));
  const owners = new Map((workbook.objects ?? []).filter(object => object.kind === "chart").map(object => [object.chartId, object.sheetId]));
  const sheets = new Map(workbook.sheets.map(sheet => [sheet.sheetId, sheet]));
  return Object.freeze({ snapshot, async read(input: WorkbookChartReadRequest, signal: AbortSignal): Promise<WorkbookChartReadResponse> {
    cancelled(signal);
    if (input.documentId !== snapshot.manifest.documentId || input.revision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "Chart data belongs to another document or revision.");
    const budget = limits(input);
    if (input.chartIds.length > budget.maxCharts) throw new DocumentError("ResourceLimitExceeded", "The chart batch exceeds its chart budget.");
    const ids = [...input.chartIds];
    if (new Set(ids).size !== ids.length || ids.some(id => !charts.has(id))) throw new DocumentError("InvalidOperation", "A chart batch must select unique existing charts.");
    let count = 0, points = 0, units = 0;
    const account = (amount: number): void => {
      units += amount;
      if (units > budget.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Chart source data exceeds its text budget.");
    };
    const selected = ids.map(chartId => {
      const chart = charts.get(chartId)!, sheetId = owners.get(chartId)!;
      const entries: { slot: WorkbookChartReferenceSlot; reference: DeepReadonly<WorkbookChartVector> }[] = [];
      account(chartId.length + sheetId.length);
      references(chart, (slot, reference) => {
        if (++count > budget.maxReferences) throw new DocumentError("ResourceLimitExceeded", "The chart batch exceeds its reference budget.");
        account(("expression" in reference ? reference.expression.length : 0) + ("seriesId" in slot ? slot.seriesId.length : 0));
        entries.push({ slot, reference });
      });
      return { chartId, sheetId, entries };
    });
    type ResolvedVector = WorkbookReferenceResult | DeepReadonly<WorkbookChartLiteral>;
    const pending: { chartId: string; sheetId: string; entries: { slot: WorkbookChartReferenceSlot; resolved: ResolvedVector }[] }[] = [];
    let admitted = 0;
    for (const chart of selected) {
      const entries: { slot: WorkbookChartReferenceSlot; resolved: ResolvedVector }[] = [];
      for (const { slot, reference } of chart.entries) {
        cancelled(signal);
        const resolved = "expression" in reference ? model.resolveReference(reference.expression, { sheetId: chart.sheetId, row: 0, column: 0 }) : reference;
        if (resolved.kind === "literal") {
          points += resolved.count;
          if (points > budget.maxPoints) throw new DocumentError("ResourceLimitExceeded", "Chart literals exceed their cumulative point budget.");
          account(resolved.numberFormat?.length ?? 0);
          for (let i = 0; i < resolved.points.length; i++) {
            cancelled(signal);
            const point = resolved.points[i]!;
            account((typeof point.value === "string" ? point.value.length : 0) + (point.numberFormat?.length ?? 0));
            if (i % 256 === 255 && i + 1 < resolved.points.length) await new Promise<void>(resolve => setTimeout(resolve, 0));
          }
        }
        if (resolved.kind === "ranges") for (const area of resolved.areas) {
          points += (area.to.row - area.from.row + 1) * (area.to.column - area.from.column + 1);
          if (points > budget.maxPoints) throw new DocumentError("ResourceLimitExceeded", "Chart references exceed their cumulative cell budget.");
          account(area.sheetId.length);
        }
        entries.push({ slot, resolved });
        if (++admitted < count && admitted % 32 === 0) await new Promise<void>(resolve => setTimeout(resolve, 0));
      }
      pending.push({ chartId: chart.chartId, sheetId: chart.sheetId, entries });
    }
    let visited = 0, yielded = 0, steps = 0, readCells = 0, checkpoint = performance.now();
    let evaluator: WorkbookScalarEvaluator | undefined;
    const evaluated = new Map<string, WorkbookScalarResult>();
    const currentFormula = (sheetId: string, row: number, column: number): WorkbookScalarResult => {
      const key = JSON.stringify([sheetId, row, column]), cached = evaluated.get(key);
      if (cached) return cached;
      if (steps >= budget.maxFormulaSteps) throw new DocumentError("ResourceLimitExceeded", "Chart formulas exceed their cumulative step budget.");
      const evaluatorLimits = { maxSteps: budget.maxFormulaSteps, maxCells: budget.maxFormulaCells, maxCodeUnits: budget.maxCodeUnits, maxWindowCells: 256 };
      evaluator ??= calculation === undefined ? createWorkbookScalarEvaluator(model, evaluatorLimits) : createWorkbookCalculatedEvaluator(model, calculation, evaluatorLimits);
      const result = evaluator.cell({ sheetId, row, column }, { signal, maxSteps: budget.maxFormulaSteps - steps, maxCodeUnits: budget.maxCodeUnits - units });
      account(result.codeUnits);
      steps += result.steps; readCells += result.readCells;
      if (readCells > budget.maxFormulaCells) throw new DocumentError("ResourceLimitExceeded", "Chart formulas exceed their cumulative dependency-cell budget.");
      evaluated.set(key, result); return result;
    };
    const results: WorkbookChartData[] = [];
    try {
      for (const chart of pending) {
        const refs: (WorkbookChartReferenceSlot & { data: WorkbookChartReferenceData })[] = [];
        for (const { slot, resolved } of chart.entries) {
          if (resolved.kind !== "ranges") { refs.push(Object.freeze({ ...slot, data: resolved })); continue; }
          const areas: { range: DeepReadonly<WorkbookReferenceArea>; cells: readonly WorkbookChartCell[] }[] = [];
          for (const range of resolved.areas) {
            const sheet: DeepReadonly<WorkbookSheet> = sheets.get(range.sheetId)!;
            const cells: WorkbookChartCell[] = [];
            // Wide ranges use one row per query; narrow ranges batch rows while
            // preserving row order. Each query contains at most 256 coordinates.
            const batchColumns = Math.min(256, range.to.column - range.from.column + 1), batchRows = Math.floor(256 / batchColumns);
            for (let row = range.from.row; row <= range.to.row; row += batchRows) for (let column = range.from.column; column <= range.to.column; column += batchColumns) {
              cancelled(signal);
              const columns = Math.min(batchColumns, range.to.column - column + 1), rows = Math.min(batchRows, range.to.row - row + 1);
              const stored = new Map(model.readWindow({ sheetId: range.sheetId, row, column, rows, columns }).map(cell => [cell.row * 16_384 + cell.column, cell]));
              for (let r = row; r < row + rows; r++) for (let c = column; c < column + columns; c++) {
                const cell = stored.get(r * 16_384 + c), result = cell?.formula ? currentFormula(range.sheetId, r, c) : undefined;
                const value = result?.kind === "value" ? result.value : result ? Object.freeze({ type: "blank" as const }) : cell?.value ?? Object.freeze({ type: "blank" as const });
                const issue = result?.kind === "unavailable" ? Object.freeze({ reason: result.reason, ...(result.functionName ? { functionName: result.functionName } : {}) }) : undefined;
                const numberFormatId = workbookCellFormat(sheet, { row: r, column: c }, cell).numberFormatId;
                account((value.type === "text" || value.type === "error" ? value.value.length : 0) + (numberFormatId?.length ?? 0) + (issue?.reason.length ?? 0) + (issue?.functionName?.length ?? 0));
                cells.push(Object.freeze({ row: r, column: c, value, state: result ? result.kind === "value" ? "valid" : "unavailable" : "constant", ...(issue ? { issue } : {}),
                  ...(numberFormatId === undefined ? {} : { numberFormatId }), rowHidden: hidden(sheet.rowDimensions, r, sheet.defaultRowsHidden), columnHidden: hidden(sheet.columnDimensions, c) }));
                if (performance.now() - checkpoint >= 8) { await new Promise<void>(resolve => setTimeout(resolve, 0)); checkpoint = performance.now(); cancelled(signal); }
              }
              visited += rows * columns;
              if (visited < points && visited - yielded >= 256) { yielded = visited; await new Promise<void>(resolve => setTimeout(resolve, 0)); }
            }
            areas.push(Object.freeze({ range, cells: Object.freeze(cells) }));
          }
          refs.push(Object.freeze({ ...slot, data: Object.freeze({ kind: "cells", areas: Object.freeze(areas) }) }));
        }
        results.push(Object.freeze({ chartId: chart.chartId, sheetId: chart.sheetId, references: Object.freeze(refs) }));
      }
      cancelled(signal);
      return Object.freeze({ documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, charts: Object.freeze(results) });
    } finally { evaluator?.dispose(); evaluated.clear(); }
  } });
}
