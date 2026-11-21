import { DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS, DocumentError, readWorkbookRowVisibility } from "@doffi/core";
import type { DeepReadonly, WorkbookCalculationResult, WorkbookModel, WorkbookScalarContext, WorkbookSheet } from "@doffi/core";

type Sheet = DeepReadonly<WorkbookSheet>;
interface Interval { from: number; to: number }
export async function filteredRows(model: WorkbookModel, sheet: Sheet, calculation: WorkbookCalculationResult | undefined,
  context: WorkbookScalarContext, checkpoint: () => Promise<void>, signal?: { readonly aborted: boolean }): Promise<{
    rows?: ReadonlySet<number>; unresolved: boolean;
  }> {
  const filters = [sheet.autoFilter, ...(sheet.tables ?? []).map(table => table.autoFilter)]
    .filter(filter => filter?.columns.some(column => column.criterion));
  const intervals = filters.map(filter => ({ from: filter!.range.from.row + 1, to: filter!.range.to.row })).filter(interval => interval.from <= interval.to);
  if (!intervals.length) return { unresolved: false };
  intervals.sort((left, right) => left.from - right.from);
  const spans: Interval[] = [];
  for (const interval of intervals) {
    const last = spans.at(-1);
    if (last && interval.from <= last.to + 1) last.to = Math.max(last.to, interval.to);
    else spans.push(interval);
  }
  const limits = { ...DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS };
  if (spans.reduce((sum, { from, to }) => sum + to - from + 1, 0) > limits.maxRows)
    throw new DocumentError("ResourceLimitExceeded", "The XLSX filter row cache exceeds its native output budget.");
  const { documentId, revision } = model.snapshot().manifest, rows = new Set<number>();
  let unresolved = false;
  for (const span of spans) {
    const result = await readWorkbookRowVisibility(model, { ...limits, documentId, revision, sheetId: sheet.sheetId, ...span },
      { ...(calculation ? { calculation } : { context }), yield: checkpoint, ...(signal ? { signal } : {}) });
    for (const row of result.rows) {
      if (row.filtered === true) rows.add(row.row);
      else if (row.filtered !== false) unresolved = true;
    }
    limits.maxSteps -= result.steps;
    limits.maxReadCells -= result.readCells;
    limits.maxCodeUnits -= result.codeUnits;
    if (limits.maxSteps < 1 || limits.maxReadCells < 1 || limits.maxCodeUnits < 1)
      throw new DocumentError("ResourceLimitExceeded", "The XLSX filter row cache exceeds its calculation budget.");
  }
  return { rows, unresolved };
}
