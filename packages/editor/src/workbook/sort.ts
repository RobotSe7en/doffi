import { captureWorkbookModel, createWorkbookCalculatedEvaluator, DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS, DocumentError, readWorkbookRowVisibility, resolveWorkbookSortLayout, WorkbookSortSchema } from "@doffi/core";
import type { DeepReadonly, DocumentTransaction, WorkbookCalculationResult, WorkbookModel, WorkbookSort, WorkbookValue } from "@doffi/core";
import { workbookSortComparison } from "./sort-order.js";

export interface WorkbookSortOptions {
  readonly calculation: WorkbookCalculationResult; readonly locale?: string;
  readonly maxRows?: number; readonly maxReadCells?: number; readonly maxSteps?: number; readonly maxCodeUnits?: number;
  readonly signal?: { readonly aborted: boolean }; readonly yield?: () => Promise<void>;
  readonly onProgress?: (progress: { readonly completed: number; readonly total: number }) => void;
}
/** Stable, cancellable first-party merge sort. A current completed batch owns
 * all key values; raw formula caches never choose a record's destination. */
export async function createWorkbookSortTransaction(input: WorkbookModel, sheetId: string, definition: DeepReadonly<WorkbookSort>, settings: WorkbookSortOptions): Promise<DocumentTransaction> {
  if (!settings || typeof settings !== "object" || Array.isArray(settings)) throw new DocumentError("InvalidOperation", "Invalid workbook sorting options.");
  const options = { ...settings };
  if (!options || typeof options !== "object" || Object.keys(options).some(key => !["calculation", "locale", "maxRows", "maxReadCells", "maxSteps", "maxCodeUnits", "signal", "yield", "onProgress"].includes(key))
    || options.yield !== undefined && typeof options.yield !== "function" || options.onProgress !== undefined && typeof options.onProgress !== "function"
    || options.signal !== undefined && (!options.signal || typeof options.signal.aborted !== "boolean")) throw new DocumentError("InvalidOperation", "Invalid workbook sorting options.");
  const parsed = WorkbookSortSchema.safeParse(definition); if (!parsed.success) throw new DocumentError("InvalidOperation", "Invalid workbook sort definition.");
  const sort = parsed.data, range = sort.range, model = captureWorkbookModel(input), snapshot = model.snapshot(), sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId);
  const limits = { ...DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS, ...Object.fromEntries(["maxRows", "maxReadCells", "maxSteps", "maxCodeUnits"].flatMap(key => options[key as keyof WorkbookSortOptions] === undefined ? [] : [[key, options[key as keyof WorkbookSortOptions]]])) } as typeof DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS;
  if (!sheet || range.to.row < range.from.row || range.to.column < range.from.column || !Object.values(limits).every(value => Number.isSafeInteger(value) && value > 0)
    || sort.keys.some(key => key.column < range.from.column || key.column > range.to.column) || new Set(sort.keys.map(key => key.column)).size !== sort.keys.length)
    throw new DocumentError("InvalidOperation", "Sorting requires an existing worksheet, valid keys and positive budgets.");
  if ((range.to.row - range.from.row + 1) * sort.keys.length > limits.maxReadCells || range.to.row - range.from.row + 1 > limits.maxRows) throw new DocumentError("ResourceLimitExceeded", "Sorting exceeds its key/row budget.");
  const { rowsPerRecord, columnsPerCell } = resolveWorkbookSortLayout(sheet, range);
  let steps = 0, readCells = 0, codeUnits = 0;
  function guard(): void {
    if (options.signal?.aborted) throw new DocumentError("OperationCancelled", "Workbook sorting was cancelled.");
    if (captureWorkbookModel(input) !== model) throw new DocumentError("RevisionConflict", "Sorting source changed before application.");
    if (++steps > limits.maxSteps) throw new DocumentError("ResourceLimitExceeded", "Sorting exceeds its step budget.");
  }
  function progress(completed: number, total: number): void { if (options.onProgress) try { options.onProgress(Object.freeze({ completed, total })); } catch { /* Reporting does not own the command. */ } }
  let locale = options.locale ?? snapshot.metadata.locale ?? "zh-CN";
  try { locale = new Intl.Locale(locale, sort.method === "none" ? {} : { collation: sort.method === "pinYin" ? "pinyin" : "stroke" }).toString(); }
  catch { throw new DocumentError("InvalidOperation", "Invalid workbook sorting locale."); }
  const compare = workbookSortComparison(locale, sort.caseSensitive), evaluator = createWorkbookCalculatedEvaluator(model, options.calculation, { maxSteps: limits.maxSteps, maxCells: limits.maxReadCells, maxCodeUnits: limits.maxCodeUnits });
  try {
    guard();
    const visibility = await readWorkbookRowVisibility(input, { ...limits, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, sheetId, from: range.from.row, to: range.to.row },
      { calculation: options.calculation, ...(options.signal ? { signal: options.signal } : {}), ...(options.yield ? { yield: options.yield } : {}) });
    steps += visibility.steps; readCells += visibility.readCells; codeUnits += visibility.codeUnits; guard();
    type Record = { row: number; keys: DeepReadonly<WorkbookValue>[] };
    let records: Record[] = []; const positions: number[] = [];
    let nextProgress = 256;
    for (let offset = 0; offset < visibility.rows.length; offset += rowsPerRecord) {
      const visible = visibility.rows[offset]!; let shown = 0;
      for (let index = offset; index < offset + rowsPerRecord; index++) {
        guard(); const row = visibility.rows[index]!;
        if (typeof row.filtered !== "boolean") throw new DocumentError("UnsupportedFeature", "Sorting requires evaluable current filter results.");
        if (!row.hidden && !row.filtered) shown++;
      }
      if (shown && shown !== rowsPerRecord) throw new DocumentError("InvalidOperation", "Sorting cannot split a merged record between visible and hidden rows.");
      if (shown) {
        const keys: DeepReadonly<WorkbookValue>[] = [];
        for (const key of sort.keys) {
          const column = range.from.column + Math.floor((key.column - range.from.column) / columnsPerCell) * columnsPerCell;
          guard(); const result = evaluator.cell({ sheetId, row: visible.row, column }, { maxSteps: Math.max(1, limits.maxSteps - steps), maxCodeUnits: Math.max(0, limits.maxCodeUnits - codeUnits), ...(options.signal ? { signal: options.signal } : {}) });
          steps += result.steps; readCells += result.readCells; codeUnits += result.codeUnits;
          if (steps > limits.maxSteps || readCells > limits.maxReadCells || codeUnits > limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Sorting exceeds its evaluation budget.");
          if (result.kind !== "value") throw new DocumentError("UnsupportedFeature", "Sorting requires valid current key values."); keys.push(result.value);
        }
        records.push({ row: visible.row, keys }); positions.push(visible.row);
      }
      if (offset + rowsPerRecord >= nextProgress || offset + rowsPerRecord === visibility.rows.length) {
        progress(offset + rowsPerRecord, visibility.rows.length); nextProgress = offset + rowsPerRecord + 256;
        if (options.yield) await options.yield(); guard();
      }
    }
    let work = 0;
    for (let width = 1; width < records.length; width *= 2) {
      const sorted: Record[] = [];
      for (let start = 0; start < records.length; start += 2 * width) {
        let left = start, right = Math.min(records.length, start + width); const middle = right, end = Math.min(records.length, start + width * 2);
        while (left < middle || right < end) {
          guard(); let order = 0;
          if (left < middle && right < end) for (const [index, key] of sort.keys.entries()) { guard(); order = compare(records[left]!.keys[index]!, records[right]!.keys[index]!, key.descending); if (order) break; }
          sorted.push(right >= end || left < middle && order <= 0 ? records[left++]! : records[right++]!);
          if (++work % 256 === 0) { progress(work, 0); if (options.yield) await options.yield(); guard(); }
        }
      }
      records = sorted;
    }
    const rows = Array.from({ length: range.to.row - range.from.row + 1 }, (_, offset) => range.from.row + offset);
    records.forEach((record, index) => { for (let offset = 0; offset < rowsPerRecord; offset++) rows[record.row - range.from.row + offset] = positions[index]! + offset; }); guard();
    return { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations: [{ op: "reorderWorkbookRows", sheetId, range, rows }] };
  } finally { evaluator.dispose(); }
}
