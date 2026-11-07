import { captureWorkbookModel, copyWorkbookFormula, DocumentError, readWorkbookCellSource, resolveWorkbookCellFormat, WorkbookRangeSchema } from "@doffi/core";
import type { DeepReadonly, DocumentTransaction, WorkbookCellInput, WorkbookCellSource, WorkbookCoordinate, WorkbookModel, WorkbookRange } from "@doffi/core";

export interface WorkbookFillOptions {
  readonly mode?: "copy" | "series";
  readonly axis?: "row" | "column";
  readonly format?: "source" | "preserve";
  readonly maxCells?: number;
  readonly maxCodeUnits?: number;
  readonly signal?: { readonly aborted: boolean };
  readonly yield?: () => Promise<void>;
  readonly onProgress?: (progress: { readonly completed: number; readonly total: number }) => void;
}
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
function overlap(a: DeepReadonly<WorkbookRange>, b: DeepReadonly<WorkbookRange>): boolean {
  return a.from.row <= b.to.row && b.from.row <= a.to.row && a.from.column <= b.to.column && b.from.column <= a.to.column;
}
/** Build one revision-bound fill from immutable seeds. Overlapping targets
 * never read earlier outputs, and UI/Worker hosts own yielding and cancellation. */
export async function createWorkbookFillTransaction(input: WorkbookModel, sheetId: string, source: DeepReadonly<WorkbookRange>, target: DeepReadonly<WorkbookRange>, options: WorkbookFillOptions = {}): Promise<DocumentTransaction> {
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["mode", "axis", "format", "maxCells", "maxCodeUnits", "signal", "yield", "onProgress"].includes(key))) invalid("Invalid workbook fill options.");
  const config = { mode: "copy", format: "source", maxCells: 250_000, maxCodeUnits: 32 * 1024 * 1024, ...options };
  if (!["copy", "series"].includes(config.mode) || !["source", "preserve"].includes(config.format)
    || config.axis !== undefined && !["row", "column"].includes(config.axis)
    || ![config.maxCells, config.maxCodeUnits].every(value => Number.isSafeInteger(value) && value > 0)
    || config.signal !== undefined && (!config.signal || typeof config.signal.aborted !== "boolean")
    || config.yield !== undefined && typeof config.yield !== "function" || config.onProgress !== undefined && typeof config.onProgress !== "function") invalid("Invalid workbook fill mode or budget.");
  const range = (value: DeepReadonly<WorkbookRange>): WorkbookRange => {
    const parsed = WorkbookRangeSchema.safeParse(value);
    if (!parsed.success || parsed.data.from.row > parsed.data.to.row || parsed.data.from.column > parsed.data.to.column) invalid("Workbook fill requires ordered source rectangles."); return parsed.data;
  };
  const seeds = range(source), destination = range(target), model = captureWorkbookModel(input), snapshot = model.snapshot();
  const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId); if (!sheet) invalid("Workbook fill requires an existing worksheet.");
  const rows = seeds.to.row - seeds.from.row + 1, columns = seeds.to.column - seeds.from.column + 1;
  const total = (destination.to.row - destination.from.row + 1) * (destination.to.column - destination.from.column + 1);
  if (total > config.maxCells || rows * columns > config.maxCells) throw new DocumentError("ResourceLimitExceeded", "Workbook fill exceeds its cell budget.");
  if (sheet.merges.some(merge => overlap(merge, seeds) || overlap(merge, destination))) invalid("Range fill requires unmerged source and destination cells.");
  const guard = (): void => {
    if (config.signal?.aborted) throw new DocumentError("OperationCancelled", "Workbook fill was cancelled.");
    if (captureWorkbookModel(input) !== model) throw new DocumentError("RevisionConflict", "The workbook changed while preparing fill.");
  };
  const cache = new Map<number, { readonly at: WorkbookCoordinate; readonly source: WorkbookCellSource; readonly cssClass: string; readonly numberFormatId?: string }>();
  let units = 0;
  const account = (amount: number): void => { units += amount; if (units > config.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook fill exceeds its source-text budget."); };
  function read(row: number, column: number) {
    guard(); const key = row * 16_384 + column, known = cache.get(key); if (known) return known;
    const at = { row, column }, source = readWorkbookCellSource(model, { sheetId, ...at }), format = resolveWorkbookCellFormat(sheet!, at, model.getCell(sheetId, row, column));
    account(JSON.stringify(source).length + format.cssClass.length + (format.numberFormatId?.length ?? 0));
    const value = { at, source, ...format }; cache.set(key, value); return value;
  }
  const progress = (completed: number, total: number): void => { try { config.onProgress?.(Object.freeze({ completed, total })); } catch { /* Host reporting cannot alter source. */ } };
  let base = 0, increment = 0;
  const axis = config.axis ?? (columns === 1 ? "row" : "column");
  if (config.mode === "series") {
    if (axis === "row" ? columns !== 1 || destination.from.column !== seeds.from.column || destination.to.column !== seeds.to.column
      : rows !== 1 || destination.from.row !== seeds.from.row || destination.to.row !== seeds.to.row) invalid("Numeric series fill requires aligned one-dimensional seeds and targets.");
    const count = axis === "row" ? rows : columns;
    const values: number[] = [];
    for (let index = 0; index < count; index++) {
      const value = read(seeds.from.row + (axis === "row" ? index : 0), seeds.from.column + (axis === "column" ? index : 0)).source;
      if (value.kind !== "value" || value.value.type !== "number") invalid("Series fill requires numeric constants as seeds."); values.push(value.value.value);
      if (index % 256 === 255 && config.yield) { await config.yield(); guard(); }
    }
    base = values[0]!; increment = values.length > 1 ? values[1]! - base : 1;
    if (!Number.isFinite(increment)) invalid("The numeric series step exceeds its finite range.");
    for (let index = 0; index < values.length; index++) if (Math.abs(values[index]! - (base + index * increment)) > Number.EPSILON * 16 * Math.max(1, Math.abs(values[index]!), Math.abs(base), Math.abs(index * increment))) invalid("The fill seeds do not form an arithmetic series.");
  }
  let general = snapshot.metadata.workbook.numberFormats.find(format => "builtinId" in format ? format.builtinId === 0 : format.code.toUpperCase() === "GENERAL")?.formatId;
  let registerGeneral = false;
  const modulo = (value: number, size: number) => (value % size + size) % size;
  const cells: WorkbookCellInput[] = []; let completed = 0;
  guard();
  for (let row = destination.from.row; row <= destination.to.row; row++) for (let column = destination.from.column; column <= destination.to.column; column++) {
    guard(); const template = read(seeds.from.row + modulo(row - seeds.from.row, rows), seeds.from.column + modulo(column - seeds.from.column, columns));
    let source = template.source;
    if (config.mode === "series") {
      const value = base + ((axis === "row" ? row - seeds.from.row : column - seeds.from.column)) * increment;
      if (!Number.isFinite(value)) invalid("A filled series value exceeds its finite range."); source = { kind: "value", value: { type: "number", value } };
    } else if (source.kind === "formula") source = { kind: "formula", expression: copyWorkbookFormula(source.expression, template.at, { row, column }) };
    if (row !== template.at.row || column !== template.at.column) {
      const cell: WorkbookCellInput = { row, column, source };
      let unchangedBlank = false;
      if (config.format === "source") {
        const targetCell = model.getCell(sheetId, row, column), targetFormat = resolveWorkbookCellFormat(sheet, { row, column }, targetCell);
        if (template.numberFormatId === undefined && general === undefined) {
          const identifiers = new Set(snapshot.metadata.workbook.numberFormats.map(format => format.formatId)); let serial = 1;
          do { general = "fill-general-" + serial++; } while (identifiers.has(general)); registerGeneral = true;
        }
        cell.cssClass = template.cssClass; cell.formatId = template.numberFormatId ?? general ?? null;
        unchangedBlank = !targetCell && source.kind === "value" && source.value.type === "blank" && targetFormat.cssClass === template.cssClass
          && (targetFormat.numberFormatId ?? general) === (template.numberFormatId ?? general);
      } else unchangedBlank = source.kind === "value" && source.value.type === "blank" && !model.getCell(sheetId, row, column);
      if (!unchangedBlank) { account(JSON.stringify(cell).length); cells.push(cell); }
    }
    if (++completed % 256 === 0 || completed === total) { progress(completed, total); guard(); if (completed < total && config.yield) { await config.yield(); guard(); } }
  }
  guard();
  const operations: DocumentTransaction["operations"] = [];
  if (registerGeneral && cells.some(cell => cell.formatId === general)) {
    const definition = { op: "setWorkbookNumberFormat" as const, format: { formatId: general!, builtinId: 0 } }; account(JSON.stringify(definition).length); operations.push(definition);
  }
  if (cells.length) operations.push({ op: "setWorkbookCells", sheetId, cells });
  return { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision,
    operations };
}
