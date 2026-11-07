import { captureWorkbookModel, copyWorkbookFormula, createWorkbookCalculatedEvaluator, DocumentError, readWorkbookCellSource, replaceWorkbookCellText, resolveWorkbookCellFormat, WorkbookCoordinateSchema, WorkbookRangeSchema } from "@doffi/core";
import type { DeepReadonly, DocumentTransaction, WorkbookCalculationResult, WorkbookCellFormatInput, WorkbookCellInput, WorkbookCellSource, WorkbookCoordinate, WorkbookModel, WorkbookRange } from "@doffi/core";
import { workbookTextSplice } from "./text-projection.js";

export interface WorkbookRangeCopyOptions {
  /** Contents includes rich source and cell formats. Values and formulas keep
   * destination formats; formats leaves destination content and caches intact. */
  readonly mode?: "contents" | "values" | "formulas" | "formats";
  readonly calculation?: WorkbookCalculationResult;
  /** Only true blank constants are skipped; empty text and formulas remain. */
  readonly skipBlanks?: boolean;
  readonly maxCells?: number;
  readonly maxCodeUnits?: number;
  readonly maxSteps?: number;
  readonly signal?: { readonly aborted: boolean };
  readonly yield?: () => Promise<void>;
  readonly onProgress?: (progress: { readonly completed: number; readonly total: number }) => void;
}
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
function overlap(a: DeepReadonly<WorkbookRange>, b: DeepReadonly<WorkbookRange>): boolean {
  return a.from.row <= b.to.row && b.from.row <= a.to.row && a.from.column <= b.to.column && b.from.column <= a.to.column;
}
/** Copy one fixed source rectangle. This command handles same-sheet unmerged
 * content/format input; clipboard lifetime, links, comments and structural
 * copy are composed separately. No output is read back as a later seed. */
export async function createWorkbookRangeCopyTransaction(input: WorkbookModel, sheetId: string, range: DeepReadonly<WorkbookRange>, origin: WorkbookCoordinate, options: WorkbookRangeCopyOptions = {}): Promise<DocumentTransaction> {
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["mode", "calculation", "skipBlanks", "maxCells", "maxCodeUnits", "maxSteps", "signal", "yield", "onProgress"].includes(key))) invalid("Invalid workbook range-copy options.");
  const config = { mode: "contents", skipBlanks: false, maxCells: 250_000, maxCodeUnits: 64 * 1024 * 1024, maxSteps: 10_000_000, ...options };
  if (!["contents", "values", "formulas", "formats"].includes(config.mode) || typeof config.skipBlanks !== "boolean"
    || ![config.maxCells, config.maxCodeUnits, config.maxSteps].every(value => Number.isSafeInteger(value) && value > 0)
    || config.signal !== undefined && (!config.signal || typeof config.signal.aborted !== "boolean") || config.yield !== undefined && typeof config.yield !== "function"
    || config.onProgress !== undefined && typeof config.onProgress !== "function") invalid("Invalid workbook range-copy mode or budget.");
  const parsed = WorkbookRangeSchema.safeParse(range), coordinate = WorkbookCoordinateSchema.safeParse(origin);
  if (!parsed.success || parsed.data.from.row > parsed.data.to.row || parsed.data.from.column > parsed.data.to.column || !coordinate.success) invalid("Range copy requires an ordered source rectangle and a destination coordinate.");
  const source = parsed.data, destination = coordinate.data, model = captureWorkbookModel(input), snapshot = model.snapshot(), sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId);
  if (!sheet) invalid("Range copy requires an existing worksheet.");
  const rows = source.to.row - source.from.row + 1, columns = source.to.column - source.from.column + 1, total = rows * columns;
  if (total > config.maxCells) throw new DocumentError("ResourceLimitExceeded", "Range copy exceeds its cell budget.");
  const target = { from: destination, to: { row: destination.row + rows - 1, column: destination.column + columns - 1 } };
  if (target.to.row >= 1_048_576 || target.to.column >= 16_384) invalid("Range copy exceeds the destination worksheet bounds.");
  if (sheet.merges.some(merge => overlap(merge, source) || overlap(merge, target))) invalid("Range copy requires unmerged source and destination cells.");
  let units = 0, steps = 0, completed = 0;
  const guard = (): void => {
    if (config.signal?.aborted) throw new DocumentError("OperationCancelled", "Workbook range copy was cancelled.");
    if (captureWorkbookModel(input) !== model) throw new DocumentError("RevisionConflict", "Workbook range-copy source changed before completion.");
  };
  const account = (amount: number): void => { units += amount; if (units > config.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Range copy exceeds its text budget."); };
  const evaluator = config.calculation === undefined ? undefined : createWorkbookCalculatedEvaluator(model, config.calculation, { maxSteps: config.maxSteps, maxCells: config.maxCells, maxCodeUnits: config.maxCodeUnits });
  const cells: WorkbookCellInput[] = [], formats: WorkbookCellFormatInput[] = [];
  let general = snapshot.metadata.workbook.numberFormats.find(format => "builtinId" in format ? format.builtinId === 0 : format.code.toUpperCase() === "GENERAL")?.formatId, registerGeneral = false;
  try {
    guard();
    for (let row = 0; row < rows; row++) for (let column = 0; column < columns; column++) {
      guard(); const from = { row: source.from.row + row, column: source.from.column + column }, to = { row: destination.row + row, column: destination.column + column };
      const raw = model.getCell(sheetId, from.row, from.column), same = from.row === to.row && from.column === to.column;
      if (!(config.skipBlanks && (!raw || raw.value.type === "blank" && !raw.formula)) && (!same || config.mode === "values" && raw?.formula)) {
        const format = resolveWorkbookCellFormat(sheet, from, raw), previous = model.getCell(sheetId, to.row, to.column);
        if (config.mode === "formats" || config.mode === "contents") {
          if (format.numberFormatId === undefined && general === undefined) {
            const identifiers = new Set(snapshot.metadata.workbook.numberFormats.map(format => format.formatId)); let serial = 1;
            do { general = "copy-general-" + serial++; } while (identifiers.has(general)); registerGeneral = true;
          }
        }
        if (config.mode === "formats") {
          const targetFormat = resolveWorkbookCellFormat(sheet, to, previous);
          if (previous || targetFormat.cssClass !== format.cssClass || (targetFormat.numberFormatId ?? general) !== (format.numberFormatId ?? general)) {
            const cell = { ...to, cssClass: format.cssClass, formatId: format.numberFormatId ?? general ?? null }; account(JSON.stringify(cell).length); formats.push(cell);
          }
        } else {
          let content: WorkbookCellSource;
          if (config.mode === "values" && raw?.formula) {
            if (!evaluator) invalid("Pasting formula values requires a completed calculation batch for this source.");
            if (steps >= config.maxSteps) throw new DocumentError("ResourceLimitExceeded", "Range copy exceeds its evaluation budget.");
            const value = evaluator.cell({ sheetId, ...from }, { maxSteps: config.maxSteps - steps, maxCodeUnits: Math.max(0, config.maxCodeUnits - units), ...(config.signal ? { signal: config.signal } : {}) });
            steps += value.steps; account(value.codeUnits);
            if (value.kind !== "value") throw new DocumentError("UnsupportedFeature", "Pasting formula values requires valid current results.");
            content = { kind: "value", value: value.value };
          } else if ((config.mode === "values" || config.mode === "formulas") && !raw?.formula) content = { kind: "value", value: raw?.value ?? { type: "blank" } };
          else content = readWorkbookCellSource(model, { sheetId, ...from });
          account(JSON.stringify(content).length);
          if (content.kind === "formula") content = { kind: "formula", expression: copyWorkbookFormula(content.expression, from, to) };
          let unchanged = false;
          if (config.mode !== "contents" && content.kind === "value" && content.value.type === "text" && previous?.value.type === "text" && !previous.formula) {
            const edit = workbookTextSplice(previous.text, content.value.value);
            if (edit) content = replaceWorkbookCellText(model, { sheetId, ...to }, [edit]); else unchanged = true;
          }
          const cell: WorkbookCellInput = { ...to, source: content };
          if (config.mode === "contents") {
            cell.cssClass = format.cssClass; cell.formatId = format.numberFormatId ?? general ?? null;
            const targetFormat = resolveWorkbookCellFormat(sheet, to, previous);
            unchanged = !previous && content.kind === "value" && content.value.type === "blank" && targetFormat.cssClass === format.cssClass
              && (targetFormat.numberFormatId ?? general) === (format.numberFormatId ?? general);
          } else if (!previous && content.kind === "value" && content.value.type === "blank") unchanged = true;
          if (!unchanged) { account(JSON.stringify(cell).length); cells.push(cell); }
        }
      }
      if (++completed % 256 === 0 || completed === total) {
        try { config.onProgress?.(Object.freeze({ completed, total })); } catch { /* Host progress cannot publish source. */ }
        guard(); if (completed < total && config.yield) { await config.yield(); guard(); }
      }
    }
    guard(); const operations: DocumentTransaction["operations"] = [];
    if (registerGeneral && (cells.some(cell => cell.formatId === general) || formats.some(cell => cell.formatId === general))) {
      const operation = { op: "setWorkbookNumberFormat" as const, format: { formatId: general!, builtinId: 0 } }; account(JSON.stringify(operation).length); operations.push(operation);
    }
    if (cells.length) operations.push({ op: "setWorkbookCells", sheetId, cells });
    if (formats.length) operations.push({ op: "setWorkbookCellFormats", sheetId, cells: formats });
    return { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations };
  } finally { evaluator?.dispose(); }
}
