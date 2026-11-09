import { captureWorkbookModel, copyWorkbookFormula, createWorkbookCalculatedEvaluator, DocumentError, readWorkbookCellSource, replaceWorkbookCellText, resolveWorkbookCellFormat, WorkbookCoordinateSchema, WorkbookRangeSchema } from "@doffi/core";
import type { DeepReadonly, DocumentTransaction, WorkbookCalculationResult, WorkbookCellFormatInput, WorkbookCellInput, WorkbookCellSource, WorkbookCoordinate, WorkbookModel, WorkbookRange } from "@doffi/core";
import { workbookTextSplice } from "./text-projection.js";
import { createWorkbookHyperlinkCopy } from "./copy-hyperlinks.js";
import { createWorkbookNameCopyPlan } from "./copy-names.js";
import { copyWorkbookRules, createWorkbookCopyFootprint } from "./copy-rules.js";
import { createWorkbookTableCopyPlan } from "./copy-tables.js";

export interface WorkbookRangeCopyOptions {
  /** Defaults to the source worksheet; both belong to the same workbook. */
  readonly targetSheetId?: string;
  /** Contents includes rich source, formats, hyperlinks, drawings, comments and both rule categories.
   * Values/formulas keep destination formats and rules. Formats copies conditional
   * formats while retaining destination validation, content and formula caches. */
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
function contains(a: DeepReadonly<WorkbookRange>, b: DeepReadonly<WorkbookRange>): boolean {
  return a.from.row <= b.from.row && a.from.column <= b.from.column && a.to.row >= b.to.row && a.to.column >= b.to.column;
}
/** Copy one fixed source rectangle, including complete merged regions for
 * content/format copies; values/formulas retain target geometry. Clipboard lifetime and structural
 * copy are composed separately. No output is read back as a later seed. */
export async function createWorkbookRangeCopyTransaction(input: WorkbookModel, sheetId: string, range: DeepReadonly<WorkbookRange>, origin: WorkbookCoordinate, options: WorkbookRangeCopyOptions = {}): Promise<DocumentTransaction> {
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["targetSheetId", "mode", "calculation", "skipBlanks", "maxCells", "maxCodeUnits", "maxSteps", "signal", "yield", "onProgress"].includes(key))) invalid("Invalid workbook range-copy options.");
  const config = { mode: "contents", skipBlanks: false, maxCells: 250_000, maxCodeUnits: 64 * 1024 * 1024, maxSteps: 10_000_000, ...options };
  if (!["contents", "values", "formulas", "formats"].includes(config.mode) || typeof config.skipBlanks !== "boolean"
    || config.targetSheetId !== undefined && (typeof config.targetSheetId !== "string" || !config.targetSheetId || config.targetSheetId.length > 256)
    || ![config.maxCells, config.maxCodeUnits, config.maxSteps].every(value => Number.isSafeInteger(value) && value > 0)
    || config.signal !== undefined && (!config.signal || typeof config.signal.aborted !== "boolean") || config.yield !== undefined && typeof config.yield !== "function"
    || config.onProgress !== undefined && typeof config.onProgress !== "function") invalid("Invalid workbook range-copy mode or budget.");
  const parsed = WorkbookRangeSchema.safeParse(range), coordinate = WorkbookCoordinateSchema.safeParse(origin);
  if (!parsed.success || parsed.data.from.row > parsed.data.to.row || parsed.data.from.column > parsed.data.to.column || !coordinate.success) invalid("Range copy requires an ordered source rectangle and a destination coordinate.");
  const source = parsed.data, destination = coordinate.data, model = captureWorkbookModel(input), snapshot = model.snapshot(), sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId);
  const targetSheetId = config.targetSheetId ?? sheetId, targetSheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === targetSheetId);
  if (!sheet || !targetSheet) invalid("Range copy requires existing source and destination worksheets.");
  const rows = source.to.row - source.from.row + 1, columns = source.to.column - source.from.column + 1, total = rows * columns;
  if (total > config.maxCells) throw new DocumentError("ResourceLimitExceeded", "Range copy exceeds its cell budget.");
  const target = { from: destination, to: { row: destination.row + rows - 1, column: destination.column + columns - 1 } };
  if (target.to.row >= 1_048_576 || target.to.column >= 16_384) invalid("Range copy exceeds the destination worksheet bounds.");
  let units = 0, steps = 0, completed = 0;
  const guard = (): void => {
    if (config.signal?.aborted) throw new DocumentError("OperationCancelled", "Workbook range copy was cancelled.");
    if (captureWorkbookModel(input) !== model) throw new DocumentError("RevisionConflict", "Workbook range-copy source changed before completion.");
  };
  const account = (amount: number): void => { units += amount; if (units > config.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Range copy exceeds its text budget."); };
  const step = (): void => { if (++steps > config.maxSteps) throw new DocumentError("ResourceLimitExceeded", "Range copy exceeds its preparation work budget."); };
  const copiesTopology = config.mode === "contents" || config.mode === "formats";
  const mappedMerges: WorkbookRange[] = [], copiedMerges: WorkbookRange[] = [], targetMerges: DeepReadonly<WorkbookRange>[] = [];
  for (const merge of sheet.merges) {
    guard(); step();
    if (overlap(merge, source)) {
      if (!contains(source, merge)) invalid("Range copy cannot split a merged region.");
      const mapped = { from: { row: destination.row + merge.from.row - source.from.row, column: destination.column + merge.from.column - source.from.column },
        to: { row: destination.row + merge.to.row - source.from.row, column: destination.column + merge.to.column - source.from.column } };
      account(JSON.stringify(mapped).length); mappedMerges.push(mapped);
      let include = copiesTopology && !config.skipBlanks;
      if (copiesTopology && config.skipBlanks) {
        // Native imports retain populated covered cells too. A blank anchor
        // alone does not make the complete merged seed blank. Disjoint merges
        // keep this scan within the selected rectangle's existing cell budget.
        content: for (let row = merge.from.row; row <= merge.to.row; row++) for (let column = merge.from.column; column <= merge.to.column; column++) {
          guard(); step(); const cell = model.getCell(sheetId, row, column);
          if (cell && (cell.formula !== undefined || cell.value.type !== "blank")) { include = true; break content; }
          if (steps % 256 === 0 && config.yield) { await config.yield(); guard(); }
        }
      }
      if (include) copiedMerges.push(mapped);
    }
    if (steps % 256 === 0 && config.yield) { await config.yield(); guard(); }
  }
  const mergeKey = (range: DeepReadonly<WorkbookRange>): string => [range.from.row, range.from.column, range.to.row, range.to.column].join(":");
  const mergeKeys = new Set(mappedMerges.map(mergeKey));
  for (const merge of targetSheet.merges) {
    guard(); step();
    if (overlap(merge, target)) {
      if (!contains(target, merge)) invalid("Range copy cannot split a merged region.");
      if (!mergeKeys.has(mergeKey(merge))) invalid("Range copy requires identically sized and aligned destination merges.");
      targetMerges.push(merge);
    }
    if (steps % 256 === 0 && config.yield) { await config.yield(); guard(); }
  }
  // Matching target merges already have exactly the required geometry. Keep
  // them in place, including native covered content and skipped blank seeds;
  // only newly constructed merges require the explicit covered-content check.
  const targetMergeKeys = new Set(targetMerges.map(mergeKey)), addedMerges = copiedMerges.filter(merge => !targetMergeKeys.has(mergeKey(merge)));
  const tables = config.mode === "contents" && (sheetId !== targetSheetId || source.from.row !== destination.row || source.from.column !== destination.column)
    && sheet.tables?.some(table => contains(source, table.range))
    ? await createWorkbookTableCopyPlan(snapshot.metadata.workbook, sheet, targetSheet, source, target, { guard, account, step, ...(config.yield ? { yield: config.yield } : {}) }) : undefined;
  let names: Awaited<ReturnType<typeof createWorkbookNameCopyPlan>> | undefined;
  const copyExpression = async (expression: string, from: WorkbookCoordinate, to: WorkbookCoordinate): Promise<string> => {
    if (sheetId !== targetSheetId && snapshot.metadata.workbook.names?.length) {
      names ??= await createWorkbookNameCopyPlan(snapshot.metadata.workbook, sheet, targetSheet, { guard, account, step, ...(config.yield ? { yield: config.yield } : {}) });
      const result = await names.copy(expression, from, to); guard(); return result;
    }
    return copyWorkbookFormula(expression, from, to);
  };
  const evaluator = config.calculation === undefined ? undefined : createWorkbookCalculatedEvaluator(model, config.calculation, { maxSteps: config.maxSteps, maxCells: config.maxCells, maxCodeUnits: config.maxCodeUnits });
  const cells: WorkbookCellInput[] = [], formats: WorkbookCellFormatInput[] = [];
  let general = snapshot.metadata.workbook.numberFormats.find(format => "builtinId" in format ? format.builtinId === 0 : format.code.toUpperCase() === "GENERAL")?.formatId, registerGeneral = false;
  try {
    guard();
    const hyperlinks = config.mode === "contents" && (sheetId !== targetSheetId || source.from.row !== destination.row || source.from.column !== destination.column) && (sheet.hyperlinks?.length || targetSheet.hyperlinks?.length)
      ? createWorkbookHyperlinkCopy(sheet.hyperlinks ?? [], targetSheet.hyperlinks ?? [], source, target, step) : undefined;
    let comments = false;
    if (config.mode === "contents") for (const annotation of snapshot.annotations) {
      guard(); step();
      const anchor = annotation.anchor;
      if (anchor.kind === "cells") for (const range of anchor.ranges) {
        step();
        if (anchor.sheetId === sheetId && overlap(range, source) || anchor.sheetId === targetSheetId && overlap(range, target)) { comments = true; break; }
      }
      if (comments) break;
    }
    const drawingIds: string[] = [];
    if (config.mode === "contents" && (sheetId !== targetSheetId || source.from.row !== destination.row || source.from.column !== destination.column)) for (const object of snapshot.metadata.workbook.objects ?? []) {
      guard(); step();
      if (object.sheetId === sheetId && (object.anchor.kind === "absolute" || contains(source, { from: object.anchor.from, to: object.anchor.from }))) drawingIds.push(object.nodeId);
      if (steps % 256 === 0 && config.yield) { await config.yield(); guard(); }
    }
    const footprint = (config.mode === "contents" || config.mode === "formats") && (sheetId !== targetSheetId || source.from.row !== destination.row || source.from.column !== destination.column)
      && (sheet.conditionalRules?.length || targetSheet.conditionalRules?.length || config.mode === "contents" && (drawingIds.length || comments || sheet.dataValidations?.length || targetSheet.dataValidations?.length)) ? createWorkbookCopyFootprint() : undefined;
    for (let row = 0; row < rows; row++) for (let column = 0; column < columns; column++) {
      guard(); step(); const from = { row: source.from.row + row, column: source.from.column + column }, to = { row: destination.row + row, column: destination.column + column };
      const raw = model.getCell(sheetId, from.row, from.column), same = sheetId === targetSheetId && from.row === to.row && from.column === to.column;
      const copied = !(config.skipBlanks && (!raw || raw.value.type === "blank" && !raw.formula)) && !same;
      hyperlinks?.cell(from.row, from.column, to.row, to.column, copied);
      footprint?.cell(to.row, to.column, copied);
      if (copied || config.mode === "values" && same && raw?.formula) {
        const format = resolveWorkbookCellFormat(sheet, from, raw), previous = model.getCell(targetSheetId, to.row, to.column);
        if (config.mode === "formats" || config.mode === "contents") {
          if (format.numberFormatId === undefined && general === undefined) {
            const identifiers = new Set(snapshot.metadata.workbook.numberFormats.map(format => format.formatId)); let serial = 1;
            do { general = "copy-general-" + serial++; } while (identifiers.has(general)); registerGeneral = true;
          }
        }
        if (config.mode === "formats") {
          const targetFormat = resolveWorkbookCellFormat(targetSheet, to, previous);
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
          if (content.kind === "formula") {
            const expression = await copyExpression(content.expression, from, to);
            content = { kind: "formula", expression: tables?.bind(expression, from, to) ?? expression };
          }
          let unchanged = false;
          if (config.mode !== "contents" && content.kind === "value" && content.value.type === "text" && previous?.value.type === "text" && !previous.formula) {
            const edit = workbookTextSplice(previous.text, content.value.value);
            if (edit) content = replaceWorkbookCellText(model, { sheetId: targetSheetId, ...to }, [edit]); else unchanged = true;
          }
          const cell: WorkbookCellInput = { ...to, source: content };
          if (config.mode === "contents") {
            cell.cssClass = format.cssClass; cell.formatId = format.numberFormatId ?? general ?? null;
            const targetFormat = resolveWorkbookCellFormat(targetSheet, to, previous);
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
    const copiedRanges = footprint?.finish();
    const ruleCopy = copiedRanges ? await copyWorkbookRules(snapshot.metadata.workbook, sheet, targetSheet, source, target, copiedRanges, config.mode === "contents", copyExpression, { guard, step, account, ...(config.yield ? { yield: config.yield } : {}) }) : undefined;
    const copiedTables = await tables?.finish(copyExpression) ?? [];
    guard(); const operations: DocumentTransaction["operations"] = [];
    if (comments && copiedRanges?.length) {
      const sourceMasks = copiedRanges.map(range => ({ from: { row: range.from.row - destination.row + source.from.row, column: range.from.column - destination.column + source.from.column }, to: { row: range.to.row - destination.row + source.from.row, column: range.to.column - destination.column + source.from.column } }));
      const included = (ranges: readonly DeepReadonly<WorkbookRange>[]): boolean => {
        for (const range of ranges) for (const mask of sourceMasks) { guard(); step(); if (overlap(range, mask)) return true; } return false;
      };
      for (const annotation of snapshot.annotations) {
        guard(); step();
        if (annotation.anchor.kind !== "cells" || annotation.anchor.sheetId !== sheetId || !included(annotation.anchor.ranges)) continue;
        account(JSON.stringify(annotation).length);
        for (const body of [annotation, ...annotation.replies ?? []]) if (body.bodyNodeId) { step(); account(model.getNode(body.bodyNodeId)?.html.length ?? 0); }
      }
      const operation = { op: "copyWorkbookComments" as const, sheetId, range: source, targetSheetId, target: destination, ranges: copiedRanges };
      account(JSON.stringify(operation).length); operations.push(operation);
    }
    for (const name of names?.names ?? []) {
      const operation = { op: "setWorkbookDefinedName" as const, name }; account(JSON.stringify(operation).length); operations.push(operation);
    }
    if (registerGeneral && (cells.some(cell => cell.formatId === general) || formats.some(cell => cell.formatId === general))) {
      const operation = { op: "setWorkbookNumberFormat" as const, format: { formatId: general!, builtinId: 0 } }; account(JSON.stringify(operation).length); operations.push(operation);
    }
    if (cells.length) operations.push({ op: "setWorkbookCells", sheetId: targetSheetId, cells });
    if (formats.length) operations.push({ op: "setWorkbookCellFormats", sheetId: targetSheetId, cells: formats });
    for (const table of copiedTables) {
      const operation = { op: "setWorkbookTable" as const, sheetId: targetSheetId, table, cellPolicy: "preserve" as const }; account(JSON.stringify(operation).length); operations.push(operation);
    }
    const copiedLinks = hyperlinks?.finish();
    if (copiedLinks) { const operation = { op: "replaceWorkbookHyperlinks" as const, sheetId: targetSheetId, range: target, links: copiedLinks }; account(JSON.stringify(operation).length); operations.push(operation); }
    if (ruleCopy) { account(JSON.stringify(ruleCopy).length); operations.push(ruleCopy); }
    for (const range of addedMerges) {
      const operation = { op: "mergeWorkbookCells" as const, sheetId: targetSheetId, range, coveredContent: "reject" as const };
      account(JSON.stringify(operation).length); operations.push(operation);
    }
    if (drawingIds.length && copiedRanges?.length) {
      const operation = { op: "copyWorkbookObjects" as const, nodeIds: drawingIds, sheetId, range: source, targetSheetId, target: destination, ranges: copiedRanges,
        maxSourceCodeUnits: 0 };
      const overhead = JSON.stringify(operation).length + String(config.maxCodeUnits).length - 1;
      operation.maxSourceCodeUnits = Math.max(0, config.maxCodeUnits - units - overhead);
      account(JSON.stringify(operation).length); operations.push(operation);
    }
    return { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations };
  } finally { evaluator?.dispose(); }
}
