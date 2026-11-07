import { captureWorkbookModel, createWorkbookCalculatedEvaluator, createWorkbookConditionalFormatResolver, createWorkbookNumberFormatter, createWorkbookTableFormatResolver, DocumentError, isDocumentTextBoundary, readWorkbookRules, resolveWorkbookCellFormat, WorkbookCoordinateSchema, WorkbookRangeSchema } from "@doffi/core";
import type { DeepReadonly, WorkbookCell, WorkbookCellFormat, WorkbookModel, WorkbookRange, WorkbookRuleCell, WorkbookScalarEvaluator, WorkbookValue } from "@doffi/core";
import type { WorkbookFindDiagnostic, WorkbookFindMatch, WorkbookFindOptions, WorkbookFindQuery, WorkbookFindResult } from "./find-contracts.js";

export function invalidWorkbookFind(message: string): never { throw new DocumentError("InvalidOperation", message); }
export function workbookFindUnicode(text: string): void {
  if (typeof text !== "string") invalidWorkbookFind("Workbook search requires text.");
  for (const character of text) { const point = character.codePointAt(0)!; if (!point || point >= 0xd800 && point <= 0xdfff) invalidWorkbookFind("Workbook search requires well-formed Unicode without null characters."); }
}
export interface WorkbookFindVisit { readonly cell: WorkbookCell; readonly text: string; readonly matches: readonly WorkbookFindMatch[] }
/** Internal visitor shares one bounded scan with the replacement command. */
export async function scanWorkbookText(input: WorkbookModel, query: WorkbookFindQuery, options: WorkbookFindOptions = {}, visit?: (value: WorkbookFindVisit, account: (units: number) => void) => void, only?: WorkbookFindMatch): Promise<WorkbookFindResult> {
  if (!query || typeof query !== "object" || Array.isArray(query) || Object.keys(query).some(key => !["text", "field", "matchCase", "entireCell", "scope"].includes(key))
    || query.field !== undefined && !["source", "formula", "display"].includes(query.field)
    || query.matchCase !== undefined && typeof query.matchCase !== "boolean" || query.entireCell !== undefined && typeof query.entireCell !== "boolean") invalidWorkbookFind("Invalid workbook search query.");
  workbookFindUnicode(query.text);
  if (!query.text || query.text.length > 32767) invalidWorkbookFind("Workbook search requires a nonempty bounded query.");
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["calculation", "locale", "maxCells", "maxMatches", "maxCodeUnits", "maxWindows", "maxSteps", "maxValues", "windowCells", "signal", "yield", "onProgress"].includes(key))) invalidWorkbookFind("Invalid workbook search options.");
  const config = { maxCells: 250_000, maxMatches: 250_000, maxCodeUnits: 64 * 1024 * 1024, maxWindows: 65536, maxSteps: 100_000_000, maxValues: 250_000, windowCells: 4096, ...options };
  if (![config.maxCells, config.maxMatches, config.maxCodeUnits, config.maxWindows, config.maxSteps, config.maxValues, config.windowCells].every(value => Number.isSafeInteger(value) && value > 0) || config.windowCells > 4096
    || config.locale !== undefined && (typeof config.locale !== "string" || !config.locale || config.locale.length > 255)
    || config.signal !== undefined && (!config.signal || typeof config.signal.aborted !== "boolean") || config.yield !== undefined && typeof config.yield !== "function"
    || config.onProgress !== undefined && typeof config.onProgress !== "function") invalidWorkbookFind("Invalid workbook search limits or callbacks.");
  const model = captureWorkbookModel(input), snapshot = model.snapshot(), field = query.field ?? "source", scoped = query.scope !== undefined, entireCell = query.entireCell === true;
  const sheets = new Map(snapshot.metadata.workbook.sheets.map(sheet => [sheet.sheetId, sheet])), scopes = new Map<string, readonly DeepReadonly<WorkbookRange>[]>();
  if (query.scope !== undefined) {
    if (!Array.isArray(query.scope) || query.scope.length > 1024) invalidWorkbookFind("Workbook search scope must be a bounded list of worksheet ranges.");
    for (const area of query.scope) {
      if (!area || typeof area !== "object" || Object.keys(area).some(key => !["sheetId", "range"].includes(key)) || !sheets.has(area.sheetId)) invalidWorkbookFind("Workbook search references an unknown worksheet.");
      const parsed = WorkbookRangeSchema.safeParse(area.range);
      if (!parsed.success || parsed.data.from.row > parsed.data.to.row || parsed.data.from.column > parsed.data.to.column) invalidWorkbookFind("Workbook search requires ordered ranges.");
      scopes.set(area.sheetId, [...scopes.get(area.sheetId) ?? [], parsed.data]);
    }
  }
  const contains = (sheetId: string, row: number, column: number): boolean => (!only || only.sheetId === sheetId && only.row === row && only.column === column)
    && (!scoped || Boolean(scopes.get(sheetId)?.some(range => range.from.row <= row && row <= range.to.row && range.from.column <= column && column <= range.to.column)));
  const intersects = (sheetId: string, row: number, column: number, rows: number, columns: number): boolean => (!only || only.sheetId === sheetId && row <= only.row && only.row < row + rows && column <= only.column && only.column < column + columns)
    && (!scoped || Boolean(scopes.get(sheetId)?.some(range => range.from.row < row + rows && row <= range.to.row && range.from.column < column + columns && column <= range.to.column)));
  const pattern = query.text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), regex = new RegExp(query.entireCell ? "^(?:" + pattern + ")$" : pattern, query.matchCase ? "gu" : "giu");
  const matches: WorkbookFindMatch[] = [], diagnostics: WorkbookFindDiagnostic[] = []; let cells = 0, windows = 0, units = 0, steps = 0, values = 0;
  function account(count: number): void { units += count; if (units > config.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook search exceeds its text budget."); }
  function guard(): void {
    if (config.signal?.aborted) throw new DocumentError("OperationCancelled", "Workbook search was cancelled.");
    if (captureWorkbookModel(input) !== model) throw new DocumentError("RevisionConflict", "Workbook search source changed before completion.");
  }
  async function checkpoint(): Promise<void> {
    guard(); try { config.onProgress?.(Object.freeze({ cells, windows, matches: matches.length })); } catch { /* Reporting cannot publish source. */ }
    guard(); if (config.yield) { await config.yield(); guard(); }
  }
  const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem, ...(config.locale === undefined ? {} : { hostLocale: config.locale }), ...(snapshot.metadata.locale ? { documentLocale: snapshot.metadata.locale } : {}) });
  const formats = new Map(snapshot.metadata.workbook.numberFormats.map(format => [format.formatId, format]));
  let evaluator: WorkbookScalarEvaluator | undefined;
  function sourceText(cell: WorkbookCell, format: WorkbookCellFormat, rule?: WorkbookRuleCell): string | undefined {
    if (field === "formula") return cell.formula?.expression;
    if (field === "source") return cell.formula?.expression ?? (cell.value.type === "blank" ? undefined : cell.value.type === "boolean" ? cell.value.value ? "TRUE" : "FALSE" : String(cell.value.value));
    let value: DeepReadonly<WorkbookValue> = cell.value;
    if (rule) {
      if (rule.current.kind !== "value") { diagnostics.push(Object.freeze({ sheetId: cell.sheetId, row: cell.row, column: cell.column, reason: "unavailable" })); return; }
      value = rule.current.value;
    } else if (cell.formula) {
      if (!evaluator) { diagnostics.push(Object.freeze({ sheetId: cell.sheetId, row: cell.row, column: cell.column, reason: "calculation-required" })); return; }
      if (steps >= config.maxSteps || values >= config.maxValues) throw new DocumentError("ResourceLimitExceeded", "Workbook display search exceeds its evaluation budget.");
      const result = evaluator.cell({ sheetId: cell.sheetId, row: cell.row, column: cell.column }, { maxSteps: config.maxSteps - steps, maxCodeUnits: Math.max(0, config.maxCodeUnits - units), ...(config.signal ? { signal: config.signal } : {}) });
      steps += result.steps; values += result.readCells; account(result.codeUnits);
      if (values > config.maxValues) throw new DocumentError("ResourceLimitExceeded", "Workbook display search exceeds its value budget.");
      if (result.kind !== "value") { diagnostics.push(Object.freeze({ sheetId: cell.sheetId, row: cell.row, column: cell.column, reason: "unavailable" })); return; }
      value = result.value;
    }
    if (value.type === "blank") return;
    if (value.type === "boolean") return value.value ? "TRUE" : "FALSE";
    if (value.type === "error") return value.value;
    const definition = format.numberFormatId === undefined ? undefined : formats.get(format.numberFormatId);
    const display = formatter.format(value.value, definition);
    if (display.kind !== "parts") { diagnostics.push(Object.freeze({ sheetId: cell.sheetId, row: cell.row, column: cell.column, reason: "format-overflow" })); return; }
    return display.parts.map(part => part.kind === "text" ? part.text : "").join("");
  }
  try {
    guard(); account(query.text.length);
    if (config.calculation !== undefined) evaluator = createWorkbookCalculatedEvaluator(model, config.calculation, { maxSteps: config.maxSteps, maxCells: config.maxValues, maxCodeUnits: config.maxCodeUnits });
    for (const sheet of sheets.values()) {
      if (scoped && !scopes.has(sheet.sheetId)) continue;
      const tableFormat = field === "display" ? createWorkbookTableFormatResolver(snapshot, sheet) : undefined, conditionalFormat = createWorkbookConditionalFormatResolver(sheet);
      const formatRules = sheet.conditionalRules?.filter(rule => "numberFormatId" in rule && rule.numberFormatId !== undefined) ?? [];
      const chunks = snapshot.manifest.chunks.filter(chunk => chunk.sheetId === sheet.sheetId && intersects(sheet.sheetId, chunk.row, chunk.column, 256, 32)).sort((a, b) => a.row - b.row || a.column - b.column);
      for (let start = 0; start < chunks.length;) {
        let end = start + 1; while (end < chunks.length && chunks[end]!.row === chunks[start]!.row) end++;
        const block = chunks.slice(start, end), bandRows = Math.min(128, Math.max(1, Math.floor(config.windowCells / 32)));
        for (let band = 0; band < 256; band += bandRows) {
          const found: WorkbookCell[] = [], rows = Math.min(bandRows, 256 - band);
          for (const chunk of block) {
            const row = chunk.row + band;
            for (let offset = 0; offset < 32; offset += config.windowCells) {
              const columns = Math.min(32 - offset, Math.floor(config.windowCells / rows));
              if (!intersects(sheet.sheetId, row, chunk.column + offset, rows, columns)) continue;
              guard(); if (++windows > config.maxWindows) throw new DocumentError("ResourceLimitExceeded", "Workbook search exceeds its window budget.");
              const window = { sheetId: sheet.sheetId, row, column: chunk.column + offset, rows, columns };
              for (const cell of model.readWindow(window)) {
                if (!WorkbookCoordinateSchema.safeParse({ row: cell.row, column: cell.column }).success || cell.sheetId !== sheet.sheetId || cell.row < row || cell.row >= row + rows || cell.column < window.column || cell.column >= window.column + columns) throw new DocumentError("InvalidDocument", "Workbook provider returned a cell outside its search window.");
                if (++cells > config.maxCells) throw new DocumentError("ResourceLimitExceeded", "Workbook search exceeds its cell budget.");
                account(cell.html.length + cell.text.length + (cell.formula?.expression.length ?? 0));
                if (contains(sheet.sheetId, cell.row, cell.column)) found.push(cell);
              }
              await checkpoint();
            }
          }
          found.sort((a, b) => a.row - b.row || a.column - b.column);
          const ruleCells = new Map<number, WorkbookRuleCell>();
          if (field === "display" && config.calculation && formatRules.length) {
            const needed = found.filter(cell => formatRules.some(rule => rule.ranges.some(range => range.from.row <= cell.row && cell.row <= range.to.row && range.from.column <= cell.column && cell.column <= range.to.column)));
            for (let offset = 0; offset < needed.length; offset += 4096) {
              guard(); if (steps >= config.maxSteps || values >= config.maxValues || units >= config.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook display search exceeds its rule budget.");
              const response = await readWorkbookRules(model, { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, sheetId: sheet.sheetId,
                coordinates: needed.slice(offset, offset + 4096).map(cell => ({ row: cell.row, column: cell.column })), maxCells: 4096, maxRules: 256, maxValues: config.maxValues - values, maxSteps: config.maxSteps - steps, maxCodeUnits: config.maxCodeUnits - units },
                { calculation: config.calculation, ...(config.signal ? { signal: config.signal } : {}), yield: checkpoint });
              steps += response.steps; values += response.readCells; account(JSON.stringify(response).length); guard();
              for (const value of response.cells) ruleCells.set(value.row * 16_384 + value.column, value);
            }
          }
          for (let index = 0; index < found.length; index++) {
            const cell = found[index]!, base = resolveWorkbookCellFormat(sheet, cell, cell), rule = ruleCells.get(cell.row * 16_384 + cell.column);
            const composed = conditionalFormat(cell, rule?.conditional ?? [], tableFormat?.(cell, base) ?? base);
            if (field === "display" && (!composed.complete || !config.calculation && formatRules.some(rule => rule.ranges.some(range => range.from.row <= cell.row && cell.row <= range.to.row && range.from.column <= cell.column && cell.column <= range.to.column)))) {
              diagnostics.push(Object.freeze({ sheetId: cell.sheetId, row: cell.row, column: cell.column, reason: config.calculation ? "unavailable" : "calculation-required" })); continue;
            }
            const text = sourceText(cell, composed.format, rule); if (text === undefined) continue; account(text.length);
            regex.lastIndex = 0; const cellMatches: WorkbookFindMatch[] = [];
            for (let match = regex.exec(text); match; match = regex.exec(text)) {
              const from = match.index, to = from + match[0].length;
              if (entireCell && (from !== 0 || to !== text.length)) continue;
              if (!isDocumentTextBoundary(text, from) || !isDocumentTextBoundary(text, to)) continue;
              if (only && (from !== only.from || to !== only.to || match[0] !== only.text)) continue;
              if (matches.length >= config.maxMatches) throw new DocumentError("ResourceLimitExceeded", "Workbook search exceeds its match budget.");
              const hit = Object.freeze({ sheetId: cell.sheetId, row: cell.row, column: cell.column, from, to, text: match[0] }); account(hit.text.length); matches.push(hit); cellMatches.push(hit);
            }
            if (cellMatches.length) visit?.({ cell, text, matches: cellMatches }, account);
            if (index % 256 === 255) await checkpoint();
          }
        }
        start = end;
      }
    }
    await checkpoint(); guard();
    return Object.freeze({ documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, matches: Object.freeze(matches), diagnostics: Object.freeze(diagnostics), complete: !diagnostics.length, cells, windows });
  } finally { evaluator?.dispose(); formatter.dispose(); }
}

/** Search stored source cells in worksheet/row/column order. Implicit empty
 * grids are never materialized; limits or cancellation never return a partial
 * result disguised as complete. Display diagnostics are explicit. */
export async function findWorkbookText(model: WorkbookModel, query: WorkbookFindQuery, options?: WorkbookFindOptions): Promise<WorkbookFindResult> {
  return scanWorkbookText(model, query, options);
}
