import type { WorkbookCoordinate, WorkbookMetadata, WorkbookSheet } from "../contracts/workbook.js";
import { rewriteFormula } from "../workbook/formula-rewrite.js";
import type { FormulaLocation, FormulaReferenceChange } from "../workbook/formula-rewrite.js";
import { WorkbookCellEdit, rangeContains } from "./workbook-cells.js";

/** Every formula-bearing source participates in a structural reference edit.
 * Cached values are never rewritten into independent authoritative content. */
export function rewriteWorkbookReferences(workbook: WorkbookMetadata, cells: WorkbookCellEdit, change: FormulaReferenceChange): void {
  for (const _step of rewriteWorkbookReferenceSteps(workbook, cells, change)) { /* Synchronous driver of the same kernel. */ }
}
export function createWorkbookReferenceRewriter(workbook: WorkbookMetadata, change: FormulaReferenceChange, options: { readonly preserveRelativeAxis?: boolean } = {}) {
  const tables = workbook.sheets.flatMap(sheet => (sheet.tables ?? []).map(table => ({ name: table.name, sheet: sheet.name,
    range: table.range, header: table.header, totals: table.totals, columns: table.columns.map(column => column.name) })));
  const sheetNames = new Map(workbook.sheets.map(sheet => [sheet.sheetId, sheet.name.toUpperCase()])), globalNames = new Map<string, string>(), locals = new Map<string, Map<string, string>>(), localNames = new Set<string>();
  for (const name of workbook.names ?? []) {
    const token = name.name.toUpperCase();
    if (name.sheetId === undefined) globalNames.set(token, name.nameId);
    else { const owner = sheetNames.get(name.sheetId)!, scope = locals.get(owner) ?? new Map<string, string>(); scope.set(token, name.nameId); locals.set(owner, scope); localNames.add(token); }
  }
  const tableNames = new Set(tables.map(table => table.name.toUpperCase()));
  return function rewrite(expression: string, sheet: string, at: WorkbookCoordinate, table?: string, destination?: FormulaLocation, globalContext = false): string {
    const target = change.kind === "renameSheet" && sheet === change.from ? change.to : sheet;
    return rewriteFormula(expression, { ...at, sheet }, destination ?? { ...at, sheet: target }, change, { ...options, tables, ...(table ? { table } : {}),
      resolveName: (name, owner, explicit) => {
        const token = name.toUpperCase();
        // A workbook definition whose nested name is locally shadowed has no
        // verified caller scope. Preserve that ambiguity instead of rebinding.
        if (globalContext && localNames.has(token)) return undefined;
        return locals.get(owner.toUpperCase())?.get(token) ?? (explicit ? undefined : globalNames.get(token));
      },
      isTableName: (name, owner) => !globalNames.has(name.toUpperCase()) && !locals.get(owner.toUpperCase())?.has(name.toUpperCase()) && tableNames.has(name.toUpperCase()),
    });
  };
}
export function* rewriteWorkbookReferenceSteps(workbook: WorkbookMetadata, cells: WorkbookCellEdit, change: FormulaReferenceChange,
  options: { readonly skipTableFormulas?: ReadonlySet<string> } = {}): Generator<{ completed: number; total: number }, void> {
  let completed = 0; const rewrite = createWorkbookReferenceRewriter(workbook, change);
  for (const cell of cells.entries(undefined, true)) {
    const sheet = workbook.sheets.find(sheet => sheet.sheetId === cell.sheetId)!;
    if (change.kind === "deleteSheet" && sheet.name === change.sheet) continue;
    const destination = change.kind === "move" && sheet.name.toUpperCase() === change.sheet.toUpperCase() && rangeContains(change.range, cell.row, cell.column)
      ? { sheet: change.target.sheet, row: cell.row + change.target.row - change.range.from.row, column: cell.column + change.target.column - change.range.from.column } : undefined;
    const expression = rewrite(cell.formula!.expression, sheet.name, cell, sheet.tables?.find(table => rangeContains(table.range, cell.row, cell.column))?.name, destination);
    if (expression !== cell.formula!.expression) cells.rewriteFormula(cell.sheetId, cell.row, cell.column, expression);
    if (++completed % 256 === 0) yield { completed, total: 0 };
  }
  for (const name of workbook.names ?? []) {
    const sheet = name.sheetId === undefined ? "$workbook" : workbook.sheets.find(sheet => sheet.sheetId === name.sheetId)!.name;
    if (change.kind === "deleteSheet" && sheet === change.sheet) continue;
    name.expression = rewrite(name.expression, sheet, name.origin, undefined, undefined, name.sheetId === undefined);
  }
  function sheetReferences(sheet: WorkbookSheet): void {
    for (const rule of [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []]) {
      if ("expression" in rule) rule.expression = rewrite(rule.expression, sheet.name, rule.origin);
      if ("formulas" in rule) rule.formulas = rule.formulas.map(formula => rewrite(formula, sheet.name, rule.origin));
      if (rule.kind === "colorScale") for (const stop of rule.stops) if (stop.kind === "formula") stop.expression = rewrite(stop.expression, sheet.name, rule.origin);
    }
    for (const table of sheet.tables ?? []) if (!options.skipTableFormulas?.has(table.tableId)) for (const [index, column] of table.columns.entries()) {
      if (column.calculatedFormula) column.calculatedFormula = rewrite(column.calculatedFormula, sheet.name, { row: table.range.from.row + Number(table.header), column: table.range.from.column + index }, table.name);
      if (column.totalFormula) column.totalFormula = rewrite(column.totalFormula, sheet.name, { row: table.range.to.row, column: table.range.from.column + index }, table.name);
    }
    for (const link of sheet.hyperlinks ?? []) if (link.location) link.location = rewrite("=" + link.location, sheet.name, { row: 0, column: 0 }).slice(1);
  }
  for (const sheet of workbook.sheets) if (change.kind !== "deleteSheet" || sheet.name !== change.sheet) sheetReferences(sheet);
  function chartReferences(value: unknown, sheet: string): void {
    if (!value || typeof value !== "object") return;
    if (Array.isArray(value)) { for (const item of value) chartReferences(item, sheet); return; }
    const object = value as Record<string, unknown>;
    if (typeof object.expression === "string") {
      const expression = rewrite(object.expression, sheet, { row: 0, column: 0 });
      if (expression !== object.expression && object.cache && typeof object.cache === "object") (object.cache as { state: string }).state = "stale";
      object.expression = expression;
    }
    for (const [name, item] of Object.entries(object)) if (name !== "cache") chartReferences(item, sheet);
  }
  for (const chart of workbook.charts ?? []) {
    const owner = workbook.objects!.find(object => object.nodeId === chart.nodeId)!, sheet = workbook.sheets.find(sheet => sheet.sheetId === owner.sheetId)!;
    if (change.kind !== "deleteSheet" || sheet.name !== change.sheet) chartReferences(chart, sheet.name);
  }
  for (const link of workbook.objectHyperlinks ?? []) if (link.location) {
    // Object hyperlinks may target child text nodes, so their owning sheet is
    // supplied by the caller's absolute native location or a workbook scope.
    link.location = rewrite("=" + link.location, "$workbook", { row: 0, column: 0 }, undefined, undefined, true).slice(1);
  }
  yield { completed, total: completed };
}
