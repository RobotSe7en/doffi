import type { WorkbookCoordinate, WorkbookMetadata, WorkbookSheet } from "../contracts/workbook.js";
import { rewriteFormula } from "../workbook/formula-rewrite.js";
import type { FormulaReferenceChange } from "../workbook/formula-rewrite.js";
import { WorkbookCellEdit, rangeContains } from "./workbook-cells.js";

/** Every formula-bearing source participates in a structural reference edit.
 * Cached values are never rewritten into independent authoritative content. */
export function rewriteWorkbookReferences(workbook: WorkbookMetadata, cells: WorkbookCellEdit, change: FormulaReferenceChange): void {
  const tables = workbook.sheets.flatMap(sheet => (sheet.tables ?? []).map(table => ({ name: table.name, sheet: sheet.name,
    range: table.range, header: table.header, totals: table.totals, columns: table.columns.map(column => column.name) })));
  function rewrite(expression: string, sheet: string, at: WorkbookCoordinate, table?: string): string {
    const target = change.kind === "renameSheet" && sheet === change.from ? change.to : sheet;
    return rewriteFormula(expression, { ...at, sheet }, { ...at, sheet: target }, change, { tables, ...(table ? { table } : {}),
      isTableName: (name, owner) => !workbook.names?.some(item => item.name.toUpperCase() === name.toUpperCase() && (item.sheetId === undefined || workbook.sheets.find(sheet => sheet.sheetId === item.sheetId)?.name === owner))
        && tables.some(table => table.name.toUpperCase() === name.toUpperCase()),
    });
  }
  for (const cell of cells.entries(undefined, true)) {
    const sheet = workbook.sheets.find(sheet => sheet.sheetId === cell.sheetId)!;
    if (change.kind === "deleteSheet" && sheet.name === change.sheet) continue;
    const expression = rewrite(cell.formula!.expression, sheet.name, cell, sheet.tables?.find(table => rangeContains(table.range, cell.row, cell.column))?.name);
    if (expression !== cell.formula!.expression) cells.set(cell.sheetId, { row: cell.row, column: cell.column, source: { kind: "formula", expression } });
  }
  for (const name of workbook.names ?? []) {
    const sheet = name.sheetId === undefined ? "$workbook" : workbook.sheets.find(sheet => sheet.sheetId === name.sheetId)!.name;
    if (change.kind === "deleteSheet" && sheet === change.sheet) continue;
    name.expression = rewrite(name.expression, sheet, name.origin);
  }
  function sheetReferences(sheet: WorkbookSheet): void {
    for (const rule of [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []]) {
      if ("expression" in rule) rule.expression = rewrite(rule.expression, sheet.name, rule.origin);
      if ("formulas" in rule) rule.formulas = rule.formulas.map(formula => rewrite(formula, sheet.name, rule.origin));
      if (rule.kind === "colorScale") for (const stop of rule.stops) if (stop.kind === "formula") stop.expression = rewrite(stop.expression, sheet.name, rule.origin);
    }
    for (const table of sheet.tables ?? []) for (const [index, column] of table.columns.entries()) {
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
    link.location = rewrite("=" + link.location, "$workbook", { row: 0, column: 0 }).slice(1);
  }
}
