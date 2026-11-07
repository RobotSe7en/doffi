import type { WorkbookModel } from "../runtime.js";
import { parseFormula } from "./formula.js";
import type { FormulaOrigin } from "./formula.js";

/** Comparison only: worksheet identity survives a source spelling change.
 * String literals and coordinates remain exact; this never writes source. */
export function workbookExpressionIdentity(model: WorkbookModel, expression: string, origin: FormulaOrigin): string {
  const sheets = new Map(model.snapshot().metadata.workbook.sheets.map(sheet => [sheet.name.toUpperCase(), sheet.sheetId]));
  return JSON.stringify(parseFormula(expression, origin), (key, value: unknown) => key === "sheet" && typeof value === "string"
    ? sheets.has(value.toUpperCase()) ? ["worksheet", sheets.get(value.toUpperCase())] : ["unbound", value.toUpperCase()] : value);
}
