import { checkLimit, invalid } from "../document/validation.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "./chunk-index.js";
import { type FormulaExpression, type FormulaOrigin } from "./formula.js";
import { printStructuredReference } from "./structured-reference.js";

export function checkFormulaOrigin(origin: FormulaOrigin): void {
  if (!Number.isInteger(origin.row) || !Number.isInteger(origin.column) || origin.row < 0 || origin.column < 0
    || origin.row >= MAX_SHEET_ROWS || origin.column >= MAX_SHEET_COLUMNS) invalid("Invalid formula origin.");
}
export function formulaColumn(column: number): string {
  let output = "";
  for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) output = String.fromCharCode(65 + (value - 1) % 26) + output;
  return output;
}
const sheetPrefix = (sheet: string | undefined): string => sheet === undefined ? "" : sheet === "" ? "!" : "'" + sheet.replaceAll("'", "''") + "'!";

/** Canonical, fully parenthesized expressions preserve reference precedence. */
export function printFormula(expression: FormulaExpression, origin: FormulaOrigin): string { return printWithStyle(expression, origin, "A1"); }
export function printR1C1Formula(expression: FormulaExpression, origin: FormulaOrigin): string { return printWithStyle(expression, origin, "R1C1"); }
function printWithStyle(expression: FormulaExpression, origin: FormulaOrigin, referenceStyle: "A1" | "R1C1"): string {
  checkFormulaOrigin(origin);
  let nodes = 0;
  function print(node: FormulaExpression, depth = 0): string {
    checkLimit(depth, 128); checkLimit(++nodes, 4096);
    let output: string;
    switch (node.kind) {
      case "number": output = String(node.value); break;
      case "boolean": output = node.value ? "TRUE" : "FALSE"; break;
      case "string": output = '"' + node.value.replaceAll('"', '""') + '"'; break;
      case "error": output = node.value; break;
      case "name": output = (node.workbook === 0 ? "[0]!" : sheetPrefix(node.sheet)) + node.name; break;
      case "structured": output = printStructuredReference(node); break;
      case "cell": {
        if (referenceStyle === "R1C1") { output = sheetPrefix(node.sheet) + "R" + (node.absoluteRow ? node.row + 1 : node.row ? "[" + node.row + "]" : "") + "C" + (node.absoluteColumn ? node.column + 1 : node.column ? "[" + node.column + "]" : ""); break; }
        const row = node.row + (node.absoluteRow ? 0 : origin.row), column = node.column + (node.absoluteColumn ? 0 : origin.column);
        output = row < 0 || row >= MAX_SHEET_ROWS || column < 0 || column >= MAX_SHEET_COLUMNS ? "#REF!"
          : sheetPrefix(node.sheet) + (node.absoluteColumn ? "$" : "") + formulaColumn(column) + (node.absoluteRow ? "$" : "") + (row + 1);
        break;
      }
      case "range": {
        const from = print(node.from, depth + 1), to = print(node.to, depth + 1);
        // A deleted/out-of-grid endpoint stays explicit, rather than clipping it.
        output = from + ":" + (from !== "#REF!" && node.from.sheet === node.to.sheet && node.to.sheet !== undefined && to !== "#REF!" ? to.slice(sheetPrefix(node.to.sheet).length) : to);
        break;
      }
      case "axisRange": {
        if (referenceStyle === "R1C1") {
          const axis = node.axis === "row" ? "R" : "C", end = (position: typeof node.from): string => axis + (position.absolute ? position.index + 1 : position.index ? "[" + position.index + "]" : "");
          output = sheetPrefix(node.sheet) + end(node.from) + ":" + end(node.to); break;
        }
        const end = (position: typeof node.from): string => {
          const index = position.index + (position.absolute ? 0 : origin[node.axis]);
          return index < 0 || index >= (node.axis === "row" ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS) ? "#REF!"
            : (position.absolute ? "$" : "") + (node.axis === "row" ? index + 1 : formulaColumn(index));
        };
        const from = end(node.from), to = end(node.to);
        output = from === "#REF!" || to === "#REF!" ? "#REF!" : sheetPrefix(node.sheet) + from + ":" + to;
        break;
      }
      case "unary": {
        const value = print(node.operand, depth + 1);
        output = node.operator === "%" ? value + "%" : node.operator + (node.operand.kind === "unary" && node.operand.operator === "%" ? "(" + value + ")" : value);
        break;
      }
      case "binary": output = "(" + print(node.left, depth + 1) + node.operator + print(node.right, depth + 1) + ")"; break;
      case "call": output = node.name + "(" + node.arguments.map(argument => argument === null ? "" : print(argument, depth + 1)).join(",") + ")"; break;
    }
    checkLimit(output.length + 1, 8192); return output;
  }
  return "=" + print(expression);
}
