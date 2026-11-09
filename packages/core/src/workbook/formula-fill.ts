import type { WorkbookModel } from "../runtime.js";
import { parseFormula, type FormulaExpression, type FormulaOrigin } from "./formula.js";
import { checkFormulaOrigin, printFormula } from "./formula-print.js";
import { captureWorkbookModel } from "./source.js";

export type WorkbookFormulaFiller = (source: string, from: FormulaOrigin, to: FormulaOrigin) => string;

/** Capture table columns once for a revision-bound fill. Ordinary formula copy
 * keeps structured columns; horizontal fill advances qualified single columns.
 * Explicit column ranges (including [A]:[A]) retain their column bindings. */
export function createWorkbookFormulaFiller(input: WorkbookModel): WorkbookFormulaFiller {
  const workbook = captureWorkbookModel(input).snapshot().metadata.workbook;
  const tables = new Map(workbook.sheets.flatMap(sheet => (sheet.tables ?? []).map(table => {
    const columns = table.columns.map(column => column.name);
    return [table.name.toUpperCase(), { columns, indexes: new Map(columns.map((column, index) => [column.toUpperCase(), index])) }] as const;
  })));
  return (source, from, to) => {
    const ast = parseFormula(source, from); checkFormulaOrigin(to);
    if (from.row === to.row && from.column === to.column) return source;
    const delta = to.column - from.column;
    function visit(node: FormulaExpression): FormulaExpression {
      if (node.kind === "structured" && delta && node.table !== undefined && node.columns.length === 1) {
        const table = tables.get(node.table.toUpperCase()), index = table?.indexes.get(node.columns[0]!.toUpperCase());
        // Unresolved references retain their source error; filling must not
        // silently bind a misspelled table or column to a neighboring column.
        if (table && index !== undefined) return { ...node, columns: [table.columns[((index + delta) % table.columns.length + table.columns.length) % table.columns.length]!] };
      }
      if (node.kind === "binary") return { ...node, left: visit(node.left), right: visit(node.right) };
      if (node.kind === "unary") return { ...node, operand: visit(node.operand) };
      if (node.kind === "call") return { ...node, arguments: node.arguments.map(argument => argument === null ? null : visit(argument)) };
      return node;
    }
    const output = printFormula(delta ? visit(ast) : ast, to); parseFormula(output, to); return output;
  };
}
