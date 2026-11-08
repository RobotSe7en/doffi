import type { WorkbookMetadata } from "../contracts/workbook.js";
import type { DeepReadonly } from "../runtime.js";
import { parseFormula, type FormulaExpression, type FormulaOrigin } from "./formula.js";

/** Source restoration compares native bindings and relative/absolute intent.
 * Redundant owner qualifiers must not manufacture a revision. The index belongs
 * to this preparation only, and never caches another document's private state. */
export function createWorkbookFormulaEquivalence(workbook: DeepReadonly<WorkbookMetadata>) {
  const sheets = new Map(workbook.sheets.map(sheet => [sheet.sheetId, sheet.name.toUpperCase()])), ids = new Map([...sheets].map(([id, name]) => [name, id]));
  const scopes = new Map<string | undefined, Map<string, string>>();
  for (const name of workbook.names ?? []) { const scope = scopes.get(name.sheetId) ?? new Map<string, string>(); scope.set(name.name.toUpperCase(), name.nameId); scopes.set(name.sheetId, scope); }
  return (left: string, right: string, origin: FormulaOrigin, sheetId?: string): boolean => {
    if (left === right) return true;
    const sheetName = sheets.get(sheetId ?? "") ?? "$workbook";
    function visit(node: FormulaExpression): unknown {
      if (node.kind === "cell" || node.kind === "axisRange") return { ...node, sheet: (node.sheet ?? sheetName).toUpperCase() };
      if (node.kind === "range") return { ...node, from: visit(node.from), to: visit(node.to) };
      if (node.kind === "name") {
        const owner = ids.get((node.sheet ?? sheetName).toUpperCase()), token = node.name.toUpperCase();
        const binding = (node.workbook === 0 || owner === undefined ? undefined : scopes.get(owner)?.get(token)) ?? (node.sheet === undefined ? scopes.get(undefined)?.get(token) : undefined);
        return binding === undefined ? node : { kind: "name", nameId: binding };
      }
      if (node.kind === "binary") return { ...node, left: visit(node.left), right: visit(node.right) };
      if (node.kind === "unary") return { ...node, operand: visit(node.operand) };
      if (node.kind === "call") return { ...node, arguments: node.arguments.map(argument => argument === null ? null : visit(argument)) };
      return node;
    }
    return JSON.stringify(visit(parseFormula(left, origin))) === JSON.stringify(visit(parseFormula(right, origin)));
  };
}
