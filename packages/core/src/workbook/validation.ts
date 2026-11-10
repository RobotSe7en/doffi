import type { DeepReadonly } from "../runtime.js";
import type { WorkbookCoordinate, WorkbookDataValidation } from "../contracts/workbook.js";
import type { WorkbookReferenceArea, WorkbookReferenceResult } from "./reference.js";
import type { WorkbookRuleValue, WorkbookRuleIssue, WorkbookValidationEvaluation } from "./rule-reader.js";
import type { ScalarValue } from "./scalar-values.js";
import { scalarCompare, scalarToNumber, scalarToText } from "./scalar-values.js";
import { parseFormula } from "./formula.js";
import { printFormula } from "./formula-print.js";
type Validation = DeepReadonly<WorkbookDataValidation>;
type WorkbookComparison = Extract<WorkbookDataValidation, { kind: "whole" | "decimal" | "date" | "time" | "textLength" }>["operator"];
export interface WorkbookListValidationPlan { readonly kind: "list"; readonly ruleId: string; readonly areas: readonly DeepReadonly<WorkbookReferenceArea>[]; readonly text: string }
export interface WorkbookValidationContext {
  evaluate(expression: string, rule: Validation, coordinate: WorkbookCoordinate): WorkbookRuleValue;
  resolve(expression: string, coordinate: WorkbookCoordinate): WorkbookReferenceResult;
  account(amount: number): void;
}
export function workbookRuleIssue(result: WorkbookRuleValue): WorkbookRuleIssue | undefined {
  return result.kind === "unavailable" ? { reason: result.reason, ...(result.functionName ? { functionName: result.functionName } : {}) }
    : result.value.type === "error" ? { reason: "formula-error" } : undefined;
}
function truth(value: ScalarValue): boolean { return value.type === "boolean" ? value.value : value.type === "number" && value.value !== 0; }
export function compareWorkbookRule(operator: WorkbookComparison, value: ScalarValue, bounds: readonly ScalarValue[]): boolean {
  const a = scalarCompare(value, bounds[0]!); if (typeof a !== "number") return false;
  if (operator === "between" || operator === "notBetween") {
    const b = scalarCompare(value, bounds[1]!); if (typeof b !== "number") return false;
    return operator === "between" ? a >= 0 && b <= 0 : a < 0 || b > 0;
  }
  return operator === "equal" ? a === 0 : operator === "notEqual" ? a !== 0 : operator === "lessThan" ? a < 0 : operator === "lessThanOrEqual" ? a <= 0 : operator === "greaterThan" ? a > 0 : a >= 0;
}

export function beginWorkbookValidation(rule: Validation, coordinate: WorkbookCoordinate, value: WorkbookRuleValue, context: WorkbookValidationContext): WorkbookValidationEvaluation | WorkbookListValidationPlan {
    const base = { ruleId: rule.ruleId };
    if (rule.kind === "none") return { ...base, state: "none" };
    if (value.kind === "unavailable") return { ...base, state: "unavailable", issue: workbookRuleIssue(value)! };
    if (rule.allowBlank && value.value.type === "blank") return { ...base, state: "valid" };
    if (value.value.type === "error") return { ...base, state: "invalid", issue: { reason: "formula-error" } };
    if (rule.kind === "custom") {
      const result = context.evaluate(rule.expression, rule, coordinate);
      return result.kind === "unavailable" ? { ...base, state: "unavailable", issue: workbookRuleIssue(result)! }
        : { ...base, state: truth(result.value) ? "valid" : "invalid", ...(workbookRuleIssue(result) ? { issue: workbookRuleIssue(result)! } : {}) };
    }
    if (rule.kind === "list") {
      const ast = parseFormula(rule.expression, rule.origin), text = scalarToText(value.value); if (typeof text !== "string") return { ...base, state: "invalid" };
      if (ast.kind === "string") { context.account(ast.value.length); return { ...base, state: ast.value.split(',').some(item => item.toUpperCase() === text.toUpperCase()) ? "valid" : "invalid" }; }
      const reference = context.resolve(printFormula(ast, coordinate), coordinate);
      if (reference.kind !== "ranges") return { ...base, state: "unavailable", issue: { reason: reference.kind === "unresolved" ? reference.reason : "invalid-list-reference" } };
      if (reference.areas.length !== 1 || reference.areas[0]!.from.row !== reference.areas[0]!.to.row && reference.areas[0]!.from.column !== reference.areas[0]!.to.column)
        return { ...base, state: "unavailable", issue: { reason: "invalid-list-reference" } };
      return { kind: "list", ruleId: rule.ruleId, areas: reference.areas, text };

    }
    const scalar = rule.kind === "textLength" ? scalarToText(value.value) : value.value.type === "number" ? value.value.value : undefined;
    if (scalar === undefined || typeof scalar === "object") return { ...base, state: "invalid", issue: { reason: "type-mismatch" } };
    const number = typeof scalar === "string" ? scalar.length : scalar;
    if (rule.kind === "whole" && !Number.isInteger(number)) return { ...base, state: "invalid", issue: { reason: "type-mismatch" } };
    const bounds: ScalarValue[] = [];
    for (const expression of rule.formulas) {
      const result = context.evaluate(expression, rule, coordinate), failed = workbookRuleIssue(result);
      if (failed) return { ...base, state: "unavailable", issue: failed };
      const value = scalarToNumber((result as Extract<WorkbookRuleValue, { kind: "value" }>).value);
      if (typeof value !== "number") return { ...base, state: "unavailable", issue: { reason: "non-numeric-threshold" } }; bounds.push({ type: "number", value });
    }
    return { ...base, state: compareWorkbookRule(rule.operator, { type: "number", value: number }, bounds) ? "valid" : "invalid" };
}
