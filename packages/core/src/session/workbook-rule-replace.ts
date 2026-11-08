import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { WorkbookConditionalRule, WorkbookDataValidation, WorkbookRange, WorkbookSheet } from "../contracts/workbook.js";
import { checkLimit } from "../document/validation.js";
import { DocumentError } from "../errors.js";
import { parseFormula } from "../workbook/formula.js";
import { intersectWorkbookRange, subtractWorkbookRange } from "../workbook/range-mapping.js";
import { checkNonOverlappingRanges, checkRange, sameWorkbookRangeCoverage } from "../workbook/ranges.js";

type Rule = WorkbookConditionalRule | WorkbookDataValidation;
function template(rule: Rule): string {
  const value: Record<string, unknown> = { ...rule }; delete value.ruleId; delete value.ranges; delete value.origin; delete value.priority;
  if ("expression" in rule) value.expression = parseFormula(rule.expression, rule.origin);
  if ("formulas" in rule) value.formulas = rule.formulas.map(expression => parseFormula(expression, rule.origin));
  if (rule.kind === "colorScale") value.stops = rule.stops.map(stop => stop.kind === "formula" ? { ...stop, expression: parseFormula(stop.expression, rule.origin) } : stop);
  function canonical(input: unknown): unknown {
    if (Array.isArray(input)) return input.map(canonical);
    if (input && typeof input === "object") return Object.fromEntries(Object.entries(input).sort(([a], [b]) => a.localeCompare(b)).map(([key, value]) => [key, canonical(value)]));
    return input;
  }
  return JSON.stringify(canonical(value));
}
/** Replace a sparse paste footprint, clipping current candidate metadata rather
 * than trusting an editor's stale copy of the surrounding rule definitions. */
export function* replaceWorkbookRules(sheet: WorkbookSheet, operation: Extract<WorkbookOperation, { op: "replaceWorkbookRules" }>, maxItems: number): Generator<{ completed: number; total: number }, void> {
  if (operation.conditionalRules === undefined && operation.dataValidations === undefined) throw new DocumentError("InvalidOperation", "Rule replacement requires at least one rule category.");
  for (const range of operation.ranges) checkRange(range);
  checkNonOverlappingRanges(operation.ranges);
  const fragmentLimit = Math.min(4096, maxItems); let completed = 0;
  function* tick(): Generator<{ completed: number; total: number }, void> { if (++completed % 256 === 0) yield { completed, total: completed }; }
  function* outside(ranges: WorkbookRange[]): Generator<{ completed: number; total: number }, WorkbookRange[]> {
    let remaining = ranges;
    for (const cover of operation.ranges) {
      const next: WorkbookRange[] = [];
      for (const range of remaining) {
        next.push(...(intersectWorkbookRange(range, cover) ? subtractWorkbookRange(range, cover) : [range]));
        checkLimit(next.length, fragmentLimit); yield* tick();
      }
      remaining = next; if (!remaining.length) break;
    }
    return remaining;
  }
  for (const key of ["conditionalRules", "dataValidations"] as const) {
    const incoming = operation[key]; if (incoming === undefined) continue;
    checkLimit(incoming.length, maxItems);
    for (const rule of incoming) {
      for (const range of rule.ranges) checkRange(range);
      if ((yield* outside(rule.ranges)).length) throw new DocumentError("InvalidOperation", "Replacement rules must stay inside their declared ranges.");
      yield* tick();
    }
    const previous: Rule[] = sheet[key] ?? [], retained: Rule[] = [], removed: Rule[] = [];
    let fragments = 0, partial = false;
    for (const rule of previous) {
      const ranges = yield* outside(rule.ranges); fragments += ranges.length; checkLimit(fragments, fragmentLimit);
      if (!ranges.length) removed.push(rule);
      else if (ranges.length === rule.ranges.length && ranges.every((range, index) => range === rule.ranges[index])) retained.push(rule);
      else { retained.push({ ...rule, ranges }); partial = true; }
      yield* tick();
    }
    // A repeated paste retains original identities, priorities and exact source
    // bytes when both ordered templates and their coverage already agree.
    const order = (rules: readonly Rule[]) => [...rules].sort((a, b) => ("priority" in a ? a.priority : 0) - ("priority" in b ? b.priority : 0));
    if (!partial && removed.length === incoming.length) {
      const left = order(removed), right = order(incoming); let same = true;
      for (let index = 0; index < left.length; index++) {
        const a = left[index]!, b = right[index]!; same &&= template(a) === template(b) && sameWorkbookRangeCoverage(a.ranges, b.ranges); yield* tick();
      }
      if (same) continue;
    }
    for (const rule of incoming) { fragments += rule.ranges.length; checkLimit(fragments, fragmentLimit); retained.push(rule); yield* tick(); }
    checkLimit(retained.length, maxItems);
    if (key === "conditionalRules") sheet.conditionalRules = retained as WorkbookConditionalRule[];
    else sheet.dataValidations = retained as WorkbookDataValidation[];
  }
}
