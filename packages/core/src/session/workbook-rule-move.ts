import type { WorkbookConditionalRule, WorkbookDataValidation, WorkbookMetadata, WorkbookRange, WorkbookSheet } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import { DocumentError } from "../errors.js";
import { copyFormula } from "../workbook/formula-rewrite.js";
import { intersectWorkbookRange, subtractWorkbookRange } from "../workbook/range-mapping.js";
import { createWorkbookReferenceRewriter } from "./workbook-references.js";

type Rule = WorkbookConditionalRule | WorkbookDataValidation;
type Move = Extract<WorkbookOperation, { op: "moveWorkbookRange" }>;
type RuleLists = Pick<WorkbookSheet, "conditionalRules" | "dataValidations">;
function list(lists: RuleLists, key: keyof RuleLists): Rule[] { return lists[key] ?? (lists[key] = []); }
/** Prepare from original templates before any source reference is rewritten.
 * Each moved part gets its own origin in the original intersection; this keeps
 * relative templates precise even when the old origin lies outside the cut.
 * Source and destination coverage is subtracted as rectangles, never cells. */
export function planWorkbookRuleMove(workbook: WorkbookMetadata, operation: Move, destination: WorkbookRange,
  rewrite: ReturnType<typeof createWorkbookReferenceRewriter>): Map<WorkbookSheet, RuleLists> {
  const source = workbook.sheets.find(sheet => sheet.sheetId === operation.sheetId)!, target = workbook.sheets.find(sheet => sheet.sheetId === operation.targetSheetId)!;
  const plan = new Map<WorkbookSheet, RuleLists>(), used = new Set(workbook.sheets.flatMap(sheet => [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []].map(rule => rule.ruleId)));
  let serial = 0, fragments = 0;
  const allocate = (): string => { let id: string; do { id = "moved-rule-" + ++serial; } while (used.has(id)); used.add(id); return id; };
  function template(rule: Rule, sheet: WorkbookSheet, ranges: WorkbookRange[], moving: boolean): Rule {
    const result = JSON.parse(JSON.stringify(rule)) as Rule;
    result.ranges = ranges;
    const original = moving ? intersectWorkbookRange(rule.ranges.find(range => intersectWorkbookRange(range, operation.range))!, operation.range)!.from : rule.origin;
    const origin = moving ? { row: original.row + destination.from.row - operation.range.from.row, column: original.column + destination.from.column - operation.range.from.column } : original;
    const update = (expression: string): string => rewrite(copyFormula(expression, rule.origin, original), sheet.name, original, undefined,
      { sheet: moving ? target.name : sheet.name, ...origin });
    if ("expression" in result) result.expression = update(result.expression);
    if ("formulas" in result) result.formulas = result.formulas.map(update);
    if (result.kind === "colorScale") for (const stop of result.stops) if (stop.kind === "formula") stop.expression = update(stop.expression);
    result.origin = { ...origin }; return result;
  }
  for (const sheet of new Set([source, target])) {
    const lists: RuleLists = {}; plan.set(sheet, lists);
    for (const key of ["conditionalRules", "dataValidations"] as const) if (sheet[key] !== undefined) lists[key] = [];
  }
  for (const key of ["conditionalRules", "dataValidations"] as const) {
    const moved: Rule[] = [];
    for (const sheet of plan.keys()) for (const rule of sheet[key] ?? []) {
      const cut = sheet === source ? rule.ranges.flatMap(range => { const part = intersectWorkbookRange(range, operation.range); return part ? [part] : []; }) : [];
      const retained = rule.ranges.flatMap(range => sheet === source ? subtractWorkbookRange(range, operation.range) : [range])
        .flatMap(range => sheet === target ? subtractWorkbookRange(range, destination) : [range]);
      fragments += retained.length + cut.length;
      if (fragments > 4096) throw new DocumentError("ResourceLimitExceeded", "Cutting workbook rules exceeds the fragment budget.");
      if (retained.length) list(plan.get(sheet)!, key).push(template(rule, sheet, retained, false));
      if (cut.length) {
        const ranges = cut.map(range => ({ from: { row: range.from.row + destination.from.row - operation.range.from.row, column: range.from.column + destination.from.column - operation.range.from.column },
          to: { row: range.to.row + destination.from.row - operation.range.from.row, column: range.to.column + destination.from.column - operation.range.from.column } }));
        const item = template(rule, sheet, ranges, true); if (retained.length) item.ruleId = allocate(); moved.push(item);
      }
    }
    if (moved.length) {
      const destinationRules = list(plan.get(target)!, key);
      let priority = 0; for (const rule of destinationRules) if ("priority" in rule) priority = Math.max(priority, rule.priority);
      // Cross-sheet rules retain their original precedence amongst themselves,
      // after destination rules whose coverage was not overwritten.
      for (const item of moved.sort((a, b) => ("priority" in a ? a.priority : 0) - ("priority" in b ? b.priority : 0))) {
        if ("priority" in item) {
          if (priority === 2147483647) throw new DocumentError("ResourceLimitExceeded", "Cutting conditional rules exceeds the priority budget.");
          item.priority = ++priority;
        }
        destinationRules.push(item);
      }
    }
  }
  return plan;
}
