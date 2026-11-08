import type { DeepReadonly, TransactionPreparationProgress, WorkbookModel } from "../runtime.js";
import type { WorkbookInteractionRange } from "../contracts/workbook-interaction.js";
import type { WorkbookReferenceLocation } from "../workbook/reference.js";
import { parseFormula } from "../workbook/formula.js";
import { workbookExpressionDependencies } from "../workbook/expression-dependencies.js";
import type { WorkbookScalarKernel } from "../workbook/scalar.js";
import type { WorkbookInteractionBudget } from "./workbook-interaction-budget.js";
import { workbookContextIntersects } from "./workbook-interaction-scope.js";

/** Rule formulas use the rule's source origin and each affected coordinate.
 * Blank coordinates matter too. Work is explicit and bounded; no sampling or
 * corner approximation can silently miss an INDEX-dependent interior value. */
export function createWorkbookInteractionRuleReader(model: WorkbookModel, budget: WorkbookInteractionBudget,
  kernel: WorkbookScalarKernel, discover: (expression: string, origin: WorkbookReferenceLocation, at?: WorkbookReferenceLocation) => void,
  add: (area: { readonly sheetId: string; readonly from: { readonly row: number; readonly column: number }; readonly to: { readonly row: number; readonly column: number } }) => void) {
  const seen = new Set<string>(), workbook = model.snapshot().metadata.workbook;
  function* read(entry: DeepReadonly<WorkbookInteractionRange>): Generator<TransactionPreparationProgress> {
    const sheet = workbook.sheets.find(sheet => sheet.sheetId === entry.sheetId)!; let completed = 0;
    for (const rule of [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []]) {
      budget.step();
      if (rule.kind === "colorScale") {
        const key = JSON.stringify([sheet.sheetId, rule.ruleId, "scale"]);
        if (!seen.has(key) && rule.ranges.some(range => workbookContextIntersects(range, entry.range))) {
          seen.add(key); budget.dependency(); for (const range of rule.ranges) add({ sheetId: sheet.sheetId, ...range });
          const origin = { sheetId: sheet.sheetId, ...rule.origin };
          for (const stop of rule.stops) if (stop.kind === "formula") { discover(stop.expression, origin); kernel.evaluate(stop.expression, origin); }
        }
        continue;
      }
      const expressions = "expression" in rule ? [rule.expression] : "formulas" in rule ? rule.formulas : [];
      if (!expressions.length) continue;
      for (const range of rule.ranges) {
        budget.step(); if (!workbookContextIntersects(range, entry.range)) continue;
        for (let row = Math.max(range.from.row, entry.range.from.row); row <= Math.min(range.to.row, entry.range.to.row); row++)
          for (let column = Math.max(range.from.column, entry.range.from.column); column <= Math.min(range.to.column, entry.range.to.column); column++) {
            budget.step(); if (++completed % 128 === 0) { yield { phase: "cells", completed, total: 0 }; budget.step(); }
            const key = JSON.stringify([sheet.sheetId, rule.ruleId, row, column]); if (seen.has(key)) continue;
            seen.add(key); budget.dependency(); const at = { sheetId: sheet.sheetId, row, column }, origin = { sheetId: sheet.sheetId, ...rule.origin };
            for (const expression of expressions) { discover(expression, origin, at); kernel.evaluate(expression, origin, { at }); }
          }
      }
    }
    for (const table of sheet.tables ?? []) {
      budget.step(); if (!workbookContextIntersects(table.range, entry.range)) continue;
      for (let offset = 0; offset < table.columns.length; offset++) {
        budget.step(); const column = table.range.from.column + offset; if (column < entry.range.from.column || column > entry.range.to.column) continue;
        const definition = table.columns[offset]!;
        const regions = [
          ...(definition.calculatedFormula ? [{ expression: definition.calculatedFormula, first: table.range.from.row + Number(table.header), last: table.range.to.row - Number(table.totals), kind: "data" }] : []),
          ...(definition.totalFormula && table.totals ? [{ expression: definition.totalFormula, first: table.range.to.row, last: table.range.to.row, kind: "total" }] : []),
        ];
        for (const region of regions) for (let row = Math.max(region.first, entry.range.from.row); row <= Math.min(region.last, entry.range.to.row); row++) {
          budget.step(); if (++completed % 128 === 0) { yield { phase: "cells", completed, total: 0 }; budget.step(); }
          const key = JSON.stringify([sheet.sheetId, table.tableId, region.kind, row, column]); if (seen.has(key)) continue;
          seen.add(key); budget.dependency(); const origin = { sheetId: sheet.sheetId, row: region.first, column }, at = { sheetId: sheet.sheetId, row, column };
          discover(region.expression, origin, at); kernel.evaluate(region.expression, origin, { at });
        }
      }
    }
  }
  return read;
}

/** Rebase parsing at the source origin, then resolve at the actual coordinate. */
export function workbookInteractionExpression(model: WorkbookModel, expression: string, origin: WorkbookReferenceLocation, at: WorkbookReferenceLocation,
  step: () => void, add: Parameters<typeof workbookExpressionDependencies>[4], onName: NonNullable<Parameters<typeof workbookExpressionDependencies>[7]>["onName"], kernel?: WorkbookScalarKernel): void {
  workbookExpressionDependencies(model, parseFormula(expression, origin), at, step, add, undefined,
    (text, location) => { const bound = model.resolveReference(text, location); return bound.kind === "unresolved" && kernel ? kernel.referenceAreas(text, location) : bound; }, { includeFilterRanges: false, ...(onName ? { onName } : {}) });
}
