import type { DeepReadonly, TransactionPreparationProgress, WorkbookModel } from "../runtime.js";
import type { WorkbookReferenceArea } from "../workbook/reference.js";
import type { WorkbookScalarContext } from "../workbook/scalar.js";
import { createWorkbookScalarKernel } from "../workbook/scalar.js";
import { WorkbookDependencyIndex } from "../workbook/dependency-index.js";
import { parseFormula } from "../workbook/formula.js";
import { workbookExpressionDependencies } from "../workbook/expression-dependencies.js";
import type { WorkbookInteractionBudget } from "./workbook-interaction-budget.js";
import { workbookContextIntersects } from "./workbook-interaction-scope.js";
import { workbookSource } from "../workbook/source.js";

/** Validation and conditional-format consumers can live outside source writes.
 * They depend on source values but never make those target cells writable. */
export function* workbookInteractionRuleImpact(model: WorkbookModel, impact: readonly DeepReadonly<WorkbookReferenceArea>[], context: WorkbookScalarContext,
  budget: WorkbookInteractionBudget, add: (area: DeepReadonly<WorkbookReferenceArea>) => void): Generator<TransactionPreparationProgress> {
  const index = new WorkbookDependencyIndex(impact.map(area => ({ area, owner: "impact" })), budget.step), workbook = model.snapshot().metadata.workbook;
  let references: DeepReadonly<WorkbookReferenceArea>[] = [];
  const collect = (area: DeepReadonly<WorkbookReferenceArea>): void => { budget.dependency(); references.push(area); };
  const kernel = createWorkbookScalarKernel(model, { maxCells: budget.limits.maxCells, maxSteps: budget.limits.maxSteps, maxCodeUnits: budget.limits.maxBytes }, context,
    { step: budget.step, account: budget.account, reference: collect, read: at => collect({ sheetId: at.sheetId, from: at, to: at }) });
  try {
    let completed = 0;
    for (const sheet of workbook.sheets) for (const rule of [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []]) {
      budget.step(); if (rule.kind === "none") continue;
      if (rule.kind === "colorScale") {
        if (rule.ranges.some(range => index.query({ sheetId: sheet.sheetId, ...range }, budget.step).next().done === false)) for (const range of rule.ranges) add({ sheetId: sheet.sheetId, ...range });
        continue;
      }
      const expressions = "expression" in rule ? [rule.expression] : rule.formulas, seen = new Set<number>();
      const parsed = expressions.map(expression => parseFormula(expression, rule.origin));
      function* targets(): Generator<{ row: number; column: number }> {
        for (const range of rule.ranges) {
          if ("allowBlank" in rule && rule.allowBlank) {
            for (const chunk of model.snapshot().manifest.chunks) {
              budget.step(); if (chunk.sheetId !== sheet.sheetId) continue;
              const row = Math.max(range.from.row, chunk.row), column = Math.max(range.from.column, chunk.column), rows = Math.min(range.to.row, chunk.row + 255) - row + 1, columns = Math.min(range.to.column, chunk.column + 31) - column + 1;
              if (rows > 0 && columns > 0) for (const cell of workbookSource(model)!.chunks.get(chunk.path)!.readWindow({ row, column, rows, columns })) yield cell;
            }
          } else for (let row = range.from.row; row <= range.to.row; row++) for (let column = range.from.column; column <= range.to.column; column++) yield { row, column };
        }
      }
      for (const { row, column } of targets()) {
        budget.step(); if (++completed % 128 === 0) { yield { phase: "validation", completed, total: 0 }; budget.step(); }
        const key = row * 16_384 + column; if (seen.has(key)) continue;
        budget.dependency(); seen.add(key); references = []; let subtotal = false;
        const at = { sheetId: sheet.sheetId, row, column };
        for (const ast of parsed) workbookExpressionDependencies(model, ast, at, budget.step, collect, name => { subtotal ||= name === "SUBTOTAL"; },
          (text, location) => { const bound = model.resolveReference(text, location); return bound.kind === "unresolved" ? kernel.referenceAreas(text, location) : bound; }, { includeFilterRanges: false });
        if (subtotal) {
          const inputs = [...references];
          for (const source of workbook.sheets) for (const filter of [...source.autoFilter ? [source.autoFilter] : [], ...source.tables?.flatMap(table => table.autoFilter ? [table.autoFilter] : []) ?? []]) {
            budget.step(); if (!inputs.some(area => area.sheetId === source.sheetId && workbookContextIntersects(area, filter.range))) continue;
            for (const field of filter.columns) if (field.criterion && filter.range.from.row < filter.range.to.row) collect({ sheetId: source.sheetId, from: { row: filter.range.from.row + 1, column: field.column }, to: { row: filter.range.to.row, column: field.column } });
          }
        }
        if (references.some(area => index.query(area, budget.step).next().done === false)) add({ sheetId: sheet.sheetId, from: at, to: at });
      }
    }
  } finally { kernel.dispose(); }
}
