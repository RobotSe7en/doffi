import type { DeepReadonly, TransactionPreparationProgress, WorkbookModel } from "../runtime.js";
import type { WorkbookInteractionRange } from "../contracts/workbook-interaction.js";
import type { WorkbookScalarContext } from "../workbook/scalar.js";
import { createWorkbookScalarKernel } from "../workbook/scalar.js";
import { WorkbookDependencyIndex } from "../workbook/dependency-index.js";
import type { WorkbookDependencyEntry } from "../workbook/dependency-index.js";
import { workbookExpressionDependencies } from "../workbook/expression-dependencies.js";
import { parseFormula } from "../workbook/formula.js";
import { printFormula } from "../workbook/formula-print.js";
import type { WorkbookReferenceArea, WorkbookReferenceLocation } from "../workbook/reference.js";
import { workbookSource } from "../workbook/source.js";
import type { WorkbookInteractionBudget } from "./workbook-interaction-budget.js";
import { workbookContextContains, workbookContextCovers, workbookContextIntersects } from "./workbook-interaction-scope.js";
import { workbookInteractionRuleImpact } from "./workbook-interaction-rule-impact.js";

export const workbookInteractionCellKey = (at: WorkbookReferenceLocation): string => JSON.stringify([at.sheetId, at.row, at.column]);
export interface WorkbookInteractionFormulaDependencies {
  readonly at: WorkbookReferenceLocation;
  readonly areas: readonly DeepReadonly<WorkbookReferenceArea>[];
  readonly nameIds: ReadonlySet<string>;
}
/** Source dependencies and reverse consumers share the calculation index. The
 * graph stays private: unrelated formulas, values and names are never emitted.
 * Rectangles remain compact, including full columns and empty native tables. */
export function* workbookInteractionDependencies(model: WorkbookModel, writable: readonly DeepReadonly<WorkbookInteractionRange>[], context: WorkbookScalarContext,
  budget: WorkbookInteractionBudget, addConsumer: (area: DeepReadonly<WorkbookReferenceArea>) => void): Generator<TransactionPreparationProgress, ReadonlyMap<string, WorkbookInteractionFormulaDependencies>> {
  const source = workbookSource(model)!, workbook = model.snapshot().metadata.workbook;
  const entries: WorkbookDependencyEntry[] = [], records = new Map<string, WorkbookInteractionFormulaDependencies>();
  const editableTables = workbook.sheets.flatMap(sheet => (sheet.tables ?? []).filter(table => workbookContextCovers(writable, { sheetId: sheet.sheetId, range: table.range }, budget.step)).map(table => ({ sheetId: sheet.sheetId, table })));
  let record: { at: WorkbookReferenceLocation; areas: DeepReadonly<WorkbookReferenceArea>[]; nameIds: Set<string>; seen: Set<string> };
  const add = (area: DeepReadonly<WorkbookReferenceArea>): void => {
    budget.step(); const key = JSON.stringify([area.sheetId, area.from.row, area.from.column, area.to.row, area.to.column]);
    if (record.seen.has(key)) return; record.seen.add(key); budget.dependency(); record.areas.push(area);
  };
  const kernel = createWorkbookScalarKernel(model, { maxCells: budget.limits.maxCells, maxSteps: budget.limits.maxSteps, maxCodeUnits: budget.limits.maxBytes }, context,
    { step: budget.step, account: budget.account, reference: add, read: at => add({ sheetId: at.sheetId, from: at, to: at }), name: definition => record.nameIds.add(definition.nameId) });
  try {
    let completed = 0;
    for (const descriptor of model.snapshot().manifest.chunks) {
      budget.step();
      for (const formula of source.chunks.get(descriptor.path)!.formulaLocations()) {
        budget.dependency(); const at = { sheetId: descriptor.sheetId, row: formula.row, column: formula.column }, owner = workbookInteractionCellKey(at);
        record = { at, areas: [], nameIds: new Set(), seen: new Set() };
        const pending = [{ ast: formula.expression, at }], names = new Set<string>();
        workbookExpressionDependencies(model, formula.expression, at, budget.step, add, undefined, (text, location) => {
          const bound = model.resolveReference(text, location); return bound.kind === "unresolved" ? kernel.referenceAreas(text, location) : bound;
        }, {
          includeFilterRanges: false, onName: (definition, location) => {
            record.nameIds.add(definition.nameId); const key = JSON.stringify([definition.nameId, location.sheetId, location.row, location.column]);
            if (!names.has(key)) { names.add(key); pending.push({ ast: parseFormula(definition.expression, definition.origin), at: { ...location, sheetId: definition.sheetId ?? location.sheetId } }); }
          },
        });
        while (pending.length) {
          budget.step(); const { ast, at: location } = pending.pop()!;
          if (ast.kind === "call") {
            const name = ast.name.toUpperCase().replace(/^_XLFN\./, "");
            if (["OFFSET", "INDIRECT", "INDEX"].includes(name)) {
              const resolved = kernel.referenceAreas(printFormula(ast, location), location);
              if (resolved.kind === "ranges") for (const area of resolved.areas) add(area);
            }
            if (name === "SUBTOTAL") {
              const areas: DeepReadonly<WorkbookReferenceArea>[] = [];
              for (const argument of ast.arguments.slice(1)) if (argument) workbookExpressionDependencies(model, argument, location, budget.step, area => areas.push(area), undefined,
                (text, origin) => kernel.referenceAreas(text, origin), { includeFilterRanges: false });
              for (const sheet of workbook.sheets) for (const filter of [...sheet.autoFilter ? [sheet.autoFilter] : [], ...sheet.tables?.flatMap(table => table.autoFilter ? [table.autoFilter] : []) ?? []]) {
                budget.step(); if (!areas.some(area => area.sheetId === sheet.sheetId && workbookContextIntersects(area, filter.range))) continue;
                for (const column of filter.columns) if (column.criterion && filter.range.from.row < filter.range.to.row) add({ sheetId: sheet.sheetId, from: { row: filter.range.from.row + 1, column: column.column }, to: { row: filter.range.to.row, column: column.column } });
              }
            }
            for (const argument of ast.arguments) if (argument) pending.push({ ast: argument, at: location });
          } else if (ast.kind === "unary") pending.push({ ast: ast.operand, at: location });
          else if (ast.kind === "binary") pending.push({ ast: ast.left, at: location }, { ast: ast.right, at: location });
          else if (ast.kind === "structured") for (const entry of editableTables) {
            budget.step(); const matches = ast.table ? entry.table.name.toUpperCase() === ast.table.toUpperCase() : entry.sheetId === location.sheetId && workbookContextContains([{ sheetId: entry.sheetId, range: entry.table.range }], location.sheetId, location.row, location.column, budget.step);
            if (matches) add({ sheetId: entry.sheetId, from: entry.table.range.from, to: entry.table.range.to });
          }
        }
        records.set(owner, { at, areas: record.areas, nameIds: record.nameIds });
        for (const area of record.areas) entries.push({ owner, area });
        if (++completed % 128 === 0) { yield { phase: "cells", completed, total: 0 }; budget.step(); }
      }
    }
    const index = new WorkbookDependencyIndex(entries, budget.step), impact = writable.map(entry => ({ sheetId: entry.sheetId, ...entry.range })), visited = new Set<string>();
    for (let cursor = 0; cursor < impact.length; cursor++) {
      budget.step(); for (const owner of index.query(impact[cursor]!, budget.step)) {
        if (visited.has(owner)) continue; visited.add(owner); const at = records.get(owner)!.at, area = { sheetId: at.sheetId, from: at, to: at };
        addConsumer(area); impact.push(area);
      }
      if (cursor % 128 === 127) { yield { phase: "cells", completed: cursor + 1, total: 0 }; budget.step(); }
    }
    yield* workbookInteractionRuleImpact(model, impact, context, budget, addConsumer);
    return records;
  } finally { kernel.dispose(); }
}
