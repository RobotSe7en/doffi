import type { DeepReadonly, TransactionPrepareOptions, WorkbookModel, WorkbookRangeChange } from "../runtime.js";
import type { WorkbookDataValidation } from "../contracts/workbook.js";
import type { WorkbookReferenceLocation } from "../workbook/reference.js";
import type { WorkbookRuleValue, WorkbookValidationEvaluation, WorkbookRuleIssue } from "../workbook/rule-reader.js";
import { DEFAULT_WORKBOOK_RULE_LIMITS } from "../workbook/rule-reader.js";
import { createWorkbookScalarKernel } from "../workbook/scalar.js";
import { createWorkbookCalculator } from "../workbook/calculation.js";
import { beginWorkbookValidation, workbookRuleIssue } from "../workbook/validation.js";
import { scalarToText } from "../workbook/scalar-values.js";
import { checkLimit } from "../document/validation.js";
import { DocumentError } from "../errors.js";
import { rangeContains } from "./workbook-cells.js";
import { workbookExpressionDependencies } from "../workbook/expression-dependencies.js";
import { parseFormula } from "../workbook/formula.js";
import { WorkbookDependencyIndex } from "../workbook/dependency-index.js";
import { workbookExpressionIdentity } from "../workbook/expression-identity.js";
import type { WorkbookReferenceArea } from "../workbook/reference.js";
import type { WorkbookScalarKernel } from "../workbook/scalar.js";
import type { WorkbookCalculationResult } from "../workbook/calculation.js";
import type { WorkbookPreparation, WorkbookPreparationStep } from "./workbook-preparation.js";

/** All cell-writing entry points validate the final candidate through this
 * shared rule implementation. Unsupported stop conditions never mean pass. */
export function* validateWorkbookCandidateSteps(previous: WorkbookModel, model: WorkbookModel, sourceRanges: readonly WorkbookRangeChange[], options?: TransactionPrepareOptions): WorkbookPreparation<true> {
  let visited = 0;
  function* checkpoint(): Generator<WorkbookPreparationStep, void, WorkbookCalculationResult | undefined> {
    if (++visited % 256 === 0) yield { kind: "progress", value: { phase: "validation", completed: visited, total: 0 } };
  }
  const snapshot = model.snapshot(), rules = new Map(snapshot.metadata.workbook.sheets.map(sheet => [sheet.sheetId, sheet.dataValidations ?? []]));
  const prior = previous.snapshot(), previousRules = new Map(prior.metadata.workbook.sheets.flatMap(sheet => sheet.dataValidations?.map(rule => [rule.ruleId, rule] as const) ?? []));
  const namesIdentity = (source: WorkbookModel): string => JSON.stringify(source.snapshot().metadata.workbook.names?.map(name => ({ ...name, expression: workbookExpressionIdentity(source, name.expression, name.origin) })));
  const ruleIdentity = (source: WorkbookModel, rule: DeepReadonly<WorkbookDataValidation> | undefined): string | undefined => rule && JSON.stringify({ ...rule,
    ...("expression" in rule ? { expression: workbookExpressionIdentity(source, rule.expression, rule.origin) } : {}),
    ...("formulas" in rule ? { formulas: rule.formulas.map(expression => workbookExpressionIdentity(source, expression, rule.origin)) } : {}),
  });
  const contextChanged = namesIdentity(previous) !== namesIdentity(model) || prior.metadata.locale !== snapshot.metadata.locale;
  const changedRules = new Set([...rules.values()].flatMap(items => items.filter(rule => contextChanged || ruleIdentity(previous, previousRules.get(rule.ruleId)) !== ruleIdentity(model, rule)).map(rule => rule.ruleId)));
  const visibilityChanges = snapshot.metadata.workbook.sheets.filter(sheet => {
    const old = prior.metadata.workbook.sheets.find(item => item.sheetId === sheet.sheetId);
    const identity = (item: typeof old): string => JSON.stringify(item && [item.defaultRowsHidden ?? false, item.rowDimensions.filter(span => span.hidden !== undefined).map(span => [span.from, span.to, span.hidden]), item.autoFilter, item.tables?.map(table => table.autoFilter)]);
    return identity(old) !== identity(sheet);
  });
  if (![...rules.values()].some(items => items.length)) return true;
  const previousSheets = new Set(prior.metadata.workbook.sheets.map(sheet => sheet.sheetId)), nextSheets = new Set(rules.keys());
  const changes: WorkbookRangeChange[] = [];
  for (const change of sourceRanges) for (let row = change.range.from.row; row <= change.range.to.row; row++) for (let column = change.range.from.column; column <= change.range.to.column; column++) {
    const before = previousSheets.has(change.sheetId) ? previous.getCell(change.sheetId, row, column) : undefined, after = nextSheets.has(change.sheetId) ? model.getCell(change.sheetId, row, column) : undefined;
    const signature = (source: WorkbookModel, cell: typeof before): string => cell?.formula ? workbookExpressionIdentity(source, cell.formula.expression, cell) : JSON.stringify(cell?.value ?? { type: "blank" });
    if (signature(previous, before) !== signature(model, after)) changes.push({ sheetId: change.sheetId, range: { from: { row, column }, to: { row, column } } });
    yield* checkpoint();
  }
  if (!changes.length && !changedRules.size && !visibilityChanges.length) return true;
  const limits = DEFAULT_WORKBOOK_RULE_LIMITS;
  let steps = 0, units = 0, cells = 0;
  const step = (): void => checkLimit(++steps, limits.maxSteps);
  const account = (amount: number): void => { units += amount; checkLimit(units, limits.maxCodeUnits); };
  let currentModel = previous;
  const calculator = createWorkbookCalculator(() => currentModel, { maxSteps: limits.maxSteps, maxReadCells: limits.maxValues, maxCodeUnits: limits.maxCodeUnits });
  let disposableEvaluator: WorkbookScalarKernel | undefined, previousEvaluator: WorkbookScalarKernel | undefined;
  try {
    const beforeCalculation = (yield { kind: "calculation", calculator, ...(options?.calculationContext ? { context: options.calculationContext } : {}) })!;
    steps += beforeCalculation.steps; account(beforeCalculation.codeUnits); let readCells = beforeCalculation.readCells;
    currentModel = model;
    const calculation = (yield { kind: "calculation", calculator, ...(options?.calculationContext ? { context: options.calculationContext } : {}) })!;
    steps += calculation.steps; checkLimit(steps, limits.maxSteps); account(calculation.codeUnits); readCells += calculation.readCells; checkLimit(readCells, limits.maxValues);
    const cellKey = (at: WorkbookReferenceLocation): string => JSON.stringify([at.sheetId, at.row, at.column]);
    const beforeValues = new Map(beforeCalculation.cells.map(cell => [cellKey(cell), cell.state === "valid" ? JSON.stringify(cell.value) : cell.state === "unavailable" ? JSON.stringify(cell.issue) : "stale"]));
    const derived = calculation.cells.filter(cell => beforeValues.get(cellKey(cell)) !== (cell.state === "valid" ? JSON.stringify(cell.value) : cell.state === "unavailable" ? JSON.stringify(cell.issue) : "stale"));
    const affected = new WorkbookDependencyIndex([...changes.map(change => ({ area: { sheetId: change.sheetId, ...change.range }, owner: "source" })), ...derived.map(cell => ({ area: { sheetId: cell.sheetId, from: cell, to: cell }, owner: "calculated" }))], step);
    const visibilityRanges: WorkbookRangeChange[] = [];
    for (const sheet of visibilityChanges) {
      const old = prior.metadata.workbook.sheets.find(item => item.sheetId === sheet.sheetId);
      const left = old?.rowDimensions.filter(span => span.hidden !== undefined) ?? [], right = sheet.rowDimensions.filter(span => span.hidden !== undefined);
      const boundaries = [...new Set([0, 1_048_576, ...[...left, ...right].flatMap(span => [span.from, span.to + 1])])].sort((a, b) => { step(); return a - b; });
      let a = 0, b = 0;
      const add = (from: number, to: number): void => { visibilityRanges.push({ sheetId: sheet.sheetId, range: { from: { row: from, column: 0 }, to: { row: to, column: 16_383 } } }); };
      for (let index = 0; index + 1 < boundaries.length; index++) {
        step(); const from = boundaries[index]!, to = boundaries[index + 1]! - 1;
        while (left[a] && left[a]!.to < from) { step(); a++; } while (right[b] && right[b]!.to < from) { step(); b++; }
        const hiddenBefore = left[a] && left[a]!.from <= from ? left[a]!.hidden : old?.defaultRowsHidden ?? false;
        const hiddenAfter = right[b] && right[b]!.from <= from ? right[b]!.hidden : sheet.defaultRowsHidden ?? false;
        if (hiddenBefore !== hiddenAfter) add(from, to);
        yield* checkpoint();
      }
      const filters = (source: typeof old) => [...source?.autoFilter ? [source.autoFilter] : [], ...source?.tables?.flatMap(table => table.autoFilter ? [table.autoFilter] : []) ?? []];
      const beforeFilters = filters(old), afterFilters = filters(sheet);
      if (JSON.stringify(beforeFilters) !== JSON.stringify(afterFilters)) for (const filter of [...beforeFilters, ...afterFilters]) {
        step(); if (filter.range.from.row < filter.range.to.row) add(filter.range.from.row + 1, filter.range.to.row);
      }
    }
    const visibility = new WorkbookDependencyIndex(visibilityRanges.map(change => ({ area: { sheetId: change.sheetId, ...change.range }, owner: "visibility" })), step);
    let referenceRead: ((area: DeepReadonly<WorkbookReferenceArea>) => void) | undefined;
    const evaluator = createWorkbookScalarKernel(model, { maxSteps: limits.maxSteps, maxCells: limits.maxValues, maxCodeUnits: limits.maxCodeUnits }, options?.calculationContext, {
      step, account, read: location => { checkLimit(++readCells, limits.maxValues); referenceRead?.({ sheetId: location.sheetId, from: location, to: location }); },
      reference: area => referenceRead?.(area),
      cachedCell: location => {
        const value = calculator.getCell(location); return value.state === "valid" ? value.value : value.state === "unavailable" ? value.issue : undefined;
      },
    });
    disposableEvaluator = evaluator;
    function originalEvaluator(): WorkbookScalarKernel {
      if (previousEvaluator) return previousEvaluator;
      const values = new Map(beforeCalculation.cells.map(cell => [cellKey(cell), cell]));
      return previousEvaluator = createWorkbookScalarKernel(previous, { maxSteps: limits.maxSteps, maxCells: limits.maxValues, maxCodeUnits: limits.maxCodeUnits }, options?.calculationContext, {
        step, account, read: () => checkLimit(++readCells, limits.maxValues), cachedCell: location => {
          const value = values.get(cellKey(location)); return value?.state === "valid" ? value.value : value?.state === "unavailable" ? value.issue : undefined;
        },
      });
    }
    function reference(expression: string, at: WorkbookReferenceLocation, collect: (area: DeepReadonly<WorkbookReferenceArea>) => void) {
      const outer = referenceRead; referenceRead = collect;
      try { return evaluator.referenceAreas(expression, at); } finally { referenceRead = outer; }
    }
    function current(location: WorkbookReferenceLocation, kernel = evaluator): WorkbookRuleValue {
      step(); const value = kernel.cell(location); return value.kind === "value" ? { kind: "value", value: value.value } : { kind: "unavailable", reason: value.reason, ...(value.functionName ? { functionName: value.functionName } : {}) };
    }
    function* validate(rule: DeepReadonly<WorkbookDataValidation>, location: WorkbookReferenceLocation, source = model, kernel = evaluator): Generator<WorkbookPreparationStep, WorkbookValidationEvaluation, WorkbookCalculationResult | undefined> {
      const plan = beginWorkbookValidation(rule, location, current(location, kernel), { account,
        resolve: (expression, at) => {
          const origin = { sheetId: location.sheetId, ...at }, value = source.resolveReference(expression, origin);
          return value.kind === "unresolved" ? kernel.referenceAreas(expression, origin) : value;
        },
        evaluate: (expression, definition, coordinate) => {
          const result = kernel.evaluate(expression, { sheetId: location.sheetId, ...definition.origin }, { at: { sheetId: location.sheetId, ...coordinate } });
          return result.kind === "value" ? { kind: "value", value: result.value } : { kind: "unavailable", reason: result.reason, ...(result.functionName ? { functionName: result.functionName } : {}) };
        },
      });
      if (!("kind" in plan)) return plan;
      let unknown: WorkbookRuleIssue | undefined, count = 0;
      for (const area of plan.areas) for (const chunk of source.snapshot().manifest.chunks) {
        step(); if (chunk.sheetId !== area.sheetId) continue;
        const row = Math.max(area.from.row, chunk.row), lastRow = Math.min(area.to.row, chunk.row + 255), column = Math.max(area.from.column, chunk.column), lastColumn = Math.min(area.to.column, chunk.column + 31);
        if (row > lastRow || column > lastColumn) continue;
        const columns = lastColumn - column + 1, batchRows = Math.max(1, Math.floor(4096 / columns));
        for (let start = row; start <= lastRow; start += batchRows) for (const cell of source.readWindow({ sheetId: area.sheetId, row: start, column, rows: Math.min(batchRows, lastRow - start + 1), columns })) {
          yield* checkpoint();
          step(); checkLimit(++count, limits.maxValues); const value = current({ sheetId: area.sheetId, row: cell.row, column: cell.column }, kernel);
          if (value.kind === "unavailable") { unknown = workbookRuleIssue(value); continue; }
          const text = scalarToText(value.value); if (typeof text === "string" && text.toUpperCase() === plan.text.toUpperCase()) return { ruleId: plan.ruleId, state: "valid" };
        }
      }
      const area = plan.areas[0]!, total = (area.to.row - area.from.row + 1) * (area.to.column - area.from.column + 1);
      if (plan.text === "" && count < total) return { ruleId: plan.ruleId, state: "valid" };
      return unknown ? { ruleId: plan.ruleId, state: "unavailable", issue: unknown } : { ruleId: plan.ruleId, state: "invalid" };
    }
    const targets = new Set<string>();
    function relevant(rule: DeepReadonly<WorkbookDataValidation>, location: WorkbookReferenceLocation): boolean {
      if (rule.kind === "none") return false;
      if (changedRules.has(rule.ruleId) || affected.query({ sheetId: location.sheetId, from: location, to: location }, step).next().done === false) return true;
      let match = false, subtotal = false; const dependencies: DeepReadonly<WorkbookReferenceArea>[] = [];
      const collect = (area: DeepReadonly<WorkbookReferenceArea>): void => { dependencies.push(area); if (affected.query(area, step).next().done === false) match = true; };
      const expressions = "expression" in rule ? [rule.expression] : rule.formulas;
      for (const expression of expressions) workbookExpressionDependencies(model, parseFormula(expression, rule.origin), location, step, collect,
        name => { subtotal ||= name === "SUBTOTAL"; }, (expression, at) => reference(expression, at, collect));
      if (subtotal && dependencies.some(area => visibility.query(area, step).next().done === false)) match = true;
      return match;
    }
    function* add(location: WorkbookReferenceLocation): Generator<WorkbookPreparationStep, void, WorkbookCalculationResult | undefined> {
      const rule = rules.get(location.sheetId)?.find(rule => rule.ranges.some(range => rangeContains(range, location.row, location.column))); if (!rule) return;
      const key = cellKey(location); if (targets.has(key)) return;
      if (!relevant(rule, location)) return;
      checkLimit(++cells, limits.maxCells); targets.add(key);
      const result = yield* validate(rule, location);
      if (result.state === "valid" || result.state === "none") return;
      // A view-only metadata edit may preserve an imported violation. It must
      // still reject a previously valid target becoming invalid/unavailable.
      if (!changes.length && !changedRules.size && result.state === "invalid") {
        const oldRule = previousRules.get(rule.ruleId);
        if (oldRule && (yield* validate(oldRule, location, previous, originalEvaluator())).state === "invalid") return;
      }
      if (result.state === "invalid" && rule.errorStyle !== "stop" && options?.validation?.acceptedRuleIds.includes(rule.ruleId)) return;
      throw new DocumentError("InvalidOperation", result.state === "unavailable" ? "The candidate's data validation cannot be evaluated." : "The candidate does not satisfy its data validation.", [
        { code: rule.errorStyle === "stop" || result.state === "unavailable" ? "validation-blocked" : "validation-confirmation-required", path: [location.sheetId, location.row, location.column, rule.ruleId], message: rule.error || result.issue?.reason || "The cell does not satisfy the rule." },
      ]);
    }
    // Explicit clears must include now-absent cells. Other implicit blanks are
    // skipped only when the definition itself explicitly permits blank values.
    for (const change of changes) for (let row = change.range.from.row; row <= change.range.to.row; row++) for (let column = change.range.from.column; column <= change.range.to.column; column++) { step(); yield* add({ sheetId: change.sheetId, row, column }); yield* checkpoint(); }
    for (const chunk of snapshot.manifest.chunks) {
      if (!rules.get(chunk.sheetId)?.length) continue;
      for (let row = chunk.row; row < chunk.row + 256; row += 128) for (const cell of model.readWindow({ sheetId: chunk.sheetId, row, column: chunk.column, rows: 128, columns: 32 })) { step(); yield* add({ sheetId: chunk.sheetId, row: cell.row, column: cell.column }); yield* checkpoint(); }
    }
    for (const [sheetId, definitions] of rules) for (const rule of definitions) {
      checkLimit(definitions.length, limits.maxRules);
      if (rule.allowBlank || rule.kind === "none") continue;
      let references = false;
      const expressions = "expression" in rule ? [rule.expression] : rule.formulas;
      for (const expression of expressions) workbookExpressionDependencies(model, parseFormula(expression, rule.origin), { sheetId, ...rule.origin }, step, () => { references = true; });
      if (!changedRules.has(rule.ruleId) && !references) continue;
      for (const range of rule.ranges) {
        // Visit implicit targets without materializing their source cells.
        // Every actual target is checked, including nonlinear relative names;
        // the cumulative step budget bounds scans that cannot be completed.
        for (let row = range.from.row; row <= range.to.row; row++) for (let column = range.from.column; column <= range.to.column; column++) { step(); yield* add({ sheetId, row, column }); yield* checkpoint(); }
      }
    }
    return true;
  } finally { disposableEvaluator?.dispose(); previousEvaluator?.dispose(); calculator.dispose(); }
}
