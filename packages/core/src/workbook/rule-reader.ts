import { beginWorkbookValidation, workbookRuleIssue as issue, compareWorkbookRule as compare } from "./validation.js";
import { DocumentError } from "../errors.js";
import type { DeepReadonly, WorkbookModel } from "../runtime.js";
import type { WorkbookConditionalRule, WorkbookCoordinate, WorkbookDataValidation, WorkbookRange } from "../contracts/workbook.js";
import { checkLimit, freeze } from "../document/validation.js";
import { createWorkbookScalarEvaluator } from "./scalar.js";
import { createWorkbookCalculatedEvaluator } from "./calculated-evaluator.js";
import type { WorkbookCalculationResult } from "./calculation.js";
import type { WorkbookScalarResult, WorkbookScalarUnavailable } from "./scalar.js";
import { scalarToNumber, scalarToText } from "./scalar-values.js";
import type { ScalarValue } from "./scalar-values.js";
import type { WorkbookReferenceLocation } from "./reference.js";
import { currentWorkbookModel } from "./source.js";

export interface WorkbookRuleReadLimits {
  readonly maxCells: number; readonly maxRules: number; readonly maxValues: number; readonly maxSteps: number; readonly maxCodeUnits: number;
}
export const DEFAULT_WORKBOOK_RULE_LIMITS: WorkbookRuleReadLimits = Object.freeze({ maxCells: 25_000, maxRules: 256, maxValues: 250_000, maxSteps: 5_000_000, maxCodeUnits: 8 * 1024 * 1024 });
export interface WorkbookRuleReadRequest extends WorkbookRuleReadLimits {
  readonly documentId: string; readonly revision: number; readonly sheetId: string; readonly coordinates: readonly WorkbookCoordinate[];
  /** Include current values even at coordinates with no conditional/validation rule. */
  readonly includeCurrent?: boolean;
}
export interface WorkbookRuleReadOptions {
  readonly signal?: { readonly aborted: boolean };
  /** The host supplies a macrotask checkpoint when message-based cancellation
   * must run. Core owns no Worker, timer or event-loop implementation. */
  readonly yield?: () => Promise<void>;
  readonly calculation?: WorkbookCalculationResult;
}
export type WorkbookRuleValue = { readonly kind: "value"; readonly value: ScalarValue } | WorkbookScalarUnavailable;
export type WorkbookRuleReason = WorkbookScalarUnavailable["reason"] | "formula-error" | "non-numeric-threshold" | "empty-scale" | "invalid-threshold-order" | "invalid-list-reference" | "type-mismatch";
export interface WorkbookRuleIssue { readonly reason: WorkbookRuleReason; readonly functionName?: string }
export type WorkbookConditionalEvaluation = {
  readonly ruleId: string; readonly state: "match" | "miss" | "unavailable"; readonly issue?: WorkbookRuleIssue;
  readonly scale?: { readonly fromStop: number; readonly toStop: number; readonly ratio: number };
};
export interface WorkbookValidationEvaluation { readonly ruleId: string; readonly state: "valid" | "invalid" | "unavailable" | "none"; readonly issue?: WorkbookRuleIssue }
export interface WorkbookRuleCell {
  readonly row: number; readonly column: number; readonly current: WorkbookRuleValue;
  readonly conditional: readonly WorkbookConditionalEvaluation[]; readonly validation?: WorkbookValidationEvaluation;
}
export interface WorkbookRuleReadResponse {
  readonly documentId: string; readonly revision: number; readonly sheetId: string; readonly cells: readonly WorkbookRuleCell[];
  readonly steps: number; readonly readCells: number; readonly evaluatedCells: number;
}
type Conditional = DeepReadonly<WorkbookConditionalRule>;
type Validation = DeepReadonly<WorkbookDataValidation>;
type Rule = Conditional | Validation;
type Range = DeepReadonly<WorkbookRange>;
type Samples = { readonly kind: "samples"; readonly values: readonly number[] } | WorkbookRuleIssue;
const coordinateKey = (coordinate: WorkbookCoordinate): number => coordinate.row * 16_384 + coordinate.column;
function lowerBound<Value>(values: readonly Value[], predicate: (value: Value) => boolean): number {
  let low = 0, high = values.length;
  while (low < high) { const mid = (low + high) >>> 1; if (predicate(values[mid]!)) low = mid + 1; else high = mid; } return low;
}
function truth(value: ScalarValue): boolean { return value.type === "boolean" ? value.value : value.type === "number" && value.value !== 0; }
/** Weighted sums and ratios avoid overflowing max-min for finite endpoints. */
function interpolate(a: number, b: number, ratio: number): number {
  if (ratio === 0) return a; if (ratio === 1) return b;
  const scale = Math.max(Math.abs(a), Math.abs(b), 1), normalized = a / scale * (1 - ratio) + b / scale * ratio;
  return Math.max(-1, Math.min(1, normalized)) * scale;
}
function proportion(value: number, from: number, to: number): number {
  const scale = Math.max(Math.abs(value), Math.abs(from), Math.abs(to), 1);
  return Math.max(0, Math.min(1, (value / scale - from / scale) / (to / scale - from / scale)));
}

/** Evaluate only requested source coordinates and the dependencies needed by
 * their rules. No rule operation modifies imported values or their caches.
 * Request identity and all batch budgets are captured before the first yield. */
export async function readWorkbookRules(model: WorkbookModel, input: WorkbookRuleReadRequest, options: WorkbookRuleReadOptions = {}): Promise<WorkbookRuleReadResponse> {
  model = currentWorkbookModel(model);
  if (!input || typeof input !== "object" || !options || typeof options !== "object" || options.yield !== undefined && typeof options.yield !== "function"
    || options.signal !== undefined && (!options.signal || typeof options.signal.aborted !== "boolean")) throw new DocumentError("InvalidOperation", "Invalid workbook rule query options.");
  const snapshot = model.snapshot(), sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === input.sheetId);
  if (input.documentId !== snapshot.manifest.documentId || input.revision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "Workbook rules belong to another source revision.");
  if (!sheet || input.includeCurrent !== undefined && typeof input.includeCurrent !== "boolean" || Object.keys(input).some(key => !["documentId", "revision", "sheetId", "coordinates", "includeCurrent", ...Object.keys(DEFAULT_WORKBOOK_RULE_LIMITS)].includes(key))
    || Object.keys(options).some(key => !["signal", "yield", "calculation"].includes(key))) throw new DocumentError("InvalidOperation", "Invalid workbook rule query.");
  const limits = Object.fromEntries(Object.keys(DEFAULT_WORKBOOK_RULE_LIMITS).map(key => [key, input[key as keyof WorkbookRuleReadLimits]])) as unknown as WorkbookRuleReadLimits;
  if (!Object.values(limits).every(value => Number.isSafeInteger(value) && value > 0) || !Array.isArray(input.coordinates)) throw new DocumentError("InvalidOperation", "Workbook rule queries require coordinates and positive safe-integer budgets.");
  checkLimit(input.coordinates.length, limits.maxCells);
  const coordinates = input.coordinates.map(coordinate => {
    if (!coordinate || typeof coordinate !== "object" || Object.keys(coordinate).some(key => !["row", "column"].includes(key)) || !Number.isSafeInteger(coordinate.row) || !Number.isSafeInteger(coordinate.column)
      || coordinate.row < 0 || coordinate.row >= 1_048_576 || coordinate.column < 0 || coordinate.column >= 16_384) throw new DocumentError("InvalidOperation", "A workbook rule coordinate is outside its worksheet.");
    return { row: coordinate.row, column: coordinate.column };
  });
  if (new Set(coordinates.map(coordinateKey)).size !== coordinates.length) throw new DocumentError("InvalidOperation", "Workbook rule coordinates must be unique.");
  const sheetId = sheet.sheetId, documentId = snapshot.manifest.documentId, revision = snapshot.manifest.revision, signal = options.signal, checkpoint = options.yield;
  const calculation = options.calculation, includeCurrent = input.includeCurrent === true;
  let steps = 0, units = 0, readCells = 0, retainedSamples = 0, checkpoints = 0;
  const step = (): void => { if (signal?.aborted) throw new DocumentError("OperationCancelled", "Workbook rule reading was cancelled."); checkLimit(++steps, limits.maxSteps); };
  const account = (amount: number): void => { units += amount; checkLimit(units, limits.maxCodeUnits); };
  const yieldIfNeeded = async (): Promise<void> => { step(); if (++checkpoints % 32 === 0 && checkpoint) { await checkpoint(); step(); } };
  const matched = new Map(coordinates.map(coordinate => [coordinateKey(coordinate), new Set<Rule>()]));
  const rowGroups = new Map<number, WorkbookCoordinate[]>();
  for (const coordinate of coordinates) {
    let group = rowGroups.get(coordinate.row); if (!group) { group = []; rowGroups.set(coordinate.row, group); } group.push(coordinate);
  }
  const rows = [...rowGroups].sort((a, b) => a[0] - b[0]).map(([row, values]) => ({ row, coordinates: values.sort((a, b) => a.column - b.column) }));
  const selectedRules = new Set<Rule>();
  for (const rule of [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []]) for (const range of rule.ranges) {
    step();
    for (let index = lowerBound(rows, item => item.row < range.from.row); index < rows.length && rows[index]!.row <= range.to.row; index++) {
      step(); const values = rows[index]!.coordinates;
      for (let column = lowerBound(values, item => item.column < range.from.column); column < values.length && values[column]!.column <= range.to.column; column++) {
        step(); matched.get(coordinateKey(values[column]!))!.add(rule); selectedRules.add(rule); checkLimit(selectedRules.size, limits.maxRules); await yieldIfNeeded();
      }
    }
  }
  const evaluatorLimits = { maxSteps: limits.maxSteps, maxCells: limits.maxValues, maxCodeUnits: limits.maxCodeUnits };
  const evaluator = calculation === undefined ? createWorkbookScalarEvaluator(model, evaluatorLimits) : createWorkbookCalculatedEvaluator(model, calculation, evaluatorLimits);
  const currentValues = new Map<string, WorkbookRuleValue>(), samples = new Map<string, Samples>();
  const location = (coordinate: WorkbookCoordinate): WorkbookReferenceLocation => ({ sheetId, row: coordinate.row, column: coordinate.column });
  function accept(result: WorkbookScalarResult): WorkbookRuleValue {
    account(result.codeUnits);
    steps += result.steps; readCells += result.readCells; checkLimit(steps, limits.maxSteps); checkLimit(readCells, limits.maxValues);
    if (result.kind === "value") { if (result.value.type === "text") account(result.value.value.length); return { kind: "value", value: result.value }; }
    account(result.functionName?.length ?? 0); return { kind: "unavailable", reason: result.reason, ...(result.functionName ? { functionName: result.functionName } : {}) };
  }
  const remainingUnits = (): number => { checkLimit(units, limits.maxCodeUnits); return limits.maxCodeUnits - units; };
  const remainingSteps = (): number => { checkLimit(steps + 1, limits.maxSteps); return limits.maxSteps - steps; };
  function current(coordinate: WorkbookCoordinate, owner = sheetId): WorkbookRuleValue {
    step(); const key = JSON.stringify([owner, coordinateKey(coordinate)]), known = currentValues.get(key); if (known) return known;
    checkLimit(currentValues.size + 1, limits.maxValues);
    const value = accept(evaluator.cell({ ...location(coordinate), sheetId: owner }, { ...(signal ? { signal } : {}), maxSteps: remainingSteps(), maxCodeUnits: remainingUnits() })); currentValues.set(key, value); return value;
  }
  function evaluate(expression: string, rule: Rule, coordinate: WorkbookCoordinate = rule.origin): WorkbookRuleValue {
    step();
    return accept(evaluator.evaluate(expression, location(rule.origin), { at: location(coordinate), ...(signal ? { signal } : {}), maxSteps: remainingSteps(), maxCodeUnits: remainingUnits() }));
  }
  async function* stored(ranges: readonly (Range & { readonly sheetId?: string })[]): AsyncGenerator<WorkbookCoordinate & { readonly sheetId: string }> {
    const seen = new Set<string>();
    for (const range of ranges) for (const chunk of snapshot.manifest.chunks) {
      if (chunk.sheetId !== (range.sheetId ?? sheetId)) continue;
      step(); const firstRow = Math.max(range.from.row, chunk.row), lastRow = Math.min(range.to.row, chunk.row + 255), firstColumn = Math.max(range.from.column, chunk.column), lastColumn = Math.min(range.to.column, chunk.column + 31);
      if (firstRow > lastRow || firstColumn > lastColumn) continue;
      const columns = lastColumn - firstColumn + 1, rows = Math.max(1, Math.floor(4096 / columns));
      for (let row = firstRow; row <= lastRow; row += rows) {
        step(); const values = model.readWindow({ sheetId: chunk.sheetId, row, column: firstColumn, rows: Math.min(rows, lastRow - row + 1), columns });
        for (const value of values) {
          step(); const key = JSON.stringify([chunk.sheetId, coordinateKey(value)]); if (seen.has(key)) continue;
          checkLimit(seen.size + 1, limits.maxValues); seen.add(key); yield { sheetId: chunk.sheetId, row: value.row, column: value.column };
        }
        await yieldIfNeeded();
      }
    }
  }
  async function scaleSamples(rule: Extract<Conditional, { kind: "colorScale" }>): Promise<Samples> {
    const cached = samples.get(rule.ruleId); if (cached) return cached;
    const values: number[] = [];
    for await (const coordinate of stored(rule.ranges)) {
      const value = current(coordinate); if (value.kind === "unavailable") { const result = issue(value)!; samples.set(rule.ruleId, result); return result; }
      if (value.value.type === "number") { checkLimit(++retainedSamples, limits.maxValues); values.push(value.value.value); }
      await yieldIfNeeded();
    }
    if (!values.length) { const result = { reason: "empty-scale" } as const; samples.set(rule.ruleId, result); return result; }
    values.sort((a, b) => { step(); return a < b ? -1 : a > b ? 1 : 0; });
    const result = { kind: "samples", values } as const; samples.set(rule.ruleId, result); return result;
  }
  async function conditional(rule: Conditional, coordinate: WorkbookCoordinate, value: WorkbookRuleValue): Promise<WorkbookConditionalEvaluation> {
    const base = { ruleId: rule.ruleId };
    if (rule.kind === "expression") {
      const result = evaluate(rule.expression, rule, coordinate);
      return result.kind === "unavailable" ? { ...base, state: "unavailable", issue: issue(result)! } : { ...base, state: truth(result.value) ? "match" : "miss", ...(issue(result) ? { issue: issue(result)! } : {}) };
    }
    if (value.kind === "unavailable") return { ...base, state: "unavailable", issue: issue(value)! };
    if (value.value.type === "error") return { ...base, state: "miss", issue: { reason: "formula-error" } };
    if (rule.kind === "cellIs") {
      const bounds = rule.formulas.map(expression => evaluate(expression, rule, coordinate)), failed = bounds.find(bound => issue(bound));
      if (failed) return { ...base, state: failed.kind === "unavailable" ? "unavailable" : "miss", issue: issue(failed)! };
      return { ...base, state: compare(rule.operator, value.value, bounds.map(bound => (bound as Extract<WorkbookRuleValue, { kind: "value" }>).value)) ? "match" : "miss" };
    }
    if (value.value.type !== "number") return { ...base, state: "miss" };
    const data = await scaleSamples(rule); if (!("kind" in data)) return { ...base, state: "unavailable", issue: data };
    const minimum = data.values[0]!, maximum = data.values.at(-1)!, thresholds: number[] = [];
    for (const stop of rule.stops) {
      step();
      if (stop.kind === "min") thresholds.push(minimum);
      else if (stop.kind === "max") thresholds.push(maximum);
      else if (stop.kind === "number") thresholds.push(stop.value);
      else if (stop.kind === "percent") thresholds.push(interpolate(minimum, maximum, stop.value / 100));
      else if (stop.kind === "percentile") {
        const index = (data.values.length - 1) * stop.value / 100, lower = Math.floor(index), higher = Math.ceil(index);
        thresholds.push(interpolate(data.values[lower]!, data.values[higher]!, index - lower));
      } else if (stop.kind === "formula") {
        const result = evaluate(stop.expression, rule), failed = issue(result);
        if (failed) return { ...base, state: "unavailable", issue: failed };
        const number = scalarToNumber((result as Extract<WorkbookRuleValue, { kind: "value" }>).value);
        if (typeof number !== "number") return { ...base, state: "unavailable", issue: { reason: "non-numeric-threshold" } }; thresholds.push(number);
      }
    }
    if (thresholds.some((number, index) => index > 0 && number < thresholds[index - 1]!)) return { ...base, state: "unavailable", issue: { reason: "invalid-threshold-order" } };
    const number = value.value.value, last = thresholds.length - 1;
    let fromStop = 0, toStop = 0, ratio = 0;
    if (last === 1 && number === thresholds[0] && thresholds[0] === thresholds[1]) fromStop = toStop = 0;
    else if (number >= thresholds[last]!) fromStop = toStop = last;
    else if (number > thresholds[0]!) {
      for (let index = 0; index < last; index++) if (number < thresholds[index + 1]!) { fromStop = index; toStop = index + 1; ratio = proportion(number, thresholds[index]!, thresholds[index + 1]!); break; }
    }
    return { ...base, state: "match", scale: { fromStop, toStop, ratio } };
  }
  async function validate(rule: Validation, coordinate: WorkbookCoordinate, value: WorkbookRuleValue): Promise<WorkbookValidationEvaluation> {
    const plan = beginWorkbookValidation(rule, coordinate, value, { evaluate, account, resolve: (expression, at) => model.resolveReference(expression, location(at)) });
    if (!("kind" in plan)) return plan;
    const base = { ruleId: plan.ruleId }, text = plan.text, reference = { areas: plan.areas };
      let unknown: WorkbookRuleIssue | undefined, count = 0;
      for await (const candidate of stored(reference.areas)) {
        count++; const candidateValue = current(candidate, candidate.sheetId);
        if (candidateValue.kind === "unavailable") { unknown = issue(candidateValue); continue; }
        const candidateText = scalarToText(candidateValue.value);
        if (typeof candidateText === "string" && candidateText.toUpperCase() === text.toUpperCase()) return { ...base, state: "valid" };
        await yieldIfNeeded();
      }
      const area = reference.areas[0]!, total = (area.to.row - area.from.row + 1) * (area.to.column - area.from.column + 1);
      if (text === "" && count < total) return { ...base, state: "valid" };
      return unknown ? { ...base, state: "unavailable", issue: unknown } : { ...base, state: "invalid" };
  }

  try {
    const cells: WorkbookRuleCell[] = [];
    for (const coordinate of coordinates) {
      step(); const definitions = [...matched.get(coordinateKey(coordinate))!]; if (!definitions.length && !includeCurrent) continue;
      const value = current(coordinate), conditions = definitions.filter((rule): rule is Conditional => "priority" in rule).sort((a, b) => a.priority - b.priority);
      const conditionalResults: WorkbookConditionalEvaluation[] = [];
      for (const rule of conditions) {
        const result = await conditional(rule, coordinate, value); account(result.ruleId.length + (result.issue?.functionName?.length ?? 0)); conditionalResults.push(result);
        if (rule.stopIfTrue && result.state !== "miss") break;
        await yieldIfNeeded();
      }
      const validation = definitions.find((rule): rule is Validation => !("priority" in rule)), validationResult = validation ? await validate(validation, coordinate, value) : undefined;
      cells.push({ row: coordinate.row, column: coordinate.column, current: value, conditional: conditionalResults, ...(validationResult ? { validation: validationResult } : {}) });
      await yieldIfNeeded();
    }
    step(); return freeze({ documentId, revision, sheetId, cells, steps, readCells, evaluatedCells: currentValues.size });
  } finally { evaluator.dispose(); currentValues.clear(); samples.clear(); }
}
