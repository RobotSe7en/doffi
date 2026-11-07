import type { DeepReadonly, WorkbookModel, WorkbookRangeChange } from "../runtime.js";
import type { WorkbookValue } from "../contracts/workbook.js";
import type { WorkbookReferenceArea, WorkbookReferenceLocation } from "./reference.js";
import type { FormulaExpression } from "./formula.js";
import { workbookExpressionDependencies } from "./expression-dependencies.js";
import { copyWorkbookScalarContext, createWorkbookScalarKernel } from "./scalar.js";
import type { ScalarResolved, WorkbookScalarContext, WorkbookScalarUnavailable } from "./scalar.js";
import { WorkbookDependencyIndex } from "./dependency-index.js";
import type { WorkbookDependencyEntry } from "./dependency-index.js";
import { currentWorkbookModel, workbookSource } from "./source.js";
import type { ChunkIndex } from "./chunk-index.js";
import { checkLimit, freeze } from "../document/validation.js";
import { DocumentError } from "../errors.js";
import { compareWorkbooks } from "./source-changes.js";
import { registerWorkbookCalculation } from "./calculation-source.js";

export const WORKBOOK_CALCULATION_VERSION = "doffi-scalar-1";
export interface WorkbookCalculationLimits {
  readonly maxFormulas: number; readonly maxDependencies: number; readonly maxSteps: number; readonly maxReadCells: number; readonly maxCodeUnits: number;
}
export const DEFAULT_WORKBOOK_CALCULATION_LIMITS: WorkbookCalculationLimits = Object.freeze({ maxFormulas: 25_000, maxDependencies: 250_000, maxSteps: 100_000_000, maxReadCells: 250_000, maxCodeUnits: 64 * 1024 * 1024 });
export interface WorkbookCalculationOptions {
  readonly signal?: { readonly aborted: boolean };
  readonly yield?: () => Promise<void>;
  readonly context?: WorkbookScalarContext;
  readonly onProgress?: (progress: { readonly phase: "index" | "calculate"; readonly completed: number; readonly total: number }) => void;
}
export type WorkbookCalculatedCell = WorkbookReferenceLocation & {
  readonly documentId: string; readonly revision: number; readonly batchId: number; readonly engineVersion: string;
} & ({ readonly state: "valid"; readonly value: DeepReadonly<WorkbookValue> } | { readonly state: "stale" } | { readonly state: "unavailable"; readonly issue: WorkbookScalarUnavailable });
export interface WorkbookCalculationResult {
  readonly documentId: string; readonly revision: number; readonly batchId: number; readonly engineVersion: string;
  readonly cells: readonly WorkbookCalculatedCell[];
  readonly calculated: number; readonly reused: number; readonly steps: number; readonly dependencies: number;
  readonly codeUnits: number; readonly readCells: number;
}
export interface WorkbookCalculator {
  calculate(options?: WorkbookCalculationOptions): Promise<WorkbookCalculationResult>;
  calculateSync(options?: Omit<WorkbookCalculationOptions, "yield">): WorkbookCalculationResult;
  getCell(location: WorkbookReferenceLocation): WorkbookCalculatedCell;
  dispose(): void;
}
interface FormulaRecord { readonly key: string; readonly location: WorkbookReferenceLocation; readonly expression: string; readonly ast: FormulaExpression;
  readonly dependencies: readonly DeepReadonly<WorkbookReferenceArea>[]; readonly lazy: boolean; readonly volatile: boolean; readonly subtotal: boolean }
interface Graph { readonly model: WorkbookModel; readonly namespace: string; readonly records: Map<string, FormulaRecord>; readonly chunks: Map<ChunkIndex, readonly FormulaRecord[]>;
  readonly reverse: WorkbookDependencyIndex; readonly points: WorkbookDependencyIndex; readonly dependencies: number }
const key = (location: WorkbookReferenceLocation): string => JSON.stringify([location.sheetId, location.row, location.column]);
class DependencyPending { constructor(readonly record: FormulaRecord) {} }

/** A calculator retains derived results for one authorized source/provider.
 * Results are installed atomically after cancellation and revision guards. */
export function createWorkbookCalculator(source: WorkbookModel | (() => WorkbookModel), input: Partial<WorkbookCalculationLimits> = {}): WorkbookCalculator {
  if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).some(name => !Object.hasOwn(DEFAULT_WORKBOOK_CALCULATION_LIMITS, name)) || Object.values(input).some(value => !Number.isSafeInteger(value) || value! < 1)) throw new DocumentError("InvalidOperation", "Invalid workbook calculation limits.");
  const limits = { ...DEFAULT_WORKBOOK_CALCULATION_LIMITS, ...input };
  let provider: (() => WorkbookModel) | undefined = typeof source === "function" ? source : () => source;
  let graph: Graph | undefined, values = new Map<string, ScalarResolved>(), contextKey = "", localeKey = "", batch = 0, publishedBatch = 0;
  function current(): WorkbookModel { if (!provider) throw new DocumentError("InvalidOperation", "This workbook calculator has been disposed."); return currentWorkbookModel(provider()); }
  function recordResult(model: WorkbookModel, location: WorkbookReferenceLocation, result?: ScalarResolved, resultBatch = publishedBatch): WorkbookCalculatedCell {
    const manifest = model.snapshot().manifest, identity = { ...location, documentId: manifest.documentId, revision: manifest.revision, batchId: resultBatch, engineVersion: WORKBOOK_CALCULATION_VERSION };
    return freeze(result === undefined ? { ...identity, state: "stale" } : "type" in result ? { ...identity, state: "valid", value: result } : { ...identity, state: "unavailable", issue: result });
  }
  function* run(options: WorkbookCalculationOptions = {}): Generator<void, WorkbookCalculationResult> {
      if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(name => !["signal", "yield", "context", "onProgress"].includes(name))
        || options.yield !== undefined && typeof options.yield !== "function" || options.onProgress !== undefined && typeof options.onProgress !== "function"
        || options.signal !== undefined && (!options.signal || typeof options.signal.aborted !== "boolean")) throw new DocumentError("InvalidOperation", "Invalid workbook calculation options.");
      const model = current(), manifest = model.snapshot().manifest, mine = ++batch, snapshot = model.snapshot();
      const privateSource = workbookSource(model); if (!privateSource) throw new DocumentError("InvalidOperation", "Calculation requires a validated workbook model.");
      const calculationContext = copyWorkbookScalarContext(options.context);
      const resolvedLocale = calculationContext.locale ?? snapshot.metadata.locale ?? "zh-CN", changedLocale = resolvedLocale !== localeKey;
      let steps = 0, units = 0, calculated = 0, reused = 0, yields = 0;
      const readCells = new Set<string>();
      function guard(): void {
        if (options.signal?.aborted || mine !== batch || !provider) throw new DocumentError("OperationCancelled", "Workbook calculation was cancelled.");
        const active = current(); if (active !== model || active.snapshot().manifest.revision !== manifest.revision || active.snapshot().manifest.documentId !== manifest.documentId) throw new DocumentError("RevisionConflict", "Calculation belongs to a previous workbook source.");
      }
      const step = (): void => { checkLimit(++steps, limits.maxSteps); if ((steps & 255) === 1) guard(); };
      const account = (amount: number): void => { units += amount; checkLimit(units, limits.maxCodeUnits); };
      const progress = (phase: "index" | "calculate", completed: number, total: number): void => { if (options.onProgress) try { options.onProgress(freeze({ phase, completed, total })); } catch { /* Host owns progress callback errors. */ } };
      function* checkpoint(): Generator<void> { if (++yields % 64 === 0) yield; guard(); }
      guard(); yield; guard();
      const namespace = JSON.stringify([snapshot.metadata.workbook.dateSystem, snapshot.metadata.workbook.sheets.map(sheet => [sheet.sheetId, sheet.name, sheet.tables]), snapshot.metadata.workbook.names]);
      const records = new Map<string, FormulaRecord>(), chunks = new Map<ChunkIndex, readonly FormulaRecord[]>(), dependencies: WorkbookDependencyEntry[] = [];
      for (const descriptor of manifest.chunks) {
        step(); const index = privateSource.chunks.get(descriptor.path)!, known = graph?.namespace === namespace ? graph.chunks.get(index) : undefined;
        const items: FormulaRecord[] = [];
        if (known) items.push(...known);
        else for (const formula of index.formulaLocations()) {
          step(); checkLimit(records.size + items.length + 1, limits.maxFormulas);
          const location = { sheetId: descriptor.sheetId, row: formula.row, column: formula.column }, cell = model.getCell(location.sheetId, location.row, location.column)!;
          const areas: DeepReadonly<WorkbookReferenceArea>[] = [];
          let lazy = false, volatile = false, subtotal = false;
          workbookExpressionDependencies(model, formula.expression, location, step, area => {
            checkLimit(dependencies.length + areas.length + 1, limits.maxDependencies); areas.push(area);
          }, name => {
              lazy ||= ["IF", "IFS", "IFERROR", "IFNA", "INDEX", "MATCH", "VLOOKUP", "HLOOKUP", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", "SUBTOTAL", "COUNTA", "ISBLANK", "ROW", "COLUMN", "ROWS", "COLUMNS"].includes(name);
              volatile ||= name === "TODAY" || name === "NOW" || name === "SUBTOTAL"; subtotal ||= name === "SUBTOTAL";
          });
          // A changed chunk may contain an unchanged formula. Keep its learned
          // reference areas (for example SUMIF's expanded sum_range), so edits
          // in that same chunk still invalidate the formula before evaluation.
          const previous = graph?.namespace === namespace ? graph.records.get(key(location)) : undefined;
          const retainedAreas = previous?.expression === cell.formula!.expression ? previous.dependencies : areas;
          account(cell.formula!.expression.length); items.push({ key: key(location), location, expression: cell.formula!.expression, ast: formula.expression, dependencies: retainedAreas, lazy, volatile, subtotal });
          yield* checkpoint();
        }
        chunks.set(index, items);
        for (const record of items) { step(); checkLimit(records.size + 1, limits.maxFormulas); records.set(record.key, record); for (const area of record.dependencies) { checkLimit(dependencies.length + 1, limits.maxDependencies); dependencies.push({ area, owner: record.key }); } }
        progress("index", records.size, limits.maxFormulas); yield* checkpoint();
      }
      const reverse = new WorkbookDependencyIndex(dependencies, step), points = new WorkbookDependencyIndex([...records.values()].map(record => ({ owner: record.key, area: { sheetId: record.location.sheetId, from: record.location, to: record.location } })), step);
      let next: Graph = { model, namespace, records, chunks, reverse, points, dependencies: dependencies.length };
      const calculatedValues = new Map<string, ScalarResolved>();
      const dirty = new Set<string>(), changedContext = JSON.stringify(calculationContext) !== contextKey;
      if (graph && graph.namespace === namespace && graph.model.snapshot().manifest.documentId === manifest.documentId) {
        const changes: readonly WorkbookRangeChange[] = graph.model === model ? [] : compareWorkbooks(graph.model, model).changedRanges;
        const queue = [...changes];
        for (const record of records.values()) {
          const old = graph.records.get(record.key);
          if (!old || old.expression !== record.expression || changedLocale || record.volatile && changedContext || record.subtotal && graph.model !== model) { dirty.add(record.key); queue.push({ sheetId: record.location.sheetId, range: { from: record.location, to: record.location } }); }
        }
        for (let cursor = 0; cursor < queue.length; cursor++) {
          step(); const changed = queue[cursor]!;
          for (const owner of reverse.query({ sheetId: changed.sheetId, ...changed.range }, step)) if (!dirty.has(owner)) { dirty.add(owner); const record = records.get(owner)!; queue.push({ sheetId: record.location.sheetId, range: { from: record.location, to: record.location } }); }
        }
        for (const record of records.values()) if (!dirty.has(record.key) && values.has(record.key)) { calculatedValues.set(record.key, values.get(record.key)!); reused++; }
      }
      const active = new Set<string>();
      let evaluating: FormulaRecord | undefined, dependencyCount = dependencies.length, newDependencies = false;
      function rememberDependency(area: DeepReadonly<WorkbookReferenceArea>): void {
        if (!evaluating) return;
        const record = records.get(evaluating.key)!;
        for (const known of record.dependencies) {
          step(); if (known.sheetId === area.sheetId && known.from.row <= area.from.row && known.to.row >= area.to.row && known.from.column <= area.from.column && known.to.column >= area.to.column) return;
        }
        checkLimit(++dependencyCount, limits.maxDependencies); account(area.sheetId.length + 64);
        records.set(record.key, { ...record, dependencies: [...record.dependencies, area] }); newDependencies = true;
      }
      function refreshDependencies(): void {
        if (!newDependencies) return;
        const entries = [...records.values()].flatMap(record => record.dependencies.map(area => ({ area, owner: record.key })));
        next = { ...next, reverse: new WorkbookDependencyIndex(entries, step), dependencies: entries.length,
          chunks: new Map([...chunks].map(([index, items]) => [index, items.map(record => records.get(record.key)!)])) };
        newDependencies = false;
      }
      const evaluator = createWorkbookScalarKernel(model, { maxSteps: limits.maxSteps, maxCells: limits.maxReadCells, maxCodeUnits: limits.maxCodeUnits }, calculationContext, {
        step, account,
        reference: rememberDependency,
        read: location => { readCells.add(key(location)); checkLimit(readCells.size, limits.maxReadCells); rememberDependency({ sheetId: location.sheetId, from: location, to: location }); },
        cachedCell: location => {
          const id = key(location), known = calculatedValues.get(id); if (known) return known;
          if (active.has(id)) return { kind: "unavailable", reason: "circular-reference" };
          const record = records.get(id); if (record) throw new DependencyPending(record); return undefined;
        },
      });
      function evaluateRecord(record: FormulaRecord): ScalarResolved {
        evaluating = record;
        try {
          const result = evaluator.evaluate(record.expression, record.location, { ...(options.signal ? { signal: options.signal } : {}) });
          return result.kind === "value" ? result.value : { kind: "unavailable", reason: result.reason, ...(result.functionName ? { functionName: result.functionName } : {}) };
        } finally { evaluating = undefined; }
      }
      function* prerequisites(record: FormulaRecord): Generator<string> { if (!record.lazy) for (const area of record.dependencies) yield* points.query(area, step); }
      try {
        for (const root of records.values()) {
          if (calculatedValues.has(root.key)) continue;
          const stack = [{ record: root, dependencies: prerequisites(root) }]; active.add(root.key);
          while (stack.length) {
            step(); const frame = stack.at(-1)!, dependency = frame.dependencies.next();
            if (!dependency.done) {
              const record = records.get(dependency.value)!;
              if (!calculatedValues.has(record.key) && !active.has(record.key)) { active.add(record.key); stack.push({ record, dependencies: prerequisites(record) }); } continue;
            }
            try {
              const value = evaluateRecord(frame.record);
              calculatedValues.set(frame.record.key, value); active.delete(frame.record.key); stack.pop(); calculated++;
              progress("calculate", calculated + reused, records.size); yield* checkpoint();
            } catch (error) {
              if (!(error instanceof DependencyPending)) throw error;
              active.add(error.record.key); stack.push({ record: error.record, dependencies: prerequisites(error.record) });
            }
          }
        }
        refreshDependencies();
        // Static prerequisites are an optimization, and may visit an edge
        // that an earlier formula error/condition never actually reads. A
        // provisional circular result must be reconciled with the now-stable
        // values before publication; genuine cycles remain unavailable.
        const provisional = new Set([...calculatedValues].filter(([, value]) => !("type" in value) && value.reason === "circular-reference").map(([id]) => id));
        const retry = new Set(provisional);
        while (retry.size) {
          step(); const id = retry.values().next().value!; retry.delete(id);
          const record = records.get(id)!, value = evaluateRecord(record);
          if (!("type" in value) && value.reason === "circular-reference") continue;
          calculatedValues.set(id, value); provisional.delete(id); refreshDependencies();
          for (const owner of next.reverse.query({ sheetId: record.location.sheetId, from: record.location, to: record.location }, step)) if (provisional.has(owner)) retry.add(owner);
          yield* checkpoint();
        }
        refreshDependencies();
        const cells = [...records.values()].map(record => {
          const value = calculatedValues.get(record.key)!;
          account(record.key.length + ("type" in value && value.type === "text" ? value.value.length : 0));
          return recordResult(model, record.location, value, mine);
        });
        const result = freeze({ documentId: manifest.documentId, revision: manifest.revision, batchId: mine, engineVersion: WORKBOOK_CALCULATION_VERSION, cells, calculated, reused, steps, dependencies: next.dependencies, codeUnits: units, readCells: readCells.size });
        guard(); graph = next; values = calculatedValues; contextKey = JSON.stringify(calculationContext); localeKey = resolvedLocale; publishedBatch = mine;
        registerWorkbookCalculation(result, { model, context: calculationContext, values: calculatedValues });
        return result;
      } finally { evaluator.dispose(); active.clear(); }
  }
  return Object.freeze({
    getCell(location: WorkbookReferenceLocation): WorkbookCalculatedCell {
      if (!location || typeof location !== "object" || Object.keys(location).some(name => !["sheetId", "row", "column"].includes(name))) throw new DocumentError("InvalidOperation", "Invalid calculated cell coordinate.");
      const model = current(), cell = model.getCell(location.sheetId, location.row, location.column);
      return recordResult(model, location, !cell?.formula ? cell?.value ?? { type: "blank" } : graph?.model === model ? values.get(key(location)) : undefined);
    },
    async calculate(options: WorkbookCalculationOptions = {}): Promise<WorkbookCalculationResult> {
      const iterator = run(options);
      try {
        for (let result = iterator.next();; result = iterator.next()) {
          if (result.done) return result.value;
          if (options.yield) await options.yield();
        }
      } finally { const cleanup: Iterator<unknown> = iterator; cleanup.return?.(); }
    },
    calculateSync(options: Omit<WorkbookCalculationOptions, "yield"> = {}): WorkbookCalculationResult {
      const iterator = run(options);
      try { for (let result = iterator.next();; result = iterator.next()) if (result.done) return result.value; }
      finally { const cleanup: Iterator<unknown> = iterator; cleanup.return?.(); }
    },
    dispose(): void { provider = undefined; graph = undefined; values = new Map(); contextKey = ""; localeKey = ""; batch++; },
  });
}
