import type { DocumentTransaction } from "../contracts/operations.js";
import type { DocumentEvent } from "../contracts/events.js";
import type { CandidateChange, ChangeMapping, DocumentModel, DocumentSessionOptions, SessionLifecycle, TransactionPreparationExecution, TransactionPreparationProgress, TransactionPrepareOptions, TransactionReceipt } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { checkLimit, freeze, readLimits } from "../document/validation.js";
import { parseContract } from "../schema-registry.js";
import { checkTransactionSize, sessionLimits } from "./limits.js";
import { invertWorkbookMapping } from "../workbook/range-mapping.js";

export type SessionDiff = Pick<CandidateChange<DocumentModel>, "affectedNodeIds" | "affectedSlideIds" | "affectedSheetIds" | "changedRanges" | "changes" | "changedAnnotationIds" | "mapping">;
export interface SessionPrepared<State> { readonly state: State; readonly mapping?: ChangeMapping; readonly diff?: SessionDiff }
export interface SessionPreparationControl {
  readonly signal?: { readonly aborted: boolean };
  guard(): void;
  checkpoint(progress: TransactionPreparationProgress): Promise<void>;
}
export interface SessionAdapter<State, Model extends DocumentModel> {
  model(state: State): Model;
  prepare(base: State, transaction: DocumentTransaction, allocate: () => string, resolvedIds?: readonly string[], options?: TransactionPrepareOptions): SessionPrepared<State>;
  prepareAsync?(base: State, transaction: DocumentTransaction, allocate: () => string, resolvedIds: readonly string[] | undefined, options: TransactionPrepareOptions | undefined, control: SessionPreparationControl): Promise<SessionPrepared<State>>;
  rebase(state: State, revision: number): State;
  compare(before: State, after: State, mapping?: ChangeMapping): SessionDiff;
  same(before: State, after: State): boolean;
  retained(states: Iterable<State>): number;
  identities(state: State): Iterable<string>;
}

/** The only candidate ownership, revision, publication and undo/redo boundary. */
export function createSessionLifecycle<State, Model extends DocumentModel>(initial: State, adapter: SessionAdapter<State, Model>, options: DocumentSessionOptions) {
  const limits = readLimits(options.limits), budgets = sessionLimits(options.sessionLimits);
  type Candidate = CandidateChange<Model>;
  interface Entry { before: State; after: State; mapping: ChangeMapping }
  let current: State | undefined = initial, undo: Entry[] = [], redo: Entry[] = [];
  const active = new Map<Candidate, State>(), listeners = new Set<import("../runtime.js").DocumentEventListener>();
  let issued = new WeakMap<Candidate, { baseRevision: number; consumed: boolean }>();
  const reserved = new Set(adapter.identities(initial));
  let sequence = 0, publishing = false, preparing = 0;
  function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
  function state(): State { if (current === undefined) invalid("This document session has been disposed."); return current; }
  function editable(): void { state(); if (publishing) invalid("A session cannot be changed while its events are being delivered."); }
  function emit(event: DocumentEvent): void {
    publishing = true;
    try { const value = freeze(event); for (const listener of [...listeners]) try { listener(value); } catch { /* Host owns callback errors. */ } }
    finally { publishing = false; }
  }
  function execute<Value>(action: () => Value): Value {
    editable();
    try { return action(); } catch (error) {
      const bounded = error instanceof DocumentError ? error.code === "InvalidDocument" ? new DocumentError("InvalidOperation", error.message) : error
        : new DocumentError("InvalidOperation", "The document operation could not be prepared or committed.");
      const manifest = adapter.model(state()).snapshot().manifest;
      emit({ protocolVersion: 1, documentId: manifest.documentId, revision: manifest.revision, type: "error", error: bounded.toJSON() }); throw bounded;
    }
  }
  function reserve(value: unknown, key = ""): void {
    if (typeof value === "string") { if (/(?:Ids?|Style|Class)$/.test(key)) reserved.add(value); return; }
    if (Array.isArray(value)) { for (const item of value) reserve(item, key); return; }
    if (value && typeof value === "object") for (const [name, item] of Object.entries(value)) reserve(item, name);
  }
  function allocate(): string {
    let id: string;
    do { if (sequence >= Number.MAX_SAFE_INTEGER) throw new DocumentError("ResourceLimitExceeded", "Node identity allocation is exhausted."); id = "edit-" + ++sequence; } while (reserved.has(id));
    reserved.add(id); return id;
  }
  function historyStates(entries: readonly Entry[]): State[] { return entries.flatMap(entry => [entry.before, entry.after]); }
  function receipt(diff: SessionDiff, changed: boolean): TransactionReceipt {
    const manifest = adapter.model(state()).snapshot().manifest;
    return freeze({ documentId: manifest.documentId, revision: manifest.revision, changed, affectedNodeIds: diff.affectedNodeIds, mapping: diff.mapping,
      ...(diff.affectedSlideIds ? { affectedSlideIds: diff.affectedSlideIds } : {}), ...(diff.affectedSheetIds ? { affectedSheetIds: diff.affectedSheetIds } : {}),
      ...(diff.changedRanges ? { changedRanges: diff.changedRanges } : {}) });
  }
  function empty(): TransactionReceipt { return receipt(adapter.compare(state(), state()), false); }
  function revision(): number {
    const value = adapter.model(state()).snapshot().manifest.revision;
    if (value >= Number.MAX_SAFE_INTEGER) throw new DocumentError("ResourceLimitExceeded", "The document revision is exhausted."); return value + 1;
  }
  function publish(next: State, mapping: ChangeMapping): TransactionReceipt {
    const diff = adapter.compare(state(), next, mapping); current = next; active.clear();
    const result = receipt(diff, true), identity = { protocolVersion: 1 as const, documentId: result.documentId, revision: result.revision };
    emit({ ...identity, type: "documentChanged", affectedNodeIds: [...diff.affectedNodeIds],
      ...(diff.affectedSlideIds ? { affectedSlideIds: [...diff.affectedSlideIds] } : {}), ...(diff.affectedSheetIds ? { affectedSheetIds: [...diff.affectedSheetIds] } : {}) });
    if (diff.changedAnnotationIds.length) emit({ ...identity, type: "annotationChanged", annotationIds: [...diff.changedAnnotationIds] }); return result;
  }
  function preparationInput(input: DocumentTransaction): { readonly transaction: DocumentTransaction; readonly base: State } {
    if (input && Array.isArray(input.operations)) {
      let count = input.operations.length;
      for (const operation of input.operations) { if (operation?.op === "reconcileFragment" && Array.isArray(operation.textEdits)) count += operation.textEdits.length; checkLimit(count, budgets.maxOperations); }
    }
    checkTransactionSize(input, budgets.maxTransactionBytes, limits.maxDepth);
    const transaction = parseContract("document-transaction", input), base = state(), snapshot = adapter.model(base).snapshot();
    if (transaction.documentId !== snapshot.manifest.documentId) invalid("The transaction belongs to another document.");
    if (transaction.baseRevision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "The transaction is based on a different document revision.");
    checkLimit(active.size + preparing + 1, budgets.maxCandidates); reserve(snapshot.metadata); reserve(snapshot.annotations); reserve(transaction);
    return { transaction, base };
  }
  function preparedCandidate(transaction: DocumentTransaction, base: State, prepared: SessionPrepared<State>): Candidate {
    const model = adapter.model(prepared.state);
    checkLimit(active.size + 1, budgets.maxCandidates);
    checkLimit(adapter.retained([...active.values(), prepared.state]), budgets.maxCandidateBytes);
    const candidate = freeze({ documentId: transaction.documentId, baseRevision: transaction.baseRevision,
      preview: model.snapshot(), model, ...(prepared.diff ?? adapter.compare(base, prepared.state, prepared.mapping)) }) as unknown as Candidate;
    issued.set(candidate, { baseRevision: transaction.baseRevision, consumed: false }); active.set(candidate, prepared.state); return candidate;
  }
  function prepareTransaction(input: DocumentTransaction, resolvedIds?: readonly string[], preparation?: TransactionPrepareOptions): Candidate {
    const { transaction, base } = preparationInput(input);
    return preparedCandidate(transaction, base, adapter.prepare(base, transaction, allocate, resolvedIds, preparation));
  }
  async function prepareAsyncTransaction(input: DocumentTransaction, resolvedIds?: readonly string[], preparation?: TransactionPrepareOptions, execution: TransactionPreparationExecution = {}): Promise<Candidate> {
    editable(); let entered = false;
    try {
      if (!adapter.prepareAsync) invalid("This document format has no asynchronous preparation adapter.");
      if (!execution || typeof execution !== "object" || Array.isArray(execution) || Object.keys(execution).some(key => !["signal", "yield", "onProgress"].includes(key))
        || execution.signal !== undefined && (!execution.signal || typeof execution.signal.aborted !== "boolean")
        || execution.yield !== undefined && typeof execution.yield !== "function" || execution.onProgress !== undefined && typeof execution.onProgress !== "function") invalid("Invalid transaction preparation execution options.");
      const config = { ...execution };
      if (config.signal?.aborted) throw new DocumentError("OperationCancelled", "Transaction preparation was cancelled.");
      if (preparing) throw new DocumentError("ResourceLimitExceeded", "Another asynchronous preparation is already in progress for this session.");
      // The structural parser captures the input before any host callback or
      // checkpoint. No mutable caller transaction is retained across a yield.
      const { transaction, base } = preparationInput(input); preparing++; entered = true;
      const guard = (): void => {
        if (current === undefined || config.signal?.aborted) throw new DocumentError("OperationCancelled", "Transaction preparation was cancelled or disposed.");
        if (current !== base) throw new DocumentError("RevisionConflict", "The document changed while preparing the transaction.");
      };
      const control: SessionPreparationControl = { ...(config.signal ? { signal: config.signal } : {}), guard,
        async checkpoint(progress) {
          guard(); try { config.onProgress?.(freeze({ ...progress })); } catch { /* Host reporting cannot publish source. */ }
          guard(); await config.yield?.(); guard();
        } };
      const prepared = await adapter.prepareAsync(base, transaction, allocate, resolvedIds, preparation, control);
      guard(); return preparedCandidate(transaction, base, prepared);
    } catch (error) {
      const bounded = error instanceof DocumentError ? error.code === "InvalidDocument" ? new DocumentError("InvalidOperation", error.message) : error
        : new DocumentError("InvalidOperation", "The document operation could not be prepared or committed.");
      if (current !== undefined) {
        const manifest = adapter.model(current).snapshot().manifest;
        emit({ protocolVersion: 1, documentId: manifest.documentId, revision: manifest.revision, type: "error", error: bounded.toJSON() });
      }
      throw bounded;
    } finally { if (entered) preparing--; }
  }
  const lifecycle: SessionLifecycle<Candidate> = {
    getAnnotations: () => adapter.model(state()).snapshot().annotations,
    prepare: (input, preparation) => execute(() => prepareTransaction(input, undefined, preparation)),
    commit: candidate => execute(() => {
      const record = issued.get(candidate); if (!record || record.consumed) invalid("The candidate is foreign, discarded or already consumed.");
      const base = state();
      if (record.baseRevision !== adapter.model(base).snapshot().manifest.revision) throw new DocumentError("RevisionConflict", "The candidate is based on an older document revision.");
      const prepared = active.get(candidate); if (prepared === undefined) invalid("The candidate is no longer available.");
      if (adapter.same(base, prepared)) { record.consumed = true; active.delete(candidate); return empty(); }
      const next = adapter.rebase(prepared, revision()), entries = [...undo, { before: base, after: next, mapping: candidate.mapping }];
      checkLimit(entries.length, budgets.maxHistorySteps); checkLimit(adapter.retained(historyStates(entries)), budgets.maxHistoryBytes);
      record.consumed = true; undo = entries; redo = []; return publish(next, candidate.mapping);
    }),
    discard: candidate => execute(() => { const record = issued.get(candidate); if (!record || record.consumed) invalid("The candidate is foreign, discarded or already consumed."); record.consumed = true; active.delete(candidate); }),
    undo: () => execute(() => {
      const entry = undo.at(-1); if (!entry) return empty(); const next = adapter.rebase(entry.before, revision());
      const mapping: ChangeMapping = { nodes: entry.mapping.nodes, textEdits: [...entry.mapping.textEdits].reverse().map(item => ({ nodeId: item.nodeId, from: item.from, to: item.from + item.insertedLength, insertedLength: item.to - item.from })),
        ...(entry.mapping.workbook ? { workbook: invertWorkbookMapping(entry.mapping.workbook) } : {}) };
      undo = undo.slice(0, -1); redo = [...redo, entry]; return publish(next, mapping);
    }),
    redo: () => execute(() => { const entry = redo.at(-1); if (!entry) return empty(); const next = adapter.rebase(entry.after, revision()); redo = redo.slice(0, -1); undo = [...undo, entry]; return publish(next, entry.mapping); }),
    history: () => { state(); return freeze({ undoSteps: undo.length, redoSteps: redo.length, retainedBytes: adapter.retained(historyStates([...undo, ...redo])) }); },
    clearHistory: () => execute(() => { undo = []; redo = []; }),
    subscribe: listener => { state(); listeners.add(listener); return () => { listeners.delete(listener); }; },
    dispose: () => { if (current === undefined) return; editable(); current = undefined; undo = []; redo = []; active.clear(); issued = new WeakMap(); reserved.clear(); listeners.clear(); },
  };
  return { lifecycle, state, execute, prepareTransaction, prepareAsyncTransaction };
}
