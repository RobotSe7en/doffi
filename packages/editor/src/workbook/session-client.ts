import { DocumentError, DocumentErrorDataSchema, WorkbookCoordinateSchema, WorkbookRangeSchema } from "@doffi/core";
import type { TransactionPrepareOptions, WorkbookChangeImpact, WorkbookCoordinate } from "@doffi/core";
import { connectWorkbookReads, linkWorkbookPreviewSource } from "@doffi/render";
import type { WorkbookReadConnection } from "@doffi/render";
import { checkSessionMessage, freezeSessionValue, invalidSessionTransport, keys, positiveId, record, sessionTransportConfig, WORKBOOK_SESSION_PROTOCOL } from "./session-transport.js";
import type { WorkbookSessionCandidate, WorkbookSessionCommand, WorkbookSessionConnection, WorkbookSessionExecution, WorkbookSessionFindOptions, WorkbookSessionInput, WorkbookSessionProgress, WorkbookSessionReceipt, WorkbookSessionState, WorkbookSessionTransportOptions, WorkbookSessionViewOptions } from "./session-transport.js";
import type { WorkbookFindQuery, WorkbookFindResult } from "./find-contracts.js";

interface Reply { readonly value: unknown; readonly port?: MessagePort }
interface Pending { readonly operation: string; cancelled?: boolean; resolve(reply: Reply): void; reject(error: unknown): void; clean(): void; progress?(value: WorkbookSessionProgress): void }
/** Owns an authorized port and bounded view state. Core candidates, source
 * chunks, calculation and history stay in the host's sole Worker session. */
export async function connectWorkbookSession(port: MessagePort, options: WorkbookSessionTransportOptions = {}, signal?: AbortSignal): Promise<WorkbookSessionConnection> {
  const config = sessionTransportConfig(options), pending = new Map<number, Pending>(), listeners = new Set<(state: WorkbookSessionState) => void>(), views = new Set<WorkbookReadConnection>();
  let disposed = false, nextId = 0, current: WorkbookSessionState;
  const candidates = new WeakMap<WorkbookSessionCandidate, { handle: number; consumed: boolean }>();
  const viewHandles = new WeakMap<WorkbookReadConnection, number>();
  function cancelled(message = "The workbook session operation was cancelled."): DocumentError { return new DocumentError("OperationCancelled", message); }
  function check(): void { if (disposed) invalidSessionTransport("This workbook session connection has been disposed."); }
  function closeTransferred(port: MessagePort | undefined): void { if (port) void connectWorkbookReads(port, config.reads, AbortSignal.abort()).catch(() => {}); }
  function dispose(reason: DocumentError = cancelled("The workbook session connection was closed.")): void {
    if (disposed) return; disposed = true;
    try { port.postMessage({ protocol: WORKBOOK_SESSION_PROTOCOL, event: "close" }); } catch { /* Its host may already have closed the port. */ }
    for (const request of pending.values()) { request.clean(); request.reject(reason); } pending.clear();
    for (const view of views) view.dispose(); views.clear(); listeners.clear();
    port.removeEventListener("message", receive); port.removeEventListener("messageerror", invalidMessage); port.close();
  }
  function invalidMessage(): void { dispose(new DocumentError("InvalidDocument", "The workbook session returned an invalid response.")); }
  function acceptState(value: unknown, notify = false): WorkbookSessionState {
    if (!record(value) || !keys(value, ["documentId", "revision", "history"]) || typeof value.documentId !== "string" || !value.documentId
      || !Number.isSafeInteger(value.revision) || (value.revision as number) < 0 || !record(value.history) || !keys(value.history, ["undoSteps", "redoSteps", "retainedBytes"])
      || ![value.history.undoSteps, value.history.redoSteps, value.history.retainedBytes].every(item => typeof item === "number" && Number.isSafeInteger(item) && item >= 0)) invalidSessionTransport("Invalid workbook session state.");
    const result = freezeSessionValue(value as unknown as WorkbookSessionState);
    if (current && result.documentId !== current.documentId) invalidSessionTransport("A workbook connection cannot switch its document identity.");
    if (!current || result.revision >= current.revision) {
      current = result;
      if (notify) for (const listener of [...listeners]) try { listener(current); } catch { /* Host owns callback errors. */ }
    }
    return result;
  }
  function receive(event: MessageEvent<unknown>): void {
    const message = event.data;
    if (!record(message) || message.protocol !== WORKBOOK_SESSION_PROTOCOL) return;
    try {
      if (message.event === "close" && keys(message, ["protocol", "event"])) { dispose(); return; }
      if (message.event === "state" && keys(message, ["protocol", "event", "value"])) { checkSessionMessage(message.value, config); acceptState(message.value, true); return; }
      if (!positiveId(message.id)) throw new Error("Missing response identity");
      const request = pending.get(message.id);
      if (message.event === "cancelled") {
        if (!keys(message, ["protocol", "id", "event"]) || request && !request.cancelled) throw new Error("Invalid cancellation acknowledgment");
        if (request) { pending.delete(message.id); request.clean(); request.reject(cancelled()); } return;
      }
      if (message.event === "progress") {
        if (!keys(message, ["protocol", "id", "event", "value"])) throw new Error("Invalid progress envelope");
        checkSessionMessage(message.value, config);
        const value = message.value;
        if (!record(value) || !keys(value, ["phase", "completed", "total"]) || !["operations", "cells", "chunks", "validation", "calculation", "copy", "fill", "find"].includes(String(value.phase))
          || ![value.completed, value.total].every(item => typeof item === "number" && Number.isSafeInteger(item) && item >= 0)) throw new Error("Invalid progress result");
        try { if (!request?.cancelled) request?.progress?.(freezeSessionValue(value as unknown as WorkbookSessionProgress)); } catch { /* Progress is reporting only. */ } return;
      }
      if (typeof message.ok !== "boolean" || !keys(message, message.ok ? ["protocol", "id", "ok", "value", "port"] : ["protocol", "id", "ok", "error"])) throw new Error("Invalid result envelope");
      const transferred = message.port === undefined ? undefined : message.port instanceof MessagePort ? message.port : undefined;
      if (message.port !== undefined && !transferred) throw new Error("Invalid view port");
      if (!request) { closeTransferred(transferred); return; }
      if (Boolean(transferred) !== (message.ok && request.operation === "view")) { closeTransferred(transferred); throw new Error("Unexpected view port"); }
      checkSessionMessage(message.ok ? message.value : message.error, config);
      if (request.cancelled) { pending.delete(message.id); request.clean(); closeTransferred(transferred); request.reject(cancelled()); return; }
      if (message.ok) { pending.delete(message.id); request.clean(); request.resolve({ value: message.value, ...(transferred ? { port: transferred } : {}) }); }
      else {
        const parsed = DocumentErrorDataSchema.safeParse(message.error); if (!parsed.success) throw new Error("Invalid error result");
        pending.delete(message.id); request.clean(); request.reject(new DocumentError(parsed.data.code, parsed.data.message, parsed.data.details));
      }
    } catch { if (message.port instanceof MessagePort) closeTransferred(message.port); invalidMessage(); }
  }
  function request(operation: string, payload: unknown, execution: WorkbookSessionExecution = {}, mutation = false): Promise<Reply> {
    check();
    if (!execution || typeof execution !== "object" || Array.isArray(execution) || Object.keys(execution).some(key => !["signal", "onProgress"].includes(key)) || execution.onProgress !== undefined && typeof execution.onProgress !== "function"
      || execution.signal !== undefined && (!execution.signal || typeof execution.signal.aborted !== "boolean" || typeof execution.signal.addEventListener !== "function" || typeof execution.signal.removeEventListener !== "function")) invalidSessionTransport("Invalid workbook session execution options.");
    const supplied = { ...execution };
    if (supplied.signal?.aborted) return Promise.reject(cancelled());
    checkSessionMessage(payload, config);
    if (pending.size >= config.maxPendingRequests || nextId >= Number.MAX_SAFE_INTEGER) throw new DocumentError("ResourceLimitExceeded", "Too many pending workbook session requests.");
    const id = ++nextId;
    return new Promise((resolve, reject) => {
      const cancel = (): void => {
        const owned = pending.get(id); if (!owned || owned.cancelled) return;
        if (operation === "prepare") {
          owned.cancelled = true; supplied.signal?.removeEventListener("abort", cancel); clearTimeout(timer);
          timer = setTimeout(() => dispose(cancelled("The remote workbook preparation did not acknowledge cancellation.")), config.requestTimeoutMs);
          try { port.postMessage({ protocol: WORKBOOK_SESSION_PROTOCOL, id, operation: "cancel" }); } catch { dispose(cancelled()); }
          return;
        }
        pending.delete(id); clean();
        try { port.postMessage({ protocol: WORKBOOK_SESSION_PROTOCOL, id, operation: "cancel" }); } catch { /* Local cancellation still settles. */ }
        reject(cancelled());
      };
      let timer = setTimeout(() => {
        if (mutation) dispose(new DocumentError("InvalidOperation", "The workbook command response timed out; its commit outcome is unknown. Inspect the host session before retrying.")); else cancel();
      }, config.requestTimeoutMs);
      const clean = (): void => { clearTimeout(timer); supplied.signal?.removeEventListener("abort", cancel); };
      pending.set(id, { operation, resolve, reject, clean, ...(supplied.onProgress ? { progress: supplied.onProgress } : {}) });
      supplied.signal?.addEventListener("abort", cancel, { once: true });
      try { port.postMessage({ protocol: WORKBOOK_SESSION_PROTOCOL, id, operation, payload }); }
      catch { pending.delete(id); clean(); reject(new DocumentError("InvalidOperation", "The workbook session could not send this command.")); }
    });
  }
  function candidateRecord(candidate: WorkbookSessionCandidate): { handle: number; consumed: boolean } {
    check(); const owned = candidates.get(candidate);
    if (!owned || owned.consumed) invalidSessionTransport("The candidate is foreign, discarded or already consumed."); return owned;
  }
  function identity(): { documentId: string; revision: number } { check(); return { documentId: current.documentId, revision: current.revision }; }
  function receipt(reply: Reply): WorkbookSessionReceipt {
    if (!record(reply.value) || !keys(reply.value, ["receipt", "state"]) || !record(reply.value.receipt)) invalidSessionTransport("Invalid workbook mutation result.");
    acceptState(reply.value.state);
    const value = reply.value.receipt;
    if (value.documentId !== current.documentId || value.revision !== current.revision || typeof value.changed !== "boolean" || !Array.isArray(value.affectedNodeIds) || !record(value.mapping)) invalidSessionTransport("Invalid workbook mutation receipt.");
    return freezeSessionValue(value as unknown as WorkbookSessionReceipt);
  }
  port.addEventListener("message", receive); port.addEventListener("messageerror", invalidMessage); port.start();
  try {
    acceptState((await request("state", {}, signal ? { signal } : {})).value);
    return Object.freeze({
      state() { check(); return current; },
      async prepare(command: WorkbookSessionCommand, preparation?: TransactionPrepareOptions, execution?: WorkbookSessionExecution) {
        const { value } = await request("prepare", { command, ...(preparation === undefined ? {} : { options: preparation }) }, execution);
        if (!record(value) || !keys(value, ["handle", "candidate"]) || !positiveId(value.handle) || !record(value.candidate)
          || value.candidate.documentId !== current.documentId || !Number.isSafeInteger(value.candidate.baseRevision) || !Array.isArray(value.candidate.changes)
          || !Array.isArray(value.candidate.affectedNodeIds) || !record(value.candidate.mapping)) invalidSessionTransport("Invalid workbook candidate response.");
        const candidate = freezeSessionValue(value.candidate as unknown as WorkbookSessionCandidate); candidates.set(candidate, { handle: value.handle, consumed: false }); return candidate;
      },
      async readInput(sheetId: string, at: WorkbookCoordinate, candidate?: WorkbookSessionCandidate, execution?: WorkbookSessionExecution) {
        const coordinate = WorkbookCoordinateSchema.safeParse(at);
        if (!coordinate.success || typeof sheetId !== "string" || !sheetId || sheetId.length > 256) invalidSessionTransport("Invalid workbook input coordinate.");
        const point = coordinate.data, expected = candidate ? { documentId: candidate.documentId, revision: candidate.baseRevision } : identity(), owned = candidate ? candidateRecord(candidate) : undefined;
        const { value } = await request("inputText", { ...expected, sheetId, at: point, ...(owned ? { handle: owned.handle } : {}) }, execution);
        if (!record(value) || !keys(value, ["documentId", "revision", "sheetId", "at", "text"]) || value.documentId !== expected.documentId || value.revision !== expected.revision
          || value.sheetId !== sheetId || !record(value.at) || !keys(value.at, ["row", "column"]) || value.at.row !== point.row || value.at.column !== point.column
          || typeof value.text !== "string" || value.text.length > 32768) invalidSessionTransport("Invalid workbook input result.");
        for (const character of value.text) { const code = character.codePointAt(0)!; if (code === 0 || code >= 0xd800 && code <= 0xdfff) invalidSessionTransport("Workbook input returned malformed Unicode."); }
        if (current.revision !== expected.revision || owned?.consumed) throw new DocumentError("RevisionConflict", "Workbook input source changed before it became available.");
        return freezeSessionValue(value as unknown as WorkbookSessionInput);
      },
      async view(candidate?: WorkbookSessionCandidate, options: WorkbookSessionViewOptions = {}) {
        if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["signal", "onProgress", "calculate", "calculationContext", "previous"].includes(key))) invalidSessionTransport("Invalid workbook view options.");
        const { calculate, calculationContext, previous, ...execution } = options;
        const previousView = previous === undefined ? undefined : viewHandles.get(previous);
        if (previous !== undefined && (previousView === undefined || !views.has(previous))) invalidSessionTransport("The previous view must be retained by this workbook connection.");
        const expected = candidate ? { documentId: candidate.documentId, revision: candidate.baseRevision } : identity(), owned = candidate ? candidateRecord(candidate) : undefined;
        if (views.size >= config.maxViews) throw new DocumentError("ResourceLimitExceeded", "Too many retained workbook views.");
        const reply = await request("view", { ...expected, ...(owned ? { handle: owned.handle } : {}), ...(calculate === undefined ? {} : { calculate }), ...(calculationContext === undefined ? {} : { calculationContext }), ...(previousView === undefined ? {} : { previousView }) }, execution);
        if (!reply.port || !record(reply.value) || !keys(reply.value, ["documentId", "revision", "viewHandle", "impact"]) || !positiveId(reply.value.viewHandle)
          || reply.value.documentId !== expected.documentId || reply.value.revision !== expected.revision) { closeTransferred(reply.port); invalidSessionTransport("Invalid workbook view identity."); }
        let impact: WorkbookChangeImpact | undefined;
        if (reply.value.impact !== undefined) {
          const value = reply.value.impact;
          if (!previous || !record(value) || !keys(value, ["documentId", "fromRevision", "revision", "sourceRanges", "calculatedRanges"])
            || value.documentId !== expected.documentId || value.fromRevision !== previous.cells.snapshot.manifest.revision || value.revision !== expected.revision
            || !Array.isArray(value.sourceRanges) || !Array.isArray(value.calculatedRanges)) { closeTransferred(reply.port); invalidSessionTransport("Invalid workbook view invalidation."); }
          for (const ranges of [value.sourceRanges, value.calculatedRanges]) for (const item of ranges) {
            const range = record(item) ? WorkbookRangeSchema.safeParse(item.range) : undefined;
            if (!record(item) || !keys(item, ["sheetId", "range"]) || typeof item.sheetId !== "string" || !item.sheetId || item.sheetId.length > 256
              || !range?.success || range.data.from.row > range.data.to.row || range.data.from.column > range.data.to.column) { closeTransferred(reply.port); invalidSessionTransport("Invalid workbook view invalidation rectangle."); }
          }
          impact = freezeSessionValue(value as unknown as WorkbookChangeImpact);
        }
        const source = await connectWorkbookReads(reply.port, config.reads, execution?.signal);
        if (disposed || current.revision !== expected.revision || source.cells.snapshot.manifest.documentId !== expected.documentId || source.cells.snapshot.manifest.revision !== expected.revision || owned?.consumed) {
          source.dispose(); throw new DocumentError("RevisionConflict", "Workbook view source changed before it became available.");
        }
        const result: WorkbookReadConnection = Object.freeze({ ...source, ...(impact ? { impact } : {}), dispose() { views.delete(result); source.dispose(); } });
        views.add(result); viewHandles.set(result, reply.value.viewHandle);
        if (impact && previous) linkWorkbookPreviewSource(result, previous, impact);
        return result;
      },
      async commit(candidate: WorkbookSessionCandidate) {
        const owned = candidateRecord(candidate);
        if (candidate.baseRevision !== current.revision) throw new DocumentError("RevisionConflict", "The candidate belongs to an older workbook revision.");
        const result = receipt(await request("commit", { handle: owned.handle }, {}, true)); owned.consumed = true; return result;
      },
      async discard(candidate: WorkbookSessionCandidate) {
        const owned = candidateRecord(candidate);
        if (candidate.baseRevision === current.revision) await request("discard", { handle: owned.handle }, {}, true);
        owned.consumed = true;
      },
      async undo() { return receipt(await request("undo", identity(), {}, true)); },
      async redo() { return receipt(await request("redo", identity(), {}, true)); },
      async find(query: WorkbookFindQuery, options: WorkbookSessionFindOptions = {}) {
        if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["signal", "onProgress", "calculationContext", "locale"].includes(key))) invalidSessionTransport("Invalid workbook find options.");
        const { calculationContext, locale, ...execution } = options;
        const expected = identity(), { value } = await request("find", { ...expected, query, ...(calculationContext === undefined ? {} : { calculationContext }), ...(locale === undefined ? {} : { locale }) }, execution);
        if (!record(value) || value.documentId !== expected.documentId || value.baseRevision !== expected.revision || current.revision !== expected.revision || !Array.isArray(value.matches) || !Array.isArray(value.diagnostics) || typeof value.complete !== "boolean") invalidSessionTransport("Invalid workbook find result.");
        return freezeSessionValue(value as unknown as WorkbookFindResult);
      },
      subscribe(listener: (value: WorkbookSessionState) => void) { check(); if (typeof listener !== "function") invalidSessionTransport("Workbook session listeners must be functions."); listeners.add(listener); return () => { listeners.delete(listener); }; },
      dispose: () => dispose(),
    });
  } catch (reason) { dispose(); throw reason; }
}
