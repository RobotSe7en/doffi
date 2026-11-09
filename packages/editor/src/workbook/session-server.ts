import { captureWorkbookModel, copyWorkbookScalarContext, createWorkbookCalculator, DocumentError, readWorkbookChangeImpact, WorkbookCoordinateSchema, WorkbookRangeSchema } from "@doffi/core";
import type { WorkbookCandidateChange, WorkbookCalculationResult, WorkbookModel, WorkbookSession, TransactionPrepareOptions, WorkbookInteractionRequestInput } from "@doffi/core";
import { serveWorkbookReads } from "@doffi/render";
import type { WorkbookReadServer } from "@doffi/render";
import { createWorkbookRangeCopyTransaction } from "./copy-range.js";
import { createWorkbookSortTransaction } from "./sort.js";
import { createWorkbookTextFormatTransaction } from "./text-format.js";
import { createWorkbookFillTransaction } from "./fill.js";
import { createWorkbookReplaceTransaction } from "./replace.js";
import { findWorkbookText } from "./find.js";
import { createWorkbookInputTransaction, readWorkbookInputText } from "./input-source.js";
import { createWorkbookPasteTransaction } from "./text-input.js";
import { createWorkbookHtmlPastePlan } from "./html-paste.js";
import { createWorkbookClipboardData, WORKBOOK_CLIPBOARD_MAX_CELLS } from "./clipboard-data.js";
import { createWorkbookTaskQueue } from "./task-queue.js";
import { calculationOptions, checkSessionMessage, invalidSessionTransport, keys, positiveId, record, sessionMapping, sessionReceipt, sessionTransportConfig, WORKBOOK_SESSION_PROTOCOL } from "./session-transport.js";
import type { WorkbookSessionCommand, WorkbookSessionProgress, WorkbookSessionServer, WorkbookSessionState, WorkbookSessionTransportOptions } from "./session-transport.js";
import type { WorkbookFindQuery } from "./find-contracts.js";
import { requestWorkbookLayout } from "./session-layout-server.js";

/** One exclusive host-authorized endpoint, bound to the host's sole session.
 * Closing it releases its candidates/views, never the host session or Worker. */
export function serveWorkbookSession(port: MessagePort, session: WorkbookSession, options: WorkbookSessionTransportOptions = {}): WorkbookSessionServer {
  if (session.kind !== "workbook") invalidSessionTransport("Workbook session endpoints require their owning core session.");
  const config = sessionTransportConfig(options), pending = new Map<number, AbortController>(), jobs = new Map<number, Promise<void>>(), views = new Set<WorkbookReadServer>();
  const viewRequests = new Map<number, { server: WorkbookReadServer; model: WorkbookModel; calculation?: WorkbookCalculationResult }>();
  const candidates = new Map<number, { candidate: WorkbookCandidateChange; requestId: number; views: Set<WorkbookReadServer>; preparation?: TransactionPrepareOptions }>();
  let disposed = false, highestId = 0, nextCandidate = 0, preparing = false;
  const state = (): WorkbookSessionState => { const manifest = session.snapshot().manifest; return { documentId: manifest.documentId, revision: manifest.revision, history: session.history() }; };
  function send(message: Record<string, unknown>, transfer: Transferable[] = []): void {
    if (disposed) return;
    try { port.postMessage({ protocol: WORKBOOK_SESSION_PROTOCOL, ...message }, transfer); } catch { dispose(); }
  }
  function post(id: number, value: unknown): void { checkSessionMessage(value, config); send({ id, ok: true, value }); }
  function error(id: number, reason: unknown): void {
    const bounded = reason instanceof DocumentError ? reason : new DocumentError("InvalidOperation", "The authorized workbook session could not complete this command.");
    send({ id, ok: false, error: bounded.toJSON() });
  }
  function identity(payload: unknown): void {
    const current = state();
    if (!record(payload) || payload.documentId !== current.documentId || payload.revision !== current.revision) throw new DocumentError("RevisionConflict", "Workbook command belongs to another document revision.");
  }
  function release(id: number, discard = true): void {
    const owned = candidates.get(id); if (!owned) return; candidates.delete(id);
    for (const view of owned.views) { view.dispose(); views.delete(view); }
    if (discard) try { session.discard(owned.candidate); } catch { /* A host commit or disposal may have released this candidate. */ }
  }
  const unsubscribe = session.subscribe(event => {
    if (event.type !== "documentChanged") return;
    for (const id of [...candidates.keys()]) release(id);
    send({ event: "state", value: state() });
  });
  const tasks = createWorkbookTaskQueue(), yieldTask = (): Promise<void> => tasks.yield();
  async function calculate(model: WorkbookModel, abort: AbortController, progress: (value: WorkbookSessionProgress) => void, preparation?: TransactionPrepareOptions, guard?: () => void): Promise<WorkbookCalculationResult> {
    const calculator = captureWorkbookModel(session) === model ? undefined : createWorkbookCalculator(model);
    progress({ phase: "calculation", completed: 0, total: 0 });
    try { return await (calculator?.calculate ?? session.calculate)({ signal: abort.signal, yield: async () => { await yieldTask(); guard?.(); }, ...calculationOptions(preparation?.calculationContext),
      onProgress: value => progress({ phase: "calculation", completed: value.completed, total: value.total }) }); }
    finally { calculator?.dispose(); }
  }
  async function prepare(payload: Record<string, unknown>, abort: AbortController, progress: (value: WorkbookSessionProgress) => void, requestId: number): Promise<void> {
    if (!keys(payload, ["command", "options", "measureLayout"]) || !record(payload.command) || payload.measureLayout !== undefined && payload.measureLayout !== true) invalidSessionTransport("Invalid workbook preparation request.");
    if (preparing) throw new DocumentError("ResourceLimitExceeded", "Another workbook command is already preparing.");
    preparing = true; let owned: WorkbookCandidateChange | undefined, handle: number | undefined;
    try {
      const command = payload.command as unknown as WorkbookSessionCommand;
      let preparation = payload.options as TransactionPrepareOptions | undefined;
      const base = captureWorkbookModel(session), guard = (): void => {
        if (abort.signal.aborted || disposed) throw new DocumentError("OperationCancelled", "Workbook command was cancelled.");
        if (captureWorkbookModel(session) !== base) throw new DocumentError("RevisionConflict", "Workbook command source changed while preparing.");
      };
      if (payload.measureLayout) preparation = { ...preparation, measureWorkbookLayout: async request => {
        guard();
        const calculation = request.model.snapshot().metadata.workbook.calculationMode === "auto" ? await calculate(request.model, abort, progress,
          request.calculationContext ? { calculationContext: request.calculationContext } : undefined, guard) : undefined;
        guard();
        return requestWorkbookLayout(request, calculation, abort.signal, config, (value, readPort, replyPort) => {
          send({ id: requestId, event: "layout", value, readPort, replyPort }, [readPort, replyPort]);
        });
      } };
      let transaction, clipboardDiagnostics: readonly string[] | undefined;
      const common = { signal: abort.signal, yield: yieldTask };
      if (command.kind === "proposal") {
        if (!keys(payload.command, ["kind", "request", "proposal"])) invalidSessionTransport("Invalid workbook proposal command.");
        owned = await session.prepareProposalAsync(command.request, command.proposal, preparation, { ...common, onProgress: progress });
        const frozen = command.request.context.calculationContext;
        preparation = { ...preparation, calculationContext: { ...(frozen.now ? { now: frozen.now } : {}), ...(frozen.locale !== undefined ? { locale: frozen.locale } : {}) } };
      } else if (command.kind === "transaction") {
        if (!keys(payload.command, ["kind", "transaction"])) invalidSessionTransport("Invalid workbook transaction command.");
        transaction = command.transaction;
      } else {
        identity(command);
        if (command.kind === "input") {
          if (!keys(payload.command, ["kind", "documentId", "revision", "sheetId", "at", "text", "mode"])) invalidSessionTransport("Invalid workbook input command.");
          transaction = createWorkbookInputTransaction(base, command.sheetId, command.at, command.text, command.mode);
        } else if (command.kind === "formatText") {
          if (!keys(payload.command, ["kind", "documentId", "revision", "sheetId", "at", "selection", "properties"])) invalidSessionTransport("Invalid workbook text font command.");
          transaction = createWorkbookTextFormatTransaction(base, command.sheetId, command.at, command.selection, command.properties);
        } else if (command.kind === "paste") {
          if (!keys(payload.command, ["kind", "documentId", "revision", "sheetId", "at", "text", "html", "options"]) || typeof command.text !== "string") invalidSessionTransport("Invalid workbook paste command.");
          if (command.html !== undefined) {
            const pasted = createWorkbookHtmlPastePlan(base.snapshot(), command.sheetId, command.at, command.html, command.options);
            transaction = pasted.transaction; clipboardDiagnostics = pasted.diagnostics;
          } else transaction = createWorkbookPasteTransaction(base.snapshot(), command.sheetId, command.at, command.text, command.options);
        } else if (command.kind === "copy") {
          if (!keys(payload.command, ["kind", "documentId", "revision", "sheetId", "source", "target", "targetSheetId", "mode", "skipBlanks"])) invalidSessionTransport("Invalid workbook copy command.");
          const calculation = command.mode === "values" ? await calculate(base, abort, progress, preparation, guard) : undefined; guard();
          transaction = await createWorkbookRangeCopyTransaction(session, command.sheetId, command.source, command.target, { ...common,
            ...(command.targetSheetId === undefined ? {} : { targetSheetId: command.targetSheetId }),
            ...(command.mode === undefined ? {} : { mode: command.mode }), ...(command.skipBlanks === undefined ? {} : { skipBlanks: command.skipBlanks }),
            ...(calculation ? { calculation } : {}), onProgress: value => progress({ phase: "copy", ...value }) });
        } else if (command.kind === "fill") {
          if (!keys(payload.command, ["kind", "documentId", "revision", "sheetId", "source", "target", "mode", "axis"])) invalidSessionTransport("Invalid workbook fill command.");
          transaction = await createWorkbookFillTransaction(session, command.sheetId, command.source, command.target, { ...common,
            ...(command.mode === undefined ? {} : { mode: command.mode }), ...(command.axis === undefined ? {} : { axis: command.axis }), onProgress: value => progress({ phase: "fill", ...value }) });
        } else if (command.kind === "sort") {
          if (!keys(payload.command, ["kind", "documentId", "revision", "sheetId", "sort", "locale"])) invalidSessionTransport("Invalid workbook sort command.");
          const calculation = await calculate(base, abort, progress, preparation, guard); guard();
          transaction = await createWorkbookSortTransaction(session, command.sheetId, command.sort, { ...common, calculation,
            ...(command.locale === undefined ? {} : { locale: command.locale }), onProgress: value => progress({ phase: "sort", ...value }) });
        } else if (command.kind === "replace") {
          if (!keys(payload.command, ["kind", "documentId", "revision", "query", "replacement", "valueMode", "only"])) invalidSessionTransport("Invalid workbook replacement command.");
          transaction = await createWorkbookReplaceTransaction(session, command.query, command.replacement, { ...common,
            ...(command.valueMode === undefined ? {} : { valueMode: command.valueMode }), ...(command.only === undefined ? {} : { only: command.only }), onProgress: value => progress({ phase: "find", completed: value.cells, total: 0 }) });
        } else invalidSessionTransport("Unknown workbook preparation command.");
      }
      guard(); if (!owned) { if (!transaction) invalidSessionTransport("Workbook preparation requires a typed transaction."); owned = await session.prepareAsync(transaction, preparation, { ...common, onProgress: progress }); } guard();
      if (nextCandidate >= Number.MAX_SAFE_INTEGER) throw new DocumentError("ResourceLimitExceeded", "Workbook candidate handles are exhausted.");
      const retainedPreparation = preparation === undefined ? undefined : { ...preparation }; if (retainedPreparation) delete retainedPreparation.measureWorkbookLayout;
      handle = ++nextCandidate; candidates.set(handle, { candidate: owned, requestId, views: new Set(), ...(retainedPreparation === undefined ? {} : { preparation: retainedPreparation }) });
      const { documentId, baseRevision, affectedNodeIds, affectedSheetIds, changedRanges, changes, changedAnnotationIds, mapping } = owned;
      const review = session.getProposalReview(owned);
      post(requestId, { handle, candidate: { documentId, baseRevision, affectedNodeIds, ...(affectedSheetIds ? { affectedSheetIds } : {}), ...(changedRanges ? { changedRanges } : {}), changes, changedAnnotationIds, mapping: sessionMapping(mapping), ...(review ? { proposalReview: review } : {}), ...(clipboardDiagnostics ? { clipboardDiagnostics } : {}) } });
    } catch (reason) {
      if (handle !== undefined) release(handle); else if (owned) try { session.discard(owned); } catch { /* Disposed or invalidated by its host. */ }
      throw reason;
    } finally { preparing = false; }
  }
  function mutationResult(id: number, action: () => import("@doffi/core").TransactionReceipt): void {
    const receipt = action(), current = state();
    try { post(id, { receipt: sessionReceipt(receipt), state: current }); }
    catch { throw new DocumentError("ResourceLimitExceeded", `The workbook change was committed at revision ${current.revision}; its receipt exceeds the connection budget. Inspect the host session before retrying.`,
      [{ code: "committed-result-too-large", path: [current.documentId, current.revision], message: "The committed source remains authoritative." }]); }
  }
  async function handle(id: number, operation: string, payload: unknown, abort: AbortController): Promise<void> {
    if (!record(payload)) invalidSessionTransport("Workbook commands require an object payload.");
    const progress = (value: WorkbookSessionProgress): void => { if (!abort.signal.aborted) send({ id, event: "progress", value }); };
    if (operation === "state") { if (!keys(payload, [])) invalidSessionTransport("State requests have no fields."); post(id, state()); }
    else if (operation === "prepare") await prepare(payload, abort, progress, id);
    else if (operation === "interactionRequest") {
      if (!keys(payload, ["documentId", "revision", "input"]) || !record(payload.input)) invalidSessionTransport("Invalid workbook interaction request.");
      identity(payload);
      const request = await session.buildInteractionRequestAsync(payload.input as unknown as WorkbookInteractionRequestInput, { signal: abort.signal, yield: yieldTask, onProgress: progress });
      if (abort.signal.aborted || disposed) throw new DocumentError("OperationCancelled", "Workbook interaction request was cancelled.");
      identity(payload); post(id, request);
    }
    else if (operation === "commit" || operation === "discard") {
      if (!keys(payload, ["handle"]) || !positiveId(payload.handle)) invalidSessionTransport("Invalid workbook candidate handle.");
      const owned = candidates.get(payload.handle); if (!owned) invalidSessionTransport("The candidate does not belong to this connection or is no longer available.");
      if (operation === "discard") { release(payload.handle); post(id, null); }
      else mutationResult(id, () => { const receipt = session.commit(owned.candidate); release(payload.handle as number, false); return receipt; });
    } else if (operation === "undo" || operation === "redo") {
      if (!keys(payload, ["documentId", "revision"])) invalidSessionTransport("Invalid workbook history request."); identity(payload);
      mutationResult(id, () => session[operation]());
    } else if (operation === "reference") {
      if (!keys(payload, ["documentId", "revision", "expression", "origin"]) || typeof payload.expression !== "string" || !record(payload.origin)
        || !keys(payload.origin, ["sheetId", "row", "column"]) || typeof payload.origin.sheetId !== "string") invalidSessionTransport("Invalid workbook reference request.");
      identity(payload); const model = captureWorkbookModel(session);
      const coordinate = WorkbookCoordinateSchema.safeParse({ row: payload.origin.row, column: payload.origin.column });
      if (!coordinate.success) invalidSessionTransport("Invalid workbook reference origin.");
      const { documentId, revision } = model.snapshot().manifest;
      post(id, { documentId, revision, reference: model.resolveReference(payload.expression, { sheetId: payload.origin.sheetId, ...coordinate.data }) });
    } else if (operation === "clipboard") {
      if (!keys(payload, ["documentId", "revision", "sheetId", "range", "context", "view"]) || typeof payload.sheetId !== "string") invalidSessionTransport("Invalid workbook clipboard request.");
      identity(payload); const model = captureWorkbookModel(session), range = WorkbookRangeSchema.parse(payload.range);
      if (range.from.row > range.to.row || range.from.column > range.to.column || !model.snapshot().metadata.workbook.sheets.some(sheet => sheet.sheetId === payload.sheetId)) invalidSessionTransport("Clipboard output requires an ordered range on an existing worksheet.");
      if ((range.to.row - range.from.row + 1) * (range.to.column - range.from.column + 1) > WORKBOOK_CLIPBOARD_MAX_CELLS) throw new DocumentError("ResourceLimitExceeded", "Clipboard output exceeds its cell budget.");
      const context = payload.context === undefined ? undefined : copyWorkbookScalarContext(payload.context as never);
      const view = payload.view === undefined ? undefined : positiveId(payload.view) ? viewRequests.get(payload.view) : undefined;
      if (payload.view !== undefined && (!view || view.model !== model)) invalidSessionTransport("Clipboard results require a retained view of the current source.");
      const calculation = view ? view.calculation : await calculate(model, abort, progress, context === undefined ? undefined : { calculationContext: context });
      identity(payload);
      const data = await createWorkbookClipboardData(session, payload.sheetId, range, { ...(calculation ? { calculation } : {}), signal: abort.signal, yield: yieldTask, ...(context?.locale === undefined ? {} : { locale: context.locale }) });
      identity(payload); post(id, { ...state(), ...data });
    } else if (operation === "inputText") {
      if (!keys(payload, ["documentId", "revision", "sheetId", "at", "handle"])) invalidSessionTransport("Invalid workbook input read.");
      const owned = payload.handle === undefined ? undefined : positiveId(payload.handle) ? candidates.get(payload.handle) : undefined;
      if (payload.handle !== undefined && !owned) invalidSessionTransport("The input candidate is unavailable.");
      if (!owned) identity(payload);
      const model = owned?.candidate.model ?? captureWorkbookModel(session), snapshot = model.snapshot(), coordinate = WorkbookCoordinateSchema.safeParse(payload.at);
      if (payload.documentId !== snapshot.manifest.documentId || payload.revision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "Workbook input belongs to another source revision.");
      if (!coordinate.success || !snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === payload.sheetId)) invalidSessionTransport("Workbook input requires an existing worksheet coordinate.");
      const at = coordinate.data, sheetId = payload.sheetId as string;
      post(id, { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, sheetId, at, text: readWorkbookInputText(model.getCell(sheetId, at.row, at.column)) });
    } else if (operation === "view") {
      if (!keys(payload, ["documentId", "revision", "handle", "calculate", "calculationContext", "previousView"]) || payload.calculate !== undefined && typeof payload.calculate !== "boolean") invalidSessionTransport("Invalid workbook view request.");
      const previous = payload.previousView === undefined ? undefined : positiveId(payload.previousView) ? viewRequests.get(payload.previousView) : undefined;
      if (payload.previousView !== undefined && (!previous || previous.server.disposed)) invalidSessionTransport("The previous workbook view is unavailable or belongs to another connection.");
      const owned = payload.handle === undefined ? undefined : positiveId(payload.handle) ? candidates.get(payload.handle) : undefined;
      if (payload.handle !== undefined && !owned) invalidSessionTransport("The preview candidate is unavailable.");
      if (!owned) identity(payload);
      const model = owned?.candidate.model ?? captureWorkbookModel(session);
      const guard = (): void => { if (owned ? candidates.get(payload.handle as number) !== owned : captureWorkbookModel(session) !== model) throw new DocumentError("RevisionConflict", "Workbook view source changed before completion."); };
      const context = payload.calculationContext === undefined ? owned?.preparation?.calculationContext : copyWorkbookScalarContext(payload.calculationContext as never);
      const calculation = (payload.calculate ?? model.snapshot().metadata.workbook.calculationMode === "auto") ? await calculate(model, abort, progress,
        context === undefined ? undefined : { calculationContext: context }, guard) : undefined;
      if (abort.signal.aborted) throw new DocumentError("OperationCancelled", "Workbook view was cancelled.");
      if (owned ? candidates.get(payload.handle as number) !== owned : captureWorkbookModel(session) !== model) throw new DocumentError("RevisionConflict", "Workbook view source changed before completion.");
      for (const view of views) if (view.disposed) views.delete(view);
      for (const [requestId, view] of viewRequests) if (view.server.disposed) viewRequests.delete(requestId);
      if (views.size >= config.maxViews) throw new DocumentError("ResourceLimitExceeded", "Too many retained workbook views.");
      const value: Record<string, unknown> = { documentId: model.snapshot().manifest.documentId, revision: model.snapshot().manifest.revision, viewHandle: id };
      checkSessionMessage(value, config);
      if (previous && !previous.server.disposed) {
        const impact = readWorkbookChangeImpact(previous.model, model, previous.calculation, calculation);
        try { checkSessionMessage({ ...value, impact }, config); value.impact = impact; }
        catch (reason) { if (!(reason instanceof DocumentError) || reason.code !== "ResourceLimitExceeded") throw reason; }
      }
      const channel = new MessageChannel(); let view: WorkbookReadServer;
      try { view = serveWorkbookReads(channel.port1, model, config.reads, calculation); }
      catch (reason) { channel.port1.close(); channel.port2.close(); throw reason; }
      views.add(view); viewRequests.set(id, { server: view, model, ...(calculation ? { calculation } : {}) }); owned?.views.add(view);
      send({ id, ok: true, value, port: channel.port2 }, [channel.port2]);
    } else if (operation === "find") {
      if (!keys(payload, ["documentId", "revision", "query", "calculationContext", "locale", "viewHandle"])) invalidSessionTransport("Invalid workbook find request."); identity(payload);
      const model = captureWorkbookModel(session), query = payload.query as WorkbookFindQuery;
      const view = payload.viewHandle === undefined ? undefined : positiveId(payload.viewHandle) ? viewRequests.get(payload.viewHandle) : undefined;
      if (payload.viewHandle !== undefined && (!view || view.server.disposed || view.model !== model)) throw new DocumentError("RevisionConflict", "Workbook find requires a retained view of its actual current source.");
      if (view && payload.calculationContext !== undefined) invalidSessionTransport("Workbook find cannot replace a retained view's calculation context.");
      if (view && query?.field === "display" && !view.calculation) invalidSessionTransport("Display search requires a view with completed calculation.");
      const context = payload.calculationContext === undefined ? undefined : copyWorkbookScalarContext(payload.calculationContext as never);
      const locale = payload.locale === undefined ? context?.locale : copyWorkbookScalarContext({ locale: payload.locale as string }).locale;
      const calculation = query?.field === "display" ? view?.calculation ?? await calculate(model, abort, progress, context === undefined ? undefined : { calculationContext: context }, () => { if (captureWorkbookModel(session) !== model) throw new DocumentError("RevisionConflict", "Workbook find source changed while calculating."); }) : undefined;
      const result = await findWorkbookText(session, query, { signal: abort.signal, yield: yieldTask, ...(calculation ? { calculation } : {}),
        ...(locale === undefined ? {} : { locale }),
        onProgress: value => progress({ phase: "find", completed: value.cells, total: 0 }) });
      if (captureWorkbookModel(session) !== model) throw new DocumentError("RevisionConflict", "Workbook find source changed before completion."); post(id, result);
    } else invalidSessionTransport("Unknown workbook session command.");
  }
  function receive(event: MessageEvent<unknown>): void {
    const message = event.data;
    if (!record(message) || message.protocol !== WORKBOOK_SESSION_PROTOCOL) return;
    if (message.event === "close" && keys(message, ["protocol", "event"])) { dispose(); return; }
    if (!positiveId(message.id)) return;
    const id = message.id;
    if (message.operation === "cancel" && keys(message, ["protocol", "id", "operation"])) {
      pending.get(id)?.abort(); for (const [handle, owned] of candidates) if (owned.requestId === id) release(handle);
      const view = viewRequests.get(id)?.server; view?.dispose(); if (view) views.delete(view); viewRequests.delete(id);
      const work = jobs.get(id); if (work) void work.then(() => send({ id, event: "cancelled" })); else send({ id, event: "cancelled" }); return;
    }
    try {
      if (!keys(message, ["protocol", "id", "operation", "payload"]) || typeof message.operation !== "string") invalidSessionTransport("Invalid workbook session envelope.");
      if (id <= highestId) invalidSessionTransport("Workbook session request IDs must increase."); highestId = id;
      checkSessionMessage(message.payload, config);
      if (pending.size >= config.maxPendingRequests) throw new DocumentError("ResourceLimitExceeded", "Too many pending workbook session requests.");
      const abort = new AbortController(); pending.set(id, abort);
      const timer = setTimeout(() => abort.abort(), config.requestTimeoutMs);
      const work = handle(id, message.operation, message.payload, abort).catch(reason => error(id, reason)).finally(() => { clearTimeout(timer); pending.delete(id); jobs.delete(id); });
      jobs.set(id, work);
    } catch (reason) { error(id, reason); }
  }
  function dispose(): void {
    if (disposed) return; disposed = true; unsubscribe(); tasks.dispose();
    for (const abort of pending.values()) abort.abort(); pending.clear();
    for (const handle of [...candidates.keys()]) release(handle);
    for (const view of views) view.dispose(); views.clear(); viewRequests.clear();
    try { port.postMessage({ protocol: WORKBOOK_SESSION_PROTOCOL, event: "close" }); } catch { /* A failed port is still closed. */ }
    port.removeEventListener("message", receive); port.removeEventListener("messageerror", dispose); port.close();
  }
  port.addEventListener("message", receive); port.addEventListener("messageerror", dispose); port.start();
  return Object.freeze({ get disposed() { return disposed; }, dispose });
}
