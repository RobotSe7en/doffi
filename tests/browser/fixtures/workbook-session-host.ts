import { DEFAULT_WORKBOOK_RULE_LIMITS } from "@doffi/core";
import { connectWorkbookSession } from "@doffi/editor";
import type { WorkbookSessionConnection } from "@doffi/editor";
import { mountDocumentPreview } from "@doffi/render";
import type { WorkbookPreview, WorkbookReadConnection } from "@doffi/render";
import { decodeTestBytes } from "./package-transport.js";

export async function editWorker(code: string, input: string, documentId: string, capacity = false, crossSheet = false) {
  const url = URL.createObjectURL(new Blob([code], { type: "text/javascript" })), worker = new Worker(url), channel = new MessageChannel(), abort = new AbortController();
  const container = document.createElement("div"); document.body.append(container);
  let client: WorkbookSessionConnection | undefined, source: WorkbookReadConnection | undefined, preview: WorkbookPreview | undefined;
  let ticks = 0, largestControlReplyUnits = 0, largestViewReplyUnits = 0, copied = 0, prepared = 0;
  const failures: string[] = [], timer = setInterval(() => ticks++, 10), states: number[] = [];
  let auditResolve: (value: Record<string, unknown>) => void = () => {};
  const audited = new Promise<Record<string, unknown>>(resolve => { auditResolve = resolve; });
  worker.onmessage = event => {
    if (event.data.kind === "error") { failures.push(JSON.stringify(event.data.error)); abort.abort(); }
    if (event.data.kind === "audit") auditResolve(event.data);
  };
  worker.onerror = () => { failures.push("Editing Worker failed"); abort.abort(); };
  channel.port2.addEventListener("message", event => {
    largestControlReplyUnits = Math.max(largestControlReplyUnits, JSON.stringify(event.data).length);
    if (event.data.port instanceof MessagePort) event.data.port.addEventListener("message", (event: MessageEvent<unknown>) => { largestViewReplyUnits = Math.max(largestViewReplyUnits, JSON.stringify(event.data).length); });
  });
  try {
    const bytes = decodeTestBytes(input); worker.postMessage({ kind: "open", bytes, documentId, port: channel.port1 }, [bytes.buffer, channel.port1]);
    client = await connectWorkbookSession(channel.port2, {}, abort.signal); client.subscribe(state => states.push(state.revision));
    source = await client.view(undefined, { signal: abort.signal });
    const initial = source.cells.snapshot, sheetId = initial.metadata.workbook.activeSheetId;
    preview = mountDocumentPreview(container, source, { widthPt: 420, heightPt: 240 }); await preview.ready;
    const first = preview.layout(), start = performance.now();
    const seed = await client.prepare({ kind: "transaction", transaction: { protocolVersion: 1, documentId, baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId, cells: [
      { row: 12000, column: 32, source: { kind: "value", value: { type: "number", value: 1 } } },
      { row: 12001, column: 32, source: { kind: "formula", expression: "=AG12001+1" } },
    ] }, ...(crossSheet ? [{ op: "insertWorkbookSheet" as const, name: "Copy target" }] : [])] } }, undefined, { signal: abort.signal });
    const seedPreparedMs = performance.now() - start, isolated = client.state().revision === 0 && client.state().history.undoSteps === 0;
    await client.commit(seed);
    let next = await client.view(undefined, { signal: abort.signal, previous: source });
    await preview.update(next); await preview.settled(); source.dispose(); source = next;
    const seededChunks = source.cells.snapshot.manifest.chunks;
    const targetSheetId = crossSheet ? source.cells.snapshot.metadata.workbook.sheets.find(sheet => sheet.name === "Copy target")!.sheetId : sheetId;
    const cancelling = new AbortController(); let cancelled = "", cancellationPoint = 0;
    try {
      const candidate = await client.prepare({ kind: "fill", documentId, revision: 1, sheetId, source: { from: { row: 12000, column: 32 }, to: { row: 12000, column: 32 } },
        target: { from: { row: 12002, column: 32 }, to: { row: 12601, column: 32 } } }, undefined, { signal: cancelling.signal, onProgress: value => {
          if (value.phase === "cells" && value.completed === 256) { cancellationPoint = value.completed; cancelling.abort(); }
        } });
      await client.discard(candidate); cancelled = "accepted";
    } catch (reason) { cancelled = (reason as { code: string }).code; }
    const cancellationState = client.state();
    const copyStart = performance.now(), copy = await client.prepare({ kind: "copy", documentId, revision: 1, sheetId,
      ...(crossSheet ? { targetSheetId } : {}),
      source: capacity ? { from: { row: 0, column: 0 }, to: { row: 9999, column: 19 } } : { from: { row: 12000, column: 32 }, to: { row: 12001, column: 32 } },
      target: capacity ? { row: 0, column: 1 } : { row: 12000, column: 33 } }, undefined, { signal: abort.signal, onProgress: value => { if (value.phase === "copy") copied++; else prepared++; } });
    const copyPreparedMs = performance.now() - copyStart; await client.commit(copy);
    next = await client.view(undefined, { signal: abort.signal, previous: source });
    const at = capacity ? { row: 9999, column: 20 } : { row: 12001, column: 33 };
    const current = await next.rules!.read({ documentId, revision: 2, sheetId: targetSheetId, coordinates: [at], includeCurrent: true, ...DEFAULT_WORKBOOK_RULE_LIMITS }, abort.signal);
    await preview.update(next); await preview.settled(); source.dispose(); source = next;
    if (crossSheet) await preview.showSheet(targetSheetId);
    await preview.selectRange({ from: at, to: at }); const shown = preview.layout(), selection = preview.selection();
    await client.undo(); next = await client.view(undefined, { signal: abort.signal });
    const copyUndo = JSON.stringify(next.cells.snapshot.manifest.chunks) === JSON.stringify(seededChunks); next.dispose();
    await client.undo(); client.dispose(); client = undefined;
    const auditTimeout = setTimeout(() => { failures.push("Worker close audit timed out"); auditResolve({}); }, 15_000);
    worker.postMessage({ kind: "closed-audit" }); const audit = await audited; clearTimeout(auditTimeout);
    return { first, shown, targetSheetId, selection, current: current.cells[0]?.current, isolated, seedPreparedMs, copyPreparedMs, cancelled, cancellationPoint, cancellationState,
      copyUndo, originalChunksRestored: JSON.stringify(audit.chunks) === JSON.stringify(initial.manifest.chunks), audit,
      detachedInput: bytes.byteLength === 0, largestControlReplyUnits, largestViewReplyUnits, copied, prepared, ticks, states, failures, userAgent: navigator.userAgent };
  } finally {
    clearInterval(timer); preview?.dispose(); source?.dispose(); client?.dispose(); channel.port2.close(); worker.terminate(); URL.revokeObjectURL(url); container.remove();
  }
}
Object.assign(window, { workbookSessionHost: { editWorker } });
