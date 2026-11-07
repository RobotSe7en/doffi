import { DEFAULT_WORKBOOK_RULE_LIMITS } from '@doffi/core';
import { bindWorkbookAsyncTextInput, connectWorkbookSession, createWorkbookRemoteInputController } from '@doffi/editor';
import type { WorkbookSessionConnection } from '@doffi/editor';
import type { WorkbookAsyncTextBinding, WorkbookRemoteInputController } from '@doffi/editor';
import { mountDocumentPreview } from '@doffi/render';
import type { WorkbookPreview, WorkbookReadConnection } from '@doffi/render';
import { decodeTestBytes } from './package-transport.js';

/** Public Worker/session/preview path. Samples are not a fixed-resource SLA. */
export async function ordinaryWorkerEdits(code: string, encoded: string) {
  const documentId = 'ordinary-editing-worker', url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' }));
  const worker = new Worker(url), channel = new MessageChannel(), abort = new AbortController(), host = document.createElement('div'); document.body.append(host);
  let client: WorkbookSessionConnection | undefined, source: WorkbookReadConnection | undefined, preview: WorkbookPreview | undefined;
  let binding: WorkbookAsyncTextBinding | undefined, controller: WorkbookRemoteInputController | undefined;
  let captureRows = false, largestReplyUnits = 0, maximumReadRow = -1, ticks = 0, sameFrame = true;
  const samples: { inputMs: number; viewMs: number; previewMs: number; totalMs: number }[] = [], failures: unknown[] = [];
  const timer = setInterval(() => ticks++, 10); let auditResolve!: (value: Record<string, unknown>) => void;
  const audited = new Promise<Record<string, unknown>>(resolve => { auditResolve = resolve; });
  worker.onmessage = event => {
    if (event.data.kind === 'audit') auditResolve(event.data);
    if (event.data.kind === 'error') { failures.push(event.data.error); abort.abort(); }
  };
  worker.onerror = () => { failures.push('Worker failure'); abort.abort(); };
  channel.port2.addEventListener('message', event => {
    largestReplyUnits = Math.max(largestReplyUnits, JSON.stringify(event.data).length);
    if (event.data.port instanceof MessagePort) event.data.port.addEventListener('message', (reply: MessageEvent<{ value?: { cells?: { row: number }[] } }>) => {
      largestReplyUnits = Math.max(largestReplyUnits, JSON.stringify(reply.data).length);
      if (captureRows && Array.isArray(reply.data.value?.cells)) for (const cell of reply.data.value.cells) if (typeof cell.row === 'number') maximumReadRow = Math.max(maximumReadRow, cell.row);
    });
  });
  try {
    const started = performance.now(), bytes = decodeTestBytes(encoded);
    worker.postMessage({ kind: 'open', documentId, bytes, port: channel.port1 }, [bytes.buffer, channel.port1]);
    client = await connectWorkbookSession(channel.port2, {}, abort.signal); source = await client.view(undefined, { signal: abort.signal });
    const initial = source.cells.snapshot, sheetId = initial.metadata.workbook.activeSheetId, at = { row: 0, column: 2 };
    const attach = Element.prototype.attachShadow; let chrome!: ShadowRoot;
    Element.prototype.attachShadow = function(init) { const root = attach.call(this, init); chrome ??= root; return root; };
    try { preview = mountDocumentPreview(host, source, { widthPt: 420, heightPt: 240, retainSelection: true }); }
    finally { Element.prototype.attachShadow = attach; }
    await preview.ready; await preview.selectRange({ from: at, to: at });
    const input = document.createElement('textarea'); input.setAttribute('aria-label', 'Capacity input'); host.append(input);
    controller = createWorkbookRemoteInputController(client); await controller.select({ documentId, baseRevision: client.state().revision,
      anchor: { kind: 'cells', sheetId, ranges: [{ from: at, to: at }] } });
    binding = bindWorkbookAsyncTextInput(input, controller, { onError(error) { failures.push(error.toJSON()); } });
    const firstScreenMs = performance.now() - started, originalFrame = chrome.querySelector('iframe'); captureRows = true;
    for (let index = 0; index < 100; index++) {
      input.value = String(100 + index); const start = performance.now(), receipt = await binding.apply();
      if (!receipt) throw Error('Ordinary native input did not commit.');
      const committed = performance.now();
      const next = await client.view(undefined, { previous: source, signal: abort.signal }), viewed = performance.now();
      preview.update(next); await preview.settled(); await preview.selectRange({ from: at, to: at }, false);
      const shown = performance.now(); source.dispose(); source = next;
      const result = await source.rules!.read({ documentId, revision: client.state().revision, sheetId, includeCurrent: true,
        coordinates: [{ row: 0, column: 19 }], ...DEFAULT_WORKBOOK_RULE_LIMITS }, abort.signal);
      const current = result.cells[0]?.current;
      if (current?.kind !== 'value' || current.value.type !== 'number' || current.value.value !== 570 + index) throw Error('Ordinary edit did not calculate its dependent formula.');
      sameFrame &&= chrome.querySelector('iframe') === originalFrame;
      samples.push({ inputMs: committed - start, viewMs: viewed - committed, previewMs: shown - viewed, totalMs: performance.now() - start });
    }
    captureRows = false; const fullHistory = client.state().history;
    for (let index = 0; index < 100; index++) await client.undo();
    const undoneHistory = client.state().history;
    for (let index = 0; index < 100; index++) await client.redo();
    const redoneHistory = client.state().history;
    for (let index = 0; index < 100; index++) await client.undo();
    const restored = await client.view(undefined, { previous: source, signal: abort.signal });
    const originalChunksRestored = JSON.stringify(restored.cells.snapshot.manifest.chunks) === JSON.stringify(initial.manifest.chunks);
    preview.update(restored); await preview.settled(); await controller.select({ documentId, baseRevision: client.state().revision,
      anchor: { kind: 'cells', sheetId, ranges: [{ from: at, to: at }] } }); binding.refresh(); const restoredInput = input.value;
    binding.dispose(); controller.dispose(); binding = undefined; controller = undefined; restored.dispose();
    preview.dispose(); preview = undefined; source.dispose(); source = undefined; client.dispose(); client = undefined;
    const timeout = setTimeout(() => { failures.push('Audit timeout'); auditResolve({}); }, 15_000);
    worker.postMessage({ kind: 'closed-audit' }); const audit = await audited; clearTimeout(timeout);
    const sorted = samples.map(sample => sample.totalMs).sort((a, b) => a - b);
    return { scope: 'Unrestricted actual browser/Worker diagnostic; fixed resources and default editor UI remain separate.', firstScreenMs, samples,
      editAndPreviewP95Ms: sorted[Math.ceil(sorted.length * .95) - 1]!, maximumMs: sorted.at(-1)!, maximumReadRow,
      largestReplyUnits, sameFrame, fullHistory, undoneHistory, redoneHistory, originalChunksRestored, restoredInput, audit, ticks, failures, userAgent: navigator.userAgent };
  } finally {
    clearInterval(timer); binding?.dispose(); controller?.dispose(); preview?.dispose(); source?.dispose(); client?.dispose(); channel.port2.close(); worker.terminate(); URL.revokeObjectURL(url); host.remove();
  }
}
Object.assign(window, { workbookEditingPerformance: { ordinaryWorkerEdits } });
