import type { WorkbookSnapshot } from '@doffi/core';
import { connectWorkbookSession, mountWorkbookEditor } from '@doffi/editor';
import type { WorkbookRemoteEditorView, WorkbookSessionConnection } from '@doffi/editor';
import { decodeTestBytes } from './package-transport.js';

/** Observe the test's closed surfaces without opening production shadows or
 * retaining every prior window. The active and preparing surfaces are bounded. */
function observeSurfaces() {
  const parents: ShadowRoot[] = [], frames = new Map<HTMLIFrameElement, { roots: ShadowRoot[]; release(): void }>();
  let maximumFragments = 0;
  const attach = Element.prototype.attachShadow;
  const observer = new MutationObserver(scan);
  Element.prototype.attachShadow = function(init) { const root = attach.call(this, init); parents.push(root); observer.observe(root, { childList: true, subtree: true }); return root; };
  function scan() {
    for (const [frame, record] of frames) if (!frame.isConnected) { record.release(); frames.delete(frame); }
    for (const parent of parents) for (const frame of parent.querySelectorAll('iframe')) {
      if (frames.has(frame)) continue;
      let restore = () => {}; const roots: ShadowRoot[] = [];
      const loaded = () => {
        const prototype = (frame.contentWindow as Window & typeof globalThis).Element.prototype, original = prototype.attachShadow;
        prototype.attachShadow = function(init) {
          const root = original.call(this, init);
          if (this.localName === 'office-workbook-view-surface') {
            for (let index = roots.length - 1; index >= 0; index--) if (!roots[index]!.host.isConnected) roots.splice(index, 1);
            roots.push(root);
          }
          return root;
        };
        restore = () => { prototype.attachShadow = original; };
      };
      frame.addEventListener('load', loaded, { capture: true, once: true });
      frames.set(frame, { roots, release() { frame.removeEventListener('load', loaded, true); restore(); roots.length = 0; } });
    }
    const fragments = [...frames.values()].reduce((sum, record) => sum + record.roots.filter(root => root.host.isConnected).reduce((total, root) => total + root.querySelectorAll('[data-office-view="workbook-cell"],[data-office-view="workbook-text-overflow"]').length, 0), 0);
    maximumFragments = Math.max(maximumFragments, fragments);
  }
  return {
    fragments() { scan(); return maximumFragments; },
    frame() { scan(); const active = [...frames.keys()].filter(frame => frame.isConnected && frame.style.visibility === 'visible'); if (active.length !== 1) throw Error('Expected one accepted editor iframe'); return active[0]!; },
    cell(row: number, column: number) {
      scan();
      for (const { roots } of frames.values()) for (const root of roots) if (root.host.isConnected) {
        const cell = root.querySelector(`[data-office-view="workbook-cell"][data-row="${row}"][data-column="${column}"]`);
        if (cell) return cell.textContent?.trim();
      }
      return undefined;
    },
    point(index: number) {
      const frame = this.frame(), root = frames.get(frame)!.roots.find(root => root.host.isConnected && root.querySelector('[role="grid"]'))!;
      frame.scrollIntoView({ block: 'center' });
      const candidates = [...root.querySelectorAll<HTMLElement>('[role="gridcell"]')].filter(cell => {
        const rect = cell.getBoundingClientRect(), hit = root.elementFromPoint(rect.x + rect.width / 2, rect.y + rect.height / 2);
        return rect.width > 0 && rect.height > 0 && hit !== null && cell.contains(hit);
      });
      if (candidates.length < 3) throw Error('Native feedback requires unobscured grid cells.');
      const cell = candidates[(index + 1) % Math.min(5, candidates.length)]!, local = cell.getBoundingClientRect(), outer = frame.getBoundingClientRect();
      return { x: outer.x + local.x + local.width / 2, y: outer.y + local.y + local.height / 2, row: Number(cell.dataset.row), column: Number(cell.dataset.column) };
    },
    input() { const input = parents.flatMap(root => [...root.querySelectorAll('textarea')])[0]; if (!input) throw Error('Editor input is missing'); return input; },
    button(label: string) { const button = parents.flatMap(root => [...root.querySelectorAll('button')]).find(button => button.textContent === label); if (!button || button.disabled) throw Error('Editor control is unavailable: ' + label); return button; },
    dispose() { observer.disconnect(); Element.prototype.attachShadow = attach; for (const record of frames.values()) record.release(); frames.clear(); parents.length = 0; },
  };
}
interface NativeFeedback { selections: number[]; inputs: number[]; restored: boolean }
let feedback: { point(index: number): { x: number; y: number }; sample(kind: 'selection' | 'input', index: number): Promise<number>; focus(): Promise<void>; complete(): Promise<void>; cancel(): void } | undefined;
let status: (() => Record<string, unknown>) | undefined, stop: (() => void) | undefined, lastStatus: Record<string, unknown> = { phase: 'idle' };
export function capacityStatus() { return status?.() ?? lastStatus; }
export function capacityCancel() { stop?.(); }
export const feedbackReady = () => Boolean(feedback);
export function feedbackPoint(index: number) { if (!feedback) throw Error('Native feedback is not ready.'); return feedback.point(index); }
export function feedbackSample(kind: 'selection' | 'input', index: number) { if (!feedback) throw Error('Native feedback is not ready.'); return feedback.sample(kind, index); }
export async function feedbackFocus() { if (!feedback) throw Error('Native feedback is not ready.'); await feedback.focus(); }
export async function feedbackComplete() { if (!feedback) throw Error('Native feedback is not ready.'); await feedback.complete(); }
export function feedbackCancel() { feedback?.cancel(); }

/** Actual mounted editor and visible dependent formula. This diagnostic does
 * not certify the fixed-resource stage's CPU, memory or latency limits. */
export async function ordinaryMountedWorkerEdits(code: string, encoded: string, options: { performance?: boolean; cancelBulk?: boolean; iterations?: number; nativeFeedback?: boolean } = {}) {
  const iterations = options.iterations ?? 100;
  if (!Number.isInteger(iterations) || iterations < 1 || iterations > 100) throw Error('Invalid diagnostic sample count.');
  const documentId = 'ordinary-mounted-editor', url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' })), worker = new Worker(url), channel = new MessageChannel();
  const host = document.createElement('div'); document.body.append(host);
  let client: WorkbookSessionConnection | undefined, editor: WorkbookRemoteEditorView | undefined, snapshot: WorkbookSnapshot | undefined;
  let captureRows = false, maximumReadRow = -1, largestReplyUnits = 0, sameFrame = true, ticks = 0;
  const failures: unknown[] = [], samples: { inputMs: number; previewMs: number; totalMs: number }[] = [], surfaces = observeSurfaces();
  let opening: Record<string, number> | undefined;
  let selectionChanged = () => {};
  let phase = 'native-open', fail!: (error: unknown) => void, progress: unknown;
  const fatal = new Promise<never>((_resolve, reject) => { fail = reject; }); void fatal.catch(() => {});
  status = () => ({ phase, opening, progress, samples: samples.length, failures: [...failures] });
  stop = () => { feedback?.cancel(); fail(Error('Capacity measurement cancelled.')); editor?.dispose(); client?.dispose(); worker.terminate(); };
  const timer = setInterval(() => ticks++, 10); let auditResolve!: (value: Record<string, unknown>) => void;
  const audited = new Promise<Record<string, unknown>>(resolve => { auditResolve = resolve; });
  worker.onmessage = event => { if (event.data.kind === 'progress') progress = event.data.progress; if (event.data.kind === 'opening') opening = event.data; if (event.data.kind === 'audit') auditResolve(event.data); if (event.data.kind === 'error') { failures.push(event.data.error); fail(Error(JSON.stringify(event.data.error))); } };
  worker.onerror = () => { failures.push('Worker failure'); fail(Error('Capacity Worker failed.')); };
  channel.port2.addEventListener('message', event => {
    largestReplyUnits = Math.max(largestReplyUnits, JSON.stringify(event.data).length);
    if (event.data.port instanceof MessagePort) event.data.port.addEventListener('message', (reply: MessageEvent<{ value?: { cells?: { row: number }[] } }>) => {
      largestReplyUnits = Math.max(largestReplyUnits, JSON.stringify(reply.data).length);
      if (captureRows && Array.isArray(reply.data.value?.cells)) for (const cell of reply.data.value.cells) maximumReadRow = Math.max(maximumReadRow, cell.row);
    });
  });
  try {
    const start = performance.now(), bytes = decodeTestBytes(encoded);
    worker.postMessage({ kind: 'open', documentId, bytes, performance: options.performance, port: channel.port1 }, [bytes.buffer, channel.port1]);
    client = await Promise.race([connectWorkbookSession(channel.port2), fatal]); phase = 'mounting-editor';
    editor = mountWorkbookEditor(host, client, { widthPt: 420, heightPt: 240,
      prepareOptions: () => ({ calculationContext: { now: { unixMilliseconds: Date.now(), utcOffsetMinutes: -new Date().getTimezoneOffset() } } }),
      onSelectionChange() { selectionChanged(); },
      onViewChange(value) { snapshot = value; }, onError(error) { failures.push({ message: error.message, code: (error as { code?: string }).code }); fail(error); } });
    await Promise.race([editor.ready, fatal]); phase = 'native-feedback'; const firstScreenMs = performance.now() - start, initial = snapshot!, originalFrame = surfaces.frame(), sheetId = initial.metadata.workbook.activeSheetId;
    if (options.iterations !== undefined) console.info('Editing diagnostic ready', firstScreenMs, opening);
    let nativeFeedback: NativeFeedback | undefined;
    if (options.nativeFeedback) {
      const selections: number[] = [], inputs: number[] = [], listeners = new AbortController(), pending = new Map<string, { resolve(value: number): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> }>();
      let start: number | undefined, keyStart: number | undefined, expected: { row: number; column: number } | undefined, finish!: (value: NativeFeedback) => void, cancel!: (error: Error) => void, pointers = 0;
      const done = new Promise<NativeFeedback>((resolve, reject) => { finish = resolve; cancel = reject; });
      const publish = (kind: 'selection' | 'input', elapsed: number) => {
        const samples = kind === 'selection' ? selections : inputs, key = kind + samples.push(elapsed), waiter = pending.get(key);
        if (waiter) { clearTimeout(waiter.timer); pending.delete(key); waiter.resolve(elapsed); }
      };
      const timestamp = (event: Event, view: Window) => {
        const delay = view.performance.now() - event.timeStamp;
        if (!event.isTrusted || delay < -2 || delay > 60_000) throw Error('Native feedback timestamp is not calibrated.');
        return performance.now() - Math.max(0, delay);
      };
      const document = originalFrame.contentDocument!, input = surfaces.input();
      document.addEventListener('pointerdown', event => { pointers++; start = timestamp(event, document.defaultView!); }, { capture: true, signal: listeners.signal });
      input.addEventListener('keydown', event => { if (event.isTrusted && event.key === '1') keyStart = timestamp(event, window); }, { capture: true, signal: listeners.signal });
      input.addEventListener('input', event => {
        if (!event.isTrusted) return; if (keyStart === undefined) { cancel(Error('Native input feedback requires its trusted key event.')); return; }
        const before = keyStart; keyStart = undefined;
        requestAnimationFrame(() => { const state = editor!.inputState(); if (input.value !== '1' || state.text !== '1' || !state.dirty) cancel(Error('Native edit feedback did not retain the entered draft.')); else publish('input', performance.now() - before); });
      }, { signal: listeners.signal });
      selectionChanged = () => {
        if (start === undefined) return; const before = start; start = undefined;
        const anchor = editor!.selection()?.anchor, actual = anchor?.kind === 'cells' ? anchor.ranges[0]!.from : undefined;
        if (!actual || !expected || actual.row !== expected.row || actual.column !== expected.column) { cancel(Error('Native feedback selected a different source cell: ' + JSON.stringify({ expected, actual }))); return; }
        requestAnimationFrame(() => publish('selection', performance.now() - before));
      };
      feedback = { point(index) { const point = surfaces.point(index); expected = point; return point; }, sample(kind, index) {
        const values = kind === 'selection' ? selections : inputs; if (values[index] !== undefined) return Promise.resolve(values[index]!);
        const key = kind + (index + 1);
        return new Promise((resolve, reject) => { const timer = setTimeout(() => { pending.delete(key); reject(Error('Native feedback timed out: ' + JSON.stringify({ kind, index, pointers, selections, inputs, selection: editor!.selection(), point: surfaces.point(index), frame: surfaces.frame().getBoundingClientRect().toJSON() }))); }, 15000); pending.set(key, { resolve, reject, timer }); });
      }, async focus() { await editor!.selectRange(sheetId, { from: { row: 0, column: 2 }, to: { row: 0, column: 2 } }); await editor!.settled(); editor!.focusInput(); },
      async complete() { await editor!.settled(); finish({ selections, inputs, restored: input.value === '22' && client!.state().history.undoSteps === 0 }); }, cancel() { cancel(Error('Native feedback was cancelled.')); } };
      try { nativeFeedback = await done; }
      finally { feedback = undefined; selectionChanged = () => {}; listeners.abort(); for (const waiter of pending.values()) { clearTimeout(waiter.timer); waiter.reject(Error('Native feedback closed.')); } pending.clear(); }
    }
    const at = { row: 0, column: 2 }, dependent = { row: 0, column: 19 };
    // Leave T1 on screen while the formula bar edits C1. Each sample verifies
    // the editor's own rendered batch without a second verification calculation.
    await editor.selectRange(sheetId, { from: dependent, to: dependent }); const viewport = editor.layout().viewport;
    await editor.selectRange(sheetId, { from: at, to: at }); await editor.setViewport(viewport); await editor.settled();
    if (surfaces.cell(0, 19) !== '492') throw Error('Initial dependent formula is not visible and current');
    const input = surfaces.input(), setupMs = performance.now() - start - firstScreenMs; captureRows = true;
    for (let index = 0; index < iterations; index++) {
      phase = 'edit-' + index;
      if (options.iterations !== undefined) performance.mark('doffi-edit-' + index + '-start');
      const before = performance.now(); input.value = String(100 + index); input.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
      const receipt = await editor.applyInput(); if (!receipt?.changed) throw Error('Mounted editor input did not commit: ' + JSON.stringify({ index, input: editor.inputState(), revision: client.state().revision, failures })); const committed = performance.now();
      await editor.settled(); const shown = performance.now();
      if (surfaces.cell(0, 19) !== String(570 + index)) throw Error('Mounted editor displayed an incorrect dependent formula');
      sameFrame &&= surfaces.frame() === originalFrame;
      samples.push({ inputMs: committed - before, previewMs: shown - committed, totalMs: shown - before });
      if (options.iterations !== undefined) performance.mark('doffi-edit-' + index + '-end');
      if (options.iterations !== undefined) console.info('Editing diagnostic sample', index, samples.at(-1));
    }
    if (options.iterations !== undefined) console.info('Editing diagnostic edits complete');
    captureRows = false; const fullHistory = client.state().history;
    async function history(label: string) {
      phase = 'history-' + label;
      for (let index = 0; index < iterations; index++) {
        const revision = client!.state().revision; surfaces.button(label).click(); await editor!.settled();
        if (client!.state().revision !== revision + 1) throw Error('Mounted editor history did not advance exactly one revision');
      }
    }
    await history('撤销'); const undoneHistory = client.state().history, undoneText = surfaces.cell(0, 19);
    await history('重做'); const redoneHistory = client.state().history, redoneText = surfaces.cell(0, 19);
    await history('撤销'); const originalChunksRestored = JSON.stringify(snapshot!.manifest.chunks) === JSON.stringify(initial.manifest.chunks), restoredInput = input.value, restoredText = surfaces.cell(0, 19);
    const cancellations: { command: string; error: string; progress: number; unchanged: boolean }[] = [];
    if (options.cancelBulk) {
      const commands = [
        { kind: 'sort' as const, documentId, revision: client.state().revision, sheetId, sort: { range: { from: { row: 0, column: 0 }, to: { row: 9999, column: 19 } }, method: 'none' as const, caseSensitive: false, keys: [{ column: 2, descending: true }] } },
        { kind: 'transaction' as const, transaction: { protocolVersion: 1 as const, documentId, baseRevision: client.state().revision, operations: [{ op: 'setWorkbookCells' as const, sheetId,
          cells: Array.from({ length: 10000 }, (_, row) => ({ row, column: 0, source: { kind: 'value' as const, value: { type: 'number' as const, value: row + 1 } } })) }] } },
      ];
      for (const command of commands) {
        const before = JSON.stringify(client.state()), abort = new AbortController(); let error = '', progress = 0;
        try {
          const candidate = await client.prepare(command, undefined, { signal: abort.signal, onProgress(value) {
            if ((value.phase === 'sort' || value.phase === 'cells') && value.completed >= 256) { progress = value.completed; abort.abort(); }
          } });
          await client.discard(candidate); error = 'accepted';
        } catch (reason) { error = (reason as { code: string }).code; }
        cancellations.push({ command: command.kind, error, progress, unchanged: before === JSON.stringify(client.state()) });
      }
    }
    phase = 'closed-audit'; const finalRevision = snapshot!.manifest.revision, maximumFragments = surfaces.fragments(); editor.dispose(); editor = undefined; client.dispose(); client = undefined;
    const timeout = setTimeout(() => { failures.push('Audit timeout'); auditResolve({}); }, 15_000);
    worker.postMessage({ kind: 'closed-audit' }); const audit = await audited; clearTimeout(timeout);
    const sorted = samples.map(sample => sample.totalMs).sort((a, b) => a - b);
    return { scope: 'Actual mounted public Worker editor and visible dependent formulas; its owner records execution resource constraints separately.',
      firstScreenMs, opening, setupMs, samples, editAndPreviewP95Ms: sorted[Math.ceil(sorted.length * .95) - 1]!, maximumMs: sorted.at(-1)!,
      maximumReadRow, maximumFragments, largestReplyUnits, sameFrame, originalChunksRestored, restoredInput, undoneText, redoneText, restoredText, finalRevision,
      fullHistory, undoneHistory, redoneHistory, nativeFeedback, cancellations, audit, ticks, failures, userAgent: navigator.userAgent };
  } finally { lastStatus = status!(); status = undefined; stop = undefined; clearInterval(timer); editor?.dispose(); client?.dispose(); surfaces.dispose(); channel.port2.close(); worker.terminate(); URL.revokeObjectURL(url); host.remove(); }
}
Object.assign(window, { workbookEditorPerformance: { ordinaryMountedWorkerEdits, feedbackReady, feedbackPoint, feedbackSample, feedbackFocus, feedbackComplete, feedbackCancel, capacityStatus, capacityCancel } });
