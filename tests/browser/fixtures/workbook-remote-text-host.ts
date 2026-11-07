import { bindWorkbookAsyncTextInput, connectWorkbookSession, createWorkbookRemoteInputController } from '@doffi/editor';
import { mountDocumentPreview } from '@doffi/render';
import { decodeTestBytes } from './package-transport.js';

let cleanup = () => {};
let actions: {
  inspect(): Promise<unknown>; settled(): Promise<void>; select(row: number, column: number): Promise<void>;
  external(text: string): Promise<void>; reload(): Promise<void>; undo(): Promise<void>; setMode(mode: 'auto' | 'text' | 'number'): void;
  releaseBinding(): void; hostStillUsable(): Promise<boolean>; finish(): Promise<unknown>;
};
export async function mountRemoteText(code: string, encoded: string) {
  cleanup(); const host = document.createElement('div'), input = document.createElement('textarea'), canvas = document.createElement('div');
  input.setAttribute('aria-label', 'Worker 单元格输入'); host.append(input, canvas); document.body.append(host);
  const url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' })), worker = new Worker(url), channel = new MessageChannel();
  const bytes = decodeTestBytes(encoded); worker.postMessage({ kind: 'open', format: 'package', bytes, port: channel.port1 }, [bytes.buffer, channel.port1]);
  const client = await connectWorkbookSession(channel.port2); let source = await client.view(), released = false, work = Promise.resolve();
  const preview = mountDocumentPreview(canvas, source, { widthPt: 420, heightPt: 240 }); await preview.ready;
  const sheetId = source.cells.snapshot.metadata.workbook.activeSheetId, errors: { code: string; text: string }[] = [], navigation: string[] = [];
  const selection = (row = 0, column = 0) => ({ documentId: client.state().documentId, baseRevision: client.state().revision,
    anchor: { kind: 'cells' as const, sheetId, ranges: [{ from: { row, column }, to: { row, column } }] } });
  const controller = createWorkbookRemoteInputController(client, { onChange() { const state = controller.state(); input.disabled = state.applying || state.selecting; } });
  const report = (error: unknown) => { if (!released) errors.push({ code: (error as { code?: string }).code ?? 'InvalidOperation', text: input.value }); };
  const queue = (action: () => Promise<void>) => { work = work.catch(report).then(async () => { if (!released) await action(); }); void work.catch(report); };
  async function refresh() {
    const next = await client.view(undefined, { previous: source }); preview.update(next); await preview.settled(); source.dispose(); source = next;
    const state = controller.state();
    if (!state.dirty && !state.composing && !state.applying && state.selection?.anchor.kind === 'cells') {
      const at = state.selection.anchor.ranges[0]!.from; await controller.select(selection(at.row, at.column));
    }
    binding.refresh();
  }
  const binding = bindWorkbookAsyncTextInput(input, controller, {
    onApplied() { queue(refresh); }, onNavigate(direction) { navigation.push(direction); queue(async () => { binding.focus(); }); },
    onCancelled() { queue(async () => { const state = controller.state(), at = state.selection?.anchor.kind === 'cells' ? state.selection.anchor.ranges[0]!.from : { row: 0, column: 0 }; await controller.select(selection(at.row, at.column)); binding.refresh(); binding.focus(); }); },
    onError(error, text) { errors.push({ code: error.code, text }); },
  });
  await controller.select(selection()); binding.refresh(); binding.focus();
  const release = () => { if (released) return; released = true; binding.dispose(); controller.dispose(); preview.dispose(); source.dispose(); client.dispose(); };
  cleanup = () => { release(); worker.terminate(); URL.revokeObjectURL(url); host.remove(); };
  actions = {
    async inspect() { return { ...client.state(), input: controller.state(), text: input.value, source: await client.readInput(sheetId, { row: 0, column: 0 }), navigation: [...navigation], errors: [...errors] }; },
    async settled() { for (;;) { const current = work; await current; await new Promise<void>(resolve => setTimeout(resolve, 0)); if (current === work && !controller.state().applying && !controller.state().selecting) return; } },
    async select(row, column) { await controller.select(selection(row, column)); binding.refresh(); binding.focus(); },
    async external(text) { const candidate = await client.prepare({ kind: 'input', documentId: client.state().documentId, revision: client.state().revision, sheetId, at: { row: 0, column: 0 }, text, mode: 'text' }); await client.commit(candidate); await refresh(); },
    async reload() { controller.cancel(); const state = controller.state(); const at = state.selection?.anchor.kind === 'cells' ? state.selection.anchor.ranges[0]!.from : { row: 0, column: 0 }; await controller.select(selection(at.row, at.column)); binding.refresh(); binding.focus(); },
    async undo() { await client.undo(); await refresh(); binding.focus(); },
    setMode(mode) { controller.setMode(mode); binding.refresh(); }, releaseBinding() { binding.dispose(); },
    async hostStillUsable() { const candidate = await client.prepare({ kind: 'input', documentId: client.state().documentId, revision: client.state().revision, sheetId, at: { row: 0, column: 0 }, text: 'host remains usable', mode: 'text' }); await client.commit(candidate); return (await client.readInput(sheetId, { row: 0, column: 0 })).text === 'host remains usable'; },
    async finish() {
      await work; release(); let resolve!: (value: unknown) => void; const response = new Promise<unknown>(done => { resolve = done; });
      const timeout = setTimeout(() => resolve({ error: 'audit timeout' }), 10_000);
      worker.onmessage = event => { if (event.data.kind === 'audit' || event.data.kind === 'error') resolve(event.data); };
      worker.postMessage({ kind: 'closed-audit' }); const audit = await response; clearTimeout(timeout); cleanup(); return audit;
    },
  };
}
export function inspect() { return actions.inspect(); }
export function settled() { return actions.settled(); }
export function select(row: number, column: number) { return actions.select(row, column); }
export function external(text: string) { return actions.external(text); }
export function reload() { return actions.reload(); }
export function undo() { return actions.undo(); }
export function setMode(mode: 'auto' | 'text' | 'number') { return actions.setMode(mode); }
export function releaseBinding() { actions.releaseBinding(); }
export function hostStillUsable() { return actions.hostStillUsable(); }
export function finish() { return actions.finish(); }
Object.assign(window, { workbookRemoteText: { mountRemoteText, inspect, settled, select, external, reload, undo, setMode, releaseBinding, hostStillUsable, finish } });
