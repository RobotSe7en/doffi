import { connectWorkbookSession, mountWorkbookEditor } from '@doffi/editor';
import type { WorkbookRemoteEditorView } from '@doffi/editor';
import { DEFAULT_WORKBOOK_RULE_LIMITS } from '@doffi/core';
import type { TransactionPrepareOptions, WorkbookCell, WorkbookSnapshot } from '@doffi/core';
import { decodeTestBytes } from './package-transport.js';
import { measureWorkbookCellAppearance, measureWorkbookChart, measureWorkbookDrawingText, measureWorkbookDrawingBounds, observeWorkbookViews } from './workbook-view-observer.js';

let cleanup = () => {};
let actions: { inspect(): Promise<unknown>; settled(): Promise<unknown>; external(text: string, wait?: boolean, at?: { row: number; column: number }): Promise<void>; value(row: number, column: number): Promise<string>; disposeEditor(): Promise<boolean>; finish(): Promise<unknown>;
  tableStyleSource(): Promise<{ tableStyles: WorkbookSnapshot["metadata"]["workbook"]["tableStyles"]; defaultTableStyle: WorkbookSnapshot["metadata"]["workbook"]["defaultTableStyle"] }>;
  sourceCell(row: number, column: number, sheetId?: string): Promise<WorkbookCell | undefined>;
  displayedCellAppearance(row: number, column: number): ReturnType<typeof measureWorkbookCellAppearance>;
  displayedChart(): ReturnType<typeof measureWorkbookChart>;
  displayedDrawingText(nodeId: string): ReturnType<typeof measureWorkbookDrawingText>;
  displayedDrawingBounds(): ReturnType<typeof measureWorkbookDrawingBounds>;
  displayedCell(row: number, column: number): string | undefined;
  holdReference(): void; referenceHeld(): boolean; releaseReference(): void };
export async function mount(code: string, encoded: string, observeView = false, prepareOptions?: TransactionPrepareOptions, confirmWarnings = false) {
  cleanup(); const host = document.createElement('div'); document.body.append(host);
  const url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' })), worker = new Worker(url), channel = new MessageChannel();
  let holdReference = false, heldReference: unknown;
  channel.port2.addEventListener('message', event => {
    if (holdReference && event.data.value?.reference) { holdReference = false; heldReference = event.data; event.stopImmediatePropagation(); }
  });
  const bytes = decodeTestBytes(encoded); worker.postMessage({ kind: 'open', format: 'package', bytes, port: channel.port1 }, [bytes.buffer, channel.port1]);
  const client = await connectWorkbookSession(channel.port2), errors: string[] = [];
  const attach = Element.prototype.attachShadow;
  Element.prototype.attachShadow = function(init) { return attach.call(this, { ...init, mode: 'open' }); };
  let editor: WorkbookRemoteEditorView;
  let observed: ReturnType<typeof observeWorkbookViews> | undefined;
  try {
    editor = mountWorkbookEditor(host, client, { widthPt: 420, heightPt: 240, ...(prepareOptions ? { prepareOptions: () => prepareOptions } : {}), ...(confirmWarnings ? { confirmValidationWarnings: () => window.confirm("应用验证警告？") } : {}), onError(error) { errors.push((error as { code?: string }).code ?? error.message); } });
    if (observeView) {
      Element.prototype.attachShadow = attach;
      observed = observeWorkbookViews(editor.element, { openHostShadows: true });
    }
    await editor.ready;
  }
  finally { if (!observed) Element.prototype.attachShadow = attach; }
  let disposed = false;
  const release = () => { if (!disposed) { editor.dispose(); disposed = true; } };
  cleanup = () => { observed?.dispose(); release(); client.dispose(); worker.terminate(); URL.revokeObjectURL(url); host.remove(); };
  actions = {
    async inspect() {
      const source = await client.view();
      try {
        const current = await source.rules!.read({ ...DEFAULT_WORKBOOK_RULE_LIMITS, documentId: client.state().documentId, revision: client.state().revision,
          sheetId: 'sheet-1', includeCurrent: true, coordinates: [{ row: 0, column: 1 }] }, new AbortController().signal);
        return { ...client.state(), source: await client.readInput('sheet-1', { row: 0, column: 0 }), input: editor.inputState(), selection: editor.selection(), current: current.cells[0]?.current, merges: source.cells.snapshot.metadata.workbook.sheets[0]!.merges, frozen: source.cells.snapshot.metadata.workbook.sheets[0]!.frozen,
          drawings: source.cells.snapshot.metadata.workbook.objects ?? [], charts: source.cells.snapshot.metadata.workbook.charts ?? [], content: source.cells.snapshot.content?.html,
          tables: source.cells.snapshot.metadata.workbook.sheets.map(sheet => ({ sheetId: sheet.sheetId, tables: sheet.tables ?? [] })),
          names: source.cells.snapshot.metadata.workbook.names ?? [],
          sheets: source.cells.snapshot.metadata.workbook.sheets,
          pending: editor.hasPendingInput(),
          annotations: source.cells.snapshot.annotations, comments: source.cells.snapshot.metadata.workbook.comments ?? [], stories: source.cells.snapshot.metadata.workbook.stories ?? [], errors: [...errors] };
      } finally { source.dispose(); }
    },
    async tableStyleSource() { const source = await client.view(); try { const workbook = source.cells.snapshot.metadata.workbook; return { tableStyles: workbook.tableStyles, defaultTableStyle: workbook.defaultTableStyle }; } finally { source.dispose(); } },
    settled: () => editor.settled(),
    displayedCellAppearance: (row, column) => measureWorkbookCellAppearance(observed?.current(), row, column),
    displayedChart: () => measureWorkbookChart(observed?.current()),
    displayedDrawingText: nodeId => measureWorkbookDrawingText(observed?.current(), nodeId),
    displayedDrawingBounds: () => measureWorkbookDrawingBounds(observed?.current()),
    displayedCell: (row, column) => observed?.current()?.querySelector(`[data-office-view="workbook-cell"][data-row="${row}"][data-column="${column}"]`)?.textContent ?? undefined,
    async value(row, column) { return (await client.readInput('sheet-1', { row, column })).text; },
    async sourceCell(row, column, sheetId = 'sheet-1') {
      const source = await client.view(), current = client.state();
      try { return (await source.cells.read({ documentId: current.documentId, revision: current.revision, maxCells: 1, maxCodeUnits: 65536,
        windows: [{ sheetId, row, column, rows: 1, columns: 1 }] }, new AbortController().signal)).cells[0]; }
      finally { source.dispose(); }
    },
    async external(text, wait = true, at = { row: 0, column: 0 }) { const state = client.state(); const candidate = await client.prepare({ kind: 'input', documentId: state.documentId, revision: state.revision, sheetId: 'sheet-1', at, text, mode: 'text' }); await client.commit(candidate); if (wait) await editor.settled(); },
    holdReference() { holdReference = true; heldReference = undefined; }, referenceHeld() { return heldReference !== undefined; },
    releaseReference() { if (heldReference === undefined) throw Error('No reference reply is held'); const data = heldReference; heldReference = undefined; channel.port2.dispatchEvent(new MessageEvent('message', { data })); },
    async disposeEditor() { release(); const state = client.state(), candidate = await client.prepare({ kind: 'input', documentId: state.documentId, revision: state.revision, sheetId: 'sheet-1', at: { row: 0, column: 0 }, text: 'after editor disposal', mode: 'text' }); await client.commit(candidate); return (await client.readInput('sheet-1', { row: 0, column: 0 })).text === 'after editor disposal'; },
    async finish() {
      release(); client.dispose(); let resolve!: (value: unknown) => void; const response = new Promise<unknown>(done => { resolve = done; });
      const timeout = setTimeout(() => resolve({ error: 'audit timeout' }), 10_000); worker.onmessage = event => { if (event.data.kind === 'audit' || event.data.kind === 'error') resolve(event.data); };
      worker.postMessage({ kind: 'closed-audit' }); const result = await response; clearTimeout(timeout); cleanup();
      return { ...(result as object), hostShadowsRestored: Element.prototype.attachShadow === attach };
    },
  };
}
export function inspect() { return actions.inspect(); }
export function settled() { return actions.settled(); }
export function displayedCellAppearance(row: number, column: number) { return actions.displayedCellAppearance(row, column); }
export function displayedChart() { return actions.displayedChart(); }
export function displayedDrawingText(nodeId: string) { return actions.displayedDrawingText(nodeId); }
export function displayedDrawingBounds() { return actions.displayedDrawingBounds(); }
export function displayedCell(row: number, column: number) { return actions.displayedCell(row, column); }
export function external(text: string, wait = true, at = { row: 0, column: 0 }) { return actions.external(text, wait, at); }
export function holdReference() { actions.holdReference(); }
export function referenceHeld() { return actions.referenceHeld(); }
export function releaseReference() { actions.releaseReference(); }
export function value(row: number, column: number) { return actions.value(row, column); }
export function sourceCell(row: number, column: number, sheetId?: string) { return actions.sourceCell(row, column, sheetId); }
export function disposeEditor() { return actions.disposeEditor(); }
export function finish() { return actions.finish(); }
export const tableStyleSource = () => actions.tableStyleSource();
Object.assign(window, { workbookRemoteEditor: { mount, inspect, tableStyleSource, settled, displayedCellAppearance, displayedChart, displayedDrawingText, displayedDrawingBounds, displayedCell, external, value, sourceCell, holdReference, referenceHeld, releaseReference, disposeEditor, finish } });
