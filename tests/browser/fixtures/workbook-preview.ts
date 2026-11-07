import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createDocumentSession, createWorkbookNumberFormatter, DocumentError, openDocument } from '@doffi/core';
import type { DocumentSelection, WorkbookDocumentMetadata, WorkbookModel } from '@doffi/core';
import { createWorkbookPreviewSource, mountDocumentPreview } from '@doffi/render';
import type { WorkbookPreview, WorkbookPreviewSource, WorkbookPreviewOptions } from '@doffi/render';
import { modelWindowSource } from '../../../packages/render/src/workbook/window-source.js';
import { modelAssetSource } from '../../../packages/render/src/workbook/assets.js';
import { modelChartSource } from '../../../packages/render/src/workbook/chart-source.js';
import { modelRuleSource } from '../../../packages/render/src/workbook/rule-source.js';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookWindowController } from '../../../packages/render/src/workbook/window-controller.js';
import { decorateWorkbookPreview } from '../../../packages/render/src/workbook/preview-decoration.js';
import { incrementalPreview } from './workbook-incremental.js';
export { incrementalPreview };
import type { WorkbookPreviewDecoration } from '../../../packages/render/src/workbook/preview-decoration.js';

type Entries = TestPackage;
function modelOf(entries: Entries, rtl = false): WorkbookModel {
  const files = decodeTestPackage(entries);
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.rightToLeft = rtl;
  metadata.workbook.sheets.push({ sheetId: 'preview-empty', name: 'Empty', visibility: 'visible', defaultRowHeightPt: 20, defaultColumnWidth: 8.43,
    defaultCssClass: metadata.workbook.sheets[0]!.defaultCssClass, rowDimensions: [], columnDimensions: [], merges: [], frozen: { rows: 0, columns: 0 }, showGridLines: true, rightToLeft: rtl });
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  return openDocument(files, { kind: 'workbook' });
}
function host() { const container = document.createElement('div'); container.style.cssText = 'position:fixed;left:0;top:0;z-index:2147483647;background:white'; document.body.append(container); return container; }
// Capture only our own chrome shadow for inspecting public UI. The production
// component continues to expose neither its source DOM nor its private surface.
function mount(container: HTMLElement, input: WorkbookModel | WorkbookPreviewSource, callbacks: WorkbookPreviewOptions = {}) {
  const original = Element.prototype.attachShadow; let chrome!: ShadowRoot;
  Element.prototype.attachShadow = function (init) { const root = original.call(this, init); chrome ??= root; return root; };
  try { const preview = mountDocumentPreview(container, input, { widthPt: 420, heightPt: 300, ...callbacks }); return { preview, chrome }; }
  finally { Element.prototype.attachShadow = original; }
}
let cleanup = () => {}, inspect = () => ({}), action = async (_action: string) => ({});
export function previewInspect() { return inspect(); }
export function previewAction(name: string) { return action(name); }
export async function imperativeRangeSelection(entries: Entries, rtl: boolean) {
  cleanup(); const model = modelOf(entries, rtl), before = model.serialize(), container = host();
  const { preview } = mount(container, model);
  try {
    await preview.ready;
    const range = { from: { row: 0, column: 0 }, to: { row: 3, column: 3 } }, first = await preview.selectRange(range);
    let invalid = ""; try { await preview.selectRange({ from: { row: 2, column: 2 }, to: { row: 1, column: 1 } }); } catch (error) { invalid = (error as DocumentError).code; }
    const retained = preview.selection() === first;
    const far = await preview.selectRange({ from: { row: 500, column: 40 }, to: { row: 502, column: 42 } }), moved = preview.layout().viewport;
    const unscrolled = await preview.selectRange(range, false), sameViewport = JSON.stringify(preview.layout().viewport) === JSON.stringify(moved);
    const pending = preview.selectRange(range, false); void pending.catch(() => {});
    await preview.showSheet("preview-empty"); let race = ""; try { await pending; } catch (error) { race = (error as DocumentError).code; }
    const after = model.serialize(), unchanged = [...before].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index]));
    return { first, far, unscrolled, invalid, retained, moved, sameViewport, race, unchanged, frozen: Object.isFrozen(first.anchor) && first.anchor.kind === "cells" && Object.isFrozen(first.anchor.ranges[0]!.from) };
  } finally { preview.dispose(); container.remove(); }
}
export async function calculatedPreview(entries: Entries) {
  cleanup(); const session = createDocumentSession(modelOf(entries)), before = session.serialize(), container = host(), errors: string[] = [];
  const result = await session.calculate({ yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)) });
  const { preview, chrome } = mount(container, createWorkbookPreviewSource(session, result), { retainSelection: true, onError: error => errors.push(error.code) });
  const views = new WeakMap<HTMLIFrameElement, ShadowRoot>(), watched = new Set<HTMLIFrameElement>(), restores: (() => void)[] = [];
  const watch = () => { for (const frame of chrome.querySelectorAll<HTMLIFrameElement>('iframe')) {
    if (watched.has(frame)) continue; watched.add(frame);
    frame.addEventListener('load', () => {
      const prototype = (frame.contentWindow as Window & typeof globalThis).Element.prototype, attach = prototype.attachShadow;
      prototype.attachShadow = function(init) { const root = attach.call(this, init); if (this.localName === 'office-workbook-view-surface') views.set(frame, root); return root; };
      restores.push(() => { prototype.attachShadow = attach; });
    }, { capture: true, once: true });
  } };
  const observer = new MutationObserver(watch); observer.observe(chrome, { childList: true, subtree: true }); watch();
  try {
    await preview.ready; const at = { row: 149, column: 0 };
    const selected = await preview.selectRange({ from: at, to: at });
    const cell = views.get(chrome.querySelector<HTMLIFrameElement>('iframe')!)!.querySelector('[data-office-view="workbook-cell"][data-row="149"][data-column="0"]')?.textContent;
    await preview.setViewport({ ...preview.layout().viewport, scrollTopPt: 0 }); const retained = JSON.stringify(preview.selection()) === JSON.stringify(selected);
    await preview.showSheet("preview-empty"); preview.update(createWorkbookPreviewSource(session, result)); await preview.settled();
    const sheet = preview.layout().sheetId, after = session.serialize();
    return { cell, retained, sheet, errors, unchanged: [...before].every(([path, bytes]) => bytes.every((byte, index) => byte === after.get(path)![index])) };
  } finally { observer.disconnect(); for (const restore of restores) restore(); preview.dispose(); session.dispose(); container.remove(); }
}
export async function mountPublicPreview(entries: Entries, rtl: boolean) {
  cleanup(); const model = modelOf(entries, rtl), before = model.serialize(), container = host(), errors: string[] = [], selections: (DocumentSelection | undefined)[] = [];
  let serializations = 0;
  const direct = { ...model, serialize() { serializations++; throw Error('A preview must not serialize the entire workbook'); } };
  const { preview, chrome } = mount(container, direct, { onError: error => errors.push(error.code), onSelectionChange: selection => selections.push(selection) });
  try {
    await preview.ready;
    const frame = chrome.querySelector('iframe')!, rect = frame.getBoundingClientRect(), sheetId = preview.layout().sheetId;
    let wheelPanel: HTMLElement | undefined;
    const bytesUnchanged = () => { const after = model.serialize(); return [...before].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, i) => byte === after.get(path)![i])); };
    inspect = () => ({ report: preview.layout(), address: chrome.querySelector<HTMLInputElement>('[aria-label="单元格地址"]')!.value,
      value: chrome.querySelector<HTMLInputElement>('[aria-label="单元格值或公式"]')!.value, status: chrome.querySelector('[role="status"]')!.textContent,
      validation: chrome.querySelector('[aria-label="数据验证"]')!.textContent,
      tabs: [...chrome.querySelectorAll('[role="tab"]')].map(tab => ({ text: tab.textContent, selected: tab.getAttribute('aria-selected') })),
      selection: preview.selection(), selections, errors, serializations, unchanged: bytesUnchanged(), frames: chrome.querySelectorAll('iframe').length });
    action = async name => {
      if (name.startsWith('point-')) {
        const [, column, row] = name.split('-'), headers = [...chrome.querySelectorAll('.header>div')];
        const x = headers.find(node => node.textContent === column)!.getBoundingClientRect(), y = headers.find(node => node.textContent === row)!.getBoundingClientRect();
        return { x: x.x + x.width / 2, y: y.y + y.height / 2 };
      }
      if (name === 'empty') chrome.querySelector<HTMLButtonElement>('[data-sheet-id="preview-empty"]')!.click();
      if (name === 'first') await preview.showSheet(sheetId);
      if (name === 'comment-scroll') {
        const document = frame.contentDocument!, create = document.createElement;
        document.createElement = ((tag: string, options?: ElementCreationOptions) => {
          const element = create.call(document, tag, options); if (tag === 'aside') wheelPanel = element; return element;
        }) as Document['createElement'];
        try { preview.openComment(model.snapshot().annotations[0]!.annotationId); } finally { document.createElement = create; }
        wheelPanel!.style.maxHeight = '40px'; const bounds = wheelPanel!.getBoundingClientRect();
        return { x: rect.x + bounds.x + bounds.width / 2, y: rect.y + bounds.y + bounds.height / 2, max: wheelPanel!.scrollHeight - wheelPanel!.clientHeight };
      }
      if (name === 'comment-scroll-state') return { top: wheelPanel!.scrollTop, viewport: preview.layout().viewport };
      if (name === 'comment') preview.openComment(model.snapshot().annotations[0]!.annotationId);
      if (name === 'close') preview.closeComment();
      if (name === 'bad-sheet') { try { await preview.showSheet('missing'); } catch { /* inspected below */ } return inspect(); }
      if (name === 'dispose') { const fields = [...chrome.querySelectorAll('input')]; preview.dispose(); cleanup(); return { removed: !container.isConnected, frames: chrome.querySelectorAll('iframe').length, unchanged: bytesUnchanged(), serializations, fieldsCleared: fields.every(field => !field.value) }; }
      await preview.settled(); return inspect();
    };
    cleanup = () => { preview.dispose(); container.remove(); cleanup = () => {}; };
    return { x: rect.x + (rtl ? rect.width - 16 : 16), y: rect.y + 8, ...inspect() };
  } catch (error) { preview.dispose(); container.remove(); throw error; }
}

export async function publicPreviewLifecycle(entries: Entries) {
  const model = modelOf(entries), cells = modelWindowSource(model), assets = modelAssetSource(model), charts = modelChartSource(model), container = host();
  let gate: { begin(): void; release(fail: boolean): void; signal: AbortSignal | undefined } | undefined, wrongRevision = false;
  const source: WorkbookPreviewSource = { kind: 'workbook', assets, charts, cells: { ...cells, async read(request, signal) {
    if (gate) { const held = gate; gate = undefined; held.signal = signal; await new Promise<void>((resolve, reject) => { held.release = fail => fail ? reject(Error('Late private failure')) : resolve(); held.begin(); }); }
    const response = await cells.read(request, new AbortController().signal);
    return wrongRevision ? { ...response, revision: response.revision + 1 } : response;
  } } };
  const stall = () => {
    let begin!: () => void; const started = new Promise<void>(resolve => { begin = resolve; });
    const held = { begin, release: (_fail: boolean) => {}, signal: undefined as AbortSignal | undefined }; gate = held;
    return { started, release: (fail = false) => held.release(fail), aborted: () => held.signal?.aborted };
  };
  const errors: string[] = [], unhandled: string[] = [], rejection = (event: PromiseRejectionEvent) => { unhandled.push(String(event.reason)); event.preventDefault(); };
  window.addEventListener('unhandledrejection', rejection);
  const created = new Set<string>(), revoked = new Set<string>(), originalCreate = URL.createObjectURL, originalRevoke = URL.revokeObjectURL;
  URL.createObjectURL = blob => { const url = originalCreate.call(URL, blob); created.add(url); return url; }; URL.revokeObjectURL = url => { revoked.add(url); originalRevoke.call(URL, url); };
  let preview: WorkbookPreview | undefined;
  try {
    const mounted = mount(container, source, { onError: error => errors.push(error.code) }); preview = mounted.preview; const chrome = mounted.chrome;
    await preview.ready; const first = preview.layout(), oldFrame = chrome.querySelector('iframe')!;
    const scroll = stall(), obsolete = preview.setViewport({ ...first.viewport, scrollTopPt: 30 }); await scroll.started;
    await preview.showSheet('preview-empty'); const newest = await obsolete;
    scroll.release(true); await new Promise<void>(resolve => setTimeout(resolve, 0));
    const scrollRace = newest.sheetId === 'preview-empty' && scroll.aborted() && preview.layout().sheetId === newest.sheetId;
    await preview.showSheet(first.sheetId);
    wrongRevision = true; preview.update(source); let failed = '';
    try { await preview.settled(); } catch (error) { failed = (error as { code: string }).code; }
    const retained = preview.layout().sheetId === first.sheetId && chrome.querySelector('iframe') === oldFrame && chrome.querySelectorAll('iframe').length === 1;
    wrongRevision = false;
    const replacement = stall(); preview.update(source); await replacement.started;
    const pending = preview.settled(); preview.update(model); await preview.settled(); await pending;
    replacement.release(true); await new Promise<void>(resolve => setTimeout(resolve, 0));
    const updateRace = replacement.aborted() && !oldFrame.isConnected && chrome.querySelectorAll('iframe').length === 1;
    const stopped = stall(); preview.update(source); await stopped.started;
    const last = preview.settled().then(() => 'accepted', error => (error as { code: string }).code); preview.dispose(); const disposed = await last;
    stopped.release(true); await new Promise<void>(resolve => setTimeout(resolve, 0));
    const providerAlive = (await cells.read({ documentId: first.documentId, revision: first.revision, windows: [], maxCells: 1, maxCodeUnits: 1 }, new AbortController().signal)).revision === first.revision;
    return { scrollRace, failed, retained, updateRace, disposed, errors, unhandled, providerAlive, released: [...created].every(url => revoked.has(url)), removed: !mounted.preview.element.isConnected };
  } finally { preview?.dispose(); container.remove(); window.removeEventListener('unhandledrejection', rejection); URL.createObjectURL = originalCreate; URL.revokeObjectURL = originalRevoke; }
}
export async function previewDecorationMeasurement(entries: Entries) {
  const model = modelOf(entries), snapshot = model.snapshot(), container = host();
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 }), formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
  let mode = 'ok', released = 0, controls: WorkbookPreviewDecoration | undefined;
  const controller = new WorkbookWindowController(surface, { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model) }, formatter, {}, (window, content, maxNodes) => {
    if (mode === 'throw') throw new DocumentError('UnsupportedFeature', 'Deliberate control construction failure.');
    if (mode === 'budget') return { nodes: maxNodes + 1, dispose() { released++; } };
    const next = decorateWorkbookPreview(surface, snapshot, window, content, maxNodes, action => action());
    if (mode === 'cancel') controller.cancelPending(); else controls = next;
    return { get nodes() { return next.nodes; }, get fonts() { return next.fonts; }, dispose() { next.dispose(); released++; } };
  });
  const viewport = { widthPt: 420, heightPt: 300, scrollLeftPt: 0, scrollTopPt: 0 }, errors: string[] = [];
  try {
    await controller.render(snapshot.metadata.workbook.activeSheetId, viewport);
    const accepted = controller.current!, html = surface.view.innerHTML;
    for (mode of ['throw', 'budget', 'cancel']) {
      try { errors.push(await controller.render(accepted.frame.sheetId, { ...viewport, heightPt: 320 }) ? 'accepted' : 'cancelled'); } catch (error) { errors.push((error as { code: string }).code); }
      if (controller.current !== accepted || surface.view.innerHTML !== html) throw Error('Control failure changed the accepted window');
    }
    controls!.openComment(snapshot.annotations[0]!.annotationId); controls!.selectCell(0, 0);
    const walker = surface.document.createTreeWalker(accepted.element, 0xffff_ffff); let actualNodes = 1; while (walker.nextNode()) actualNodes++;
    const result = { errors, nodes: accepted.nodes, actualNodes, fonts: accepted.fonts.length, panel: accepted.element.querySelector('[data-office-view="workbook-preview-comments"]')?.textContent,
      highlights: accepted.element.querySelector('[data-office-view="workbook-cell-selection"]')?.childNodes.length, releasedCandidates: released };
    controls!.closeComment(); controller.dispose();
    return { ...result, releasedAll: released, empty: !surface.view.childNodes.length && !surface.source.childNodes.length };
  } finally { controller.dispose(); formatter.dispose(); surface.dispose(); container.remove(); }
}
export async function initialPreviewFailure(entries: Entries) {
  const model = modelOf(entries), container = host(), errors: string[] = [];
  const missing = mount(container, model, { sheetId: 'missing', onError: error => errors.push(error.code) });
  let rejected = '';
  try { await missing.preview.ready; rejected = 'accepted'; } catch (error) { rejected = (error as { code: string }).code; }
  const frames = missing.chrome.querySelectorAll('iframe').length; missing.preview.dispose();
  const cells = modelWindowSource(model); let fail = true;
  const source: WorkbookPreviewSource = { kind: 'workbook', assets: modelAssetSource(model), charts: modelChartSource(model), cells: { ...cells, async read(request, signal) {
    if (fail) { fail = false; throw Error('Private transport details'); } return cells.read(request, signal);
  } } };
  let preview!: WorkbookPreview;
  const retry = mount(container, source, { onError: error => { errors.push(error.code); preview.update(model); } }); preview = retry.preview;
  try { await preview.ready; const report = await preview.settled(); return { rejected, frames, errors, retried: report.documentId === model.snapshot().manifest.documentId, finalFrames: retry.chrome.querySelectorAll('iframe').length }; }
  finally { preview.dispose(); container.remove(); }
}
export async function ruleWindowLifecycle(entries: Entries) {
  const model = modelOf(entries), snapshot = model.snapshot(), container = host(), rules = modelRuleSource(model), formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 });
  let mode = 'ok', begin!: () => void, release!: () => void, aborted = false;
  const controller = new WorkbookWindowController(surface, { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model),
    rules: { ...rules, async read(request, signal) {
      if (mode === 'held') { mode = 'ok'; signal.addEventListener('abort', () => { aborted = true; }, { once: true }); await new Promise<void>(resolve => { release = resolve; begin(); }); }
      const value = await rules.read(request, new AbortController().signal);
      return mode === 'empty' ? { ...value, cells: [] } : mode === 'revision' ? { ...value, revision: value.revision + 1 } : value;
    } } }, formatter);
  const viewport = { widthPt: 420, heightPt: 300, scrollTopPt: 0, scrollLeftPt: 0 }, failures: string[] = [];
  try {
    await controller.render('sheet-1', viewport); const original = controller.current!, before = surface.view.innerHTML;
    for (mode of ['empty', 'revision']) {
      try { await controller.render('sheet-1', viewport); failures.push('accepted'); } catch (error) { failures.push((error as { code: string }).code); }
      if (controller.current !== original || surface.view.innerHTML !== before) throw Error('A failed rule read changed the accepted window');
    }
    mode = 'held'; const started = new Promise<void>(resolve => { begin = resolve; }), old = controller.render('sheet-1', viewport); await started;
    await controller.render('preview-empty', viewport); const cancelled = !await old;
    release(); await new Promise<void>(resolve => setTimeout(resolve, 0));
    return { failures, cancelled, aborted, retainedLatest: controller.current!.frame.sheetId === 'preview-empty' };
  } finally { controller.dispose(); surface.dispose(); formatter.dispose(); container.remove(); }
}
(window as unknown as { workbookPreview: unknown }).workbookPreview = { incrementalPreview, mountPublicPreview, previewInspect, previewAction, imperativeRangeSelection, calculatedPreview, publicPreviewLifecycle, previewDecorationMeasurement, initialPreviewFailure, ruleWindowLifecycle, publicReadAhead };

export async function publicReadAhead(entries: Entries, allow: boolean, rows: number, mode: 'normal' | 'hold' | 'failure' | 'revision' | 'active' | 'preference' = 'normal') {
  const model = modelOf(entries), before = model.serialize(), snapshot = model.snapshot(), source = modelWindowSource(model), container = host();
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 240, heightPt: 60 }), formatter = createWorkbookNumberFormatter({ dateSystem: '1900' });
  let observing = false, foreground = false, reads = 0, largestSlots = 0, held: { signal: AbortSignal; release(): void } | undefined;
  let began!: () => void; const started = new Promise<void>(resolve => { began = resolve; });
  const unhandled: string[] = [], onUnhandled = (event: PromiseRejectionEvent) => { unhandled.push(String(event.reason)); event.preventDefault(); };
  window.addEventListener('unhandledrejection', onUnhandled);
  const cells = { ...source, get allowReadAhead() { if (mode === 'preference' && foreground) throw Error('Optional preference failed'); return allow; }, async read(request: Parameters<typeof source.read>[0], signal: AbortSignal) {
    const background = observing && !foreground;
    if (background) {
      reads++; largestSlots = Math.max(largestSlots, request.windows.reduce((sum, item) => sum + item.rows * item.columns, 0)); began();
      if (mode === 'hold' && !held) await new Promise<void>(resolve => { held = { signal, release: resolve }; });
      if (mode === 'failure') throw Error('Private optional read failure');
    }
    const response = await source.read(request, held?.signal === signal ? new AbortController().signal : signal);
    if (background && mode === 'revision') return { ...response, revision: response.revision + 1 };
    if (background && mode === 'active' && response.cells.length) return { ...response, cells: response.cells.map((cell, index) => index ? cell : { ...cell, html: cell.html.replace('<td', '<td onclick="window.readAheadExecuted=true"') }) };
    return response;
  } };
  const providers = { cells, assets: modelAssetSource(model), charts: modelChartSource(model) }, invalid: string[] = [];
  for (const readAheadRows of [-1, 1.5, 65]) try { new WorkbookWindowController(surface, providers, formatter, { readAheadRows }); invalid.push('accepted'); } catch (error) { invalid.push((error as { code: string }).code); }
  const controller = new WorkbookWindowController(surface, providers, formatter, { readAheadRows: rows });
  const pause = (ms: number) => new Promise<void>(resolve => setTimeout(resolve, ms));
  try {
    const viewport = { widthPt: 240, heightPt: 60, scrollLeftPt: 0, scrollTopPt: 0 };
    await controller.render('sheet-1', viewport); observing = true;
    const first = controller.current!, original = first.element.innerHTML;
    let aborted = false, whileHeld = 0;
    if (allow && rows && mode === 'hold') {
      await Promise.race([started, pause(2000).then(() => { throw Error('Read ahead did not start'); })]);
      surface.document.dispatchEvent(new surface.document.defaultView!.PointerEvent('pointerdown', { bubbles: true }));
      foreground = true; await controller.render('sheet-1', { ...viewport, scrollTopPt: 30 }); foreground = false;
      aborted = held!.signal.aborted; await pause(80); whileHeld = reads;
      held!.release();
    }
    await pause(300);
    const accepted = controller.current!, html = accepted.element.innerHTML, optionalReads = reads;
    foreground = true; await controller.render('sheet-1', { ...viewport, scrollTopPt: 60 }); foreground = false;
    const last = controller.current!; controller.dispose(); await pause(80);
    const stopped = reads; await pause(100);
    return { invalid, reads: optionalReads, largestSlots, aborted, whileHeld, unhandled,
      firstRetained: mode === 'hold' ? accepted !== first : accepted === first && html === original,
      foreground: last.frame.layout.scrollTopPt === 60, stopped: stopped === reads,
      clean: !surface.view.childNodes.length && !surface.source.childNodes.length,
      activeContent: Boolean((surface.document.defaultView as unknown as { readAheadExecuted?: boolean }).readAheadExecuted),
      unchanged: [...before].every(([path, bytes]) => { const after = model.serialize().get(path)!; return bytes.length === after.length && bytes.every((byte, index) => byte === after[index]); }) };
  } finally { held?.release(); controller.dispose(); formatter.dispose(); surface.dispose(); container.remove(); window.removeEventListener('unhandledrejection', onUnhandled); }
}

let matrix: {preview:WorkbookPreview;chrome:ShadowRoot;model:WorkbookModel;container:HTMLElement;before:ReadonlyMap<string,Uint8Array>;serializations:()=>number;errors:string[]} | undefined;
export async function matrixStart(entries:readonly [string,string][]){
  matrix?.preview.dispose();matrix?.container.remove();matrix=undefined;
  const model=openDocument(new Map(entries.map(([path,encoded])=>{const binary=atob(encoded),bytes=new Uint8Array(binary.length);for(let index=0;index<binary.length;index++)bytes[index]=binary.charCodeAt(index);return [path,bytes];})),{kind:'workbook'}),before=model.serialize(),container=host(),errors:string[]=[];
  let serializations=0;const direct={...model,serialize(){serializations++;throw Error('Public preview serialized the workbook');}};
  const mounted=mount(container,direct,{widthPt:880,heightPt:450,onError:error=>errors.push(error.code)});
  matrix={...mounted,model,container,before,serializations:()=>serializations,errors};
  try{await mounted.preview.ready;return {snapshot:model.snapshot(),report:mounted.preview.layout()};}
  catch(error){mounted.preview.dispose();container.remove();matrix=undefined;throw error;}
}
export async function matrixWindow(sheetId:string,last=false){
  const context=matrix!;await context.preview.showSheet(sheetId);
  if(last){const current=context.preview.layout();await context.preview.setViewport({...current.viewport,scrollLeftPt:current.maxScrollLeftPt,scrollTopPt:current.maxScrollTopPt});}
  else {const current=context.preview.layout();await context.preview.setViewport({...current.viewport,scrollLeftPt:0,scrollTopPt:0});}
  return {report:context.preview.layout(),tabs:[...context.chrome.querySelectorAll('[role=tab]')].map(node=>node.textContent),errors:context.errors,serializations:context.serializations()};
}
export function matrixFinish(){
  const context=matrix!,after=context.model.serialize(),unchanged=context.before.size===after.size&&[...context.before].every(([path,bytes])=>bytes.length===after.get(path)?.length&&bytes.every((byte,index)=>byte===after.get(path)![index]));
  context.preview.dispose();context.container.remove();matrix=undefined;
  return {unchanged,serializations:context.serializations(),errors:context.errors,frames:context.chrome.querySelectorAll('iframe').length,removed:!context.container.isConnected};
}

Object.assign(window.workbookPreview,{matrixStart,matrixWindow,matrixFinish});


export async function autoRowReadPipeline(entries: Entries, mode: 'sequential' | 'ahead' | 'cancel' | 'invalid' | 'preference' | 'late-cancel' | 'sequential-late-cancel') {
  const files = decodeTestPackage(entries);
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.defaultRowHeightCustom = false;
  const styles = metadata.workbook.sheets[0]!.rowDimensions; styles.splice(0);
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), provider = modelWindowSource(model), container = host();
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 240, heightPt: 60 });
  const { measureWorkbookColumnFont } = await import('../../../packages/render/src/workbook/column-font.js');
  const { measureWorkbookAutoRows } = await import('../../../packages/render/src/workbook/auto-rows.js');
  const formatter = createWorkbookNumberFormatter({ dateSystem: '1900' }), normal = measureWorkbookColumnFont(surface.source, snapshot.metadata.workbook, 'sheet-1');
  const abort = new AbortController(), view = surface.document.defaultView!, computed = view.getComputedStyle;
  let reads = 0, overlap = false, held: AbortSignal | undefined, release = () => {};
  const unhandled: unknown[] = [], reject = (event: PromiseRejectionEvent) => { unhandled.push(event.reason); event.preventDefault(); };
  window.addEventListener('unhandledrejection', reject);
  view.getComputedStyle = function(element, pseudo) {
    if (reads && element.closest('office-workbook-source')) {
      overlap ||= reads === 2;
      if (mode === 'cancel') abort.abort();
    }
    return computed.call(this, element, pseudo);
  };
  const source = { snapshot, get allowReadAhead() { if (mode === 'preference') throw Error('Optional preference'); return mode !== 'sequential' && mode !== 'sequential-late-cancel'; },
    async read(request: Parameters<typeof provider.read>[0], signal: AbortSignal) {
      const index = ++reads;
      if (index === 2 && (mode === 'cancel' || mode === 'invalid' || mode.includes('late-cancel'))) {
        held = signal; await new Promise<void>(resolve => { release = resolve; });
      }
      const response = await provider.read(request, held === signal ? new AbortController().signal : signal);
      if (index === 1 && mode === 'invalid') return { ...response, cells: response.cells.map((cell, index) => index ? cell : { ...cell, html: cell.html.replace('<td', '<td onclick="void 0"') }) };
      return response;
    } };
  try {
    let error: string | undefined, cells = 0, waitingObserved = false;
    const measuring = measureWorkbookAutoRows(surface, source, 'sheet-1', normal, formatter, abort.signal); void measuring.catch(() => {});
    if (mode.includes('late-cancel')) {
      await new Promise<void>(resolve => setTimeout(resolve, 100));
      waitingObserved = !!held && !!surface.source.querySelector('[data-office-view="workbook-source-measurement"]') && !surface.view.childNodes.length;
      abort.abort();
    }
    let deadline: ReturnType<typeof setTimeout> | undefined;
    try { cells = (await Promise.race([measuring, new Promise<never>((_resolve, reject) => { deadline = setTimeout(() => reject(Error('Cancelled AutoFit still waits for its provider')), 1000); })])).measuredCells; }
    catch (value) { error = (value as { code: string }).code; }
    finally { clearTimeout(deadline); }
    const cleanBeforeRelease = !surface.source.querySelector('[data-office-view$="measurement"]') && !surface.view.childNodes.length;
    const aborted = held?.aborted ?? false; release(); await new Promise<void>(resolve => setTimeout(resolve, 30));
    return { error, cells, reads, overlap, aborted, unhandled, waitingObserved, cleanBeforeRelease, clean: !surface.source.querySelector('[data-office-view$="measurement"]') && !surface.view.childNodes.length };
  } finally { release(); view.getComputedStyle = computed; window.removeEventListener('unhandledrejection', reject); formatter.dispose(); surface.dispose(); container.remove(); }
}
(window as unknown as { workbookPreview: Record<string, unknown> }).workbookPreview.autoRowReadPipeline = autoRowReadPipeline;

export async function rollingCandidateLifecycle(entries: Entries, mode: 'same' | 'value' | 'active' | 'revision' | 'font' | 'budget') {
  const model = modelOf(entries), source = modelWindowSource(model), snapshot = model.snapshot(), container = host();
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 240, heightPt: 60 });
  const formatter = createWorkbookNumberFormatter({ dateSystem: '1900' });
  let foreground = true, change = false, background = 0, fresh = 0, mountedPeak = 0;
  const measure = () => { mountedPeak = Math.max(mountedPeak, surface.view.querySelectorAll('[data-office-view="workbook-cell"],[data-office-view="workbook-text-overflow"]').length); };
  const observer = new MutationObserver(measure); observer.observe(surface.view, { childList: true, subtree: true });
  const cells = { ...source, allowReadAhead: true, async read(request: Parameters<typeof source.read>[0], signal: AbortSignal) {
    const current = foreground; if (current && change) fresh++; else if (!current) background++;
    const response = await source.read(request, signal);
    if (!current || !change) return response;
    if (mode === 'revision') return { ...response, revision: response.revision + 1 };
    return { ...response, cells: response.cells.map(cell => {
      if (cell.row !== 8 || cell.column !== 2) return cell;
      if (mode === 'active') return { ...cell, html: cell.html.replace('<td', '<td onclick="window.rollingExecuted=true"') };
      if (mode === 'value') return { ...cell, text: '9999', value: { type: 'number' as const, value: 9999 },
        html: cell.html.replace('data-value="67"', 'data-value="9999"').replace('>67<', '>9999<') };
      return cell;
    }) };
  } };
  const controller = new WorkbookWindowController(surface, { cells, assets: modelAssetSource(model), charts: modelChartSource(model) }, formatter,
    mode === 'budget' ? { cells: { maxCellFragments: 40 } } : {});
  const pause = (ms: number) => new Promise<void>(resolve => setTimeout(resolve, ms));
  const viewport = { widthPt: 240, heightPt: 60, scrollLeftPt: 0, scrollTopPt: 0 };
  try {
    await controller.render('sheet-1', viewport); await controller.render('sheet-1', { ...viewport, scrollTopPt: 30 });
    foreground = false; const accepted = controller.current!, html = accepted.element.innerHTML; await pause(300);
    const retained = controller.current === accepted && accepted.element.innerHTML === html;
    if (mode === 'font') surface.document.fonts.add(new (surface.document.defaultView as Window & typeof globalThis).FontFace('Rolling font epoch', 'local("sans-serif")'));
    foreground = true; change = true; let error: string | undefined;
    try { await controller.render('sheet-1', { ...viewport, scrollTopPt: 60 }); } catch (value) { error = (value as { code: string }).code; }
    measure();
    const result = { background, fresh, retained, error, mountedPeak, current: controller.current!.frame.layout.scrollTopPt,
      updated: controller.current!.element.textContent!.includes('9999'), failedRetained: controller.current === accepted && accepted.element.innerHTML === html,
      active: Boolean((surface.document.defaultView as unknown as { rollingExecuted?: boolean }).rollingExecuted), fragments: controller.current!.cells.length };
    controller.dispose(); const stopped = background; await pause(80);
    return { ...result, stopped: background === stopped, clean: !surface.source.childNodes.length && !surface.view.childNodes.length };
  } finally { observer.disconnect(); controller.dispose(); formatter.dispose(); surface.dispose(); container.remove(); }
}
(window as unknown as { workbookPreview: Record<string, unknown> }).workbookPreview.rollingCandidateLifecycle = rollingCandidateLifecycle;

export async function responsiveViewportLifecycle(entries: Entries, escaped = false) {
  const files = decodeTestPackage(entries);
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.defaultRowHeightCustom = false;
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  const size = escaped ? String.raw`calc(2px + 5v\77)` : '5vw';
  files.set('styles.css', new TextEncoder().encode(new TextDecoder().decode(files.get('styles.css')!) + `\n.${metadata.workbook.sheets[0]!.defaultCssClass}{font-size:${size}}`));
  const model = openDocument(files, { kind: 'workbook' }), before = model.serialize(), aHost = host(), bHost = host();
  let serializations = 0, bad = false, holdNext = false, heldSignal: AbortSignal | undefined;
  let began = () => {}, release = () => {};
  const direct = { ...model, serialize() { serializations++; throw Error('Resize serialized the whole workbook'); } }, source = modelWindowSource(direct);
  const providers = { kind: 'workbook' as const, cells: { ...source, async read(request: Parameters<typeof source.read>[0], signal: AbortSignal) {
    if (holdNext) { holdNext = false; heldSignal = signal; await new Promise<void>(resolve => { release = resolve; began(); }); }
    const response = await source.read(request, heldSignal === signal ? new AbortController().signal : signal); return bad ? { ...response, revision: response.revision + 1 } : response;
  } }, assets: modelAssetSource(direct), charts: modelChartSource(direct) };
  const errors: string[] = [], a = mount(aHost, providers, { widthPt: 240, heightPt: 120, onError: error => errors.push(error.code) });
  let b: ReturnType<typeof mount> | undefined;
  try {
    await a.preview.ready; const initial = a.preview.layout();
    const resized = await a.preview.setViewport({ ...initial.viewport, widthPt: 480 });
    b = mount(bHost, direct, { widthPt: 480, heightPt: 120 }); await b.preview.ready; const fresh = b.preview.layout();
    const scrolled = await a.preview.setViewport({ ...resized.viewport, scrollTopPt: 20 }), accepted = a.preview.layout();
    bad = true; let failure: string | undefined;
    try { await a.preview.setViewport({ ...accepted.viewport, widthPt: 720 }); } catch (error) { failure = (error as { code: string }).code; }
    const retained = JSON.stringify(a.preview.layout()) === JSON.stringify(accepted);
    bad = false; holdNext = true; const started = new Promise<void>(resolve => { began = resolve; });
    const older = a.preview.setViewport({ ...accepted.viewport, widthPt: 600 }); await started;
    const recovered = await a.preview.setViewport({ ...accepted.viewport, widthPt: 720 });
    const aborted = heldSignal!.aborted; release(); const olderResult = await older;
    const superseded = aborted && recovered.viewport.widthPt === 720 && olderResult.viewport.widthPt === 720 && a.preview.layout().viewport.widthPt === 720;
    const frames = a.chrome.querySelectorAll('iframe').length;
    a.preview.dispose(); b.preview.dispose(); const after = model.serialize();
    return { initial: { left: initial.maxScrollLeftPt, top: initial.maxScrollTopPt }, resized: { left: resized.maxScrollLeftPt, top: resized.maxScrollTopPt },
      fresh: { left: fresh.maxScrollLeftPt, top: fresh.maxScrollTopPt }, scrolled: { left: scrolled.maxScrollLeftPt, top: scrolled.maxScrollTopPt },
      retained, failure, frames, errors, serializations, superseded, unchanged: [...before].every(([path, bytes]) => bytes.every((value, index) => value === after.get(path)![index])),
      clean: !a.chrome.querySelector('iframe') && !b.chrome.querySelector('iframe') };
  } finally { release(); a.preview.dispose(); b?.preview.dispose(); aHost.remove(); bHost.remove(); }
}
(window as unknown as { workbookPreview: Record<string, unknown> }).workbookPreview.responsiveViewportLifecycle = responsiveViewportLifecycle;

export async function responsiveInlineViewport(entries: Entries) {
  const files = decodeTestPackage(entries);
  const manifest = JSON.parse(new TextDecoder().decode(files.get('manifest.json')!)) as { content: string };
  const parsed = new DOMParser().parseFromString(new TextDecoder().decode(files.get(manifest.content)!), 'text/html');
  for (const node of parsed.body.querySelectorAll<HTMLElement>('p,span')) node.style.fontSize = '5vw';
  files.set(manifest.content, new TextEncoder().encode('<!doctype html>' + parsed.documentElement.outerHTML.replaceAll('5vw', '5&#118;w')));
  const model = openDocument(files, { kind: 'workbook' }), aHost = host(), bHost = host();
  const samples = new WeakMap<HTMLIFrameElement, number[]>(), seen = new Set<HTMLIFrameElement>(), restore: (() => void)[] = [], observers: MutationObserver[] = [];
  const watch = (chrome: ShadowRoot) => {
    const connect = (frame: HTMLIFrameElement) => {
      if (seen.has(frame)) return; seen.add(frame);
      frame.addEventListener('load', () => {
        const view = frame.contentWindow!, computed = view.getComputedStyle, values: number[] = []; samples.set(frame, values);
        view.getComputedStyle = function(element, pseudo) {
          const style = computed.call(this, element, pseudo);
          if ((element as HTMLElement).style?.fontSize === '5vw') values.push(parseFloat(style.fontSize));
          return style;
        };
        restore.push(() => { try { view.getComputedStyle = computed; } catch { /* disposed frame */ } });
      }, { capture: true, once: true });
    };
    const collect = () => { for (const frame of chrome.querySelectorAll<HTMLIFrameElement>('iframe')) connect(frame); };
    const observer = new MutationObserver(collect); observer.observe(chrome, { childList: true, subtree: true }); observers.push(observer); collect();
  };
  const size = (chrome: ShadowRoot) => Math.max(0, ...samples.get(chrome.querySelector<HTMLIFrameElement>('iframe')!) ?? []);
  const a = mount(aHost, model, { widthPt: 240, heightPt: 400 }); watch(a.chrome);
  let b: ReturnType<typeof mount> | undefined;
  try {
    await a.preview.ready; const initial = size(a.chrome);
    await a.preview.setViewport({ ...a.preview.layout().viewport, widthPt: 480 }); const resized = size(a.chrome);
    b = mount(bHost, model, { widthPt: 480, heightPt: 400 }); watch(b.chrome); await b.preview.ready;
    return { initial, resized, fresh: size(b.chrome) };
  } finally { for (const observer of observers) observer.disconnect(); for (const reset of restore) reset(); a.preview.dispose(); b?.preview.dispose(); aHost.remove(); bHost.remove(); }
}
(window as unknown as { workbookPreview: Record<string, unknown> }).workbookPreview.responsiveInlineViewport = responsiveInlineViewport;

export async function autoRowRulesCancellation(entries: Entries) {
  const files = decodeTestPackage(entries);
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.defaultRowHeightCustom = false;
  metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: 'range', kind: 'whole', operator: 'between', origin: { row: 0, column: 0 },
    ranges: [{ from: { row: 0, column: 0 }, to: { row: 128, column: 0 } }], formulas: ['=1', '=1000'], allowBlank: true,
    showErrorMessage: true, showInputMessage: true, errorStyle: 'stop', errorTitle: '', error: '', promptTitle: '', prompt: '' }];
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), container = host();
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 240, heightPt: 60 });
  const { measureWorkbookColumnFont } = await import('../../../packages/render/src/workbook/column-font.js');
  const { measureWorkbookAutoRows } = await import('../../../packages/render/src/workbook/auto-rows.js');
  const { DEFAULT_WORKBOOK_RULE_LIMITS } = await import('@doffi/core');
  const formatter = createWorkbookNumberFormatter({ dateSystem: '1900' }), normal = measureWorkbookColumnFont(surface.source, snapshot.metadata.workbook, 'sheet-1');
  const abort = new AbortController(), rules = modelRuleSource(model); let calls = 0, held: AbortSignal | undefined, release = () => {}, began = () => {};
  const started = new Promise<void>(resolve => { began = resolve; });
  const source = { ...rules, async read(request: Parameters<typeof rules.read>[0], signal: AbortSignal) {
    if (++calls === 2) { held = signal; await new Promise<void>(resolve => { release = resolve; began(); }); }
    return rules.read(request, held === signal ? new AbortController().signal : signal);
  } };
  try {
    const work = measureWorkbookAutoRows(surface, modelWindowSource(model), 'sheet-1', normal, formatter, abort.signal, {}, { source, limits: DEFAULT_WORKBOOK_RULE_LIMITS });
    void work.catch(() => {}); await started;
    const retainedBeforeAbort = !!surface.source.querySelector('[data-office-view="workbook-source-measurement"]'); abort.abort();
    let error: string | undefined, deadline: ReturnType<typeof setTimeout> | undefined;
    try { await Promise.race([work, new Promise<never>((_resolve, reject) => { deadline = setTimeout(() => reject(Error('Cancelled rule read still waits')), 1000); })]); }
    catch (value) { error = (value as { code: string }).code; } finally { clearTimeout(deadline); }
    const cleanBeforeRelease = !surface.source.childNodes.length && !surface.view.childNodes.length;
    release(); await new Promise<void>(resolve => setTimeout(resolve, 30));
    return { calls, retainedBeforeAbort, error, cleanBeforeRelease, aborted: held!.aborted };
  } finally { release(); formatter.dispose(); surface.dispose(); container.remove(); }
}
(window as unknown as { workbookPreview: Record<string, unknown> }).workbookPreview.autoRowRulesCancellation = autoRowRulesCancellation;
