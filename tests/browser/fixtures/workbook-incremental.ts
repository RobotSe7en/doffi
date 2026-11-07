import { captureWorkbookModel, createDocumentSession, DocumentError, openDocument, readWorkbookChangeImpact } from '@doffi/core';
import { createWorkbookPreviewSource, linkWorkbookPreviewSource, mountDocumentPreview } from '@doffi/render';
import type { WorkbookPreviewSource } from '@doffi/render';
import { decodeTestPackage, type TestPackage } from './package-transport.js';

/** Observe authorized reads, never expose source DOM from the public preview. */
export async function incrementalPreview(entries: TestPackage) {
  const session = createDocumentSession(openDocument(decodeTestPackage(entries), { kind: 'workbook' }));
  const before = captureWorkbookModel(session), firstBatch = session.calculateSync(), first = createWorkbookPreviewSource(before, firstBatch);
  const source = (value: WorkbookPreviewSource, reads: number[], reject = false, hold?: () => Promise<void>): WorkbookPreviewSource => ({ ...value,
    cells: { ...value.cells, async read(request, signal) {
      for (const window of request.windows) for (let row = window.row; row < window.row + window.rows; row++) reads.push(row);
      if (request.windows.some(window => window.row <= 1000 && window.row + window.rows > 1000)) {
        if (reject) throw new DocumentError('RevisionConflict', 'Deliberate failure during distant row measurement.');
        if (hold) await hold();
      }
      return value.cells.read(request, signal);
    } },
  });
  const firstReads: number[] = [], previous = source(first, firstReads), errors: string[] = [], unhandled: string[] = [];
  const rejection = (event: PromiseRejectionEvent) => { unhandled.push(String(event.reason)); event.preventDefault(); };
  window.addEventListener('unhandledrejection', rejection);
  function mount(input: WorkbookPreviewSource, left: number) {
    const host = document.createElement('div'); host.style.cssText = `position:fixed;left:${left}px;top:0;background:white`; document.body.append(host);
    const attach = Element.prototype.attachShadow; let chrome!: ShadowRoot;
    Element.prototype.attachShadow = function(init) { const root = attach.call(this, init); chrome ??= root; return root; };
    try {
      const preview = mountDocumentPreview(host, input, { widthPt: 420, heightPt: 300, retainSelection: true, onError: error => errors.push(error.code) });
      return { host, chrome, preview };
    }
    finally { Element.prototype.attachShadow = attach; }
  }
  const mounted = mount(previous, 0), { preview, chrome } = mounted;
  let reference: ReturnType<typeof mount> | undefined, release = () => {};
  try {
    await preview.ready; const originalFrame = chrome.querySelector('iframe')!, original = preview.layout();
    const far = { row: 1001, column: 0 }; await preview.selectRange({ from: far, to: far });
    const oldTop = preview.layout().viewport.scrollTopPt;
    await preview.setViewport({ ...preview.layout().viewport, scrollTopPt: 0 });
    const manifest = session.snapshot().manifest;
    session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision,
      operations: [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row: 0, column: 0, source: { kind: 'value', value: { type: 'number', value: 2 } } }] }] }));
    const nextModel = captureWorkbookModel(session), nextBatch = session.calculateSync(), next = createWorkbookPreviewSource(nextModel, nextBatch);
    const impact = readWorkbookChangeImpact(before, nextModel, firstBatch, nextBatch);
    const failed = linkWorkbookPreviewSource(source(next, [], true), previous, impact);
    preview.update(failed); let failure = ''; try { await preview.settled(); } catch (error) { failure = (error as DocumentError).code; }
    const retained = preview.layout().revision === original.revision && chrome.querySelector('iframe') === originalFrame && chrome.querySelectorAll('iframe').length === 1;

    let begin!: () => void; const started = new Promise<void>(resolve => { begin = resolve; });
    const gate = new Promise<void>(resolve => { release = resolve; });
    const slow = linkWorkbookPreviewSource(source(next, [], false, () => { begin(); return gate; }), previous, impact);
    preview.update(slow); await started; const pending = preview.settled();
    const reads: number[] = [], accepted = linkWorkbookPreviewSource(source(next, reads), previous, impact);
    preview.update(accepted); await preview.settled(); await pending;
    const beforeRelease = preview.layout().revision; release(); await new Promise<void>(resolve => setTimeout(resolve, 0));
    const superseded = beforeRelease === nextModel.snapshot().manifest.revision && preview.layout().revision === beforeRelease;
    const sameFrame = chrome.querySelector('iframe') === originalFrame, frames = chrome.querySelectorAll('iframe').length;
    const measuredDistant = reads.includes(1000), skippedUntouched = !reads.includes(2000);
    await preview.selectRange({ from: far, to: far }); const updated = preview.layout().viewport;
    reference = mount(next, 650); await reference.preview.ready; await reference.preview.selectRange({ from: far, to: far });
    const fresh = reference.preview.layout().viewport;
    return { failure, retained, superseded, sameFrame, frames, measuredDistant, skippedUntouched, oldTop, updated, fresh, errors, unhandled,
      initialReadRows: new Set(firstReads).size, updatedReadRows: new Set(reads).size, calculated: session.getCalculatedCell('sheet-1', 1000, 1) };
  } finally {
    release(); reference?.preview.dispose(); reference?.host.remove(); preview.dispose(); mounted.host.remove(); session.dispose(); window.removeEventListener('unhandledrejection', rejection);
  }
}
