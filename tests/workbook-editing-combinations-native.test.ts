import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { expect, test } from 'vitest';
import { createDocumentSession, openDocument } from '@doffi/core';
import type { DocumentOperation, WorkbookEditProposal, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookSession } from '@doffi/core';
import { connectWorkbookSession, createWorkbookRangeCopyTransaction, serveWorkbookSession } from '@doffi/editor';
import { importXlsx, readDocumentZip, writeDocumentZip } from '@doffi/office';
import { mixedEditingOperations } from './workbook-combination-operations.js';

const source = readFileSync('packages/office/tests/fixtures/workbook-editing-combinations.xlsx');
const evidence = JSON.parse(readFileSync('docs/evidence/workbook-editing-combinations.json', 'utf8')) as { sourceSha256: string; outputSha256: string; producerImport: { code: string }; producer: { duplicateVmlIds: string[] } };
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const tx = (s: WorkbookSession, operations: DocumentOperation[]) => ({ protocolVersion: 1 as const, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations });
// Explicit point metrics isolate core composition. Browser cases use the actual
// first-party font/layout service for wrapping, filtering and drawing bounds.
const metrics = (request: WorkbookLayoutRequest): WorkbookLayoutMeasurement => ({ sourceDigest: request.sourceDigest, fontEnvironment: 'mixed-core-point-metrics', sheets: request.sheetIds.map(sheetId => ({ sheetId, rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
async function session() { return createDocumentSession(openDocument(await importXlsx(source, { documentId: 'mixed-native' }), { kind: 'workbook' })); }
const apply = (s: WorkbookSession, operations: DocumentOperation[]) => s.commit(s.prepare(tx(s, operations), { measureWorkbookLayout: metrics }));
function current(s: WorkbookSession, row: number, column: number, value: number, sheetId = 'sheet-1') { expect(s.getCalculatedCell(sheetId, row, column)).toMatchObject({ state: 'valid', value: { type: 'number', value } }); }
function exactUndo(s: WorkbookSession, before: ReturnType<WorkbookSession['snapshot']>) {
  s.undo(); const after = s.snapshot(); expect(after.metadata).toEqual(before.metadata); expect(after.content).toEqual(before.content);
  expect(after.styles).toEqual(before.styles); expect(after.annotations.map(({ baseRevision, ...annotation }) => annotation)).toEqual(before.annotations.map(({ baseRevision, ...annotation }) => annotation));
  expect(after.manifest.chunks).toEqual(before.manifest.chunks); expect(after.manifest.assets).toEqual(before.manifest.assets);
}
function reopen(s: WorkbookSession) { expect(openDocument(new Map(readDocumentZip(writeDocumentZip(s.serialize()))), { kind: 'workbook' }).snapshot()).toEqual(s.snapshot()); }
test('mixed source is independently reproducible and retains untouched producer differences', async () => {
  expect(createHash('sha256').update(source).digest('hex')).toBe(evidence.sourceSha256);
  expect(execFileSync('python3', ['scripts/check-workbook-editing-combinations.py'], { encoding: 'utf8' })).toContain('independent mixed');
  const bytes = readFileSync('packages/office/tests/fixtures/workbook-editing-combinations-calc.xlsx'); expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.outputSha256);
  expect(evidence.producer.duplicateVmlIds).toContain('_x0000_t202');
  await expect(importXlsx(bytes, { documentId: 'mixed-producer' })).rejects.toMatchObject({ code: evidence.producerImport.code });
  const s = await session();
  try {
    const snap = s.snapshot(); expect(snap.metadata.workbook.objects).toHaveLength(4); expect(snap.metadata.workbook.charts).toHaveLength(1);
    expect(snap.metadata.workbook.comments!.map(comment => comment.format).sort()).toEqual(['classic', 'modern', 'modern']);
    expect(snap.annotations.map(root => root.status)).toEqual(['resolved', 'open', 'open']);
    expect(snap.annotations.flatMap(root => root.replies ?? [])).toHaveLength(1); expect(s.getCell('sheet-1', 1, 0)?.text).toBe('甲禰󠄀 Latin😀');
    expect(s.getCell('sheet-1', 1, 3)?.value).toEqual({ type: 'number', value: 999 }); s.calculateSync(); current(s, 5, 3, 240); current(s, 0, 0, 240, 'sheet-2'); current(s, 1, 0, 10, 'sheet-2');
    expect(s.snapshot()).toEqual(snap); reopen(s);
  } finally { s.dispose(); }
});
test('one isolated mixed command recalculates table filters, keeps rich source/resources and restores every family through history', async () => {
  const s = await session();
  try {
    const before = s.snapshot(), rich = s.getCell('sheet-1', 1, 0)!, candidate = s.prepare(tx(s, mixedEditingOperations(before)), { measureWorkbookLayout: metrics });
    expect(s.snapshot()).toEqual(before); expect(candidate.preview.metadata.workbook.sheets[0]!.tables![0]!.autoFilter!.columns).toHaveLength(2);
    s.commit(candidate); s.calculateSync(); current(s, 5, 3, 120); current(s, 0, 0, 280, 'sheet-2'); current(s, 1, 0, 14, 'sheet-2');
    expect(s.getCell('sheet-1', 1, 0)?.html.slice(s.getCell('sheet-1', 1, 0)!.html.indexOf('>'))).toBe(rich.html.slice(rich.html.indexOf('>')));
    expect(s.snapshot().metadata.workbook.objects!.map(object => object.nodeId)).toEqual(before.metadata.workbook.objects!.map(object => object.nodeId));
    expect(s.snapshot().manifest.assets).toEqual(before.manifest.assets); expect(s.snapshot().annotations.flatMap(root => root.replies ?? [])).toHaveLength(2);
    expect(s.snapshot().metadata.workbook.names).toEqual(expect.arrayContaining([expect.objectContaining({ name: 'CurrentAmounts', expression: '=SalesTable[金额]' })]));
    expect(s.history().undoSteps).toBe(1); reopen(s); const accepted = s.snapshot(); exactUndo(s, before); s.redo(); expect(s.snapshot().metadata).toEqual(accepted.metadata); reopen(s);
  } finally { s.dispose(); }
});
test('late validation and unauthorized assets roll back prior definition, style, drawing and reply operations', async () => {
  const s = await session();
  try {
    const before = s.serialize(), history = s.history(), operations = mixedEditingOperations(s.snapshot());
    const invalid: DocumentOperation[] = [
      { op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row: 2, column: 1, source: { kind: 'value', value: { type: 'number', value: 20 } } }] },
      { op: 'replaceWorkbookImage', nodeId: s.snapshot().metadata.workbook.objects!.find(object => object.kind === 'image')!.nodeId, assetPath: 'assets/not-authorized.png', alt: '未授权替换' },
    ];
    for (const operation of invalid) { expect(() => s.prepare(tx(s, [...operations, operation]), { measureWorkbookLayout: metrics })).toThrow(); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history); }
    const stop = new AbortController(); let measured = false;
    await expect(s.prepareAsync(tx(s, operations), { measureWorkbookLayout: request => { measured = true; stop.abort(); return metrics(request); } }, { signal: stop.signal })).rejects.toMatchObject({ code: 'OperationCancelled' });
    expect(measured).toBe(true); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
    const stale = s.prepare(tx(s, operations), { measureWorkbookLayout: metrics }); apply(s, [{ op: 'updateWorkbookSheet', sheetId: 'sheet-2', changes: { showGridLines: false } }]); const newer = s.serialize();
    expect(() => s.commit(stale)).toThrowError(expect.objectContaining({ code: 'RevisionConflict' })); s.discard(stale); expect(s.serialize()).toEqual(newer);
  } finally { s.dispose(); }
});
test.each(['copy', 'move'] as const)('whole table %s across sheets carries the rich source, both comment kinds and cell-owned objects without losing formula bindings', async action => {
  const s = await session();
  try {
    const before = s.snapshot(), sourceRange = range(0, 0, 5, 3), target = { row: 259, column: 7 };
    const transaction = action === 'copy' ? await createWorkbookRangeCopyTransaction(s, 'sheet-1', sourceRange, target, { targetSheetId: 'sheet-2' })
      : tx(s, [{ op: 'moveWorkbookRange', sheetId: 'sheet-1', range: sourceRange, targetSheetId: 'sheet-2', target }]);
    const prepared = s.prepare(transaction, { measureWorkbookLayout: metrics }); expect(s.snapshot()).toEqual(before); s.commit(prepared); s.calculateSync();
    const after = s.snapshot(), table = after.metadata.workbook.sheets[1]!.tables![0]!;
    expect(table.name).toBe(action === 'copy' ? 'SalesTable_2' : 'SalesTable'); expect(table.range).toEqual(range(259, 7, 264, 10));
    expect(s.getCell('sheet-2', 260, 7)?.text).toBe('甲禰󠄀 Latin😀'); current(s, 264, 10, 240, 'sheet-2'); current(s, 0, 0, 240, 'sheet-2');
    expect(after.metadata.workbook.comments!.filter(comment => after.annotations.find(annotation => annotation.annotationId === comment.annotationId)!.anchor.kind === 'cells')).toHaveLength(action === 'copy' ? 6 : 3);
    expect(after.metadata.workbook.objects!.filter(object => object.sheetId === 'sheet-2')).toHaveLength(1);
    expect(after.manifest.assets).toEqual(before.manifest.assets); expect(s.history().undoSteps).toBe(1); reopen(s); exactUndo(s, before);
  } finally { s.dispose(); }
});
test('mixed global proposal freezes every source family, previews current differences and rejects late application', async () => {
  const s = await session();
  try {
    const before = s.snapshot(), open = before.annotations.find(root => root.status === 'open')!, input = s.buildInteractionRequest({ kind: 'workbook', annotationIds: [open.annotationId], scope: { kind: 'global' } });
    const frozen = structuredClone(input), edit: WorkbookEditProposal = { protocolVersion: 1, documentId: input.documentId, baseRevision: input.baseRevision, requestId: input.requestId, proposalId: 'mixed', annotationIds: [open.annotationId], summary: '复核完整混合候选', operations: mixedEditingOperations(before) };
    const candidate = s.prepareProposal(input, edit, { measureWorkbookLayout: metrics }), review = s.getProposalReview(candidate)!;
    expect(input).toEqual(frozen); expect(s.snapshot()).toEqual(before); expect(candidate.preview.annotations.find(root => root.annotationId === open.annotationId)!.status).toBe('resolved');
    expect(review.calculatedChanges).toEqual(expect.arrayContaining([expect.objectContaining({ sheetId: 'sheet-2', row: 0, column: 0, after: { kind: 'value', value: { type: 'number', value: 280 } } })]));
    expect(candidate.preview.metadata).not.toEqual(before.metadata); expect(candidate.preview.styles).not.toEqual(before.styles);
    s.commit(candidate); expect(s.history().undoSteps).toBe(1); reopen(s); exactUndo(s, before);
    expect(() => s.prepareProposal(input, edit, { measureWorkbookLayout: metrics })).toThrowError(expect.objectContaining({ code: 'RevisionConflict' }));
  } finally { s.dispose(); }
});
test('sole Worker accepts a measured mixed batch and preserves package, history and host ownership', async () => {
  const s = await session(), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), command = { kind: 'transaction' as const, transaction: tx(s, mixedEditingOperations(before)) };
    const options = { measureLayout: (request: { sourceDigest: string; sheetIds: readonly string[] }) => ({ sourceDigest: request.sourceDigest, fontEnvironment: 'mixed-worker-points', sheets: request.sheetIds.map(sheetId => ({ sheetId, rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) }) };
    const discard = await client.prepare(command, undefined, options); await client.discard(discard); expect(s.snapshot()).toEqual(before);
    await client.commit(await client.prepare(command, undefined, options)); s.calculateSync(); current(s, 5, 3, 120); current(s, 0, 0, 280, 'sheet-2');
    expect(client.state().history.undoSteps).toBe(1); reopen(s); await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    await expect(client.prepare(command, undefined, options)).rejects.toMatchObject({ code: 'RevisionConflict' });
    client.dispose(); server.dispose(); apply(s, [{ op: 'updateWorkbookSheet', sheetId: 'sheet-2', changes: { showGridLines: false } }]); expect(s.history().undoSteps).toBe(1);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
