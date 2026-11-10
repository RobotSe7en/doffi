import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { expect, test } from 'vitest';
import { createDocumentSession, openDocument } from '@doffi/core';
import type { DocumentOperation, WorkbookValue } from '@doffi/core';
import { importXlsx, readDocumentZip, writeDocumentZip } from '@doffi/office';

interface Operation { name: string; index: number; count: number; range: string | null; expected: WorkbookValue }
const evidence = JSON.parse(readFileSync('docs/evidence/workbook-filter-headers.json', 'utf8')) as {
  sourceSha256: string; operations: Operation[];
  cases: { name: string; operation: string; reapply: boolean; outputFile: string; outputSha256: string;
    observed: { range: string | null; formula: string; cached: WorkbookValue }; matchesCurrent: boolean }[];
};
const bytes = readFileSync('packages/office/tests/fixtures/workbook-filter-headers.xlsx');
const range = (ref: string) => {
  const point = (value: string) => { const match = /^([A-Z]+)(\d+)$/.exec(value)!; let column = 0; for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64; return { row: Number(match[2]) - 1, column: column - 1 }; };
  const [from, to = from] = ref.split(':'); return { from: point(from!), to: point(to!) };
};
async function session() { return createDocumentSession(openDocument(await importXlsx(bytes, { documentId: 'filter-headers' }), { kind: 'workbook' })); }
const transaction = (s: Awaited<ReturnType<typeof session>>, operations: DocumentOperation[]) => {
  const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations };
};
test('independent ordinary-filter source and eight unmodified producer artifacts retain their exact hashes', () => {
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
  for (const item of evidence.cases) expect(createHash('sha256').update(readFileSync(item.outputFile)).digest('hex')).toBe(item.outputSha256);
  expect(execFileSync('python3', ['scripts/check-workbook-filter-headers.py'], { encoding: 'utf8' })).toContain('eight native');
});
test.each(evidence.operations)('independent filter deletion $name preserves exact current inputs, criteria, package and history', async item => {
  const s = await session();
  try {
    const before = s.snapshot(), sheet = before.metadata.workbook.sheets[0]!, observer = before.metadata.workbook.sheets[1]!.sheetId;
    const old = s.readWindow({ sheetId: sheet.sheetId, row: 0, column: 0, rows: 8, columns: 8 });
    const candidate = s.prepare(transaction(s, [{ op: 'editWorkbookAxis', sheetId: sheet.sheetId, axis: 'row', action: 'delete', index: item.index, count: item.count }]));
    expect(s.snapshot()).toEqual(before); s.commit(candidate);
    const filter = s.snapshot().metadata.workbook.sheets[0]!.autoFilter;
    expect(filter?.range).toEqual(item.range ? range(item.range) : undefined);
    if (filter) expect(filter.columns).toEqual(sheet.autoFilter!.columns);
    for (const cell of old) {
      if (cell.row >= item.index && cell.row < item.index + item.count) continue;
      const row = cell.row >= item.index ? cell.row - item.count : cell.row;
      expect(s.getCell(sheet.sheetId, row, cell.column)?.html).toBe(cell.html);
    }
    s.calculateSync(); expect(s.getCalculatedCell(observer, 0, 0)).toMatchObject({ state: 'valid', value: item.expected });
    expect(openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: 'workbook' }).snapshot()).toEqual(s.snapshot());
    expect(s.history().undoSteps).toBe(1); s.undo();
    expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
    s.calculateSync(); expect(s.getCalculatedCell(observer, 0, 0)).toMatchObject({ value: { type: 'number', value: 70 } });
    s.redo(); s.calculateSync(); expect(s.getCalculatedCell(observer, 0, 0)).toMatchObject({ value: item.expected });
  } finally { s.dispose(); }
});
test.each(evidence.cases)('actual producer $name reopens its own caches and independently recalculates the current filter', async item => {
  const input = readFileSync(item.outputFile), original = Buffer.from(input), model = openDocument(await importXlsx(input, { documentId: 'filter-header-producer' }), { kind: 'workbook' });
  const observer = model.snapshot().metadata.workbook.sheets[1]!.sheetId;
  expect(model.getCell(observer, 0, 0)).toMatchObject({ value: item.observed.cached, formula: { expression: '=' + item.observed.formula, cacheState: 'imported' } });
  const s = createDocumentSession(model);
  try {
    s.calculateSync(); expect(s.getCalculatedCell(observer, 0, 0)).toMatchObject({ value: evidence.operations.find(operation => operation.name === item.operation)!.expected });
    expect(openDocument(readDocumentZip(writeDocumentZip(model.serialize())), { kind: 'workbook' }).snapshot()).toEqual(model.snapshot());
    expect(input).toEqual(original);
  } finally { s.dispose(); }
});
test('a later invalid write, cancellation after structural work and a stale candidate cannot publish a partial filter deletion', async () => {
  const s = await session();
  try {
    const before = s.serialize(), sheetId = s.snapshot().metadata.workbook.sheets[0]!.sheetId, observer = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
    const axis: DocumentOperation = { op: 'editWorkbookAxis', sheetId, axis: 'row', action: 'delete', index: 1, count: 1 };
    const value = { row: 4, column: 4, source: { kind: 'value' as const, value: { type: 'number' as const, value: 9 } } };
    expect(() => s.prepare(transaction(s, [axis, { op: 'setWorkbookCells', sheetId: observer, cells: [value, value] }]))).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
    expect(s.serialize()).toEqual(before); const abort = { aborted: false }; let structural = false;
    await expect(s.prepareAsync(transaction(s, [axis]), undefined, { signal: abort, yield: async () => {}, onProgress(progress) {
      if (progress.phase === 'operations') { structural = true; abort.aborted = true; }
    } })).rejects.toMatchObject({ code: 'OperationCancelled' });
    expect(structural).toBe(true); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
    const candidate = s.prepare(transaction(s, [axis]));
    s.commit(s.prepare(transaction(s, [{ op: 'setWorkbookCells', sheetId: observer, cells: [value] }])));
    const external = s.serialize(); expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: 'RevisionConflict' })); s.discard(candidate);
    expect(s.serialize()).toEqual(external);
  } finally { s.dispose(); }
});
