import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { expect, test } from 'vitest';
import { createDocumentSession, createWorkbookCalculator, openDocument } from '@doffi/core';
import type { DocumentOperation, WorkbookValue } from '@doffi/core';
import { importXlsx, readDocumentZip, writeDocumentZip } from '@doffi/office';

interface Case { family: string; expression: string; expected: WorkbookValue; observed: WorkbookValue; matchesProducer: boolean; basis: string; address: string }
interface Evidence { dateSystem: '1900' | '1904'; files: Record<'source' | 'reference' | 'producer', { path: string; sha256: string }>; cases: Case[] }
const evidence = JSON.parse(readFileSync('docs/evidence/workbook-function-boundaries.json', 'utf8')) as { workbooks: Evidence[] };
function equalValue(actual: unknown, expected: WorkbookValue, label: string) {
  expect(actual, label).toMatchObject({ type: expected.type });
  if (expected.type === 'number') expect(Math.abs((actual as { value: number }).value - expected.value), label).toBeLessThanOrEqual(Math.max(1e-12, Math.abs(expected.value) * 1e-12));
  else expect(actual, label).toEqual(expected);
}
test('pins all independent cached/uncached inputs and untouched actual producer files for two calendars', () => {
  expect(execFileSync('python3', ['scripts/check-workbook-function-boundaries.py'], { encoding: 'utf8' })).toContain('268 independent');
  expect(evidence.workbooks.map(item => item.dateSystem)).toEqual(['1900', '1904']);
  for (const item of evidence.workbooks) for (const file of Object.values(item.files)) expect(createHash('sha256').update(readFileSync(file.path)).digest('hex')).toBe(file.sha256);
});
test.each(evidence.workbooks)('$dateSystem independent current results ignore wrong caches and remain exact after a package reopen', async item => {
  const files = await importXlsx(readFileSync(item.files.source.path), { documentId: 'boundary-' + item.dateSystem });
  const model = openDocument(files, { kind: 'workbook' }), before = model.serialize(), sheetId = model.snapshot().metadata.workbook.activeSheetId;
  expect(model.snapshot().metadata.workbook.dateSystem).toBe(item.dateSystem);
  const calculator = createWorkbookCalculator(model), result = calculator.calculateSync();
  expect(result.cells.length).toBe(item.cases.length + 1);
  const differences: unknown[] = [];
  for (const [row, sample] of item.cases.entries()) {
    expect(model.getCell(sheetId, row, 4)?.value).toEqual({ type: 'number', value: 999 });
    const value = calculator.getCell({ sheetId, row, column: 4 });
    const actual = value.state === 'valid' ? value.value : undefined;
    const matches = actual?.type === sample.expected.type && (actual.type === 'number' && sample.expected.type === 'number'
      ? Math.abs(actual.value - sample.expected.value) <= Math.max(1e-12, Math.abs(sample.expected.value) * 1e-12)
      : JSON.stringify(actual) === JSON.stringify(sample.expected));
    if (!matches) differences.push({ expression: sample.expression, actual: value, expected: sample.expected });
  }
  expect(differences).toEqual([]); expect(model.serialize()).toEqual(before);
  const reopened = openDocument(new Map(readDocumentZip(writeDocumentZip(before))), { kind: 'workbook' }), again = createWorkbookCalculator(reopened);
  const batch = again.calculateSync(); expect(batch.cells.map(cell => cell.state === 'valid' ? cell.value : cell.state)).toEqual(result.cells.map(cell => cell.state === 'valid' ? cell.value : cell.state));
  expect(reopened.snapshot()).toEqual(model.snapshot()); again.dispose(); calculator.dispose();
});
test.each(evidence.workbooks)('$dateSystem unmodified producer caches retain every independently observed value', async item => {
  const bytes = readFileSync(item.files.producer.path), model = openDocument(await importXlsx(bytes, { documentId: 'producer-boundary-' + item.dateSystem }), { kind: 'workbook' });
  const sheetId = model.snapshot().metadata.workbook.activeSheetId, before = model.serialize();
  expect(model.snapshot().metadata.workbook.dateSystem).toBe(item.dateSystem);
  for (const [row, sample] of item.cases.entries()) {
    equalValue(model.getCell(sheetId, row, 4)?.value, sample.observed, sample.expression);
    expect(sample.matchesProducer).toBe(JSON.stringify(sample.expected) === JSON.stringify(sample.observed));
  }
  expect(openDocument(new Map(readDocumentZip(writeDocumentZip(before))), { kind: 'workbook' }).snapshot()).toEqual(model.snapshot());
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(item.files.producer.sha256);
});
test.each(evidence.workbooks)('$dateSystem current typed inputs share candidate isolation, history, cancellation and stale-source protection', async item => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync(item.files.source.path), { documentId: 'boundary-edit-' + item.dateSystem }), { kind: 'workbook' }));
  const sheetId = s.snapshot().metadata.workbook.activeSheetId, row = item.cases.findIndex(sample => sample.expression === '=SUM(A1:A4)');
  const tx = (operations: DocumentOperation[]) => ({ protocolVersion: 1 as const, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations });
  const set = (value: number): DocumentOperation => ({ op: 'setWorkbookCells', sheetId, cells: [{ row: 0, column: 0, source: { kind: 'value', value: { type: 'number', value } } }] });
  try {
    const before = s.serialize(), prepared = s.prepare(tx([set(2.25)])), candidate = createWorkbookCalculator(prepared.model);
    candidate.calculateSync(); expect(candidate.getCell({ sheetId, row, column: 4 })).toMatchObject({ value: { type: 'number', value: 2.25 } }); candidate.dispose();
    expect(s.serialize()).toEqual(before); s.commit(prepared); s.calculateSync(); expect(s.getCalculatedCell(sheetId, row, 4)).toMatchObject({ value: { type: 'number', value: 2.25 } });
    expect(s.history().undoSteps).toBe(1); s.undo(); s.calculateSync(); expect(s.getCalculatedCell(sheetId, row, 4)).toMatchObject({ value: { type: 'number', value: 1.25 } });
    s.redo(); const current = s.serialize(), history = s.history();
    expect(() => s.prepare(tx([set(3.25), { op: 'setWorkbookCells', sheetId, cells: [{ row: 0, column: 1, source: { kind: 'formula', expression: '=?' } }] }]))).toThrow();
    expect(s.serialize()).toEqual(current); expect(s.history()).toEqual(history);
    const abort = new AbortController(); let reached = false;
    await expect(s.prepareAsync(tx([{ op: 'setWorkbookCells', sheetId, cells: Array.from({ length: 1000 }, (_, index) => ({ row: index, column: 0, source: { kind: 'value' as const, value: { type: 'number' as const, value: index } } })) }]), undefined,
      { signal: abort.signal, yield: async () => {}, onProgress(value) { if (value.phase === 'cells' && value.completed >= 256) { reached = true; abort.abort(); } } })).rejects.toMatchObject({ code: 'OperationCancelled' });
    expect(reached).toBe(true); expect(s.serialize()).toEqual(current); expect(s.history()).toEqual(history);
    const stale = s.prepare(tx([set(4.25)])); s.commit(s.prepare(tx([set(5.25)]))); const latest = s.serialize();
    expect(() => s.commit(stale)).toThrowError(expect.objectContaining({ code: 'RevisionConflict' })); s.discard(stale); expect(s.serialize()).toEqual(latest);
  } finally { s.dispose(); }
});
