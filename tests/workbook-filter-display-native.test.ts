import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { expect, test } from 'vitest';
import { createDocumentSession, openDocument } from '@doffi/core';
import type { WorkbookAutoFilter } from '@doffi/core';
import { importXlsx, readDocumentZip, writeDocumentZip } from '@doffi/office';
const source = readFileSync('packages/office/tests/fixtures/workbook-filter-display.xlsx');
const evidence = JSON.parse(readFileSync('docs/evidence/workbook-filter-display.json', 'utf8')) as { sourceSha256: string; outputSha256: string; cases: { sheet: string; expected: number; selected: string; table: boolean }[]; producer: { subtotal: string; type: string; hiddenRows: number[] }[] };
test('independent ordinary/table format filters reproduce every input and untouched native observation', () => {
  expect(createHash('sha256').update(source).digest('hex')).toBe(evidence.sourceSha256);
  expect(execFileSync('python3', ['scripts/check-workbook-filter-display.py'], { encoding: 'utf8' })).toContain('twelve independent');
});
test.each(evidence.cases)('native $sheet current display filter ignores source caches and reopens exact history', async item => {
  const s = createDocumentSession(openDocument(await importXlsx(source, { documentId: 'native-format-filter' }), { kind: 'workbook' }));
  try {
    const before = s.snapshot(), sheet = before.metadata.workbook.sheets.find(sheet => sheet.name === item.sheet)!;
    expect(s.getCell(sheet.sheetId, 0, 3)?.value).toEqual({ type: 'number', value: 999 }); s.calculateSync();
    expect(s.getCalculatedCell(sheet.sheetId, 0, 3), item.sheet).toMatchObject({ state: 'valid', value: { type: 'number', value: item.expected } });
    const filter = item.table ? sheet.tables![0]!.autoFilter! : sheet.autoFilter!;
    expect(filter.columns[0]!.criterion).toMatchObject({ kind: 'values', values: [item.selected] });
    const m = before.manifest, changed: WorkbookAutoFilter = { ...JSON.parse(JSON.stringify(filter)), columns: [] };
    s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: item.table
      ? [{ op: 'setWorkbookTable', sheetId: sheet.sheetId, table: { ...JSON.parse(JSON.stringify(sheet.tables![0]!)), autoFilter: changed } }]
      : [{ op: 'setWorkbookAutoFilter', sheetId: sheet.sheetId, filter: changed }] }));
    s.calculateSync(); expect(s.getCalculatedCell(sheet.sheetId, 0, 3)).toMatchObject({ value: { type: 'number', value: 30 } });
    expect(openDocument(new Map(readDocumentZip(writeDocumentZip(s.serialize()))), { kind: 'workbook' }).snapshot()).toEqual(s.snapshot());
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    s.calculateSync(); expect(s.getCalculatedCell(sheet.sheetId, 0, 3)).toMatchObject({ value: { type: 'number', value: item.expected } });
  } finally { s.dispose(); }
});
test('actual Calc formatted filters retain their own cached results and manual visibility without becoming SDK expectations', async () => {
  const bytes = readFileSync('packages/office/tests/fixtures/workbook-filter-display-calc.xlsx');
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.outputSha256);
  const model = openDocument(await importXlsx(bytes, { documentId: 'native-format-filter-producer' }), { kind: 'workbook' });
  for (const [index, sample] of evidence.producer.entries()) {
    const sheet = model.snapshot().metadata.workbook.sheets[index]!;
    expect(model.getCell(sheet.sheetId, 0, 3)?.value).toEqual(sample.type === 'e' ? { type: 'error', value: sample.subtotal } : { type: 'number', value: Number(sample.subtotal) });
    for (const row of sample.hiddenRows) expect(sheet.rowDimensions.some(span => span.from <= row && span.to >= row && span.hidden)).toBe(true);
  }
  expect(openDocument(new Map(readDocumentZip(writeDocumentZip(model.serialize()))), { kind: 'workbook' }).snapshot()).toEqual(model.snapshot());
});
