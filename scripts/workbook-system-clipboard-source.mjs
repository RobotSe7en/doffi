/** Current-source clipboard output; independent native input remains immutable. */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createDocumentSession, openDocument } from '@doffi/core';
import { importXlsx } from '@doffi/office';
import { createWorkbookClipboardData } from '@doffi/editor';

const directory = '.cache/workbook-system-clipboard-native'; mkdirSync(directory, { recursive: true });
const model = openDocument(await importXlsx(readFileSync('examples/xlsx/01-values-text.xlsx'), { documentId: 'system-clipboard-native' }), { kind: 'workbook' });
const session = createDocumentSession(model), sheetId = model.snapshot().metadata.workbook.sheets[0].sheetId;
const from = { row: 255, column: 31 }, to = { row: 257, column: 34 };
const value = (row, column, value) => ({ row: from.row + row, column: from.column + column, source: { kind: 'value', value } });
try {
  session.commit(session.prepare({ protocolVersion: 1, documentId: model.snapshot().manifest.documentId, baseRevision: 0, operations: [
    { op: 'setWorkbookNumberFormat', format: { formatId: 'system-clipboard-date', builtinId: 14 } },
    { op: 'setWorkbookNumberFormat', format: { formatId: 'system-clipboard-decimal', code: '0.00' } },
    { op: 'setWorkbookCells', sheetId, cells: [
      value(0, 0, { type: 'number', value: 12.5 }), { row: 255, column: 32, source: { kind: 'formula', expression: '=AF256*2' } },
      value(0, 2, { type: 'text', value: '001234567890123456789' }), value(0, 3, { type: 'boolean', value: true }),
      value(1, 0, { type: 'text', value: '合并😀禰󠄀' }), value(1, 2, { type: 'error', value: '#DIV/0!' }), value(1, 3, { type: 'text', value: '=1+1' }),
      value(2, 0, { type: 'text', value: '中文\r\n"甲\t乙"😀禰󠄀' }), { ...value(2, 1, { type: 'number', value: 60 }), formatId: 'system-clipboard-date' },
      { row: 257, column: 33, source: { kind: 'formula', expression: '=TRUE()' } }, { ...value(2, 3, { type: 'number', value: 12345.6789 }), formatId: 'system-clipboard-decimal' },
    ] },
    { op: 'formatWorkbookRanges', sheetId, ranges: [{ from, to: from }], properties: { 'font-weight': 'bold' } },
    { op: 'mergeWorkbookCells', sheetId, range: { from: { row: 256, column: 31 }, to: { row: 256, column: 32 } }, coveredContent: 'reject' },
  ] }));
  const calculation = await session.calculate(), data = await createWorkbookClipboardData(session, sheetId, { from, to }, { calculation });
  writeFileSync(directory + '/clipboard.html', data.html); writeFileSync(directory + '/clipboard.txt', data.text);
} finally { session.dispose(); }
