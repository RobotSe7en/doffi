import { expect, test } from 'vitest';
import { createDocumentSession, DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS, openDocument, readWorkbookRowVisibility } from '../src/index.js';
import type { DocumentOperation, WorkbookValue } from '../src/index.js';
import { scalarWorkbook } from './workbook-scalar-fixture.js';

const point = (row: number, column: number) => ({ from: { row, column }, to: { row, column } });
const number = (value: number): WorkbookValue => ({ type: 'number', value });
function setup(code: string, value: WorkbookValue, selected: string, table = false, stop = false, locale?: string, dateSystem: '1900' | '1904' = '1900') {
  return createDocumentSession(scalarWorkbook([
    { row: 0, column: 0, value: { type: 'text', value: 'Key' } }, { row: 0, column: 1, value: { type: 'text', value: 'Amount' } },
    { row: 1, column: 0, value }, { row: 1, column: 1, value: number(10) }, { row: 0, column: 3, value: number(999), formula: '=SUBTOTAL(9,B2:B2)' },
  ], metadata => {
    if (locale) metadata.locale = locale; const workbook = metadata.workbook; workbook.dateSystem = dateSystem;
    workbook.numberFormats.push({ formatId: 'display', code }, { formatId: 'other', code: '0.0' }); const sheet = workbook.sheets[0]!;
    sheet.columnDimensions = [{ from: 0, to: 0, unit: 'characters', numberFormatId: 'display' }];
    const filter = { range: { from: { row: 0, column: 0 }, to: { row: 1, column: 1 } }, columns: [{ column: 0, hiddenButton: false, showButton: true, criterion: { kind: 'values' as const, values: [selected], dates: [], blank: false } }] };
    if (table) sheet.tables = [{ tableId: 'sales', name: 'Sales', range: filter.range, autoFilter: filter, header: true, totals: false, totalsPreviouslyShown: false,
      columns: [{ columnId: 'key', name: 'Key', totalFunction: 'none' }, { columnId: 'amount', name: 'Amount', totalFunction: 'none' }], rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false }];
    else sheet.autoFilter = filter;
    if (stop) sheet.dataValidations = [{ ruleId: 'subtotal-positive', ranges: [point(0, 3)], origin: { row: 0, column: 3 }, kind: 'whole', operator: 'greaterThan', formulas: ['=0'], allowBlank: false,
      showErrorMessage: true, showInputMessage: false, errorStyle: 'stop', error: 'positive subtotal', errorTitle: '', prompt: '', promptTitle: '' }];
  }));
}
for (const table of [false, true]) test.each([
  ['0.00', number(1), '1.00', undefined, '1900'], ['0.00', number(1.004), '1.00', undefined, '1900'],
  ['0%', number(.5), '50%', undefined, '1900'], ['yyyy-mm-dd', number(45351), '2024-02-29', undefined, '1900'],
  ['yyyy-mm-dd', number(59), '1904-02-29', undefined, '1904'], ['0.00', number(1.5), '1,50', 'de-DE', '1900'],
  ['[$-409]0.00', number(1.5), '1.50', 'de-DE', '1900'], ['"ID-"@', { type: 'text', value: '001' }, 'ID-001', undefined, '1900'],
] as const)(`${table ? 'table' : 'ordinary'} list filter matches current display format: %s %j %s`, async (code, value, selected, locale, system) => {
  const s = setup(code, value, selected, table, false, locale, system);
  try {
    const source = s.serialize(), calculation = s.calculateSync(); expect(s.getCalculatedCell('sheet-1', 0, 3)).toMatchObject({ value: number(10) });
    const rows = await readWorkbookRowVisibility(s, { ...DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS, documentId: 'scalar-kernel', revision: 0, sheetId: 'sheet-1', from: 1, to: 1 }, { calculation });
    expect(rows.rows).toEqual([{ row: 1, hidden: false, filtered: false }]); expect(s.serialize()).toEqual(source);
  } finally { s.dispose(); }
});
test('retains native scalar spellings and rejects unrelated display labels', () => {
  for (const [selected, expected] of [['1', 10], ['1.01', 0]] as const) {
    const s = setup('0.00', number(1), selected); try { s.calculateSync(); expect(s.getCalculatedCell('sheet-1', 0, 3)).toMatchObject({ value: number(expected) }); } finally { s.dispose(); }
  }
});
const formatting: DocumentOperation[] = [
  { op: 'setWorkbookNumberFormat', format: { formatId: 'display', code: '0.0' } },
  { op: 'setWorkbookCellFormats', sheetId: 'sheet-1', cells: [{ row: 1, column: 0, formatId: 'other' }] },
  { op: 'updateWorkbookSheet', sheetId: 'sheet-1', changes: { rowDimensions: [{ from: 1, to: 1, numberFormatId: 'other' }] } },
  { op: 'updateWorkbookSheet', sheetId: 'sheet-1', changes: { columnDimensions: [{ from: 0, to: 0, unit: 'characters', numberFormatId: 'other' }] } },
];
test.each(formatting)('format-only %j changes filtered SUBTOTAL and cannot bypass a stop validation', operation => {
  const s = setup('0.00', number(1), '1.00', false, true);
  try {
    s.calculateSync(); expect(s.getCalculatedCell('sheet-1', 0, 3)).toMatchObject({ value: number(10) }); const before = s.serialize(), history = s.history(), m = s.snapshot().manifest;
    expect(() => s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [operation] })).toThrowError(expect.objectContaining({ code: 'InvalidOperation', details: [expect.objectContaining({ code: 'validation-blocked' })] }));
    expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  } finally { s.dispose(); }
});
test('format lookup work obeys the row-query step budget instead of building an unaccounted catalog', async () => {
  const s = setup('0.00', number(1), '1.00');
  try {
    const files = new Map(s.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!));
    metadata.workbook.numberFormats.unshift(...Array.from({ length: 150 }, (_, index) => ({ formatId: 'unused-' + index, code: '0.000' })));
    files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata))); const model = openDocument(files, { kind: 'workbook' });
    await expect(readWorkbookRowVisibility(model, { ...DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS, maxSteps: 100, documentId: 'scalar-kernel', revision: 0, sheetId: 'sheet-1', from: 1, to: 1 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    const result = await readWorkbookRowVisibility(model, { ...DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS, documentId: 'scalar-kernel', revision: 0, sheetId: 'sheet-1', from: 1, to: 1 });
    expect(result.rows[0]!.filtered).toBe(false); expect(result.steps).toBeGreaterThan(150); expect(s.history().undoSteps).toBe(0);
  } finally { s.dispose(); }
});
test('unrelated format definitions preserve an imported stop violation and unrelated row inheritance', () => {
  const s = setup('0.00', number(1), 'unmatched', false, true);
  try {
    const m = s.snapshot().manifest, before = s.snapshot();
    s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [
      { op: 'setWorkbookNumberFormat', format: { formatId: 'other', code: '0.000' } },
      { op: 'updateWorkbookSheet', sheetId: 'sheet-1', changes: { rowDimensions: [{ from: 100, to: 100, numberFormatId: 'other' }] } },
    ] }));
    expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); s.calculateSync(); expect(s.getCalculatedCell('sheet-1', 0, 3)).toMatchObject({ value: number(0) });
  } finally { s.dispose(); }
});

test('splitting identical format runs for row heights cannot revive an unrelated unavailable stop rule', () => {
  const initial = setup('0.00', number(1), '1.00'), files = new Map(initial.serialize()); initial.dispose();
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!));
  const sheet = metadata.workbook.sheets[0]; sheet.columnDimensions = []; sheet.rowDimensions = [{ from: 1, to: 2, numberFormatId: 'display' }];
  sheet.dataValidations = [{ ruleId: 'unavailable', kind: 'custom', expression: '=AND(SUBTOTAL(9,B2:B2)>0,UNKNOWN())', origin: { row: 0, column: 3 }, ranges: [point(0, 3)],
    allowBlank: false, showInputMessage: false, showErrorMessage: true, errorStyle: 'stop', error: '', errorTitle: '', prompt: '', promptTitle: '' }];
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata))); const s = createDocumentSession(openDocument(files, { kind: 'workbook' }));
  try {
    const before = s.snapshot(), m = before.manifest;
    s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [{ op: 'updateWorkbookSheet', sheetId: 'sheet-1', changes: { rowDimensions: [
      { from: 1, to: 1, numberFormatId: 'display', heightPt: 20 }, { from: 2, to: 2, numberFormatId: 'display', heightPt: 21 },
    ] } }] }));
    expect(s.history().undoSteps).toBe(1); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
