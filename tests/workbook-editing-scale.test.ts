import { expect, test } from 'vitest';
import { createDocumentSession, createWorkbook, createWorkbookCalculator, openDocument } from '@doffi/core';
import type { DocumentOperation, WorkbookSession, WorkbookTable } from '@doffi/core';
import { readDocumentZip, writeDocumentZip } from '@doffi/office';

const daily = process.env.DOFFI_CI_MODE === 'daily';
const number = (value: number) => ({ type: 'number' as const, value });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const apply = (s: WorkbookSession, operations: DocumentOperation[]) => { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations })); };
const setup = () => createDocumentSession(openDocument(createWorkbook({ documentId: 'supplemental-capacity' }), { kind: 'workbook' }));
test.skipIf(daily)('one source edit recalculates exactly 1000 fan-out dependents while retaining the unrelated result @scale', () => {
  const s = setup();
  try {
    apply(s, [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [
      { row: 0, column: 0, source: { kind: 'value', value: number(2) } },
      ...Array.from({ length: 1000 }, (_, row) => ({ row, column: 1, source: { kind: 'formula' as const, expression: `=$A$1+${row + 1}` } })),
      { row: 0, column: 2, source: { kind: 'formula', expression: '=7*6' } },
    ] }]); s.clearHistory(); const before = s.snapshot(), calculator = createWorkbookCalculator(s);
    try {
      const first = calculator.calculateSync(); expect(first).toMatchObject({ calculated: 1001, reused: 0 });
      apply(s, [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row: 0, column: 0, source: { kind: 'value', value: number(3) } }] }]);
      expect(calculator.calculateSync()).toMatchObject({ calculated: 1000, reused: 1 });
      for (let row = 0; row < 1000; row++) expect(calculator.getCell({ sheetId: 'sheet-1', row, column: 1 })).toMatchObject({ value: number(row + 4) });
      expect(first.cells.find(cell => cell.row === 999 && cell.column === 1)).toMatchObject({ value: number(1002) }); expect(s.history().undoSteps).toBe(1);
      s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(calculator.calculateSync().calculated).toBe(1000);
      expect(calculator.getCell({ sheetId: 'sheet-1', row: 999, column: 1 })).toMatchObject({ value: number(1002) });
    } finally { calculator.dispose(); }
  } finally { s.dispose(); }
});
test.skipIf(daily)('expands a native table model to 10000 actual data rows with current formulas, exceptions and exact one-step restoration @scale', async () => {
  const s = setup(), definition: WorkbookTable = { tableId: 'capacity-table', name: 'Capacity', range: range(0, 0, 2, 1), header: true, totals: false, totalsPreviouslyShown: false,
    columns: [{ columnId: 'price', name: 'Price', totalFunction: 'none' }, { columnId: 'double', name: 'Double', calculatedFormula: '=[@Price]*2', totalFunction: 'none' }],
    rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false };
  try {
    apply(s, [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: Array.from({ length: 10000 }, (_, index) => ({ row: index + 1, column: 0, source: { kind: 'value' as const, value: number(index + 1) } })) },
      { op: 'setWorkbookTable', sheetId: 'sheet-1', table: definition },
      { op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row: 2, column: 1, source: { kind: 'value', value: number(99) } }, { row: 0, column: 3, source: { kind: 'formula', expression: '=SUM(Capacity[Double])' } }] }]);
    s.clearHistory(); const before = s.snapshot(), m = before.manifest;
    const candidate = await s.prepareAsync({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision,
      operations: [{ op: 'setWorkbookTable', sheetId: 'sheet-1', table: { ...definition, range: range(0, 0, 10000, 1) } }] });
    expect(s.snapshot()).toEqual(before); s.commit(candidate); const batch = s.calculateSync();
    expect(batch.cells).toHaveLength(10000); expect(s.getCell('sheet-1', 2, 1)?.value).toEqual(number(99));
    expect(s.getCalculatedCell('sheet-1', 10000, 1)).toMatchObject({ value: number(20000) });
    expect(s.getCalculatedCell('sheet-1', 0, 3)).toMatchObject({ value: number(100010095) });
    const current = s.snapshot(); expect(current.manifest.chunks).toHaveLength(40);
    expect(openDocument(new Map(readDocumentZip(writeDocumentZip(s.serialize()))), { kind: 'workbook' }).snapshot()).toEqual(current);
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    s.redo(); s.calculateSync(); expect(s.getCalculatedCell('sheet-1', 10000, 1)).toMatchObject({ value: number(20000) });
  } finally { s.dispose(); }
});
