import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "../src/index.js";
import type { WorkbookMetadata, WorkbookDocumentMetadata } from "../src/index.js";
import { createWorkbookFormulaEquivalence } from "../src/workbook/formula-equivalence.js";

const origin = { sheetId: 'sheet-1', row: 0, column: 0 };
const area = (row: number, column: number, endRow = row, endColumn = column, sheetId = 'sheet-1') => ({ sheetId, from: { row, column }, to: { row: endRow, column: endColumn } });
const ranges = (...areas: ReturnType<typeof area>[]) => ({ kind: 'ranges', areas });
function model(edit: (workbook: WorkbookMetadata) => void = () => {}) {
  const files = new Map(createWorkbook({ documentId: 'references', sheetName: 'Data' }));
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets.push({ ...metadata.workbook.sheets[0]!, sheetId: 'second', name: '第二表' });
  edit(metadata.workbook); files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  return openDocument(files, { kind: 'workbook' });
}

describe('bounded current-source reference binding', () => {
  it('resolves mixed/cross-sheet/reversed references and whole axes without cell expansion', () => {
    const book = model();
    expect(book.resolveReference('=$B3:A$1', { ...origin, row: 2, column: 2 })).toEqual(ranges(area(0, 0, 2, 1)));
    expect(book.resolveReference("='第二表'!$XFD$1048576", origin)).toEqual(ranges(area(1_048_575, 16_383, 1_048_575, 16_383, 'second')));
    expect(book.resolveReference('=C:A', origin)).toEqual(ranges(area(0, 0, 1_048_575, 2)));
    expect(book.resolveReference('=3:1', origin)).toEqual(ranges(area(0, 0, 2, 16_383)));
    expect(book.resolveReference('=Missing!A1', origin)).toEqual({ kind: 'error', code: '#REF!' });
    expect(() => book.resolveReference('=A1', { ...origin, sheetId: 'missing' })).toThrow();
  });
  it('uses caller coordinates for relative names and explicit local scope for qualified names', () => {
    const book = model(workbook => { workbook.names = [
      { nameId: 'relative', name: 'Relative', expression: '=Data!$A1', origin: { row: 0, column: 0 } },
      { nameId: 'unqualified', name: 'Unqualified', expression: '=$A1', origin: { row: 0, column: 0 } },
      { nameId: 'global', name: 'Scoped', expression: '=$A$1', origin: { row: 0, column: 0 } },
      { nameId: 'local', name: 'Scoped', sheetId: 'sheet-1', expression: '=$A$2', origin: { row: 0, column: 0 } },
      { nameId: 'local-chain', name: 'LocalChain', sheetId: 'sheet-1', expression: '=Scoped', origin: { row: 0, column: 0 } },
    ]; });
    const caller = { sheetId: 'second', row: 3, column: 2 };
    expect(book.resolveReference('=Relative', caller)).toEqual(ranges(area(3, 0)));
    expect(book.resolveReference('=Unqualified', caller)).toEqual(ranges(area(3, 0, 3, 0, 'second')));
    expect(book.resolveReference('=Scoped', origin)).toEqual(ranges(area(1, 0)));
    expect(book.resolveReference('=Scoped', caller)).toEqual(ranges(area(0, 0, 0, 0, 'second')));
    expect(book.resolveReference('=Data!LocalChain', caller)).toEqual(ranges(area(0, 0)));
    expect(book.resolveReference('=Data!Unqualified', caller)).toEqual(ranges(area(3, 0, 3, 0, 'second')));
    const result = book.resolveReference('=Relative', caller); expect(Object.isFrozen(result)).toBe(true);
    if (result.kind === 'ranges') expect(Object.isFrozen(result.areas[0]?.from)).toBe(true);
    expect(openDocument(book.serialize(), { kind: 'workbook' }).resolveReference('=Relative', caller)).toEqual(result);
  });
  it('preserves union multiplicity and resolves intersections and range operators', () => {
    const book = model();
    expect(book.resolveReference('=(A1:A2,A2:A3)', origin)).toEqual(ranges(area(0, 0, 1, 0), area(1, 0, 2, 0)));
    expect(book.resolveReference('=A1:A3 A2:A4', origin)).toEqual(ranges(area(1, 0, 2, 0)));
    expect(book.resolveReference('=(A1:A2,A4:A5):A6', origin)).toEqual(ranges(area(0, 0, 5, 0)));
    expect(book.resolveReference('=A1 B1', origin)).toEqual({ kind: 'error', code: '#NULL!' });
    expect(book.resolveReference('=(A1,第二表!A1):A3', origin)).toEqual({ kind: 'error', code: '#VALUE!' });
  });
  it('binds bare table names, structured columns and the actual containing row', () => {
    const book = model(workbook => { workbook.sheets[0]!.tables = [{ tableId: 'sales', name: 'Sales', range: { from: { row: 2, column: 0 }, to: { row: 5, column: 1 } },
      header: false, totals: true, totalsPreviouslyShown: true, showFirstColumn: false, showLastColumn: false, rowStripes: false, columnStripes: false,
      columns: [{ columnId: 'name', name: 'Name', totalFunction: 'none' }, { columnId: 'amount', name: 'Amount', totalFunction: 'sum' }] }]; });
    expect(book.resolveReference('=Sales', origin)).toEqual(ranges(area(2, 0, 4, 1)));
    expect(book.resolveReference('=Sales[[#Totals],[Amount]]', origin)).toEqual(ranges(area(5, 1)));
    expect(book.resolveReference('=Sales[#Headers]', origin)).toEqual(ranges());
    expect(book.resolveReference('=[@Amount]', { ...origin, row: 3 })).toEqual(ranges(area(3, 1)));
    expect(book.resolveReference('=[@Amount]', origin)).toEqual({ kind: 'error', code: '#VALUE!' });
    expect(book.resolveReference('=Sales[@Amount]', origin)).toEqual({ kind: 'error', code: '#VALUE!' });
    expect(book.resolveReference('=Sales[Missing]', origin)).toEqual({ kind: 'error', code: '#REF!' });
  });
  it('reports computation and cycles while binding native nested names to the workbook', () => {
    const book = model(workbook => { workbook.names = [
      { nameId: 'cycle-a', name: 'CycleA', expression: '=CycleB', origin }, { nameId: 'cycle-b', name: 'CycleB', expression: '=CycleA', origin },
      { nameId: 'chain', name: 'Chain', expression: '=Scoped', origin }, { nameId: 'scoped', name: 'Scoped', sheetId: 'sheet-1', expression: '=$A$1', origin },
      { nameId: 'relative', name: 'Relative', expression: '=A1', origin: { row: 10, column: 10 } },
    ].map(name => ({ ...name, origin: { row: name.origin.row, column: name.origin.column } })); });
    expect(book.resolveReference('=CycleA', origin)).toEqual({ kind: 'unresolved', reason: 'circular-name' });
    expect(book.resolveReference('=Chain', origin)).toEqual({ kind: 'error', code: '#NAME?' });
    expect(book.resolveReference('=OFFSET(A1,1,0)', origin)).toEqual({ kind: 'unresolved', reason: 'calculation-required' });
    expect(book.resolveReference('=MissingName', origin)).toEqual({ kind: 'error', code: '#NAME?' });
    expect(book.resolveReference('=Relative', origin)).toEqual(ranges(area(1_048_566, 16_374)));
  });
  it('bounds recursive unions before expanding an exponential name graph', () => {
    const book = model(workbook => { workbook.names = Array.from({ length: 15 }, (_, i) => ({ nameId: `n-${i}`, name: `Range_${i}`, origin: { row: 0, column: 0 }, expression: i ? `=(Range_${i - 1},Range_${i - 1})` : '=A1' })); });
    expect(() => book.resolveReference('=Range_14', origin)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    expect(book.resolveReference('=Range_0', origin)).toEqual(ranges(area(0, 0)));
  });
  it('distinguishes changing name contexts from cycles', () => {
    const book = model(workbook => { workbook.names = [
      { nameId: 'global', name: 'Parent', expression: '=!Local', origin },
      { nameId: 'local', name: 'Local', sheetId: 'sheet-1', expression: '=[0]!Parent', origin },
    ].map(name => ({ ...name, origin: { row: 0, column: 0 } })); });
    expect(book.resolveReference('=Parent', origin)).toEqual({ kind: 'error', code: '#REF!' });
    expect(book.resolveReference('=Parent', { ...origin, sheetId: 'second' })).toEqual({ kind: 'error', code: '#NAME?' });
  });
  it('source restoration distinguishes bang context, a worksheet named bang and explicitly qualified global names', () => {
    const book = model(workbook => {
      workbook.sheets[1]!.name = '!';
      workbook.names = [{ nameId: 'global', name: 'Caller', expression: '=!$A$1', origin: { row: 0, column: 0 } }];
    });
    const equal = createWorkbookFormulaEquivalence(book.snapshot().metadata.workbook);
    expect(equal('=!$A$1', "='!'!$A$1", origin, origin.sheetId)).toBe(false);
    expect(equal('=SUM(!A:A)', "=SUM('!'!A:A)", origin, origin.sheetId)).toBe(false);
    expect(equal('=Caller', '=[0]!Caller', origin, origin.sheetId)).toBe(false);
    expect(equal('=Caller', '=Data!Caller', origin, origin.sheetId)).toBe(false);
    expect(equal('=[0]!Caller', '=Data!Caller', origin, origin.sheetId)).toBe(true);
  });
});
