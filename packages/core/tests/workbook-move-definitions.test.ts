import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookSession, WorkbookTable } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const number = (value: number) => ({ type: "number" as const, value }), text = (value: string) => ({ type: "text" as const, value });
const sessions: WorkbookSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations))); }
function workbook() {
  const model = scalarWorkbook([
    { row: 0, column: 0, value: text("Amount") }, { row: 0, column: 1, value: text("Double") },
    { row: 1, column: 0, value: number(2) }, { row: 2, column: 0, value: number(3) },
    { row: 1, column: 1, value: number(999), formula: "=[@Amount]*2+$Z$1+Rate" }, { row: 2, column: 1, value: number(99) },
    { row: 3, column: 0, value: number(999), formula: "=SUBTOTAL(109,T[Amount])" }, { row: 3, column: 1, value: number(999), formula: "=SUBTOTAL(109,T[Double])" },
    { row: 0, column: 7, value: number(999), formula: "=SUM(T[Double])" }, { row: 1, column: 7, value: number(999), formula: "=A2" },
    { row: 0, column: 25, value: number(1) },
  ], metadata => {
    const sheet = metadata.workbook.sheets[0]!;
    metadata.workbook.sheets.push({ ...structuredClone(sheet), sheetId: "other", name: "Other" });
    metadata.workbook.names = [{ nameId: "source-rate", sheetId: sheet.sheetId, name: "Rate", origin: { row: 0, column: 0 }, expression: "=3" },
      { nameId: "other-rate", sheetId: "other", name: "Rate", origin: { row: 0, column: 0 }, expression: "=100" }];
    sheet.tables = [{ tableId: "sales", name: "T", objectName: "NativeT", range: range(0, 0, 3, 1), header: true, totals: true, totalsPreviouslyShown: true,
      showFirstColumn: false, showLastColumn: false, rowStripes: true, columnStripes: false, style: { kind: "builtin", name: "TableStyleMedium2" },
      columns: [{ columnId: "amount", name: "Amount", totalFunction: "sum" }, { columnId: "double", name: "Double", calculatedFormula: "=[@Amount]*2+$Z$1+Rate", totalFunction: "custom", totalFormula: "=SUBTOTAL(109,T[Double])" }],
      autoFilter: { range: range(0, 0, 2, 1), columns: [{ column: 0, hiddenButton: false, showButton: true, criterion: { kind: "custom", and: true, tests: [{ operator: "greaterThan", value: "0" }] } }], sort: { range: range(1, 0, 2, 1), keys: [{ column: 1, descending: true }], caseSensitive: false, method: "none" } },
      sort: { range: range(1, 0, 2, 1), keys: [{ column: 0, descending: false }], caseSensitive: false, method: "none" } }];
  });
  const session = createDocumentSession(model); sessions.push(session);
  apply(session, [{ op: "setWorkbookCells", sheetId: "other", cells: [{ row: 0, column: 25, source: { kind: "value", value: number(100) } }] }]); session.clearHistory(); return session;
}
const move = (targetSheetId: string, rangeInput = range(0, 0, 3, 1)): DocumentOperation => ({ op: "moveWorkbookRange", sheetId: "sheet-1", range: rangeInput, targetSheetId, target: { row: 259, column: 7 } });

test.each(["sheet-1", "other"])("whole-table cuts preserve identities, structured recipes, exceptions and filter/sort geometry (%s)", targetSheetId => {
  const s = workbook(), before = s.snapshot(), table = before.metadata.workbook.sheets[0]!.tables![0]!;
  const candidate = s.prepare(transaction(s, [move(targetSheetId)])); expect(s.snapshot()).toEqual(before); s.commit(candidate); s.calculateSync();
  const owner = s.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === targetSheetId)!, moved = owner.tables![0]!;
  expect(moved).toMatchObject({ tableId: table.tableId, name: "T", objectName: "NativeT", range: range(259, 7, 262, 8), style: table.style,
    columns: [{ columnId: "amount" }, { columnId: "double" }], autoFilter: { range: range(259, 7, 261, 8), columns: [{ column: 7 }], sort: { range: range(260, 7, 261, 8), keys: [{ column: 8 }] } },
    sort: { range: range(260, 7, 261, 8), keys: [{ column: 7 }] } });
  expect(moved.columns[1]!.calculatedFormula).toMatch(/\[@Amount\]|\[\[#This Row\],\[Amount\]\]/); expect(moved.columns[1]!.totalFormula).toContain("T[Double]");
  expect(s.getCell(targetSheetId, 260, 8)!.formula!.expression).toMatch(/\[@Amount\]|\[\[#This Row\],\[Amount\]\]/);
  expect(s.getCalculatedCell(targetSheetId, 260, 8)).toMatchObject({ state: "valid", value: number(8) });
  expect(s.getCell(targetSheetId, 261, 8)!.value).toEqual(number(99));
  expect(s.getCalculatedCell(targetSheetId, 262, 8)).toMatchObject({ state: "valid", value: number(107) });
  expect(s.getCalculatedCell("sheet-1", 0, 7)).toMatchObject({ value: number(107) }); expect(s.getCalculatedCell("sheet-1", 1, 7)).toMatchObject({ value: number(2) });
  const reopened = openDocument(s.serialize(), { kind: "workbook" }); expect(reopened.snapshot()).toEqual(s.snapshot());
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  s.redo(); expect(s.snapshot().metadata).toEqual(reopened.snapshot().metadata);
});

test("the relocated calculated template still fills only newly inserted rows in its new sheet context", () => {
  const s = workbook(); apply(s, [move("other")]);
  apply(s, [{ op: "editWorkbookAxis", sheetId: "other", axis: "row", action: "insert", index: 260, count: 1 }]); s.calculateSync();
  expect(s.getCalculatedCell("other", 260, 8)).toMatchObject({ state: "valid", value: number(4) });
  expect(s.getCalculatedCell("other", 261, 8)).toMatchObject({ state: "valid", value: number(8) }); expect(s.getCell("other", 262, 8)!.value).toEqual(number(99));
});

test("whole auto-filters move with criteria and sort but cannot replace an unrelated target definition", () => {
  const s = workbook(); apply(s, [{ op: "setWorkbookAutoFilter", sheetId: "sheet-1", filter: { range: range(10, 2, 13, 3), columns: [{ column: 2, hiddenButton: false, showButton: true }], sort: { range: range(11, 2, 13, 3), keys: [{ column: 3, descending: true }], method: "none", caseSensitive: false } } }]); s.clearHistory();
  const before = s.snapshot(); apply(s, [move("other", range(10, 2, 13, 3))]);
  expect(s.snapshot().metadata.workbook.sheets[0]!.autoFilter).toBeUndefined(); expect(s.snapshot().metadata.workbook.sheets[1]!.autoFilter).toMatchObject({ range: range(259, 7, 262, 8), columns: [{ column: 7 }], sort: { range: range(260, 7, 262, 8), keys: [{ column: 8 }] } });
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  apply(s, [{ op: "setWorkbookAutoFilter", sheetId: "other", filter: { range: range(20, 0, 22, 1), columns: [] } }]); const protectedSource = s.serialize();
  expect(() => apply(s, [move("other", range(10, 2, 13, 3))])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(protectedSource);
});

test("partial definitions, destination collisions and stale candidates retain source and full history", () => {
  const s = workbook(), before = s.serialize();
  expect(() => apply(s, [move("other", range(1, 0, 2, 1))])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before);
  const table = JSON.parse(JSON.stringify(s.snapshot().metadata.workbook.sheets[0]!.tables![0]!)) as WorkbookTable; table.tableId = "collision"; table.name = "Collision"; delete table.objectName; table.range = range(259, 7, 262, 8);
  delete table.autoFilter; delete table.sort;
  table.columns = table.columns.map(column => ({ ...column, columnId: "collision-" + column.columnId }));
  apply(s, [{ op: "setWorkbookTable", sheetId: "other", table }]); const protectedSource = s.serialize();
  expect(() => apply(s, [move("other")])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(protectedSource);
  const candidate = s.prepare(transaction(s, [move("sheet-1")])); apply(s, [{ op: "setWorkbookSettings", changes: { calculationMode: "manual" } }]);
  expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); s.discard(candidate);
});

test("moving whole definitions out and back in one batch leaves no revision, source rewrite or lost redo", () => {
  const s = workbook(); apply(s, [{ op: "setWorkbookSettings", changes: { calculationMode: "manual" } }]); s.undo();
  const before = s.serialize(), history = s.history();
  const receipt = apply(s, [move("other"), { op: "moveWorkbookRange", sheetId: "other", range: range(259, 7, 262, 8), targetSheetId: "sheet-1", target: { row: 0, column: 0 } }]);
  expect(s.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
  expect(s.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
  expect(receipt.changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
});

test("cross-sheet cuts retain workbook names and qualify original local bindings without adopting destination names", () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=Shared+Rate" }], metadata => {
    const source = metadata.workbook.sheets[0]!; metadata.workbook.sheets.push({ ...structuredClone(source), sheetId: "other", name: "Other" });
    metadata.workbook.names = [{ nameId: "shared", name: "Shared", origin: { row: 0, column: 0 }, expression: "=3" },
      { nameId: "source", name: "Rate", sheetId: source.sheetId, origin: { row: 0, column: 0 }, expression: "=5" },
      { nameId: "target", name: "Rate", sheetId: "other", origin: { row: 0, column: 0 }, expression: "=100" }];
  });
  const s = createDocumentSession(model); sessions.push(s); const before = s.serialize();
  apply(s, [{ op: "moveWorkbookRange", sheetId: "sheet-1", range: range(0, 0), targetSheetId: "other", target: { row: 2, column: 2 } }]); s.calculateSync();
  expect(s.getCell("other", 2, 2)!.formula!.expression).toBe("=(Shared+'Sheet1'!Rate)");
  expect(s.getCalculatedCell("other", 2, 2)).toMatchObject({ state: "valid", value: number(8) });
  s.undo(); expect(s.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
});

test("an unrepresentable global/local name collision rejects the complete cut rather than changing its binding", () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=Shared" }], metadata => {
    const source = metadata.workbook.sheets[0]!; metadata.workbook.sheets.push({ ...structuredClone(source), sheetId: "other", name: "Other" });
    metadata.workbook.names = [{ nameId: "shared", name: "Shared", origin: { row: 0, column: 0 }, expression: "=3" },
      { nameId: "target", name: "Shared", sheetId: "other", origin: { row: 0, column: 0 }, expression: "=100" }];
  });
  const s = createDocumentSession(model); sessions.push(s); const before = s.serialize();
  expect(() => apply(s, [{ op: "moveWorkbookRange", sheetId: "sheet-1", range: range(0, 0), targetSheetId: "other", target: { row: 2, column: 2 } }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
