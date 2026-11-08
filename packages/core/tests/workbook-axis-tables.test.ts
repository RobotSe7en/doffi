import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument, parseWorkbookFormula } from "../src/index.js";
import type { DocumentOperation, WorkbookSession, WorkbookTable } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sessions: WorkbookSession[] = [], sheetId = "sheet-1";
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const number = (value: number) => ({ type: "number" as const, value });
const table = (tableId = "sales", column = 0, row = 0): WorkbookTable => ({ tableId, name: tableId, range: range(row, column, row + 4, column + 1), header: true, totals: true, totalsPreviouslyShown: true,
  columns: [{ columnId: tableId + "-a", name: "Amount", totalFunction: "none", totalLabel: "Total" }, { columnId: tableId + "-b", name: "Double", calculatedFormula: "=A2*2", totalFunction: "sum" }],
  rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false,
  autoFilter: { range: range(row, column, row + 3, column + 1), columns: [{ column, hiddenButton: false, showButton: true, criterion: { kind: "custom", and: true, tests: [{ operator: "greaterThan", value: "0" }] } }] },
  sort: { range: range(row + 1, column, row + 3, column + 1), caseSensitive: false, method: "none", keys: [{ column, descending: true }] } });
function apply(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations })); }
const axis = (axis: "row" | "column", action: "insert" | "delete", index: number, count = 1): DocumentOperation => ({ op: "editWorkbookAxis", sheetId, axis, action, index, count, ...(action === "insert" ? { formatSource: "none" as const } : {}) });
function example() {
  const s = createDocumentSession(scalarWorkbook([1, 2, 3].map((value, index) => ({ row: index + 1, column: 0, value: number(value) })))); sessions.push(s);
  apply(s, [{ op: "setWorkbookTable", sheetId, table: table() }, { op: "setWorkbookCells", sheetId, cells: [{ row: 2, column: 1, source: { kind: "value", value: number(99) } }] }]); s.clearHistory(); return s;
}
test("inserted table data rows materialize only their own formulas, retain exceptions/caches and rebase the first-body template", () => {
  for (const index of [1, 2, 4]) {
    const s = example(), before = s.snapshot(), exception = s.getCell(sheetId, 2, 1)!;
    apply(s, [axis("row", "insert", index, 2)]);
    const current = s.snapshot().metadata.workbook.sheets[0]!.tables![0]!;
    expect(current.range).toEqual(range(0, 0, 6, 1)); expect(parseWorkbookFormula(current.columns[1]!.calculatedFormula!, { row: 1, column: 1 })).toEqual(parseWorkbookFormula("=A2*2", { row: 1, column: 1 }));
    for (const row of [index, index + 1]) expect(s.getCell(sheetId, row, 1)?.formula?.expression).toBe("=(A" + (row + 1) + "*2)");
    expect(s.getCell(sheetId, index <= 2 ? 4 : 2, 1)?.html.replace(/data-column="\d+"/, 'data-column="1"')).toBe(exception.html);
    s.calculateSync(); expect(s.getCalculatedCell(sheetId, 6, 1)).toMatchObject({ value: number(107) });
    expect(s.history().undoSteps).toBe(1); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  }
});
test("deleting a first data row rebases templates while deleting all data retains future relative recipes and invalidates absolute inputs", () => {
  const s = example(); apply(s, [axis("row", "delete", 1)]);
  expect(parseWorkbookFormula(s.snapshot().metadata.workbook.sheets[0]!.tables![0]!.columns[1]!.calculatedFormula!, { row: 1, column: 1 })).toEqual(parseWorkbookFormula("=A2*2", { row: 1, column: 1 })); expect(s.getCell(sheetId, 1, 1)?.value).toEqual(number(99));
  s.undo();
  const current = JSON.parse(JSON.stringify(s.snapshot().metadata.workbook.sheets[0]!.tables![0]!)) as WorkbookTable;
  current.columns[1]!.calculatedFormula = "=SUM(A$1:A2)+$A$2+A1"; apply(s, [{ op: "setWorkbookTable", sheetId, table: current }]); s.clearHistory();
  apply(s, [axis("row", "delete", 1, 3)]);
  const empty = s.snapshot().metadata.workbook.sheets[0]!.tables![0]!;
  expect(empty.range).toEqual(range(0, 0, 1, 1)); expect(empty.columns[1]!.calculatedFormula).toBe("=((SUM(A$1:A2)+#REF!)+A1)");
  expect(empty.autoFilter?.range).toEqual(range(0, 0, 0, 1)); expect(empty.sort).toBeUndefined();
  apply(s, [axis("row", "insert", 1)]); expect(s.getCell(sheetId, 1, 1)?.formula?.expression).toBe(empty.columns[1]!.calculatedFormula);
});
test("table column inserts allocate unique names/identities; deleted column bindings become errors without redirecting to neighbors", () => {
  const s = example(), before = s.snapshot();
  apply(s, [axis("column", "insert", 1, 2)]);
  let current = s.snapshot().metadata.workbook.sheets[0]!.tables![0]!;
  expect(current.columns.map(column => column.name)).toEqual(["Amount", "Column1", "Column2", "Double"]); expect(new Set(current.columns.map(column => column.columnId)).size).toBe(4);
  expect(current.columns[3]!.calculatedFormula).toBe("=(A2*2)"); expect(s.getCell(sheetId, 0, 1)?.text).toBe("Column1"); expect(s.getCell(sheetId, 2, 3)?.value).toEqual(number(99));
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 8, source: { kind: "formula", expression: "=SUM(sales[Amount])" } }] }, axis("column", "delete", 0)]);
  current = s.snapshot().metadata.workbook.sheets[0]!.tables![0]!;
  expect(current.columns.map(column => column.columnId)).toEqual(["sales-b"]); expect(current.columns[0]!.calculatedFormula).toContain("#REF!");
  expect(s.getCell(sheetId, 0, 7)?.formula?.expression).toBe("=SUM(#REF!)"); s.calculateSync(); expect(s.getCalculatedCell(sheetId, 0, 7)).toMatchObject({ value: { type: "error", value: "#REF!" } });
});
test("table totals deletion drops the source row without promoting data and full removal invalidates outside references", () => {
  const s = example(); apply(s, [axis("row", "delete", 4)]);
  let current = s.snapshot().metadata.workbook.sheets[0]!.tables![0]!;
  expect(current.totals).toBe(false); expect(current.totalsPreviouslyShown).toBe(true); expect(s.getCell(sheetId, 3, 1)?.formula?.expression).toBe("=(A4*2)");
  s.undo(); apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 8, column: 0, source: { kind: "formula", expression: "=SUM(sales[Amount])" } }] }, axis("row", "delete", 0, 5)]);
  expect(s.snapshot().metadata.workbook.sheets[0]!.tables).toEqual([]); expect(s.getCell(sheetId, 3, 0)?.formula?.expression).toBe("=SUM(#REF!)");
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
});
test("surviving header deletion, native bounds and cancellation reject the entire table batch", async () => {
  const s = example(), before = s.serialize();
  for (const operation of [axis("row", "delete", 0), axis("column", "insert", 1, 16_384)]) {
    expect(() => apply(s, [operation])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before);
  }
  const abort = { aborted: false }, m = s.snapshot().manifest;
  await expect(s.prepareAsync({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [axis("row", "insert", 1, 2000)] }, {},
    { signal: abort, yield: async () => { abort.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("later table column deletions are rewritten before any earlier table captures its calculated recipe", () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 11, column: 0, value: number(7) }])); sessions.push(s);
  const front = table("front", 3), later = table("later", 0, 10); front.columns[1]!.calculatedFormula = "=SUM(later[Amount])"; later.columns[1]!.calculatedFormula = "=A12*2";
  apply(s, [{ op: "setWorkbookTable", sheetId, table: front }, { op: "setWorkbookTable", sheetId, table: later }]); s.clearHistory();
  apply(s, [axis("column", "delete", 0)]);
  expect(s.snapshot().metadata.workbook.sheets[0]!.tables![0]!.columns[1]!.calculatedFormula).toBe("=SUM(#REF!)");
  expect(s.getCell(sheetId, 1, 3)?.formula?.expression).toBe("=SUM(#REF!)");
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 1, 3)).toMatchObject({ value: { type: "error", value: "#REF!" } });
});

test("inverse table axis batches retain exact templates, source bytes and redo without an extra revision", () => {
  for (const direction of ["row", "column"] as const) {
    const s = example(); apply(s, [axis("row", "insert", 2)]); s.undo(); const before = s.serialize(), history = s.history();
    expect(apply(s, [axis(direction, "insert", 1, 2), axis(direction, "delete", 1, 2)]).changed).toBe(false);
    expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  }
});
