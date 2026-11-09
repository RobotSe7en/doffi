import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession, WorkbookTable } from "@doffi/core";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { createWorkbookRangeCopyTransaction } from "../src/workbook/copy-range.js";

const range = (row: number, column = 0, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const number = (value: number) => ({ type: "number" as const, value });
const sessions: WorkbookSession[] = [];
afterEach(() => { for (const s of sessions.splice(0)) s.dispose(); });
function apply(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations })); }
function example(calculatedFormula = "=Sales[@Amount]*2") {
  const s = createDocumentSession(scalarWorkbook([1, 2, 3].map((value, index) => ({ row: index + 1, column: 0, value: number(value) })))); sessions.push(s);
  const table: WorkbookTable = { tableId: "sales", name: "Sales", objectName: "SalesNative", range: range(0, 0, 3, 1), header: true, totals: false, totalsPreviouslyShown: false,
    columns: [{ columnId: "amount", name: "Amount", totalFunction: "none" }, { columnId: "double", name: "Double", calculatedFormula, totalFunction: "none" }],
    autoFilter: { range: range(0, 0, 3, 1), columns: [{ column: 0, hiddenButton: false, showButton: true }] },
    sort: { range: range(1, 0, 3, 1), caseSensitive: false, method: "none", keys: [{ column: 1, descending: true }] },
    style: { kind: "builtin", name: "TableStyleMedium2" }, rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false };
  apply(s, [{ op: "setWorkbookTable", sheetId: "sheet-1", table }, { op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 0, column: 0, source: { kind: "richText", html: '<span class="cell-default">Amount</span>' } },
    { row: 2, column: 1, source: { kind: "value", value: number(99) } }, { row: 3, column: 1, source: { kind: "value", value: { type: "blank" } } },
    { row: 5, column: 0, source: { kind: "formula", expression: "=SUM(Sales[Amount])" } },
  ] }]); s.clearHistory(); return s;
}
test.each([false, true])("whole table copy owns fresh definitions, rich headers, blank exceptions and current inputs (other sheet=%s)", async other => {
  const s = example(); if (other) apply(s, [{ op: "insertWorkbookSheet", name: "Other" }]); s.clearHistory();
  const before = s.snapshot(), target = other ? before.metadata.workbook.sheets[1]!.sheetId : "sheet-1";
  const transaction = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 5, 1), { row: 259, column: 7 }, { targetSheetId: target });
  expect(s.snapshot()).toEqual(before); apply(s, transaction.operations); s.calculateSync();
  const table = s.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === target)!.tables!.at(-1)!;
  expect(table).toMatchObject({ name: "Sales_2", objectName: "Sales_2", range: range(259, 7, 262, 8), columns: [{ name: "Amount" }, { name: "Double", calculatedFormula: expect.stringContaining("Sales_2") }],
    autoFilter: { range: range(259, 7, 262, 8), columns: [{ column: 7 }] }, sort: { range: range(260, 7, 262, 8), keys: [{ column: 8 }] }, style: { kind: "builtin", name: "TableStyleMedium2" } });
  expect(table.tableId).not.toBe("sales"); expect(table.columns.map(column => column.columnId)).not.toContain("amount");
  expect(s.getCell(target, 259, 7)?.html).toContain('<span class="cell-default">Amount</span>');
  expect(s.getCell(target, 261, 8)?.value).toEqual(number(99)); expect(s.getCell(target, 262, 8)?.formula).toBeUndefined();
  expect(s.getCell(target, 264, 7)?.formula?.expression).toBe("=SUM(Sales[Amount])");
  expect(s.getCalculatedCell(target, 260, 8)).toMatchObject({ value: number(2) });
  expect(s.snapshot().metadata.workbook.sheets[0]!.tables![0]!).toEqual(before.metadata.workbook.sheets[0]!.tables![0]);
  expect(s.history().undoSteps).toBe(1); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  apply(s, [{ op: "setWorkbookCells", sheetId: target, cells: [{ row: 260, column: 7, source: { kind: "value", value: number(5) } }] }]); s.calculateSync();
  expect(s.getCalculatedCell(target, 260, 8)).toMatchObject({ value: number(10) }); expect(s.getCalculatedCell("sheet-1", 1, 1)).toMatchObject({ value: number(2) });
  s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});
test("partial tables and selective paste retain ordinary structured bindings without inventing a table", async () => {
  for (const mode of ["contents", "formulas", "formats", "values"] as const) {
    const s = example(), before = s.snapshot(), calculation = s.calculateSync();
    const selected = mode === "contents" ? range(1, 0, 3, 1) : range(0, 0, 3, 1);
    apply(s, (await createWorkbookRangeCopyTransaction(s, "sheet-1", selected, { row: 259, column: 7 }, { mode, calculation })).operations);
    expect(s.snapshot().metadata.workbook.sheets[0]!.tables).toEqual(before.metadata.workbook.sheets[0]!.tables);
    if (mode === "contents") expect(s.getCell("sheet-1", 259, 8)?.formula?.expression).toContain("Sales[");
  }
});
test("table name/id collisions, same-position no-op, cancellation and overlapping destinations stay atomic", async () => {
  const s = example();
  apply(s, [{ op: "setWorkbookDefinedName", name: { nameId: "copied-table-1", name: "Sales_2", expression: "=1", origin: { row: 0, column: 0 } } }]); s.clearHistory();
  const before = s.serialize();
  expect(apply(s, (await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 3, 1), { row: 0, column: 0 })).operations).changed).toBe(false);
  expect(s.serialize()).toEqual(before);
  await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 3, 1), { row: 1, column: 0 })).rejects.toMatchObject({ code: "InvalidOperation" });
  await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 3, 1), { row: 259, column: 7 }, { maxSteps: 2 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 3, 1), { row: 259, column: 7 }, { signal: { aborted: true } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  apply(s, (await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 3, 1), { row: 259, column: 7 })).operations);
  expect(s.snapshot().metadata.workbook.sheets[0]!.tables!.at(-1)!.name).toBe("Sales_3");
});

test("empty copied table recipes retain their new identity and discover cross-sheet names before transaction assembly", async () => {
  const s = example("=Sales[@Amount]*Rate"), sourceTable = structuredClone(s.snapshot().metadata.workbook.sheets[0]!.tables![0]!) as WorkbookTable;
  sourceTable.range = range(0, 0, 0, 1); delete sourceTable.autoFilter; delete sourceTable.sort;
  apply(s, [{ op: "setWorkbookTable", sheetId: "sheet-1", table: sourceTable },
    { op: "setWorkbookDefinedName", name: { nameId: "rate", name: "Rate", sheetId: "sheet-1", expression: "=7", origin: { row: 0, column: 0 } } }, { op: "insertWorkbookSheet", name: "Other" }]); s.clearHistory();
  const target = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
  const transaction = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 0, 1), { row: 259, column: 7 }, { targetSheetId: target });
  expect(transaction.operations.some(operation => operation.op === "setWorkbookDefinedName" && operation.name.sheetId === target)).toBe(true);
  apply(s, transaction.operations);
  const table = structuredClone(s.snapshot().metadata.workbook.sheets[1]!.tables![0]!) as WorkbookTable;
  expect(table.columns[1]!.calculatedFormula).toContain("Sales_2"); table.range.to.row++;
  apply(s, [{ op: "setWorkbookTable", sheetId: target, table }, { op: "setWorkbookCells", sheetId: target, cells: [{ row: 260, column: 7, source: { kind: "value", value: number(3) } }] }]);
  s.calculateSync(); expect(s.getCalculatedCell(target, 260, 8)).toMatchObject({ value: number(21) });
});

test("large definition planning yields before cell scanning and cancels or rejects a replaced revision without partial output", async () => {
  const s = example(), table = structuredClone(s.snapshot().metadata.workbook.sheets[0]!.tables![0]!) as WorkbookTable;
  table.range = range(0, 0, 0, 399); delete table.autoFilter; delete table.sort;
  table.columns = Array.from({ length: 400 }, (_, index) => ({ columnId: "wide-" + index, name: "Column" + index, totalFunction: "none" }));
  apply(s, [{ op: "setWorkbookTable", sheetId: "sheet-1", table }]); s.clearHistory(); const before = s.serialize(), signal = { aborted: false }, progress: number[] = [];
  await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 0, 399), { row: 259, column: 0 }, { signal, onProgress: value => progress.push(value.completed), yield: async () => { signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(progress).toEqual([]); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 0, 399), { row: 259, column: 0 }, {
    yield: async () => { apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 500, column: 0, source: { kind: "value", value: number(7) } }] }]); },
  })).rejects.toMatchObject({ code: "RevisionConflict" });
  expect(s.snapshot().metadata.workbook.sheets[0]!.tables).toHaveLength(1); expect(s.getCell("sheet-1", 259, 0)).toBeUndefined();
});
