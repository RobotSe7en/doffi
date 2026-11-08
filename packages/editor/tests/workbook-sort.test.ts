import { expect, test } from "vitest";
import { createDocumentSession } from "@doffi/core";
import type { WorkbookSort, WorkbookValue } from "@doffi/core";
import { createWorkbookSortTransaction } from "../src/index.js";
import { workbookSortComparison } from "../src/workbook/sort-order.js";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";

const sheetId = "sheet-1";
const definition = (rows: number, descending = false): WorkbookSort => ({ range: { from: { row: 0, column: 0 }, to: { row: rows - 1, column: 1 } },
  caseSensitive: false, method: "none", keys: [{ column: 0, descending }] });
test("Office's documented text/digit/punctuation/case order, ignored apostrophes/hyphens and blank/error rules", () => {
  const compare = workbookSortComparison("en-US", false), text = (value: string): WorkbookValue => ({ type: "text", value });
  const examples = ["0", "9", " ", "!", "+", "=", "A1", "A100", "A11", "A1X", "AB", "A-B", "B"];
  const actual = [...examples].reverse().sort((a, b) => compare(text(a), text(b), false)); expect(actual).toEqual(examples);
  expect(compare(text("O'Neil"), text("ONeil"), false)).toBe(0); expect(compare(text("Ab"), text("aB"), false)).toBe(0);
  expect(workbookSortComparison("en-US", true)(text("a"), text("A"), false)).toBeLessThan(0);
  expect(workbookSortComparison("en-US", true)(text("ab"), text("A1"), false)).toBeLessThan(0);
  expect(compare({ type: "error", value: "#N/A" }, { type: "error", value: "#REF!" }, false)).toBe(0);
  for (const descending of [false, true]) expect(compare({ type: "blank" }, { type: "number", value: -10 }, descending)).toBeGreaterThan(0);
});
test("sorting uses current formula results, stable multiple keys and keeps actual sources unchanged until one commit", async () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 999 }, formula: "=3" },
    { row: 1, column: 0, value: { type: "number", value: 0 }, formula: "=1" }, { row: 2, column: 0, value: { type: "number", value: 999 }, formula: "=1" },
    { row: 0, column: 1, value: { type: "text", value: "third" } }, { row: 1, column: 1, value: { type: "text", value: "first" } }, { row: 2, column: 1, value: { type: "text", value: "second" } }]));
  try {
    const before = s.serialize(), original = s.snapshot(), calculation = s.calculateSync(), transaction = await createWorkbookSortTransaction(s, sheetId, definition(3), { calculation });
    expect(s.serialize()).toEqual(before); expect(transaction.operations).toEqual([{ op: "reorderWorkbookRows", sheetId, range: definition(3).range, rows: [2, 0, 1] }]);
    s.commit(s.prepare(transaction)); expect(s.getCell(sheetId, 0, 1)?.text).toBe("first"); expect(s.getCell(sheetId, 1, 1)?.text).toBe("second");
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
  } finally { s.dispose(); }
});
test("manual-hidden and currently filtered rows stay at their addresses; old filter hidden bytes do not drive the sort", async () => {
  const s = createDocumentSession(scalarWorkbook([9, 100, 7, 50, 3].map((value, row) => ({ row, column: 0, value: { type: "number", value } })), metadata => {
    const sheet = metadata.workbook.sheets[0]!; sheet.rowDimensions = [{ from: 1, to: 1, hidden: true }];
    sheet.autoFilter = { range: { from: { row: 0, column: 0 }, to: { row: 4, column: 1 } }, columns: [{ column: 0, hiddenButton: false, showButton: true, criterion: { kind: "custom", and: true, tests: [{ operator: "lessThan", value: "10" }] } }] };
  }));
  try {
    const sort = definition(5); sort.range.from.row = 1;
    const transaction = await createWorkbookSortTransaction(s, sheetId, sort, { calculation: s.calculateSync() });
    expect(transaction.operations[0]).toMatchObject({ rows: [1, 4, 3, 2] });
    s.commit(s.prepare(transaction)); expect([0, 1, 2, 3, 4].map(row => s.getCell(sheetId, row, 0)?.value)).toEqual([9, 100, 3, 50, 7].map(value => ({ type: "number", value })));
  } finally { s.dispose(); }
});
test("cancelled merge-sort, caller mutation, source changes and foreign batches cannot publish an order", async () => {
  const s = createDocumentSession(scalarWorkbook(Array.from({ length: 600 }, (_, row) => ({ row, column: 0, value: { type: "number", value: 600 - row } }))));
  try {
    const before = s.serialize(), calculation = s.calculateSync(), abort = new AbortController();
    await expect(createWorkbookSortTransaction(s, sheetId, definition(600), { calculation, signal: abort.signal, yield: async () => {}, onProgress: value => { if (value.total === 0) abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
    const sort = definition(600); let changed = false;
    const transaction = await createWorkbookSortTransaction(s, sheetId, sort, { calculation, yield: async () => { if (!changed) { changed = true; sort.keys[0]!.descending = true; } } });
    expect(transaction.operations[0]).toMatchObject({ rows: Array.from({ length: 600 }, (_, row) => 599 - row) });
    await expect(createWorkbookSortTransaction(s, sheetId, definition(600), { calculation, yield: async () => {
      if (s.snapshot().manifest.revision === 0) s.commit(s.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookSettings", changes: { calculationMode: "manual" } }] }));
    } })).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
    await expect(createWorkbookSortTransaction(s, sheetId, definition(600), { calculation })).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  } finally { s.dispose(); }
});
