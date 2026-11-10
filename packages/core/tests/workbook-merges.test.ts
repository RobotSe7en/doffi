import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sheetId = "sheet-1", cleanups: WorkbookSession[] = [];
afterEach(() => { for (const session of cleanups.splice(0)) session.dispose(); });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
function session(model = openDocument(createWorkbook({ documentId: "merges" }), { kind: "workbook" })) { const value = createDocumentSession(model); cleanups.push(value); return value; }
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { return { protocolVersion: 1 as const, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations }; }
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations))); }
const value = (row: number, column: number, amount: number): DocumentOperation => ({ op: "setWorkbookCells", sheetId, cells: [{ row, column, source: { kind: "value", value: { type: "number", value: amount } } }] });

test("cross-chunk merges reject covered content by default and explicit clearing recalculates atomically with undo", () => {
  const s = session(scalarWorkbook([{ row: 255, column: 31, value: { type: "number", value: 7 } }, { row: 255, column: 32, value: { type: "number", value: 3 } },
    { row: 256, column: 31, value: { type: "number", value: 8 } }, { row: 260, column: 0, value: { type: "number", value: 999 }, formula: "=SUM(AF256:AG257)" }]));
  const area = range(255, 31, 256, 32), before = s.serialize(), original = s.snapshot();
  expect(() => s.prepare(transaction(s, [{ op: "mergeWorkbookCells", sheetId, range: area }]))).toThrowError(expect.objectContaining({ details: [expect.objectContaining({ code: "merge-covered-content" })] }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  const candidate = s.prepare(transaction(s, [{ op: "mergeWorkbookCells", sheetId, range: area, coveredContent: "clear" }]));
  expect(s.serialize()).toEqual(before); expect(candidate.model.getCell(sheetId, 255, 32)).toBeUndefined();
  expect(s.commit(candidate)).toMatchObject({ changed: true, affectedSheetIds: [sheetId] });
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 260, 0)).toMatchObject({ state: "valid", value: { type: "number", value: 7 } });
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
  expect(s.snapshot().metadata).toEqual(original.metadata); s.calculateSync(); expect(s.getCalculatedCell(sheetId, 260, 0)).toMatchObject({ value: { type: "number", value: 18 } });
  s.redo(); expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([area]);
});

test("merges retain exact leading rich source and covered blank formats; unmerge permits a later interior edit in the same batch", () => {
  const s = session(), area = range(255, 31, 256, 32);
  apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 255, column: 31, source: { kind: "richText", html: '<span class="cell-default">甲&#13;禰󠄀😀</span>' } },
    { row: 256, column: 32, source: { kind: "value", value: { type: "blank" } }, cssClass: "cell-default" }] }]); s.clearHistory();
  const leading = s.getCell(sheetId, 255, 31)!.html, blank = s.getCell(sheetId, 256, 32)!.html;
  apply(s, [{ op: "mergeWorkbookCells", sheetId, range: area }]);
  expect(s.getCell(sheetId, 255, 31)?.html).toBe(leading); expect(s.getCell(sheetId, 256, 32)?.html).toBe(blank);
  expect(() => apply(s, [value(256, 32, 4)])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  apply(s, [{ op: "unmergeWorkbookCells", sheetId, range: range(256, 32) }, value(256, 32, 4)]);
  expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([]); expect(s.getCell(sheetId, 256, 32)?.value).toEqual({ type: "number", value: 4 });
  s.undo(); expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([area]); expect(s.getCell(sheetId, 256, 32)?.html).toBe(blank);
});

test("empty text, zero, false and formulas with blank caches require an explicit covered-content choice", () => {
  for (const source of [{ kind: "value" as const, value: { type: "text" as const, value: "" } }, { kind: "value" as const, value: { type: "number" as const, value: 0 } },
    { kind: "value" as const, value: { type: "boolean" as const, value: false } }, { kind: "formula" as const, expression: "=UNIMPLEMENTED()" }]) {
    const s = session(); apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 1, source }] }]); s.clearHistory(); const before = s.serialize();
    expect(() => apply(s, [{ op: "mergeWorkbookCells", sheetId, range: range(0, 0, 0, 1) }])).toThrowError(expect.objectContaining({ details: [expect.objectContaining({ code: "merge-covered-content" })] }));
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  }
});

test("merge topology round-trips without new revision, reordered metadata, redo loss or events", () => {
  const a = range(10, 0, 11, 1), b = range(0, 0, 1, 1), s = session(scalarWorkbook([], metadata => { metadata.workbook.sheets[0]!.merges = [a, b]; }));
  apply(s, [value(20, 0, 1)]); s.undo(); const before = s.serialize(), history = s.history(), events: string[] = [];
  s.subscribe(event => events.push(event.type));
  expect(apply(s, [{ op: "unmergeWorkbookCells", sheetId, range: a }, { op: "mergeWorkbookCells", sheetId, range: a }]).changed).toBe(false);
  expect(apply(s, [{ op: "mergeWorkbookCells", sheetId, range: b, coveredContent: "clear" }]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history); expect(events).toEqual([]);
  // Equal covered area does not make two separate merges equal to one merge.
  apply(s, [{ op: "unmergeWorkbookCells", sheetId, range: b }, { op: "mergeWorkbookCells", sheetId, range: range(0, 0, 0, 1) }, { op: "mergeWorkbookCells", sheetId, range: range(1, 0, 1, 1) }]);
  expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toHaveLength(3); expect(s.history().redoSteps).toBe(0);
});

test("partial overlap, table intersection and late invalid batches cannot modify source or history", () => {
  const s = session(scalarWorkbook([], metadata => {
    const sheet = metadata.workbook.sheets[0]!; sheet.merges = [range(0, 0, 1, 1)]; sheet.tables = [{ tableId: "table", name: "Data", range: range(10, 0, 12, 1), header: false, totals: false, totalsPreviouslyShown: false,
      showFirstColumn: false, showLastColumn: false, rowStripes: false, columnStripes: false, columns: [{ columnId: "a", name: "A", totalFunction: "none" }, { columnId: "b", name: "B", totalFunction: "none" }] }];
  })), before = s.serialize();
  for (const area of [range(1, 1, 2, 2), range(10, 0, 10, 1), range(20, 0), range(2, 1, 1, 0)]) {
    expect(() => apply(s, [value(20, 0, 9), { op: "mergeWorkbookCells", sheetId, range: area, coveredContent: "clear" }])).toThrow();
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  }
  apply(s, [{ op: "mergeWorkbookCells", sheetId, range: range(0, 0, 2, 2) }]); expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([range(0, 0, 2, 2)]);
});

test("full-worksheet sparse merges and point unmerges never materialize empty cell blocks", () => {
  const s = session(), whole = range(0, 0, 1_048_575, 16_383);
  apply(s, [{ op: "mergeWorkbookCells", sheetId, range: whole }]); expect(s.snapshot().manifest.chunks).toEqual([]);
  expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([whole]);
  apply(s, [{ op: "unmergeWorkbookCells", sheetId, range: range(1_048_575, 16_383) }]);
  expect(s.snapshot().manifest.chunks).toEqual([]); expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([]);
  expect([...s.serialize()].some(([path]) => path.startsWith("content/") || path.startsWith("sheets/"))).toBe(false);
});

test("explicit clearing obeys stop validation and cannot erase an unsafe intermediate rich input", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 1, value: { type: "number", value: 5 } }], metadata => {
    metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "nonblank", kind: "whole", operator: "greaterThan", formulas: ["=0"], ranges: [range(0, 1)], origin: { row: 0, column: 1 }, allowBlank: false,
      showErrorMessage: true, showInputMessage: false, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" }];
  })), before = s.serialize();
  expect(() => apply(s, [{ op: "mergeWorkbookCells", sheetId, range: range(0, 0, 0, 1), coveredContent: "clear" }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  const other = session(), original = other.serialize();
  expect(() => apply(other, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 1, source: { kind: "richText", html: '<span onclick="bad()">unsafe</span>' } }] },
    { op: "mergeWorkbookCells", sheetId, range: range(0, 0, 0, 1), coveredContent: "clear" }])).toThrow();
  expect(other.serialize()).toEqual(original);
});

test("cancelled clearing and stale merge candidates leave their source and history intact", async () => {
  const s = session(scalarWorkbook(Array.from({ length: 601 }, (_, row) => ({ row, column: 0, value: { type: "number", value: row + 1 } })))), before = s.serialize(), abort = new AbortController();
  await expect(s.prepareAsync(transaction(s, [{ op: "mergeWorkbookCells", sheetId, range: range(0, 0, 600), coveredContent: "clear" }]), undefined, {
    signal: abort.signal, yield: async () => {}, onProgress(value) { if (value.phase === "cells" && value.completed === 256) abort.abort(); },
  })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  const candidate = s.prepare(transaction(s, [{ op: "mergeWorkbookCells", sheetId, range: range(0, 0, 600), coveredContent: "clear" }]));
  apply(s, [value(700, 0, 9)]); expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); s.discard(candidate);
  expect(s.getCell(sheetId, 600, 0)?.value).toEqual({ type: "number", value: 601 }); expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([]);
});
