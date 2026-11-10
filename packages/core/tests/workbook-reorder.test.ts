import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sheetId = "sheet-1", sessions: WorkbookSession[] = [];
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const number = (value: number) => ({ type: "number" as const, value });
function session(model = scalarWorkbook([])) { const s = createDocumentSession(model); sessions.push(s); return s; }
afterEach(() => { for (const s of sessions.splice(0)) s.dispose(); });
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations))); }
const reorder = (rows: number[], area = range(0, 0, rows.length - 1, 1)): DocumentOperation => ({ op: "reorderWorkbookRows", sheetId, range: area, rows });

test("record reordering crosses chunks, copies inside formulas and keeps outside addresses and cached rich source in one history step", () => {
  const s = session(scalarWorkbook([{ row: 255, column: 31, value: number(3) }, { row: 255, column: 32, value: number(999), formula: "=AF256+$AF$257" },
    { row: 256, column: 31, value: number(7) }, { row: 256, column: 32, value: number(999), formula: "=AF257+$AF$257" },
    { row: 600, column: 0, value: number(999), formula: "=AF256" }]));
  apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 255, column: 33, source: { kind: "richText", html: '<span class="cell-default">中文&#13;禰󠄀😀</span>' } }] }]); s.clearHistory();
  const before = s.snapshot(), rich = s.getCell(sheetId, 255, 33)!.html, originalFormula = s.getCell(sheetId, 256, 32)!.html;
  const candidate = s.prepare(transaction(s, [reorder([256, 255], range(255, 31, 256, 33))])); expect(s.snapshot()).toEqual(before);
  const receipt = s.commit(candidate); expect(receipt.mapping.workbook).toEqual([{ kind: "sort", sheetId, range: range(255, 31, 256, 33), rows: [256, 255] }]);
  expect(s.getCell(sheetId, 256, 33)?.html).toBe(rich);
  // An identical final formula/cache at its old address keeps its exact source;
  // changed precedents still invalidate its independent current calculation.
  expect(s.getCell(sheetId, 256, 32)?.html).toBe(originalFormula);
  expect(s.getCell(sheetId, 600, 0)?.formula?.expression).toBe("=AF256");
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 600, 0)).toMatchObject({ value: number(7) });
  expect(s.getCalculatedCell(sheetId, 256, 32)).toMatchObject({ value: number(6) });
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
  s.redo(); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
});

test("annotations and link fragments follow records while relative rule templates retain their meaning", () => {
  const s = session(scalarWorkbook([1, 2, 3, 4].map((value, row) => ({ row, column: 0, value: number(value) })), metadata => {
    const sheet = metadata.workbook.sheets[0]!; sheet.hyperlinks = [{ range: range(0, 0, 1), location: "Sheet1!A1" }];
    sheet.conditionalRules = [{ ruleId: "positive", kind: "expression", expression: "=A1>0", origin: { row: 0, column: 0 }, ranges: [range(0, 0, 3)], priority: 1, stopIfTrue: false }];
  }));
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "review", documentId: "scalar-kernel", baseRevision: 0, status: "resolved", instruction: "前两条",
    anchor: { kind: "cells", sheetId, ranges: [range(0, 0, 1)] } } }]); s.clearHistory(); const before = s.snapshot();
  apply(s, [reorder([3, 1, 0, 2])]);
  expect(s.getAnnotations()[0]).toMatchObject({ status: "resolved", anchor: { ranges: [range(1, 0), range(3, 0)] } });
  const sheet = s.snapshot().metadata.workbook.sheets[0]!;
  expect(sheet.hyperlinks?.map(link => link.range)).toEqual([range(1, 0), range(3, 0)]); expect(sheet.hyperlinks?.[0]?.location).toBe("Sheet1!A1");
  expect(sheet.conditionalRules?.[0]).toMatchObject({ expression: "=(A4>0)", origin: { row: 3, column: 0 }, ranges: [range(0, 0, 3)] });
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.getAnnotations()[0]?.anchor).toEqual(before.annotations[0]?.anchor);
});

test("illegal permutations and incompatible/partial merges reject atomically; matching horizontal merge records move", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(3) }, { row: 1, column: 0, value: number(7) }], metadata => {
    metadata.workbook.sheets[0]!.merges = [range(0, 0, 0, 1), range(1, 0, 1, 1)];
  })); const before = s.serialize();
  for (const operation of [reorder([1, 1]), reorder([1]), reorder([1, 0], range(0, 0, 1, 0))]) {
    expect(() => apply(s, [operation])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before);
  }
  apply(s, [reorder([1, 0])]); expect(s.getCell(sheetId, 0, 0)?.value).toEqual(number(7));
  expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([range(0, 0, 0, 1), range(1, 0, 1, 1)]);
  expect(s.history().undoSteps).toBe(1);
});

test("uniform multi-row merged records move intact with formulas and outside addresses, and reject split permutations", () => {
  const s = session(scalarWorkbook([
    { row: 0, column: 0, value: number(3) }, { row: 0, column: 1, value: number(999), formula: "=A1*2" },
    { row: 2, column: 0, value: number(1) }, { row: 2, column: 1, value: number(999), formula: "=A3*2" },
    { row: 0, column: 3, value: number(999), formula: "=A1" },
  ], metadata => { metadata.workbook.sheets[0]!.merges = [range(0, 0, 1, 0), range(0, 1, 1, 1), range(2, 0, 3, 0), range(2, 1, 3, 1)]; }));
  const before = s.snapshot(), bytes = s.serialize();
  for (const rows of [[3, 2, 0, 1], [2, 0, 3, 1]]) {
    expect(() => apply(s, [reorder(rows)])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(bytes);
  }
  apply(s, [reorder([2, 3, 0, 1])]); s.calculateSync();
  expect(s.getCell(sheetId, 0, 0)?.value).toEqual(number(1)); expect(s.getCell(sheetId, 2, 0)?.value).toEqual(number(3));
  expect(s.getCell(sheetId, 0, 1)?.formula?.expression).toBe("=A1*2"); expect(s.getCell(sheetId, 0, 3)?.formula?.expression).toBe("=A1");
  expect(s.getCalculatedCell(sheetId, 0, 1)).toMatchObject({ value: number(2) }); expect(s.getCalculatedCell(sheetId, 0, 3)).toMatchObject({ value: number(1) });
  expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual(before.metadata.workbook.sheets[0]!.merges);
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); expect(s.history().undoSteps).toBe(1);
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("Office's incompatible merged and unmerged columns reject the whole sort", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(3) }, { row: 1, column: 0, value: number(7) }], metadata => {
    metadata.workbook.sheets[0]!.merges = [range(0, 0, 0, 1), range(1, 0, 1, 1)];
  })), before = s.serialize();
  expect(() => apply(s, [reorder([1, 0], range(0, 0, 1, 2))])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
test("multi-row records straddling both chunk axes preserve their formulas, coverage and exact source on undo", () => {
  const s = session(scalarWorkbook([
    { row: 255, column: 31, value: number(3) }, { row: 255, column: 32, value: number(999), formula: "=AF256*2" },
    { row: 257, column: 31, value: number(1) }, { row: 257, column: 32, value: number(999), formula: "=AF258*2" },
    { row: 0, column: 0, value: number(999), formula: "=AF256" },
  ], metadata => { metadata.workbook.sheets[0]!.merges = [255, 257].flatMap(row => [31, 32].map(column => range(row, column, row + 1, column))); }));
  const before = s.snapshot();
  apply(s, [reorder([257, 258, 255, 256], range(255, 31, 258, 32))]); s.calculateSync();
  expect(s.getCalculatedCell(sheetId, 255, 32)).toMatchObject({ state: "valid", value: number(2) });
  expect(s.getCalculatedCell(sheetId, 257, 32)).toMatchObject({ state: "valid", value: number(6) });
  expect(s.getCell(sheetId, 0, 0)?.formula?.expression).toBe("=AF256"); expect(s.getCalculatedCell(sheetId, 0, 0)).toMatchObject({ value: number(1) });
  expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual(before.metadata.workbook.sheets[0]!.merges);
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("table sorting keeps header/total positions and only permits complete body records", () => {
  const s = session(scalarWorkbook([{ row: 1, column: 0, value: number(3) }, { row: 2, column: 0, value: number(7) }]));
  apply(s, [{ op: "setWorkbookTable", sheetId, table: { tableId: "t", name: "Data", range: range(0, 0, 3, 1), header: true, totals: true, totalsPreviouslyShown: true,
    columns: [{ columnId: "v", name: "Value", totalFunction: "sum" }, { columnId: "d", name: "Double", totalFunction: "none", calculatedFormula: "=A2*2" }],
    showFirstColumn: false, showLastColumn: false, rowStripes: true, columnStripes: false } }]); s.clearHistory(); const before = s.serialize();
  for (const operation of [reorder([3, 2, 1, 0]), reorder([2, 1], range(1, 0, 2, 0))]) {
    expect(() => apply(s, [operation])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before);
  }
  apply(s, [reorder([2, 1], range(1, 0, 2, 1))]); s.calculateSync();
  expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "text", value: "Value" });
  expect(s.getCalculatedCell(sheetId, 1, 1)).toMatchObject({ value: number(14) }); expect(s.getCalculatedCell(sheetId, 3, 0)).toMatchObject({ value: number(10) });
  expect(s.history().undoSteps).toBe(1);
});

test("implicit row formats travel without changing dimensions, identity permutations preserve redo and cancellation leaves no partial relocation", async () => {
  const s = session(scalarWorkbook([{ row: 1, column: 0, value: number(3) }], metadata => {
    metadata.workbook.sheets[0]!.rowDimensions = [{ from: 0, to: 0, numberFormatId: "general" }];
  }));
  apply(s, [reorder([1, 0])]); expect(s.getCell(sheetId, 1, 0)?.value).toEqual({ type: "blank" }); expect(s.getCell(sheetId, 1, 0)?.formatId).toBe("general");
  s.undo(); const before = s.serialize(), history = s.history(); expect(apply(s, [reorder([0, 1])]).changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  const abort = new AbortController();
  await expect(s.prepareAsync(transaction(s, [reorder(Array.from({ length: 600 }, (_, row) => 599 - row))]), undefined,
    { signal: abort.signal, yield: async () => {}, onProgress(value) { if (value.phase === "cells" && value.completed === 256) abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
});
