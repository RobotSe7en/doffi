import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sheetId = "sheet-1", sessions: WorkbookSession[] = [];
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const number = (value: number) => ({ type: "number" as const, value });
function session(model = openDocument(createWorkbook({ documentId: "axis" }), { kind: "workbook" })) { const s = createDocumentSession(model); sessions.push(s); return s; }
afterEach(() => { for (const s of sessions.splice(0)) s.dispose(); });
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations))); }
const axis = (direction: "row" | "column", action: "insert" | "delete", index: number, count = 1): DocumentOperation =>
  ({ op: "editWorkbookAxis", sheetId, axis: direction, action, index, count });
const set = (row: number, column: number, value: number): DocumentOperation => ({ op: "setWorkbookCells", sheetId, cells: [{ row, column, source: { kind: "value", value: number(value) } }] });

test("row insertion relocates exact rich source across chunks and updates inside/outside formulas in one isolated candidate", () => {
  const model = scalarWorkbook([{ row: 255, column: 31, value: number(7) }, { row: 256, column: 32, value: number(8) },
    { row: 600, column: 0, value: number(999), formula: "=SUM(AF256:AG257)" }]);
  const s = session(model);
  apply(s, [{ op: "insertWorkbookSheet", name: "Observer" }, { op: "setWorkbookCells", sheetId, cells: [{ row: 256, column: 30, source: { kind: "richText", html: '<span class="cell-default">中文&#13;禰󠄀😀</span>' } }] }]);
  const other = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
  apply(s, [{ op: "setWorkbookCells", sheetId: other, cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=Sheet1!AG257" } }] }]);
  s.clearHistory(); const before = s.snapshot(), rich = s.getCell(sheetId, 256, 30)!.html;
  const candidate = s.prepare(transaction(s, [axis("row", "insert", 256, 2)]));
  expect(s.snapshot()).toEqual(before); expect(model.getCell(sheetId, 256, 32)?.value).toEqual(number(8));
  expect(candidate.model.getCell(sheetId, 256, 32)).toBeUndefined();
  expect(candidate.model.getCell(sheetId, 258, 30)?.html).toBe(rich);
  expect(candidate.model.getCell(sheetId, 602, 0)).toMatchObject({ value: number(999), formula: { expression: "=SUM(AF256:AG259)", cacheState: "stale" } });
  const receipt = s.commit(candidate);
  expect(receipt.mapping.workbook).toEqual([{ kind: "axis", sheetId, axis: "row", action: "insert", index: 256, count: 2 }]);
  expect(s.getCell(other, 0, 0)?.formula?.expression).toBe("='Sheet1'!AG259");
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 602, 0)).toMatchObject({ state: "valid", value: number(15) });
  expect(s.getCalculatedCell(other, 0, 0)).toMatchObject({ value: number(8) });
  expect(s.history().undoSteps).toBe(1);
  expect(s.undo().mapping.workbook).toEqual([{ kind: "axis", sheetId, axis: "row", action: "delete", index: 256, count: 2 }]);
  expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  s.redo(); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
});

test("column insertion maps absolute and mixed references, changes cell coordinates and preserves blank formats", () => {
  const s = session(scalarWorkbook([{ row: 255, column: 31, value: number(7) }, { row: 0, column: 0, value: number(99), formula: "=$AF$256+AF$256" }]));
  apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 256, column: 32, source: { kind: "value", value: { type: "blank" } }, cssClass: "cell-default" }] }]);
  s.clearHistory(); const before = s.snapshot();
  apply(s, [axis("column", "insert", 31, 2)]);
  expect(s.getCell(sheetId, 255, 31)).toBeUndefined();
  expect(s.getCell(sheetId, 255, 33)).toMatchObject({ value: number(7), html: expect.stringContaining('data-column="33"') });
  expect(s.getCell(sheetId, 256, 34)).toMatchObject({ value: { type: "blank" }, cssClass: "cell-default" });
  expect(s.getCell(sheetId, 0, 0)?.formula?.expression).toBe("=($AH$256+AH$256)");
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 0, 0)).toMatchObject({ value: number(14) });
  s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("deletion shrinks references and invalidates deleted cells without changing address-like strings", () => {
  const s = session(scalarWorkbook([...[1, 2, 3, 4, 5].map((value, row) => ({ row, column: 0, value: number(value) })),
    { row: 6, column: 0, value: number(99), formula: "=SUM(A2:A5)" }, { row: 6, column: 1, value: number(99), formula: '=$A$3&"A3"' },
  ]));
  apply(s, [axis("row", "delete", 2, 2)]);
  expect(s.getCell(sheetId, 2, 0)?.value).toEqual(number(5));
  expect(s.getCell(sheetId, 4, 0)?.formula?.expression).toBe("=SUM(A2:A3)");
  expect(s.getCell(sheetId, 4, 1)?.formula?.expression).toBe('=(#REF!&"A3")');
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 4, 0)).toMatchObject({ value: number(7) });
  expect(s.getCalculatedCell(sheetId, 4, 1)).toMatchObject({ value: { type: "error", value: "#REF!" } });
});

test("dimensions retain their old identities while definitions expand through an interior insertion", () => {
  const s = session(scalarWorkbook([], metadata => {
    const sheet = metadata.workbook.sheets[0]!;
    sheet.rowDimensions = [{ from: 0, to: 4, heightPt: 20, hidden: true }];
    sheet.merges = [range(1, 1, 3, 2)]; sheet.frozen = { rows: 3, columns: 1 };
    sheet.hyperlinks = [{ range: range(1, 4, 3, 4), location: "Sheet1!A4" }];
  }));
  apply(s, [axis("row", "insert", 2, 2)]);
  const sheet = s.snapshot().metadata.workbook.sheets[0]!;
  expect(sheet.rowDimensions).toEqual([{ from: 0, to: 1, heightPt: 20, hidden: true }, { from: 4, to: 6, heightPt: 20, hidden: true }]);
  expect(sheet.merges).toEqual([range(1, 1, 5, 2)]); expect(sheet.frozen).toEqual({ rows: 5, columns: 1 });
  expect(sheet.hyperlinks).toEqual([{ range: range(1, 4, 5, 4), location: "'Sheet1'!A6" }]);
  expect(s.snapshot().manifest.chunks).toEqual([]);
});

test("deleting a rule origin rebases surviving relative templates and removes wholly deleted rules", () => {
  const s = session(scalarWorkbook([1, 2, 3, 4].map((value, row) => ({ row, column: 0, value: number(value) })), metadata => {
    const sheet = metadata.workbook.sheets[0]!;
    sheet.conditionalRules = [{ ruleId: "relative", kind: "expression", expression: "=A1>0", origin: { row: 0, column: 0 }, ranges: [range(0, 0, 3)], priority: 1, stopIfTrue: false },
      { ruleId: "removed", kind: "expression", expression: "=TRUE()", origin: { row: 0, column: 1 }, ranges: [range(0, 1)], priority: 2, stopIfTrue: false }];
    sheet.dataValidations = [{ ruleId: "positive", kind: "custom", expression: "=A1>0", origin: { row: 0, column: 0 }, ranges: [range(0, 0, 3)], allowBlank: false,
      showInputMessage: false, showErrorMessage: true, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" }];
  }));
  apply(s, [axis("row", "delete", 0)]);
  const sheet = s.snapshot().metadata.workbook.sheets[0]!;
  expect(sheet.conditionalRules).toHaveLength(1);
  expect(sheet.conditionalRules?.[0]).toMatchObject({ expression: "=A1>0", origin: { row: 0, column: 0 }, ranges: [range(0, 0, 2)] });
  expect(sheet.dataValidations?.[0]).toMatchObject({ expression: "=A1>0", origin: { row: 0, column: 0 }, ranges: [range(0, 0, 2)] });
  expect(() => apply(s, [set(0, 0, -1)])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
});

test("host annotations follow exact old cells, omit inserted gaps and retain deleted anchor bodies for reanchoring", () => {
  const s = session();
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "a", documentId: "axis", baseRevision: 0, status: "resolved", instruction: "保留正文😀",
    anchor: { kind: "cells", sheetId, ranges: [range(1, 0, 3)] } } }]); s.clearHistory();
  apply(s, [axis("row", "insert", 2, 2)]);
  expect(s.getAnnotations()[0]).toMatchObject({ status: "resolved", anchor: { ranges: [range(1, 0), range(4, 0, 5)] } });
  apply(s, [axis("row", "delete", 0, 6)]);
  expect(s.getAnnotations()[0]).toMatchObject({ status: "needs_reanchor", instruction: "保留正文😀", baseRevision: 2, anchor: { ranges: [range(1, 0), range(4, 0, 5)] } });
  s.undo(); expect(s.getAnnotations()[0]).toMatchObject({ status: "resolved", baseRevision: 4, anchor: { ranges: [range(1, 0), range(4, 0, 5)] } });
});

test("axis bounds and discarded stored formats reject a complete batch without losing bytes or history", () => {
  const s = session(); apply(s, [{ op: "setWorkbookCellFormats", sheetId, cells: [{ row: 1_048_575, column: 0, cssClass: "cell-default" }] }]); s.clearHistory();
  const before = s.serialize();
  for (const operation of [axis("row", "insert", 0), axis("column", "insert", 16_383, 2), axis("row", "delete", 1_048_575, 2)]) {
    expect(() => apply(s, [set(0, 0, 9), operation])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  }
  apply(s, [axis("row", "delete", 1_048_575)]); expect(s.snapshot().manifest.chunks).toEqual([]);
});

test("axis deletion cannot erase an unsafe intermediate rich input before validation", () => {
  const s = session(), before = s.serialize();
  expect(() => apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 1, column: 0, source: { kind: "richText", html: '<span onclick="bad()">unsafe</span>' } }] },
    axis("row", "delete", 1)])).toThrow();
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("cancelled sparse relocation releases its candidate slot and stale axis candidates cannot publish", async () => {
  const s = session(scalarWorkbook(Array.from({ length: 600 }, (_, row) => ({ row, column: 0, value: number(row) })))), before = s.serialize(), abort = new AbortController();
  await expect(s.prepareAsync(transaction(s, [axis("row", "insert", 0)]), undefined, {
    signal: abort.signal, yield: async () => {}, onProgress(value) { if (value.phase === "cells" && value.completed === 256) abort.abort(); },
  })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  const candidate = await s.prepareAsync(transaction(s, [axis("row", "insert", 0)]));
  apply(s, [set(700, 0, 9)]);
  expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); s.discard(candidate);
  expect(s.getCell(sheetId, 599, 0)?.value).toEqual(number(599));
});

test("column deletion remaps absolute filter keys and removes a fully deleted filter", () => {
  const s = session(scalarWorkbook([], metadata => {
    metadata.workbook.sheets[0]!.autoFilter = { range: range(0, 1, 10, 4), columns: [{ column: 2, hiddenButton: false, showButton: true }, { column: 4, hiddenButton: false, showButton: true }],
      sort: { range: range(1, 1, 10, 4), keys: [{ column: 2, descending: false }, { column: 4, descending: true }], caseSensitive: false, method: "none" } };
  }));
  apply(s, [axis("column", "delete", 2)]);
  expect(s.snapshot().metadata.workbook.sheets[0]!.autoFilter).toMatchObject({ range: range(0, 1, 10, 3), columns: [{ column: 3 }],
    sort: { range: range(1, 1, 10, 3), keys: [{ column: 3, descending: true }] } });
  apply(s, [axis("column", "delete", 1, 3)]); expect(s.snapshot().metadata.workbook.sheets[0]!.autoFilter).toBeUndefined();
});

test("qualified absolute defined names follow references", () => {
  const s = session(scalarWorkbook([{ row: 2, column: 0, value: number(7) }, { row: 5, column: 0, value: number(99), formula: "=Rate" }], metadata => {
    metadata.workbook.names = [{ nameId: "rate", name: "Rate", origin: { row: 0, column: 0 }, expression: "=Sheet1!$A$3" }];
  }));
  apply(s, [axis("row", "insert", 1)]);
  expect(s.snapshot().metadata.workbook.names?.[0]?.expression).toBe("='Sheet1'!$A$4");
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 6, 0)).toMatchObject({ value: number(7) });
});

test("empty full-axis operations stay sparse and insert/delete restores cached source and redo", () => {
  const empty = session(), emptySource = empty.serialize();
  expect(apply(empty, [axis("row", "delete", 0, 1_048_576)]).changed).toBe(false);
  expect(empty.serialize()).toEqual(emptySource); expect(empty.history().undoSteps).toBe(0);
  const s = session(scalarWorkbook([{ row: 255, column: 31, value: number(99), formula: "=AF255+1" }]));
  apply(s, [set(0, 0, 9)]); s.undo(); const source = s.serialize(), history = s.history();
  expect(apply(s, [axis("row", "insert", 255, 2), axis("row", "delete", 255, 2)]).changed).toBe(false);
  expect(s.serialize()).toEqual(source); expect(s.history()).toEqual(history);
});
test("an inverse axis batch restores dimension spans and metadata formula spellings without history", () => {
  const s = session(scalarWorkbook([], metadata => {
    const sheet = metadata.workbook.sheets[0]!;
    sheet.rowDimensions = [{ from: 0, to: 10, heightPt: 20, hidden: true }];
    sheet.conditionalRules = [{ ruleId: "rule", kind: "expression", expression: "=A1>0", origin: { row: 0, column: 0 }, ranges: [range(0, 0, 5)], priority: 1, stopIfTrue: false }];
    metadata.workbook.names = [{ nameId: "name", name: "Rate", expression: "=Sheet1!$A$4+1", origin: { row: 0, column: 0 } }];
  })), before = s.serialize();
  expect(apply(s, [axis("row", "insert", 2, 2), axis("row", "delete", 2, 2)]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
test("an inverse axis batch restores exact annotation grouping without filling inserted gaps", () => {
  const s = session();
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "a", documentId: "axis", baseRevision: 0, status: "open", instruction: "保留分组",
    anchor: { kind: "cells", sheetId, ranges: [range(0, 0, 5, 2)] } } }]); s.clearHistory();
  const before = s.serialize();
  expect(apply(s, [axis("row", "insert", 2, 2), axis("row", "delete", 2, 2)]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
test("explicit before/after formatting inherits only cell formats, dimensions and rules, with a bounded inserted cell count", () => {
  for (const formatSource of ["none", "before", "after"] as const) {
    const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(7) }, { row: 1, column: 0, value: number(8) }], metadata => {
      const sheet = metadata.workbook.sheets[0]!;
      sheet.rowDimensions = [{ from: 0, to: 0, heightPt: 20 }, { from: 1, to: 1, heightPt: 30 }];
      sheet.conditionalRules = [{ ruleId: "before", kind: "expression", expression: "=A1>0", ranges: [range(0, 0)], origin: { row: 0, column: 0 }, priority: 1, stopIfTrue: false }];
    }));
    apply(s, [{ op: "setWorkbookCellFormats", sheetId, cells: [{ row: 0, column: 0, cssClass: "cell-default" }] }]);
    s.clearHistory();
    apply(s, [{ op: "editWorkbookAxis", sheetId, axis: "row", action: "insert", index: 1, count: 2, formatSource }]);
    expect(s.getCell(sheetId, 0, 0)?.value).toEqual(number(7)); expect(s.getCell(sheetId, 3, 0)?.value).toEqual(number(8));
    expect(s.getCell(sheetId, 1, 0)?.value).toEqual(formatSource === "before" ? { type: "blank" } : undefined);
    expect(s.snapshot().metadata.workbook.sheets[0]!.conditionalRules?.[0]?.ranges).toEqual([range(0, 0, formatSource === "before" ? 2 : 0)]);
    expect(s.snapshot().metadata.workbook.sheets[0]!.rowDimensions.find(span => span.from === 1)?.heightPt).toBe(formatSource === "before" ? 20 : formatSource === "after" ? 30 : undefined);
    expect(() => apply(s, [{ op: "editWorkbookAxis", sheetId, axis: "row", action: "delete", index: 1, count: 1, formatSource }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  }
  const source = scalarWorkbook([{ row: 0, column: 0, value: number(7) }]), s = createDocumentSession(source, { workbookLimits: { maxCells: 3 } }); sessions.push(s);
  apply(s, [{ op: "setWorkbookCellFormats", sheetId, cells: [{ row: 0, column: 0, cssClass: "cell-default" }] }]); s.clearHistory();
  const before = s.serialize();
  expect(() => apply(s, [{ op: "editWorkbookAxis", sheetId, axis: "row", action: "insert", index: 1, count: 3, formatSource: "before" }])).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
