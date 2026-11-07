import { expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const number = (value: number) => ({ type: "number" as const, value });
function prepare(session: WorkbookSession, operations: DocumentOperation[]) {
  return session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: session.snapshot().manifest.revision, operations });
}
function commit(session: WorkbookSession, operations: DocumentOperation[]) { return session.commit(prepare(session, operations)); }
test("worksheet insertion, order and visibility share isolated candidates and undo", () => {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: "sheets" }), { kind: "workbook" }));
  const candidate = prepare(session, [{ op: "insertWorkbookSheet", name: "中文", beforeSheetId: "sheet-1" }]);
  expect(session.snapshot().metadata.workbook.sheets).toHaveLength(1);
  expect(candidate.preview.metadata.workbook.sheets.map(sheet => sheet.name)).toEqual(["中文", "Sheet1"]);
  session.commit(candidate); const added = session.snapshot().metadata.workbook.sheets[0]!.sheetId;
  commit(session, [{ op: "moveWorkbookSheet", sheetId: added }, { op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { visibility: "hidden" } }]);
  expect(session.snapshot().metadata.workbook.sheets.map(sheet => sheet.name)).toEqual(["Sheet1", "中文"]);
  expect(session.snapshot().metadata.workbook.activeSheetId).toBe(added);
  session.undo(); expect(session.snapshot().metadata.workbook.sheets[0]!.sheetId).toBe(added); expect(session.snapshot().metadata.workbook.activeSheetId).toBe("sheet-1");
  session.redo(); expect(openDocument(session.serialize(), { kind: "workbook" }).snapshot()).toEqual(session.snapshot());
});
test("rename rewrites staged formulas, names, rules and links while preserving string literals", () => {
  const session = createDocumentSession(scalarWorkbook([
    { row: 0, column: 0, value: number(5) }, { row: 0, column: 1, value: number(999), formula: "=Sheet1!A1" },
  ], metadata => {
    metadata.workbook.names = [{ nameId: "global", name: "Global", expression: "=Sheet1!$A$1", origin: { row: 0, column: 0 } }, { nameId: "local", name: "Local", sheetId: "sheet-1", expression: "=$A$1", origin: { row: 0, column: 0 } }];
    metadata.workbook.sheets[0]!.conditionalRules = [{ ruleId: "rule", kind: "expression", expression: "=Sheet1!A1>0", origin: { row: 0, column: 0 }, ranges: [{ from: { row: 0, column: 0 }, to: { row: 5, column: 0 } }], priority: 1, stopIfTrue: false }];
    metadata.workbook.sheets[0]!.hyperlinks = [{ range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }, location: "Sheet1!A1" }];
  }));
  commit(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 300, column: 0, source: { kind: "formula", expression: '=FUTUREFUNC(Sheet1!A1,"Sheet1!A1")' } }] },
    { op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { name: "O'Brien 中文" } }]);
  expect(session.getCell("sheet-1", 0, 1)?.formula?.expression).toBe("='O''Brien 中文'!A1");
  expect(session.getCell("sheet-1", 300, 0)?.formula?.expression).toBe('=FUTUREFUNC(\'O\'\'Brien 中文\'!A1,"Sheet1!A1")');
  const workbook = session.snapshot().metadata.workbook;
  expect(workbook.names?.map(name => name.expression)).toEqual(["='O''Brien 中文'!$A$1", "=$A$1"]);
  expect(workbook.sheets[0]!.conditionalRules?.[0]).toMatchObject({ expression: "=('O''Brien 中文'!A1>0)" });
  expect(workbook.sheets[0]!.hyperlinks?.[0]?.location).toBe("'O''Brien 中文'!A1");
  session.calculateSync(); expect(session.getCalculatedCell("sheet-1", 0, 1)).toMatchObject({ value: number(5) });
  session.undo(); expect(session.getCell("sheet-1", 0, 1)?.formula?.expression).toBe("=Sheet1!A1");
});
test("deleting a worksheet marks surviving formulas and anchors, and undo restores exact sources", () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(7) }]));
  commit(session, [{ op: "insertWorkbookSheet", name: "Other" }]);
  const other = session.snapshot().metadata.workbook.sheets[1]!.sheetId;
  commit(session, [{ op: "setWorkbookCells", sheetId: other, cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=Sheet1!A1+1" } }] },
    { op: "addAnnotation", annotation: { annotationId: "a", documentId: "scalar-kernel", baseRevision: 1, status: "resolved", instruction: "保留正文", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 2, column: 2 } }] } } }]);
  const before = session.getCell("sheet-1", 0, 0)!.html;
  const receipt = commit(session, [{ op: "removeWorkbookSheet", sheetId: "sheet-1" }]);
  expect(receipt.mapping.workbook).toEqual([{ kind: "deleteSheet", sheetId: "sheet-1" }]);
  expect(session.getCell(other, 0, 0)?.formula?.expression).toBe("=(#REF!+1)");
  expect(session.getAnnotations()[0]).toMatchObject({ status: "needs_reanchor", baseRevision: 2, instruction: "保留正文" });
  expect(session.undo().mapping.workbook).toEqual([{ kind: "restoreSheet", sheetId: "sheet-1" }]);
  expect(session.getCell("sheet-1", 0, 0)?.html).toBe(before); expect(session.getAnnotations()[0]).toMatchObject({ status: "resolved", baseRevision: 4 });
});
test("invalid final visibility, duplicate names, removed targets and unknown siblings fail atomically", () => {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: "guards" }), { kind: "workbook" })), before = session.serialize();
  for (const operations of [
    [{ op: "insertWorkbookSheet", name: "sheet1" }],
    [{ op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { visibility: "hidden" } }],
    [{ op: "removeWorkbookSheet", sheetId: "sheet-1" }],
    [{ op: "insertWorkbookSheet", name: "New", beforeSheetId: "foreign" }],
  ] satisfies DocumentOperation[][]) { expect(() => prepare(session, operations)).toThrow(); expect(session.serialize()).toEqual(before); }
  commit(session, [{ op: "insertWorkbookSheet", name: "Other" }]); const saved = session.serialize();
  expect(() => prepare(session, [{ op: "removeWorkbookSheet", sheetId: "sheet-1" }, { op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { name: "Gone" } }])).toThrow();
  expect(session.serialize()).toEqual(saved);
});
test("intermediate worksheet name collisions reject atomically instead of retargeting formulas during a swap", () => {
  const session = createDocumentSession(scalarWorkbook([
    { row: 0, column: 0, value: number(999), formula: "=Sheet1!C1" }, { row: 0, column: 1, value: number(999), formula: "=Other!C1" },
  ], metadata => { metadata.workbook.sheets.push({ ...metadata.workbook.sheets[0]!, sheetId: "second", name: "Other" }); }));
  const before = session.serialize();
  for (const operations of [
    [{ op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { name: "other" } }, { op: "updateWorkbookSheet", sheetId: "second", changes: { name: "Sheet1" } }],
    [{ op: "insertWorkbookSheet", name: "sheet1" }, { op: "removeWorkbookSheet", sheetId: "sheet-1" }],
  ] satisfies DocumentOperation[][]) {
    expect(() => prepare(session, operations)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
    expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  }
  commit(session, [{ op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { name: "Temporary" } },
    { op: "updateWorkbookSheet", sheetId: "second", changes: { name: "Sheet1" } }, { op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { name: "Other" } }]);
  for (const [column, sheetId] of [[0, "sheet-1"], [1, "second"]] as const) {
    const reference = session.resolveReference(session.getCell("sheet-1", 0, column)!.formula!.expression, { sheetId: "sheet-1", row: 0, column });
    expect(reference).toMatchObject({ kind: "ranges", areas: [{ sheetId }] });
  }
  expect(session.history().undoSteps).toBe(1); session.undo();
  expect(session.getCell("sheet-1", 0, 0)?.formula?.expression).toBe("=Sheet1!C1");
  expect(session.getCell("sheet-1", 0, 1)?.formula?.expression).toBe("=Other!C1");
});
