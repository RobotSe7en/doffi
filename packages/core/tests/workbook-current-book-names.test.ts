import { afterEach, expect, test } from "vitest";
import { copyWorkbookFormula, createDocumentSession, openDocument, parseWorkbookFormula } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const number = (value: number) => ({ type: "number" as const, value }), at = { row: 0, column: 0 };
const sessions: WorkbookSession[] = [];
afterEach(() => { for (const s of sessions.splice(0)) s.dispose(); });
function apply(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations })); }
function workbook() {
  const model = scalarWorkbook([{ ...at, value: number(7) }, { row: 1, column: 0, value: number(11) }, { row: 0, column: 1, value: number(3) },
    { row: 0, column: 2, value: number(999), formula: "=[0]!Shared+Shared" }, { row: 1, column: 2, value: number(999), formula: "=Alias+[0]!Shared" },
    { row: 0, column: 3, value: number(999), formula: "=SUM([0]!A:[0]!B)" }, { row: 0, column: 4, value: number(999), formula: "=[0]Sheet1!$A$1+'[0]Other'!$B$1" },
  ], metadata => {
    const source = metadata.workbook.sheets[0]!; metadata.workbook.sheets.push({ ...structuredClone(source), sheetId: "other", name: "Other" });
    metadata.workbook.names = [{ nameId: "global", name: "Shared", expression: "=Sheet1!$A$1", origin: at },
      { nameId: "source-local", name: "Shared", sheetId: source.sheetId, expression: "=Sheet1!$B$1", origin: at },
      { nameId: "other-local", name: "Shared", sheetId: "other", expression: "=Other!$B$1", origin: at },
      { nameId: "alias", name: "Alias", expression: "=[0]!Shared", origin: at },
      { nameId: "first", name: "A", expression: "=Sheet1!$A$1", origin: at }, { nameId: "last", name: "B", expression: "=Sheet1!$A$2", origin: at }];
  });
  const s = createDocumentSession(model); sessions.push(s);
  apply(s, [{ op: "setWorkbookCells", sheetId: "other", cells: [{ row: 0, column: 1, source: { kind: "value", value: number(100) } },
    { row: 0, column: 2, source: { kind: "formula", expression: "=[0]!Shared+Shared" } }] }]); s.clearHistory(); return s;
}

test("current-workbook names resolve independently of local shadows in scalar, reference and name-parameter ranges", () => {
  const s = workbook(); s.calculateSync();
  for (const [sheetId, row, column, value] of [["sheet-1", 0, 2, 10], ["sheet-1", 1, 2, 14], ["sheet-1", 0, 3, 18], ["sheet-1", 0, 4, 107], ["other", 0, 2, 107]] as const)
    expect(s.getCalculatedCell(sheetId, row, column)).toMatchObject({ state: "valid", value: number(value) });
  expect(s.resolveReference("=[0]!Shared", { sheetId: "other", ...at })).toEqual({ kind: "ranges", areas: [{ sheetId: "sheet-1", from: at, to: at }] });
  expect(parseWorkbookFormula("=[000]!Shared", at)).toEqual({ kind: "name", name: "Shared", workbook: 0 });
  expect(copyWorkbookFormula("=[0]!Shared+A1", at, { row: 1, column: 1 })).toBe("=([0]!Shared+B2)");
});

test("global-name dependencies invalidate every real consumer while unrelated local values retain their identity", () => {
  const s = workbook(); s.calculateSync(); const before = s.snapshot();
  apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ ...at, source: { kind: "value", value: number(9) } }] }]);
  expect(s.getCalculatedCell("other", 0, 2).state).toBe("stale"); s.calculateSync();
  expect(s.getCalculatedCell("other", 0, 2)).toMatchObject({ value: number(109) }); expect(s.getCalculatedCell("sheet-1", 1, 2)).toMatchObject({ value: number(18) });
  expect(s.getCell("other", 0, 1)!.value).toEqual(number(100)); expect(s.getCell("sheet-1", 0, 2)!.value).toEqual(number(999));
  s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("renaming a global name rewrites only its global bindings, including explicit aliases, with one undo", () => {
  const s = workbook(), before = s.snapshot(), name = before.metadata.workbook.names!.find(name => name.nameId === "global")!;
  apply(s, [{ op: "setWorkbookDefinedName", name: { ...name, origin: { ...name.origin }, name: "Renamed" } }]); s.calculateSync();
  expect(s.getCell("sheet-1", 0, 2)!.formula!.expression).toBe("=([0]!Renamed+Shared)");
  expect(s.snapshot().metadata.workbook.names!.find(name => name.nameId === "alias")!.expression).toBe("=[0]!Renamed");
  expect(s.getCalculatedCell("other", 0, 2)).toMatchObject({ value: number(107) }); expect(s.history().undoSteps).toBe(1);
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("request capture includes the actual global definition and its input without exposing shadowed local inputs", () => {
  const s = workbook(), manifest = s.snapshot().manifest;
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "request", documentId: manifest.documentId, baseRevision: manifest.revision, status: "open", instruction: "复核全局名称", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 1, column: 2 }, to: { row: 1, column: 2 } }] } } }]);
  const request = s.buildInteractionRequest({ kind: "workbook", annotationIds: ["request"] });
  expect(request.context.metadata.workbook.names!.map(name => name.nameId).sort()).toEqual(["alias", "global"]);
  expect(request.context.cells.find(cell => cell.row === 1 && cell.column === 2)!.current).toEqual({ kind: "value", value: number(14) });
  expect(request.context.cells.some(cell => cell.column === 1)).toBe(false);
  expect(openDocument(s.serialize(), { kind: "workbook" }).getCell("sheet-1", 1, 2)!.formula!.expression).toBe("=Alias+[0]!Shared");
});

test.each(["=[1]!Shared", "=[2]Sheet1!A1", "='[1]Sheet1'!A1", "=[0]!A1", "=[0]!1", "=[0]!"])("external or invalid workbook-qualified source %s fails before publishing any part", expression => {
  const s = workbook(), before = s.serialize();
  expect(() => apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 9, column: 9, source: { kind: "formula", expression } }] }])).toThrow();
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
