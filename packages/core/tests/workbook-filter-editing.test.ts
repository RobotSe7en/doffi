import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookAutoFilter, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sessions: WorkbookSession[] = [], sheetId = "sheet-1";
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const range = { from: { row: 0, column: 0 }, to: { row: 3, column: 1 } };
const filter: WorkbookAutoFilter = { range, columns: [{ column: 1, hiddenButton: false, showButton: true, criterion: { kind: "values", values: ["keep"], dates: [], blank: false } }] };
function model(stop = false) {
  return scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "Amount" } }, { row: 0, column: 1, value: { type: "text", value: "Choice" } },
    ...[3, 7, 5].map((value, index) => ({ row: index + 1, column: 0, value: { type: "number" as const, value } })),
    ...["keep", "hide", "keep"].map((value, index) => ({ row: index + 1, column: 1, value: { type: "text" as const, value } })),
    { row: 0, column: 2, value: { type: "number", value: 1 } }, { row: 0, column: 3, value: { type: "number", value: 999 }, formula: "=SUBTOTAL(9,A2:A4)" }], metadata => {
    if (stop) metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "subtotal", kind: "custom", expression: "=SUBTOTAL(9,$A$2:$A$4)>=10", ranges: [{ from: { row: 0, column: 2 }, to: { row: 0, column: 2 } }], origin: { row: 0, column: 2 },
      allowBlank: true, showErrorMessage: true, showInputMessage: false, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" }];
  });
}
function session(input = model()) { const value = createDocumentSession(input); sessions.push(value); return value; }
function apply(session: WorkbookSession, operations: DocumentOperation[]) { return session.commit(session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: session.snapshot().manifest.revision, operations })); }
function subtotal(session: WorkbookSession) { session.calculateSync(); return session.getCalculatedCell(sheetId, 0, 3); }

test("filter edits recalculate current subtotal without hiding or deleting source rows and share undo/redo", () => {
  const s = session(), before = s.snapshot(); expect(subtotal(s)).toMatchObject({ state: "valid", value: { type: "number", value: 15 } });
  const transaction = { protocolVersion: 1 as const, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookAutoFilter" as const, sheetId, filter }] };
  const candidate = s.prepare(transaction); expect(s.snapshot()).toEqual(before); expect(s.commit(candidate)).toMatchObject({ affectedSheetIds: [sheetId] });
  expect(subtotal(s)).toMatchObject({ value: { type: "number", value: 8 } }); expect(s.getCell(sheetId, 2, 0)?.value).toEqual({ type: "number", value: 7 });
  expect(s.snapshot().metadata.workbook.sheets[0]!.rowDimensions).toEqual(before.metadata.workbook.sheets[0]!.rowDimensions);
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(subtotal(s)).toMatchObject({ value: { type: "number", value: 15 } });
  s.redo(); expect(subtotal(s)).toMatchObject({ value: { type: "number", value: 8 } });
  apply(s, [{ op: "setWorkbookAutoFilter", sheetId, filter: null }]); expect(subtotal(s)).toMatchObject({ value: { type: "number", value: 15 } });
});

test("filter visibility changes cannot bypass dependent stop validation", () => {
  const s = session(model(true)), before = s.serialize();
  expect(() => apply(s, [{ op: "setWorkbookAutoFilter", sheetId, filter }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("invalid filter geometry, columns and overlapping merges fail without preceding edits", () => {
  const s = session(), before = s.serialize();
  const badFilters = [{ ...filter, range: { from: range.to, to: range.from } }, { ...filter, columns: [{ ...filter.columns[0]!, column: 2 }] }, { ...filter, columns: [...filter.columns, ...filter.columns] }];
  for (const invalid of badFilters) {
    expect(() => apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 20, column: 0, source: { kind: "value", value: { type: "number", value: 9 } } }] }, { op: "setWorkbookAutoFilter", sheetId, filter: invalid }])).toThrow();
    expect(s.serialize()).toEqual(before);
  }
  expect(() => apply(s, [{ op: "mergeWorkbookCells", sheetId, range: { from: { row: 0, column: 0 }, to: { row: 0, column: 1 } }, coveredContent: "clear" }, { op: "setWorkbookAutoFilter", sheetId, filter }])).toThrow();
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("no-op edits preserve valid noncanonical control JSON bytes and do not create a history step", () => {
  const files = new Map(model().serialize()), encoder = new TextEncoder(), decoder = new TextDecoder();
  const reorder = (value: unknown): unknown => Array.isArray(value) ? value.map(reorder) : value && typeof value === "object" ? Object.fromEntries(Object.entries(value).reverse().map(([key, item]) => [key, reorder(item)])) : value;
  for (const path of ["manifest.json", "metadata.json", "annotations.json"]) files.set(path, encoder.encode(JSON.stringify(reorder(JSON.parse(decoder.decode(files.get(path)!))), null, 2) + "\n"));
  const s = session(openDocument(files, { kind: "workbook" })), before = s.serialize();
  expect(apply(s, [])).toMatchObject({ changed: false, revision: 0 }); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  expect(apply(s, [{ op: "setWorkbookAutoFilter", sheetId, filter }, { op: "setWorkbookAutoFilter", sheetId, filter: null }])).toMatchObject({ changed: false, revision: 0 });
  expect(s.serialize()).toEqual(before);
});
