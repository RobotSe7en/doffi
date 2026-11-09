import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument, parseWorkbookFormula } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sessions: WorkbookSession[] = [], at = { row: 0, column: 0 }, number = (value: number) => ({ type: "number" as const, value });
afterEach(() => { for (const s of sessions.splice(0)) s.dispose(); });
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations))); }
const edit = (axis: "row" | "column", action: "insert" | "delete", index: number, sheetId = "sheet-1"): DocumentOperation => ({ op: "editWorkbookAxis", sheetId, axis, action, index, count: 1 });
function workbook() {
  const model = scalarWorkbook(Array.from({ length: 20 }, (_, row) => ({ row, column: 0, value: number(11 + row) })), metadata => {
    const sheet = metadata.workbook.sheets[0]!; sheet.name = "Data";
    metadata.workbook.sheets.push({ ...structuredClone(sheet), sheetId: "other", name: "Other" });
    metadata.workbook.names = [
      { nameId: "relative", name: "Relative", origin: at, expression: "=Data!$A1" },
      { nameId: "absolute", name: "Absolute", origin: at, expression: "=Data!$A$1" },
      { nameId: "sum", name: "WindowSum", origin: at, expression: "=SUM(Data!$A1:$A3)" },
      { nameId: "alias", name: "Alias", origin: at, expression: "=Relative+10" },
      { nameId: "local", name: "Local", sheetId: "sheet-1", origin: at, expression: "=!$A$1" },
      { nameId: "dynamic", name: "Dynamic", origin: at, expression: "=!$A$1" },
      { nameId: "bare", name: "Bare", origin: at, expression: "=$A$1" },
    ];
  });
  const s = createDocumentSession(model); sessions.push(s);
  const formulas = ["=Relative", "=Absolute", "=WindowSum", "=Alias", "=Dynamic", "=Bare", "=Data!Local"];
  for (const sheetId of ["sheet-1", "other"]) apply(s, [{ op: "setWorkbookCells", sheetId, cells: [
    ...(sheetId === "other" ? Array.from({ length: 20 }, (_, row) => ({ row, column: 0, source: { kind: "value" as const, value: number(101 + row) } })) : []),
    ...formulas.map((expression, index) => ({ row: 4, column: 2 + index, source: { kind: "formula" as const, expression } })),
  ] }]);
  s.calculateSync(); s.clearHistory(); return s;
}
function values(s: WorkbookSession, sheetId: string, row: number) {
  s.calculateSync(); return Array.from({ length: 7 }, (_, index) => s.getCalculatedCell(sheetId, row, 2 + index)).map(result => result.state === "valid" ? result.value : result.state);
}

test.each([
  { action: "insert" as const, index: 0, dataRow: 5, data: [15, 11, 48, 25, 0, 0, 11], other: [14, 11, 45, 24, 101, 101, 11] },
  { action: "insert" as const, index: 6, dataRow: 4, data: [15, 11, 31, 25, 11, 11, 11], other: [15, 11, 31, 25, 101, 101, 11] },
])("$action row $index preserves relative recipes, maps absolute bindings and retains dynamic sheet context", ({ action, index, dataRow, data, other }) => {
  const s = workbook(), before = s.snapshot(), origins = before.metadata.workbook.names!.map(name => name.origin);
  expect(values(s, "sheet-1", 4)).toEqual([15, 11, 48, 25, 11, 11, 11].map(number));
  const candidate = s.prepare(transaction(s, [edit("row", action, index)]));
  expect(s.snapshot()).toEqual(before); s.commit(candidate);
  expect(values(s, "sheet-1", dataRow)).toEqual(data.map(number));
  expect(values(s, "other", 4)).toEqual(other.map(number));
  expect(s.snapshot().metadata.workbook.names!.map(name => name.origin)).toEqual(origins);
  expect(s.snapshot().metadata.workbook.names!.filter(name => ["relative", "dynamic", "bare", "sum"].includes(name.nameId)).map(name => parseWorkbookFormula(name.expression, name.origin)))
    .toEqual(before.metadata.workbook.names!.filter(name => ["relative", "dynamic", "bare", "sum"].includes(name.nameId)).map(name => parseWorkbookFormula(name.expression, name.origin)));
  expect(s.history().undoSteps).toBe(1); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  s.redo(); expect(values(s, "other", 4)).toEqual(other.map(number));
});

test("deletion invalidates bound absolute components without discarding relative and unbound names", () => {
  const s = workbook(); apply(s, [edit("row", "delete", 0)]);
  const error = { type: "error", value: "#REF!" };
  expect(values(s, "sheet-1", 3)).toEqual([number(15), error, number(48), number(25), number(12), number(12), error]);
  expect(values(s, "other", 4)).toEqual([number(16), error, number(51), number(26), number(101), number(101), error]);
  expect(s.snapshot().metadata.workbook.names!.find(name => name.nameId === "relative")!.expression).toBe("=Data!$A1");
});

test("mixed cells and whole-axis name recipes use the stable definition origin in both directions", () => {
  const s = workbook(), origin = { row: 9, column: 5 }, expressions = ["=Data!B$2", "=Data!$B2", "=Data!$B$2", "=Data!B2", "=Data!$B:$D", "=Data!B:D", "=Data!$2:4", "=!$B$2"];
  apply(s, expressions.map((expression, index) => ({ op: "setWorkbookDefinedName", name: { nameId: "mixed-" + index, name: "Mixed_" + index, origin, expression, sheetId: "sheet-1" } })));
  const before = s.snapshot(); s.clearHistory();
  apply(s, [edit("column", "insert", 1), edit("row", "insert", 1)]);
  const expected = ["=Data!B$3", "=Data!$C2", "=Data!$C$3", "=Data!B2", "=Data!$C:$E", "=Data!B:D", "=Data!$3:4", "=!$C$3"];
  const names = s.snapshot().metadata.workbook.names!.filter(name => name.nameId.startsWith("mixed-"));
  expect(names.map(name => parseWorkbookFormula(name.expression, origin))).toEqual(expected.map(expression => parseWorkbookFormula(expression, origin)));
  expect(names.every(name => name.origin.row === origin.row && name.origin.column === origin.column)).toBe(true);
  apply(s, [edit("row", "delete", 1), edit("column", "delete", 1)]);
  expect(s.snapshot().metadata.workbook.names!.map(name => parseWorkbookFormula(name.expression, name.origin)))
    .toEqual(before.metadata.workbook.names!.map(name => parseWorkbookFormula(name.expression, name.origin)));
  s.undo(); s.undo();
  expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("a deleted fixed endpoint clips to surviving mixed ranges while relative endpoints retain their recipes", () => {
  const expressions = ["=Data!$A$1:$A3", "=Data!$A3:$A$1", "=Data!$1:3", "=Data!3:$1"];
  const s = workbook();
  apply(s, expressions.map((expression, index) => ({ op: "setWorkbookDefinedName", name: { nameId: "clip-" + index, name: "Clip_" + index, origin: at, expression } })));
  apply(s, [edit("row", "delete", 0)]);
  for (const [index, expression] of expressions.entries()) {
    const name = s.snapshot().metadata.workbook.names!.find(name => name.nameId === "clip-" + index)!;
    expect(parseWorkbookFormula(name.expression, at)).toEqual(parseWorkbookFormula(expression, at));
    expect(s.resolveReference("=" + name.name, { sheetId: "sheet-1", ...at })).toMatchObject({ kind: "ranges", areas: [{ sheetId: "sheet-1", from: { row: 0, column: 0 }, to: { row: 2 } }] });
  }
});

test("inverse edits are a no-op and a later invalid operation cannot publish rewritten names", () => {
  const s = workbook(), before = s.serialize();
  expect(apply(s, [edit("row", "insert", 0), edit("row", "delete", 0)]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  expect(() => apply(s, [edit("row", "insert", 0), { op: "setWorkbookCells", sheetId: "missing", cells: [{ ...at, source: { kind: "value", value: number(1) } }] }])).toThrow();
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("large contextual name batches yield, cancel privately and release the preparation slot", async () => {
  const s = createDocumentSession(scalarWorkbook([], metadata => {
    metadata.workbook.names = Array.from({ length: 600 }, (_, index) => ({ nameId: "name-" + index, name: "Name_" + index,
      origin: at, expression: "=Sheet1!$A1+Sheet1!$A$1" }));
  })); sessions.push(s);
  const before = s.serialize(), snapshot = s.snapshot(), abort = new AbortController();
  await expect(s.prepareAsync(transaction(s, [edit("row", "insert", 0)]), undefined, {
    signal: abort.signal, yield: async () => {}, onProgress(progress) { if (progress.phase === "cells" && progress.completed === 256) abort.abort(); },
  })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  const candidate = await s.prepareAsync(transaction(s, [edit("row", "insert", 0)]));
  expect(s.serialize()).toEqual(before); s.discard(candidate);
  apply(s, [edit("row", "insert", 0)]); expect(s.history().undoSteps).toBe(1); s.undo();
  expect(s.snapshot().metadata).toEqual(snapshot.metadata); expect(s.snapshot().manifest.chunks).toEqual(snapshot.manifest.chunks);
});
