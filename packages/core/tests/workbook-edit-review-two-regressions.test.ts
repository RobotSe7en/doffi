import { expect, test } from "vitest";
import { createDocumentSession, createWorkbookCalculator, createWorkbookScalarEvaluator, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookDataValidation, WorkbookDocumentMetadata, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const number = (value: number) => ({ type: "number" as const, value });
const point = (row: number, column: number) => ({ from: { row, column }, to: { row, column } });
const rule = (row: number, column: number, expression: string): WorkbookDataValidation => ({ ruleId: "stop", ranges: [point(row, column)], origin: { row, column },
  kind: "decimal", operator: "lessThanOrEqual", formulas: [expression], allowBlank: true, showErrorMessage: true, showInputMessage: false, errorStyle: "stop", errorTitle: "", error: "blocked", promptTitle: "", prompt: "" });
const write = (row: number, column: number, value: number): DocumentOperation => ({ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row, column, source: { kind: "value", value: number(value) } }] });
function prepare(session: WorkbookSession, operations: DocumentOperation[]) { return session.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: session.snapshot().manifest.revision, operations }); }
test("review: row visibility changes cannot bypass SUBTOTAL-dependent stop rules", () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 2, value: number(15) }, { row: 1, column: 1, value: number(10) }, { row: 2, column: 1, value: number(10) }], metadata => {
    metadata.workbook.sheets[0]!.dataValidations = [rule(0, 2, "=SUBTOTAL(109,$B$2:$B$3)")];
  })), before = session.serialize();
  expect(() => prepare(session, [{ op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { rowDimensions: [{ from: 1, to: 1, hidden: true }] } }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(session.serialize()).toEqual(before);
});
test("review: computed INDEX names retain SUMIF's expanded reference dependencies", () => {
  const session = createDocumentSession(scalarWorkbook([
    ...[0, 1, 2].map(row => ({ row, column: 0, value: number(1) })), { row: 0, column: 1, value: number(2) }, { row: 2, column: 1, value: number(5) }, { row: 0, column: 2, value: number(5) },
  ], metadata => {
    metadata.workbook.names = [{ nameId: "start", name: "Start", expression: "=INDEX(Sheet1!$B$1:$B$1,1,1)", origin: { row: 0, column: 0 } }];
    metadata.workbook.sheets[0]!.dataValidations = [rule(0, 2, "=SUMIF($A$1:$A$3,1,Start)")];
  })), before = session.serialize();
  expect(() => prepare(session, [write(2, 1, -20)])).toThrow(); expect(session.serialize()).toEqual(before);
});
test("review: reference selection parameters participate in validation dependencies", () => {
  const session = createDocumentSession(scalarWorkbook([
    ...[0, 1, 2].map(row => ({ row, column: 0, value: number(1) })), ...[10, 10, 10, -100].map((value, row) => ({ row, column: 1, value: number(value) })),
    { row: 0, column: 2, value: number(5) }, { row: 0, column: 3, value: number(1) },
  ], metadata => {
    metadata.workbook.names = [{ nameId: "start", name: "Start", expression: "=INDEX(Sheet1!$B$1:$B$4,Sheet1!$D$1,1)", origin: { row: 0, column: 0 } }];
    metadata.workbook.sheets[0]!.dataValidations = [rule(0, 2, "=SUMIF($A$1:$A$3,1,Start)")];
  }));
  expect(() => prepare(session, [write(0, 3, 2)])).toThrow(); expect(session.getCell("sheet-1", 0, 3)?.value).toEqual(number(1));
});
test("review: computed reference names retain nested SUBTOTAL visibility dependencies", () => {
  const session = createDocumentSession(scalarWorkbook([
    { row: 0, column: 1, value: number(10) }, { row: 1, column: 1, value: number(1) }, { row: 1, column: 2, value: number(1) }, { row: 0, column: 3, value: number(5) },
  ], metadata => {
    metadata.workbook.names = [{ nameId: "limit", name: "Limit", expression: "=INDEX(Sheet1!$B$1:$B$2,IF(SUBTOTAL(109,Sheet1!$C$2)>0,1,2),1)", origin: { row: 0, column: 0 } }];
    metadata.workbook.sheets[0]!.dataValidations = [rule(0, 3, "=Limit")];
  }));
  expect(() => prepare(session, [{ op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { rowDimensions: [{ from: 1, to: 1, hidden: true }] } }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(session.snapshot().metadata.workbook.sheets[0]!.rowDimensions).toEqual([]);
});
test.each([["=$B$1", 3], ["=SUM($B$1,$D$1)", 2]] as const)("review: unrelated edits between separated dependencies remain outside a large non-blank rule: %s", (formula, column) => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(3) }, { row: 0, column: 1, value: number(5) }, { row: 0, column: 3, value: number(5) }], metadata => {
    metadata.workbook.sheets[0]!.dataValidations = [{ ...rule(0, 0, formula), ranges: [{ from: { row: 0, column: 0 }, to: { row: 25_000, column: 0 } }], allowBlank: false }];
  }));
  expect(session.commit(prepare(session, [write(0, column, 1)])).changed).toBe(true);
  expect(session.getCell("sheet-1", 0, column)?.value).toEqual(number(1));
});
test("review: nonlinear relative references check middle implicit targets rather than rectangle corners", () => {
  const seed = createDocumentSession(scalarWorkbook(Array.from({ length: 6 }, (_, row) => ({ row, column: 1, value: number(1) })), metadata => {
    metadata.workbook.sheets.push({ ...metadata.workbook.sheets[0]!, sheetId: "other", name: "Other" });
    metadata.workbook.names = [{ nameId: "start", name: "Start", expression: "=INDEX(Sheet1!$B$1:$B$4,Other!D1,1)", origin: { row: 0, column: 0 } }];
  }));
  seed.commit(prepare(seed, [{ op: "setWorkbookCells", sheetId: "other", cells: [
    ...[0, 1, 2].map(row => ({ row, column: 4, source: { kind: "value" as const, value: number(1) } })),
    ...[[0, 1], [1, 4], [25_000, 1]].map(([row, value]) => ({ row: row!, column: 3, source: { kind: "value" as const, value: number(value!) } })),
  ] }]));
  const files = new Map(seed.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  const numeric = rule(0, 0, "=1");
  if (!("formulas" in numeric)) throw new Error("Expected numeric test rule");
  const { formulas: _formulas, kind: _kind, operator: _operator, ...common } = numeric;
  metadata.workbook.sheets[0]!.dataValidations = [{ ...common, kind: "custom", expression: "=SUMIF(Other!$E$1:$E$3,1,Start)>0",
    ranges: [{ from: { row: 0, column: 0 }, to: { row: 25_000, column: 0 } }], allowBlank: false }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" }));
  expect(() => prepare(session, [write(5, 1, -100)])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(session.getCell("sheet-1", 1, 0)).toBeUndefined(); expect(session.getCell("sheet-1", 5, 1)?.value).toEqual(number(1));
});
test("review: unrelated visibility and unchanged SUBTOTAL results preserve imported violations", () => {
  for (const [code, row] of [[109, 1000], [9, 1000], [9, 1]]) {
    const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 2, value: number(25) }, { row: 1, column: 1, value: number(10) }, { row: 2, column: 1, value: number(10) }], metadata => {
      metadata.workbook.sheets[0]!.dataValidations = [rule(0, 2, `=SUBTOTAL(${code},$B$2:$B$3)`)];
    }));
    expect(session.commit(prepare(session, [{ op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { rowDimensions: [{ from: row!, to: row!, hidden: true }] } }])).changed).toBe(true);
  }
});
test("review: pure worksheet rename retains pre-existing invalid literal and formula targets", () => {
  for (const formula of [undefined, "=Sheet1!$B$1+4"]) {
    const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(9), ...(formula ? { formula } : {}) }, { row: 0, column: 1, value: number(5) }], metadata => {
      metadata.workbook.sheets[0]!.dataValidations = [rule(0, 0, "=Sheet1!$B$1")];
    }));
    const rename: DocumentOperation = { op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { name: "Renamed" } };
    expect(session.commit(prepare(session, [rename])).changed).toBe(true); session.undo();
    expect(() => prepare(session, [rename, write(0, 1, 4)])).toThrow(); expect(session.snapshot().metadata.workbook.sheets[0]!.name).toBe("Sheet1");
  }
});
test("review: deleting a table's worksheet makes surviving structured and bare references explicit", () => {
  const session = createDocumentSession(scalarWorkbook([
    { row: 0, column: 0, value: { type: "text", value: "Amount" } }, ...[1, 2, 3].map((value, index) => ({ row: index + 1, column: 0, value: number(value) })),
  ], metadata => {
    const first = metadata.workbook.sheets[0]!;
    metadata.workbook.sheets.push({ ...first, sheetId: "other", name: "Other" });
    first.tables = [{ tableId: "data", name: "DataTable", range: { from: { row: 0, column: 0 }, to: { row: 3, column: 0 } }, header: true, totals: false, totalsPreviouslyShown: false,
      columns: [{ columnId: "amount", name: "Amount", totalFunction: "none" }], showFirstColumn: false, showLastColumn: false, rowStripes: false, columnStripes: false }];
    metadata.workbook.names = [{ nameId: "amount", name: "Amounts", expression: "=DataTable[Amount]", origin: { row: 0, column: 0 } }];
  }));
  session.commit(prepare(session, [{ op: "setWorkbookCells", sheetId: "other", cells: [
    { row: 0, column: 0, source: { kind: "formula", expression: "=SUM(DataTable[Amount])" } }, { row: 0, column: 1, source: { kind: "formula", expression: "=SUM(DataTable)" } },
    { row: 0, column: 2, source: { kind: "formula", expression: "=SUM(Other!DataTable)" } },
  ] }]));
  session.calculateSync(); expect(session.getCalculatedCell("other", 0, 0)).toMatchObject({ value: number(6) });
  session.commit(prepare(session, [{ op: "removeWorkbookSheet", sheetId: "sheet-1" }]));
  for (const column of [0, 1]) expect(session.getCell("other", 0, column)?.formula?.expression).toBe("=SUM(#REF!)");
  expect(session.getCell("other", 0, 2)?.formula?.expression).toBe("=SUM(Other!DataTable)");
  expect(session.snapshot().metadata.workbook.names?.[0]?.expression).toBe("=#REF!");
  session.calculateSync(); expect(session.getCalculatedCell("other", 0, 0)).toMatchObject({ value: { type: "error", value: "#REF!" } });
  session.undo(); session.calculateSync(); expect(session.getCalculatedCell("other", 0, 0)).toMatchObject({ value: number(6) });
});
test("review: filter predicate recursion preserves the caller's depth budget", () => {
  for (const [length, maxDepth] of [[10, 4], [600, 128]] as const) {
    const model = scalarWorkbook([
      { row: 0, column: 0, value: { type: "text", value: "Filter" } }, ...Array.from({ length: length! }, (_, index) => ({ row: index + 1, column: 0, value: number(999), formula: `=SUBTOTAL(9,C${index + 3})` })),
      { row: length! + 1, column: 0, value: number(1) }, ...Array.from({ length: length! + 1 }, (_, index) => ({ row: index + 1, column: 2, value: number(1) })),
    ], metadata => { metadata.workbook.sheets[0]!.autoFilter = { range: { from: { row: 0, column: 0 }, to: { row: length! + 1, column: 2 } }, columns: [{ column: 0, hiddenButton: false, showButton: true, criterion: { kind: "custom", and: false, tests: [{ operator: "greaterThan", value: "0" }] } }] }; });
    const evaluator = createWorkbookScalarEvaluator(model, { maxDepth });
    expect(() => evaluator.evaluate("=SUBTOTAL(9,C2)", { sheetId: "sheet-1", row: 0, column: 3 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); evaluator.dispose();
    const calculator = createWorkbookCalculator(model); expect(calculator.calculateSync().calculated).toBe(length);
    expect(calculator.getCell({ sheetId: "sheet-1", row: 1, column: 0 })).toMatchObject({ state: "valid", value: number(1) }); calculator.dispose();
  }
});
