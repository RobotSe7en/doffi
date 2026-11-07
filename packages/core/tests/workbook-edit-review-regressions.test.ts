import { expect, test } from "vitest";
import { createDocumentSession, createWorkbook, createWorkbookCalculator, createWorkbookScalarEvaluator, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const number = (value: number) => ({ type: "number" as const, value });
const text = (value: string) => ({ type: "text" as const, value });
const at = (row: number, column = 0) => ({ sheetId: "sheet-1", row, column });
const write = (row: number, column: number, value: number): DocumentOperation => ({ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row, column, source: { kind: "value", value: number(value) } }] });
function commit(session: WorkbookSession, operations: DocumentOperation[]) {
  return session.commit(session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: session.snapshot().manifest.revision, operations }));
}
test("review: file budgets count final new/removed chunks rather than touched existing chunks", () => {
  const seed = createDocumentSession(openDocument(createWorkbook({ documentId: "quota" }), { kind: "workbook" }));
  commit(seed, [write(0, 0, 1), write(256, 0, 2)]); expect(seed.serialize().size).toBe(6);
  const session = createDocumentSession(openDocument(seed.serialize(), { kind: "workbook" }), { limits: { maxFiles: 7 } });
  expect(commit(session, [write(0, 0, 3), write(256, 0, 4), write(512, 0, 5)]).changed).toBe(true); expect(session.serialize().size).toBe(7);
  const exact = createDocumentSession(openDocument(seed.serialize(), { kind: "workbook" }), { limits: { maxFiles: 6 } });
  commit(exact, [{ op: "clearWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 255, column: 31 } }], mode: "all" }, write(512, 0, 5)]);
  expect(exact.serialize().size).toBe(6);
  const reversed = createDocumentSession(openDocument(seed.serialize(), { kind: "workbook" }), { limits: { maxFiles: 6 } });
  commit(reversed, [write(512, 0, 5), { op: "clearWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 255, column: 31 } }], mode: "all" }]);
  expect(reversed.serialize().size).toBe(6);
  const blank = createDocumentSession(openDocument(createWorkbook({ documentId: "blank" }), { kind: "workbook" }), { limits: { maxFiles: 4 } });
  expect(commit(blank, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 999, column: 999, source: { kind: "value", value: { type: "blank" } } }] }]).changed).toBe(false);
});
test("review: locale context invalidates formatted values, parsing and dependent closures", () => {
  const calculator = createWorkbookCalculator(scalarWorkbook([
    { row: 0, column: 0, value: number(999), formula: '=VALUE("1,234")' },
    { row: 0, column: 1, value: number(999), formula: "=A1*2" },
    { row: 1, column: 0, value: number(999), formula: '=TEXT(1234.5,"#,##0.00")' },
  ]));
  calculator.calculateSync({ context: { locale: "en-US" } }); expect(calculator.getCell(at(0))).toMatchObject({ value: number(1234) });
  expect(calculator.calculateSync({ context: { locale: "de-DE" } }).reused).toBe(0);
  expect(calculator.getCell(at(0))).toMatchObject({ value: number(1.234) }); expect(calculator.getCell(at(0, 1))).toMatchObject({ value: number(2.468) });
  expect(calculator.getCell(at(1))).toMatchObject({ value: text("1.234,50") }); calculator.dispose();
});
test("review: ordered text ranges preserve implicit blank positions and cross-chunk row order", () => {
  const evaluator = createWorkbookScalarEvaluator(scalarWorkbook([
    { row: 0, column: 0, value: text("a") }, { row: 2, column: 0, value: text("c") },
    { row: 0, column: 31, value: text("x") }, { row: 1, column: 31, value: text("y") },
    { row: 0, column: 32, value: text("z") }, { row: 1, column: 32, value: text("w") },
  ]));
  expect(evaluator.evaluate('=TEXTJOIN("-",FALSE,A1:A3)', at(0))).toMatchObject({ value: text("a--c") });
  expect(evaluator.evaluate('=CONCAT(AF1:AG2)', at(0))).toMatchObject({ value: text("xzyw") });
  expect(evaluator.evaluate('=TEXTJOIN("-",FALSE,AF1:AG2)', at(0))).toMatchObject({ value: text("x-z-y-w") }); evaluator.dispose();
});
test("review: actual expanded SUMIF/AVERAGEIF ranges invalidate edits, new data and undo", () => {
  const session = createDocumentSession(scalarWorkbook([
    ...[1, 1, 1].map((value, row) => ({ row, column: 0, value: number(value) })), ...[10, 20, 30].map((value, row) => ({ row, column: 1, value: number(value) })),
    { row: 0, column: 2, value: number(999), formula: "=SUMIF(A1:A3,1,B1)" }, { row: 0, column: 3, value: number(999), formula: "=AVERAGEIF(A1:A3,1,B1)" },
  ]));
  session.calculateSync(); commit(session, [write(2, 1, 300)]); expect(session.calculateSync().calculated).toBe(2);
  expect(session.getCalculatedCell("sheet-1", 0, 2)).toMatchObject({ value: number(330) }); expect(session.getCalculatedCell("sheet-1", 0, 3)).toMatchObject({ value: number(110) });
  session.undo(); session.calculateSync(); expect(session.getCalculatedCell("sheet-1", 0, 2)).toMatchObject({ value: number(60) });
  const sparse = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(1) }, { row: 0, column: 2, value: number(999), formula: "=SUMIF(A1:A3,0,B1)" }]));
  sparse.calculateSync(); commit(sparse, [write(2, 1, 10)]); sparse.calculateSync(); expect(sparse.getCalculatedCell("sheet-1", 0, 2)).toMatchObject({ value: number(10) });
});
test("review: static prerequisite cycles cannot survive an earlier actual formula error", () => {
  const calculator = createWorkbookCalculator(scalarWorkbook([
    { row: 0, column: 0, value: number(999), formula: "=SUM(1/0,B1)" }, { row: 0, column: 1, value: number(999), formula: "=IFERROR(A1,2)" },
    { row: 0, column: 2, value: number(999), formula: "=B1+1" },
  ]));
  calculator.calculateSync(); expect(calculator.getCell(at(0))).toMatchObject({ state: "valid", value: { type: "error", value: "#DIV/0!" } });
  expect(calculator.getCell(at(0, 1))).toMatchObject({ state: "valid", value: number(2) }); expect(calculator.getCell(at(0, 2))).toMatchObject({ state: "valid", value: number(3) });
  expect(calculator.calculateSync().reused).toBe(3); calculator.dispose();
});
test("review: SUBTOTAL COUNTA counts source presence without evaluating unsupported formulas", () => {
  const evaluator = createWorkbookScalarEvaluator(scalarWorkbook([
    { row: 0, column: 0, value: number(999), formula: "=FUTUREFUNC()" }, { row: 1, column: 0, value: number(999), formula: "=1/0" }, { row: 2, column: 0, value: text(""), formula: '=""' },
  ]));
  for (const code of [3, 103]) expect(evaluator.evaluate(`=SUBTOTAL(${code},A1:A3)`, at(0, 2))).toMatchObject({ value: number(3) }); evaluator.dispose();
});
test("review: quoted error criteria match error cells in all conditional aggregates", () => {
  const evaluator = createWorkbookScalarEvaluator(scalarWorkbook([{ row: 0, column: 0, value: { type: "error", value: "#N/A" } }, { row: 0, column: 1, value: number(10) }]));
  for (const formula of ['=COUNTIF(A1,"#N/A")', '=COUNTIFS(A1,"#N/A")']) expect(evaluator.evaluate(formula, at(0, 2))).toMatchObject({ value: number(1) });
  expect(evaluator.evaluate('=COUNTIF(A1,"<>#N/A")', at(0, 2))).toMatchObject({ value: number(0) });
  for (const formula of ['=SUMIF(A1,"#N/A",B1)', '=AVERAGEIF(A1,"#N/A",B1)']) expect(evaluator.evaluate(formula, at(0, 2))).toMatchObject({ value: number(10) }); evaluator.dispose();
});
test("review: AVERAGEIF ignores logical criteria while AVERAGEIFS converts them to numbers", () => {
  const evaluator = createWorkbookScalarEvaluator(scalarWorkbook([
    { row: 0, column: 0, value: { type: "boolean", value: true } }, { row: 1, column: 0, value: number(1) }, { row: 0, column: 1, value: number(100) }, { row: 1, column: 1, value: number(10) },
  ]));
  expect(evaluator.evaluate('=AVERAGEIF(A1:A2,">0",B1:B2)', at(0, 2))).toMatchObject({ value: number(10) });
  expect(evaluator.evaluate('=AVERAGEIFS(B1:B2,A1:A2,1)', at(0, 2))).toMatchObject({ value: number(55) }); evaluator.dispose();
});
