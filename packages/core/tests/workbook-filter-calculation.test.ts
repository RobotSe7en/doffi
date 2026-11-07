import { expect, test } from "vitest";
import { createDocumentSession, createWorkbookCalculator, createWorkbookScalarEvaluator } from "../src/index.js";
import type { WorkbookFilterCriterion } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const number = (value: number) => ({ type: "number" as const, value });
const text = (value: string) => ({ type: "text" as const, value });
const at = { sheetId: "sheet-1", row: 10, column: 3 };
function filtered(values: number[], criterion: WorkbookFilterCriterion) {
  return scalarWorkbook([{ row: 0, column: 0, value: text("Header") }, ...values.map((value, index) => ({ row: index + 1, column: 0, value: number(value) }))], metadata => {
    metadata.workbook.sheets[0]!.autoFilter = { range: { from: { row: 0, column: 0 }, to: { row: values.length, column: 0 } }, columns: [{ column: 0, hiddenButton: false, showButton: true, criterion }] };
  });
}
test("SUBTOTAL filters entire rows while manual hidden rows depend on the function code", () => {
  const model = scalarWorkbook([
    { row: 0, column: 0, value: text("Header") }, { row: 1, column: 0, value: text("keep") }, { row: 2, column: 0, value: text("skip") }, { row: 3, column: 0, value: text("keep") },
    { row: 1, column: 1, value: number(1) }, { row: 2, column: 1, value: number(10) }, { row: 3, column: 1, value: number(100) },
  ], metadata => {
    const sheet = metadata.workbook.sheets[0]!; sheet.rowDimensions = [{ from: 1, to: 1, hidden: true }];
    sheet.autoFilter = { range: { from: { row: 0, column: 0 }, to: { row: 3, column: 1 } }, columns: [{ column: 0, hiddenButton: false, showButton: true, criterion: { kind: "values", blank: false, values: ["KEEP"], dates: [] } }] };
  });
  const evaluator = createWorkbookScalarEvaluator(model);
  expect(evaluator.evaluate("=SUBTOTAL(9,B2:B4)", at)).toMatchObject({ value: number(101) });
  expect(evaluator.evaluate("=SUBTOTAL(109,B2:B4)", at)).toMatchObject({ value: number(100) }); evaluator.dispose();
});
test.each([
  [{ kind: "top", top: true, percent: false, count: 2 }, 8],
  [{ kind: "top", top: false, percent: false, count: 2 }, 5],
  [{ kind: "top", top: true, percent: true, count: 25 }, 4],
  [{ kind: "dynamic", type: "aboveAverage" }, 4],
  [{ kind: "dynamic", type: "belowAverage" }, 5],
  [{ kind: "custom", and: true, tests: [{ operator: "greaterThanOrEqual", value: "2" }, { operator: "lessThan", value: "4" }] }, 4],
] satisfies [WorkbookFilterCriterion, number][])("top/average/custom criteria use current values and preserve ties: %j", (criterion, expected) => {
  const evaluator = createWorkbookScalarEvaluator(filtered([1, 2, 2, 4], criterion));
  expect(evaluator.evaluate("=SUBTOTAL(9,A2:A5)", at)).toMatchObject({ value: number(expected) }); evaluator.dispose();
});
test("date grouping honors Excel's synthetic leap day and host-supplied dynamic dates", () => {
  const evaluator = createWorkbookScalarEvaluator(filtered([59, 60, 61], { kind: "values", blank: false, values: [], dates: [{ grouping: "day", year: 1900, month: 2, day: 29 }] }));
  expect(evaluator.evaluate("=SUBTOTAL(9,A2:A4)", at)).toMatchObject({ value: number(60) }); evaluator.dispose();
  const model = filtered([45351, 45352, 45353], { kind: "dynamic", type: "today" });
  const noClock = createWorkbookScalarEvaluator(model); expect(noClock.evaluate("=SUBTOTAL(9,A2:A4)", at)).toMatchObject({ kind: "unavailable", reason: "calculation-required" }); noClock.dispose();
  const clock = createWorkbookScalarEvaluator(model, {}, { now: { unixMilliseconds: Date.UTC(2024, 1, 29, 18), utcOffsetMinutes: 480 } });
  expect(clock.evaluate("=SUBTOTAL(9,A2:A4)", at)).toMatchObject({ value: number(45352) }); clock.dispose();
});
test("filter inputs become calculation dependencies, including formula-valued predicates", () => {
  const session = createDocumentSession(scalarWorkbook([
    { row: 0, column: 0, value: text("Header") }, { row: 1, column: 0, value: number(999), formula: "=C2" }, { row: 2, column: 0, value: number(2) },
    { row: 1, column: 1, value: number(10) }, { row: 2, column: 1, value: number(20) }, { row: 1, column: 2, value: number(1) },
    { row: 10, column: 0, value: number(999), formula: "=SUBTOTAL(9,B2:B3)" }, { row: 10, column: 1, value: number(999), formula: "=A11*2" },
  ], metadata => {
    metadata.workbook.sheets[0]!.autoFilter = { range: { from: { row: 0, column: 0 }, to: { row: 2, column: 2 } }, columns: [{ column: 0, hiddenButton: false, showButton: true, criterion: { kind: "custom", and: false, tests: [{ operator: "greaterThan", value: "1" }] } }] };
  }));
  session.calculateSync(); expect(session.getCalculatedCell("sheet-1", 10, 1)).toMatchObject({ value: number(40) });
  session.commit(session.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 1, column: 2, source: { kind: "value", value: number(3) } }] }] }));
  session.calculateSync(); expect(session.getCalculatedCell("sheet-1", 10, 1)).toMatchObject({ value: number(60) });
  session.undo(); session.calculateSync(); expect(session.getCalculatedCell("sheet-1", 10, 1)).toMatchObject({ value: number(40) });
});
test("a new host date invalidates SUBTOTAL's dynamic filter and its dependents", () => {
  const source = filtered([45351, 45352], { kind: "dynamic", type: "today" });
  const session = createDocumentSession(source);
  session.commit(session.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 10, column: 0, source: { kind: "formula", expression: "=SUBTOTAL(9,A2:A3)" } }] }] }));
  const calculator = createWorkbookCalculator(session);
  calculator.calculateSync({ context: { now: { unixMilliseconds: Date.UTC(2024, 1, 29), utcOffsetMinutes: 0 } } }); expect(calculator.getCell({ ...at, column: 0 })).toMatchObject({ value: number(45351) });
  expect(calculator.calculateSync({ context: { now: { unixMilliseconds: Date.UTC(2024, 2, 1), utcOffsetMinutes: 0 } } }).calculated).toBe(1);
  expect(calculator.getCell({ ...at, column: 0 })).toMatchObject({ value: number(45352) }); calculator.dispose(); session.dispose();
});
