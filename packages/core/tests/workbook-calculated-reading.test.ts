import { expect, test } from "vitest";
import { captureWorkbookModel, createDocumentSession, createWorkbookCalculatedEvaluator, createWorkbookCalculator, DEFAULT_WORKBOOK_RULE_LIMITS, openDocument, readWorkbookRules } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const at = (row: number, column = 0) => ({ sheetId: "sheet-1", row, column });
const number = (value: number) => ({ type: "number" as const, value });
test("calculated reading reuses a completed long chain and exact batch context without rewriting source", async () => {
  const model = scalarWorkbook([...Array.from({ length: 1200 }, (_, row) => ({ row, column: 0, value: number(9999), formula: row ? `=A${row}+1` : "=1" })),
    { row: 0, column: 1, value: number(9999), formula: "=TODAY()" }], metadata => {
    metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "today", kind: "whole", operator: "equal", formulas: ["=TODAY()"], origin: { row: 0, column: 1 }, ranges: [{ from: { row: 0, column: 1 }, to: { row: 0, column: 1 } }],
      allowBlank: false, showErrorMessage: true, showInputMessage: true, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" }];
  }), before = model.serialize(), calculator = createWorkbookCalculator(model);
  const result = calculator.calculateSync({ context: { now: { unixMilliseconds: 0, utcOffsetMinutes: 0 } } });
  const evaluator = createWorkbookCalculatedEvaluator(model, result, { maxDepth: 8 });
  expect(evaluator.evaluate("=A1200+TODAY()", at(1199))).toMatchObject({ value: number(1200 + 25569), readCells: 1 });
  const read = await readWorkbookRules(model, { ...DEFAULT_WORKBOOK_RULE_LIMITS, documentId: "scalar-kernel", revision: 0, sheetId: "sheet-1", includeCurrent: true,
    coordinates: [{ row: 1199, column: 0 }, { row: 0, column: 1 }, { row: 1, column: 1 }] }, { calculation: result });
  expect(read.cells).toMatchObject([{ current: { kind: "value", value: number(1200) }, conditional: [] },
    { current: { kind: "value", value: number(25569) }, validation: { state: "valid" } }, { current: { kind: "value", value: { type: "blank" } }, conditional: [] }]);
  expect(model.serialize()).toEqual(before); calculator.dispose();
  expect(evaluator.cell(at(1199))).toMatchObject({ value: number(1200) }); evaluator.dispose();
});
test("calculation ownership rejects fabricated batches, equal IDs in another source and changed sessions", () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=1+1" }]));
  const original = captureWorkbookModel(session), result = session.calculateSync(), evaluator = createWorkbookCalculatedEvaluator(session, result);
  expect(() => createWorkbookCalculatedEvaluator(session, { ...result })).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(() => createWorkbookCalculatedEvaluator(openDocument(session.serialize(), { kind: "workbook" }), result)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  session.commit(session.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=10" } }] }] }));
  expect(() => createWorkbookCalculatedEvaluator(session, result)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  session.dispose(); expect(evaluator.cell(at(0))).toMatchObject({ revision: 0, value: number(2) }); evaluator.dispose();
  const retained = createWorkbookCalculatedEvaluator(original, result); expect(retained.cell(at(0))).toMatchObject({ revision: 0, value: number(2) }); retained.dispose();
});
test("calculated reads retain unsupported/cyclic states, source presence and text budgets", async () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=UNKNOWN()" }, { row: 1, column: 0, value: number(999), formula: "=A2" },
    { row: 2, column: 0, value: { type: "text", value: "old" }, formula: '=REPT_UNSUPPORTED()' }, { row: 3, column: 0, value: { type: "text", value: "old" }, formula: '="中文😀"' }]);
  const calculator = createWorkbookCalculator(model), result = calculator.calculateSync(), evaluator = createWorkbookCalculatedEvaluator(model, result);
  expect(evaluator.cell(at(0))).toMatchObject({ kind: "unavailable", reason: "unsupported-function", functionName: "UNKNOWN" });
  expect(evaluator.cell(at(1))).toMatchObject({ kind: "unavailable", reason: "circular-reference" });
  expect(evaluator.evaluate("=ISBLANK(A1)", at(0))).toMatchObject({ value: { type: "boolean", value: false } });
  expect(() => evaluator.cell(at(3), { maxCodeUnits: 1 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  const response = await readWorkbookRules(model, { ...DEFAULT_WORKBOOK_RULE_LIMITS, documentId: "scalar-kernel", revision: 0, sheetId: "sheet-1", coordinates: [{ row: 3, column: 0 }] }, { calculation: result });
  expect(response.cells).toEqual([]); evaluator.dispose(); calculator.dispose();
});
