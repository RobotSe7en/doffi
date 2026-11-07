import { expect, test } from "vitest";
import { createDocumentSession, createWorkbookCalculator, createWorkbookScalarEvaluator, parseWorkbookFormula } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const location = (row: number, column = 0) => ({ sheetId: "sheet-1", row, column });
const number = (value: number) => ({ type: "number" as const, value });
test("ISBLANK reads source presence even for unsupported, empty-string and cyclic formulas", () => {
  const model = scalarWorkbook([
    { row: 0, column: 0, value: number(999), formula: "=FUTUREFUNC()" }, { row: 1, column: 0, value: number(999), formula: '=ISBLANK(A2)' },
    { row: 2, column: 0, value: { type: "text", value: "" }, formula: '=""' },
    { row: 0, column: 1, value: number(999), formula: "=ISBLANK(A1)" }, { row: 2, column: 1, value: number(999), formula: "=ISBLANK(A3)" },
    { row: 3, column: 1, value: number(999), formula: "=ISBLANK(A4)" },
  ]);
  const evaluator = createWorkbookScalarEvaluator(model);
  expect(evaluator.evaluate("=ISBLANK(A1)", location(0, 1))).toMatchObject({ value: { type: "boolean", value: false } }); evaluator.dispose();
  const calculator = createWorkbookCalculator(model); calculator.calculateSync();
  for (const at of [location(0, 1), location(1), location(2, 1)]) expect(calculator.getCell(at)).toMatchObject({ state: "valid", value: { type: "boolean", value: false } });
  expect(calculator.getCell(location(3, 1))).toMatchObject({ state: "valid", value: { type: "boolean", value: true } }); calculator.dispose();
});
test("scalar queries capture an immutable session revision and formula arity is bounded at 255", () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(2) }])), evaluator = createWorkbookScalarEvaluator(session);
  session.commit(session.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: number(10) } }] }] }));
  expect(evaluator.evaluate("=A1", location(0))).toMatchObject({ revision: 0, kind: "value", value: number(2) });
  const current = createWorkbookScalarEvaluator(session); expect(current.evaluate("=A1", location(0))).toMatchObject({ revision: 1, kind: "value", value: number(10) });
  expect(parseWorkbookFormula("=SUM(" + Array(255).fill("1").join(",") + ")", location(0)).kind).toBe("call");
  expect(() => parseWorkbookFormula("=SUM(" + Array(256).fill("1").join(",") + ")", location(0))).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  evaluator.dispose(); current.dispose(); session.dispose();
});
test("evaluates long dependency chains iteratively and never trusts imported caches", async () => {
  const size = 1200, model = scalarWorkbook(Array.from({ length: size }, (_, row) => ({ row, column: 0, value: number(9999), formula: row ? `=A${row}+1` : "=1" })));
  const before = model.serialize(), calculator = createWorkbookCalculator(model), progress: string[] = [];
  const result = await calculator.calculate({ onProgress: value => progress.push(value.phase) });
  expect(result).toMatchObject({ calculated: size, reused: 0, dependencies: size - 1, documentId: "scalar-kernel", revision: 0 });
  expect(calculator.getCell(location(size - 1))).toMatchObject({ state: "valid", value: number(size), batchId: result.batchId });
  expect(model.serialize()).toEqual(before); expect(progress).toContain("calculate");
  expect(calculator.calculateSync()).toMatchObject({ calculated: 0, reused: size }); calculator.dispose();
});
test("does not confuse unvisited conditional branches with actual cycles", () => {
  const model = scalarWorkbook([
    { row: 0, column: 0, value: number(999), formula: "=IF(FALSE,B1,0)" },
    { row: 0, column: 1, value: number(999), formula: "=A1+1" },
    { row: 1, column: 0, value: number(999), formula: "=B2+1" },
    { row: 1, column: 1, value: number(999), formula: "=A2+1" },
    { row: 2, column: 0, value: number(999), formula: "=IFERROR(1/0,7)" },
  ]), calculator = createWorkbookCalculator(model); calculator.calculateSync();
  expect(calculator.getCell(location(0))).toMatchObject({ state: "valid", value: number(0) });
  expect(calculator.getCell(location(0, 1))).toMatchObject({ state: "valid", value: number(1) });
  expect(calculator.getCell(location(1))).toMatchObject({ state: "unavailable", issue: { reason: "circular-reference" } });
  expect(calculator.getCell(location(2))).toMatchObject({ state: "valid", value: number(7) }); calculator.dispose();
});
test("selective lookups and presence functions do not eagerly evaluate unused cyclic dependents", () => {
  const model = scalarWorkbook([
    { row: 0, column: 0, value: number(999), formula: "=MATCH(1,A2:A3,0)" }, { row: 1, column: 0, value: number(1) },
    { row: 2, column: 0, value: number(999), formula: "=A1+1" },
    { row: 0, column: 1, value: number(999), formula: "=COUNTA(B1:B2)" }, { row: 1, column: 1, value: number(999), formula: "=B1+1" },
  ]), calculator = createWorkbookCalculator(model); calculator.calculateSync();
  expect(calculator.getCell(location(0))).toMatchObject({ state: "valid", value: number(1) });
  expect(calculator.getCell(location(2))).toMatchObject({ state: "valid", value: number(2) });
  expect(calculator.getCell(location(0, 1))).toMatchObject({ state: "valid", value: number(2) });
  expect(calculator.getCell(location(1, 1))).toMatchObject({ state: "valid", value: number(3) }); calculator.dispose();
});
test("invalidates exact dependent closures while reusing unaffected results across session revisions", () => {
  const session = createDocumentSession(scalarWorkbook([
    { row: 0, column: 0, value: number(2) }, { row: 0, column: 1, value: number(999), formula: "=A1+1" },
    { row: 0, column: 2, value: number(999), formula: "=B1*2" }, { row: 1, column: 1, value: number(999), formula: "=1+1" },
  ])), calculator = createWorkbookCalculator(session);
  calculator.calculateSync();
  session.commit(session.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: number(10) } }] }] }));
  expect(calculator.getCell(location(0, 2))).toMatchObject({ state: "stale", revision: 1 });
  expect(calculator.calculateSync()).toMatchObject({ calculated: 2, reused: 1, revision: 1 });
  expect(calculator.getCell(location(0, 2))).toMatchObject({ state: "valid", value: number(22), revision: 1 });
  expect(session.history().undoSteps).toBe(1); session.undo(); calculator.calculateSync();
  expect(calculator.getCell(location(0, 2))).toMatchObject({ state: "valid", value: number(6), revision: 2 });
  expect(session.history()).toMatchObject({ undoSteps: 0, redoSteps: 1 }); calculator.dispose(); session.dispose();
});
test("whole-axis dependencies remain rectangular and include sparse distant data", () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(2) }, { row: 1_048_575, column: 0, value: number(5) },
    { row: 0, column: 1, value: number(999), formula: "=SUM(A:A)" }]), calculator = createWorkbookCalculator(model, { maxDependencies: 1 });
  expect(calculator.calculateSync()).toMatchObject({ dependencies: 1, calculated: 1 });
  expect(calculator.getCell(location(0, 1))).toMatchObject({ state: "valid", value: number(7) }); calculator.dispose();
});
test("cancellation and late revisions never install partial results", async () => {
  const session = createDocumentSession(scalarWorkbook(Array.from({ length: 200 }, (_, row) => ({ row, column: 0, value: number(999), formula: "=1+1" }))));
  const calculator = createWorkbookCalculator(session), signal = { aborted: false }; let checkpoints = 0;
  await expect(calculator.calculate({ signal, yield: async () => { checkpoints++; signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(checkpoints).toBe(1); expect(calculator.getCell(location(0))).toMatchObject({ state: "stale" });
  const result = calculator.calculateSync(); expect(result.calculated).toBe(200);
  await expect(calculator.calculate({ yield: async () => {
    session.commit(session.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: session.snapshot().manifest.revision, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=3" } }] }] }));
  } })).rejects.toMatchObject({ code: "RevisionConflict" });
  expect(calculator.getCell(location(0))).toMatchObject({ state: "stale", revision: 1 });
  calculator.calculateSync(); expect(calculator.getCell(location(0))).toMatchObject({ state: "valid", value: number(3) }); calculator.dispose(); session.dispose();
});
test("batch budgets span graph, dependency evaluation and all output, and disposal cancels active work", async () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=1+1" }, { row: 1, column: 0, value: number(999), formula: "=A1+1" }]);
  for (const input of [{ maxSteps: 1 }, { maxCodeUnits: 1 }, { maxFormulas: 1 }]) {
    const calculator = createWorkbookCalculator(model, input); expect(() => calculator.calculateSync()).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(calculator.getCell(location(0))).toMatchObject({ state: "stale" }); calculator.dispose();
  }
  const calculator = createWorkbookCalculator(scalarWorkbook(Array.from({ length: 100 }, (_, row) => ({ row, column: 0, value: number(999), formula: "=1" }))));
  await expect(calculator.calculate({ yield: async () => { calculator.dispose(); } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(() => calculator.getCell(location(0))).toThrow(); calculator.dispose();
});
