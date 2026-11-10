import { expect, test } from "vitest";
import { createDocumentSession, createWorkbookCalculator, createWorkbookScalarEvaluator, parseWorkbookFormula } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const location = (row: number, column = 0) => ({ sheetId: "sheet-1", row, column });
const number = (value: number) => ({ type: "number" as const, value });
test('incremental output charges retain untouched strings and exact UTF-16 totals across growth, shrinkage and scalar types', () => {
  const other = '乙𠀀'.repeat(2000), expression = '=A1', untouched = '=D1';
  const text = (value: string) => ({ type: 'text' as const, value });
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: text('initial') },
    { row: 0, column: 1, value: number(999), formula: expression }, { row: 0, column: 2, value: number(999), formula: untouched },
    { row: 0, column: 3, value: text(other) }]));
  const calculator = createWorkbookCalculator(session), original = calculator.calculateSync();
  for (const value of [text('甲𠀀禰󠄀'.repeat(64)), number(42), { type: 'error' as const, value: '#N/A' as const }, text('Z'), text('甲𠀀禰󠄀'.repeat(90))]) {
    const manifest = session.snapshot().manifest;
    session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision,
      operations: [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row: 0, column: 0, source: { kind: 'value', value } }] }] }));
    expect(calculator.calculateSync()).toMatchObject({ calculated: 1, reused: 1 });
    const expected = expression.length + untouched.length + JSON.stringify(['sheet-1', 0, 1]).length + (value.type === 'text' ? value.value.length : 0)
      + JSON.stringify(['sheet-1', 0, 2]).length + other.length;
    const seed = calculator.calculateSync(); expect(seed).toMatchObject({ calculated: 0, reused: 2, codeUnits: expected });
    const bounded = createWorkbookCalculator(session, { maxCodeUnits: expected - 1 }, seed);
    expect(() => bounded.calculateSync()).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' })); bounded.dispose();
    expect(seed.cells.map(cell => cell.state === 'valid' && cell.value)).toEqual([value, text(other)]);
  }
  expect(original.cells[0]).toMatchObject({ value: text('initial') }); calculator.dispose(); session.dispose();
});
test('immutable chunk summaries keep complete budgets and rebuild after cross-chunk formula changes', () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(2) },
    ...[0, 256, 512].map(row => ({ row, column: 1, value: number(999), formula: '=A1+' + row }))]));
  const original = createWorkbookCalculator(session), seed = original.calculateSync();
  const set = (row: number, column: number, source: { kind: 'value'; value: ReturnType<typeof number> } | { kind: 'formula'; expression: string }) => {
    const manifest = session.snapshot().manifest;
    session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision,
      operations: [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row, column, source }] }] }));
  };
  set(0, 0, { kind: 'value', value: number(7) });
  const probe = createWorkbookCalculator(session, {}, seed), changed = probe.calculateSync();
  expect(changed).toMatchObject({ calculated: 3, reused: 0, dependencies: 3 });
  expect(changed.cells.map(cell => cell.state === 'valid' && cell.value)).toEqual([number(7), number(263), number(519)]);
  for (const limits of [{ maxFormulas: 2 }, { maxDependencies: 2 }, { maxSteps: changed.steps - 1 }]) {
    const bounded = createWorkbookCalculator(session, limits, seed);
    expect(() => bounded.calculateSync()).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' })); bounded.dispose();
  }
  set(512, 1, { kind: 'formula', expression: '=A1*2' });
  const next = probe.calculateSync(); expect(next).toMatchObject({ calculated: 1, reused: 2, dependencies: 3 });
  expect(probe.getCell(location(512, 1))).toMatchObject({ value: number(14) });
  set(256, 1, { kind: 'value', value: number(8) });
  const reduced = createWorkbookCalculator(session, { maxFormulas: 2, maxDependencies: 2 }, next);
  expect(reduced.calculateSync()).toMatchObject({ calculated: 0, reused: 2, dependencies: 2 });
  expect(seed.cells.map(cell => cell.state === 'valid' && cell.value)).toEqual([number(2), number(258), number(514)]);
  original.dispose(); probe.dispose(); reduced.dispose(); session.dispose();
});
test('completed immutable seeds retain full step and UTF-16 output budgets on constant-time reuse', () => {
  const text = '甲𠀀'.repeat(80), model = scalarWorkbook([{ row: 0, column: 0, value: { type: 'text', value: text } },
    { row: 0, column: 1, value: number(999), formula: '=A1' }, { row: 0, column: 2, value: number(999), formula: '=B1&"尾"' }]);
  const original = createWorkbookCalculator(model), initial = original.calculateSync(), seed = original.calculateSync();
  expect(seed).toMatchObject({ calculated: 0, reused: 2 }); expect(seed.codeUnits).toBeGreaterThan(2 * text.length);
  for (const limits of [{ maxCodeUnits: seed.codeUnits - 1 }, { maxSteps: seed.steps - 1 }]) {
    const bounded = createWorkbookCalculator(model, limits, seed);
    expect(() => bounded.calculateSync()).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' })); bounded.dispose();
  }
  const exact = createWorkbookCalculator(model, { maxCodeUnits: seed.codeUnits, maxSteps: seed.steps }, seed);
  expect(exact.calculateSync()).toMatchObject({ calculated: 0, reused: 2, codeUnits: seed.codeUnits, steps: seed.steps });
  expect(initial.cells[1]).toMatchObject({ value: { type: 'text', value: text + '尾' } }); exact.dispose(); original.dispose();
});
test('incremental dirty roots can enter and leave a real lazy cycle while retaining unrelated values and old seeds', () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(0) },
    { row: 0, column: 1, value: number(999), formula: '=IF(A1,C1,5)' }, { row: 0, column: 2, value: number(999), formula: '=B1+1' },
    { row: 0, column: 3, value: number(999), formula: '=7' }]));
  const calculator = createWorkbookCalculator(session), seed = calculator.calculateSync();
  for (const value of [1, 0]) {
    const m = session.snapshot().manifest;
    session.commit(session.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [
      { op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row: 0, column: 0, source: { kind: 'value', value: number(value) } }] }] }));
    expect(calculator.calculateSync()).toMatchObject({ calculated: 2, reused: 1 });
    for (const column of [1, 2]) expect(calculator.getCell(location(0, column))).toMatchObject(value
      ? { state: 'unavailable', issue: { reason: 'circular-reference' } } : { state: 'valid', value: number(column + 4) });
    expect(calculator.getCell(location(0, 3))).toMatchObject({ value: number(7) });
  }
  expect(seed.cells[0]).toMatchObject({ revision: 0, value: number(5) }); calculator.dispose(); session.dispose();
});
test('cancelling a reused graph after learning an expanded lookup range cannot corrupt a completed seed', async () => {
  const session = createDocumentSession(scalarWorkbook([
    ...[0, 1].map(row => ({ row, column: 0, value: { type: 'text' as const, value: 'keep' } })),
    ...[[0, 10], [1, 20], [3, 40], [4, 50]].map(([row, value]) => ({ row: row!, column: 1, value: number(value!) })),
    { row: 5, column: 3, value: number(1) }, { row: 0, column: 4, value: number(999), formula: '=SUMIF(A1:A2,"keep",INDEX(B1:D4,D6,1))' },
  ]));
  const original = createWorkbookCalculator(session), seed = original.calculateSync(); expect(seed.cells[0]).toMatchObject({ value: number(30) });
  const set = (row: number, column: number, value: number) => {
    const m = session.snapshot().manifest;
    session.commit(session.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision,
      operations: [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row, column, source: { kind: 'value', value: number(value) } }] }] }));
  };
  set(5, 3, 4); const cancelled = createWorkbookCalculator(session, {}, seed), signal = { aborted: false }; let reached = false;
  await expect(cancelled.calculate({ signal, onProgress(value) { if (value.phase === 'calculate' && value.completed) { reached = true; signal.aborted = true; } } })).rejects.toMatchObject({ code: 'OperationCancelled' });
  expect(reached).toBe(true); expect(seed.cells[0]).toMatchObject({ value: number(30), revision: 0 });
  const current = createWorkbookCalculator(session, {}, seed), learned = current.calculateSync();
  expect(learned.dependencies).toBeGreaterThan(seed.dependencies); expect(current.getCell(location(0, 4))).toMatchObject({ value: number(90) });
  set(4, 1, 90); expect(current.calculateSync().calculated).toBe(1); expect(current.getCell(location(0, 4))).toMatchObject({ value: number(130) });
  original.dispose(); cancelled.dispose(); current.dispose(); session.dispose();
});
test('reused dependency topology still honors tighter edge budgets and a later formula removal rebuilds its exact graph', () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(2) },
    { row: 0, column: 1, value: number(999), formula: '=A1+1' }, { row: 0, column: 2, value: number(999), formula: '=B1*2' }]));
  const original = createWorkbookCalculator(session), seed = original.calculateSync();
  const set = (column: number, source: { kind: 'value'; value: ReturnType<typeof number> } | { kind: 'formula'; expression: string }) => {
    const m = session.snapshot().manifest;
    session.commit(session.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision,
      operations: [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row: 0, column, source }] }] }));
  };
  set(0, { kind: 'value', value: number(3) }); const before = session.serialize();
  const bounded = createWorkbookCalculator(session, { maxDependencies: 1 }, seed);
  expect(() => bounded.calculateSync()).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  expect(bounded.getCell(location(0, 2)).state).toBe('stale'); expect(session.serialize()).toEqual(before);
  const current = createWorkbookCalculator(session, { maxDependencies: 2 }, seed), changed = current.calculateSync();
  expect(changed).toMatchObject({ dependencies: 2, calculated: 2 }); expect(current.getCell(location(0, 2))).toMatchObject({ value: number(8) });
  set(1, { kind: 'formula', expression: '=7' }); const smaller = createWorkbookCalculator(session, { maxDependencies: 1 }, changed);
  expect(smaller.calculateSync()).toMatchObject({ dependencies: 1, calculated: 2 }); expect(smaller.getCell(location(0, 2))).toMatchObject({ value: number(14) });
  original.dispose(); bounded.dispose(); current.dispose(); smaller.dispose(); session.dispose();
});
test("completed result projections remain immutable and bound to their original batch after later calculation and disposal", () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(2) }, { row: 0, column: 1, value: number(999), formula: '=A1+1' }]));
  const calculator = createWorkbookCalculator(session), first = calculator.calculateSync();
  session.commit(session.prepare({ protocolVersion: 1, documentId: 'scalar-kernel', baseRevision: 0, operations: [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row: 0, column: 0, source: { kind: 'value', value: number(9) } }] }] }));
  const next = calculator.calculateSync(); calculator.dispose(); session.dispose();
  expect(first.cells).toEqual([expect.objectContaining({ revision: 0, batchId: first.batchId, value: number(3) })]);
  expect(first.cells).toBe(first.cells); expect(next.cells[0]).toMatchObject({ revision: 1, batchId: next.batchId, value: number(10) });
  expect(Object.isFrozen(first.cells)).toBe(true); expect(Object.isFrozen(first.cells[0])).toBe(true);
  expect(Object.isFrozen(first.cells[0]!.state === 'valid' && first.cells[0]!.value)).toBe(true);
});
test("completed calculation seeds fork isolated candidate results without altering the live calculator or its budgets", async () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(2) },
    { row: 0, column: 1, value: number(999), formula: "=A1+1" }, { row: 1, column: 1, value: number(999), formula: "=4" }]));
  const live = createWorkbookCalculator(session), seed = live.calculateSync(), m = session.snapshot().manifest;
  const restored = createWorkbookCalculator(session, {}, seed);
  expect(restored.getCell(location(0, 1))).toMatchObject({ state: "valid", value: number(3), batchId: seed.batchId }); restored.dispose();
  const candidate = session.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: number(9) } }] }] });
  const fork = createWorkbookCalculator(candidate.model, {}, seed);
  expect(fork.calculateSync()).toMatchObject({ calculated: 1, reused: 1 }); expect(fork.getCell(location(0, 1))).toMatchObject({ value: number(10) });
  expect(live.getCell(location(0, 1))).toMatchObject({ value: number(3) }); expect(live.calculateSync()).toMatchObject({ calculated: 0, reused: 2 });
  const aborted = { aborted: false };
  await expect(fork.calculate({ signal: aborted, yield: async () => { aborted.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(fork.getCell(location(0, 1))).toMatchObject({ value: number(10) });
  const bounded = createWorkbookCalculator(candidate.model, { maxFormulas: 1 }, seed);
  expect(() => bounded.calculateSync()).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => createWorkbookCalculator(candidate.model, {}, JSON.parse(JSON.stringify(seed)))).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  bounded.dispose(); fork.dispose(); live.dispose(); session.discard(candidate); session.dispose();
});
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
