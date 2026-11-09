import { afterEach, expect, test } from "vitest";
import { createDocumentSession } from "../src/index.js";
import type { DocumentOperation } from "../src/index.js";
import type { WorkbookInteractionRequestInput, WorkbookModel, WorkbookSession } from "../src/runtime.js";
import type { WorkbookInteractionRequest } from "../src/contracts/workbook-interaction.js";
import { buildWorkbookRequestSteps } from "../src/session/workbook-interaction-context.js";
import { captureWorkbookModel } from "../src/workbook/source.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sessions: WorkbookSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const number = (value: number) => ({ type: "number" as const, value });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
function apply(session: WorkbookSession, operations: DocumentOperation[]): void { const manifest = session.snapshot().manifest; session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations })); }
function setup(model: WorkbookModel, selected = range(0, 0)) {
  const session = createDocumentSession(model); sessions.push(session);
  const manifest = session.snapshot().manifest;
  apply(session, [{ op: "addAnnotation", annotation: { annotationId: "request", documentId: manifest.documentId, baseRevision: manifest.revision, status: "open", instruction: "复核当前值😀禰󠄀", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [selected] } } }]);
  session.clearHistory(); return session;
}
function build(session: WorkbookSession, options: Partial<WorkbookInteractionRequestInput> = {}): WorkbookInteractionRequest {
  const steps = buildWorkbookRequestSteps(captureWorkbookModel(session), { kind: "workbook", annotationIds: ["request"], ...options });
  for (;;) { const result = steps.next(); if (result.done) return result.value; }
}

test("context keeps exact historical source while current formulas and complete transitive dependencies are read-only", () => {
  const session = setup(scalarWorkbook([
    { row: 0, column: 0, value: number(999), formula: "=IF(1,A2,A3)" }, { row: 1, column: 0, value: number(7) },
    { row: 2, column: 0, value: number(888), formula: "=A257*2" }, { row: 256, column: 0, value: number(12) },
    { row: 9, column: 7, value: { type: "text", value: "unrelated private cell" } },
  ])), before = session.serialize(), request = build(session);
  expect(request.context.cells.map(cell => [cell.row, cell.column, cell.access])).toEqual([[0, 0, "editable"], [1, 0, "read-only"], [2, 0, "read-only"], [256, 0, "read-only"]]);
  expect(request.context.cells[0]).toMatchObject({ value: number(999), html: session.getCell("sheet-1", 0, 0)!.html, current: { kind: "value", value: number(7) } });
  expect(request.context.cells[2]).toMatchObject({ value: number(888), current: { kind: "value", value: number(24) } });
  expect(JSON.stringify(request)).not.toContain("unrelated private cell"); expect(Object.isFrozen(request.context.cells[0]!.current)).toBe(true);
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
});

test("cross-sheet names include only the source definitions and worksheets used by the frozen request", () => {
  const session = setup(scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=Rate+A257" }, { row: 256, column: 0, value: number(35) }]));
  apply(session, [{ op: "insertWorkbookSheet", name: "Rates" }, { op: "insertWorkbookSheet", name: "Unrelated" }]);
  const rates = session.snapshot().metadata.workbook.sheets[1]!.sheetId, unrelated = session.snapshot().metadata.workbook.sheets[2]!.sheetId;
  apply(session, [{ op: "setWorkbookCells", sheetId: rates, cells: [{ row: 0, column: 0, source: { kind: "value", value: number(7) } }] },
    { op: "setWorkbookCells", sheetId: unrelated, cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "unrelated secret" } } }] },
    { op: "setWorkbookDefinedName", name: { nameId: "rate", name: "Rate", expression: "=Rates!$A$1", origin: { row: 0, column: 0 } } },
    { op: "setWorkbookDefinedName", name: { nameId: "unused", name: "Unused", expression: '= "unrelated name literal"', origin: { row: 0, column: 0 } } }]);
  const request = build(session);
  expect(request.context.cells.find(cell => cell.access === "editable")!.current).toEqual({ kind: "value", value: number(42) });
  expect(request.context.metadata.workbook.sheets.map(sheet => sheet.name)).toEqual(["Sheet1", "Rates"]);
  expect(request.context.metadata.workbook.names?.map(name => name.nameId)).toEqual(["rate"]);
  expect(request.context.cells.filter(cell => cell.sheetId === rates)).toMatchObject([{ row: 0, column: 0, access: "read-only", current: { value: number(7) } }]);
  expect(JSON.stringify(request)).not.toMatch(/unrelated secret|unrelated name literal/);
});

test("a sparse whole-column dependency declares blanks without expanding a million coordinates", () => {
  const session = setup(scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=SUM(B:B)" }, { row: 256, column: 1, value: number(12) }, { row: 1_048_575, column: 1, value: number(30) }]));
  const request = build(session, { limits: { maxCells: 3, maxNodes: 6, maxBytes: 65536 } });
  expect(request.context.cells).toHaveLength(3); expect(request.context.dependencies).toEqual([{ sheetId: "sheet-1", range: range(0, 1, 1_048_575, 1) }]);
  expect(request.context.cells[0]!.current).toEqual({ kind: "value", value: number(42) });
});

test("one copied time and locale context governs all derived values without reading a clock", () => {
  const session = setup(scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=NOW()" }]));
  const clock = { unixMilliseconds: 1_700_000_000_000, utcOffsetMinutes: 480 }, steps = buildWorkbookRequestSteps(captureWorkbookModel(session), { kind: "workbook", annotationIds: ["request"], calculationContext: { now: clock, locale: "zh-CN" } });
  expect(steps.next().done).toBe(false); clock.unixMilliseconds += 86_400_000;
  let request!: WorkbookInteractionRequest; for (;;) { const result = steps.next(); if (result.done) { request = result.value; break; } }
  expect(request.context.calculationContext.now!.unixMilliseconds).toBe(1_700_000_000_000); expect(Object.isFrozen(request.context.calculationContext.now)).toBe(true);
  expect(request.context.cells[0]!.current.kind).toBe("value"); expect(build(session).context.cells[0]!.current).toMatchObject({ kind: "unavailable", reason: "calculation-required" });
});

test("reverse consumers are read-only and their transitive inputs are disclosed without unrelated formula chains", () => {
  const session = setup(scalarWorkbook([
    { row: 0, column: 0, value: number(4) },
    { row: 0, column: 1, value: number(900), formula: "=A1+D1" },
    { row: 0, column: 2, value: number(901), formula: "=B1*2" },
    { row: 0, column: 3, value: number(7) },
    { row: 8, column: 7, value: number(100), formula: "=I9+1" },
    { row: 8, column: 8, value: { type: "text", value: "unrelated consumer input" } },
  ]));
  const request = build(session);
  expect(request.context.cells.map(cell => [cell.column, cell.access, cell.current])).toEqual([
    [0, "editable", { kind: "value", value: number(4) }], [1, "read-only", { kind: "value", value: number(11) }],
    [2, "read-only", { kind: "value", value: number(22) }], [3, "read-only", { kind: "value", value: number(7) }],
  ]);
  expect(JSON.stringify(request)).not.toContain("unrelated consumer input");
});

test("dynamic reverse references bind against current source and disclose their parameter cells", () => {
  const session = setup(scalarWorkbook([
    { row: 0, column: 0, value: number(8) },
    { row: 0, column: 1, value: number(999), formula: '=INDEX(A1:A2,D1,1)*2' },
    { row: 0, column: 3, value: number(1) },
    { row: 1, column: 0, value: number(10) },
  ]));
  const request = build(session);
  expect(request.context.cells.find(cell => cell.column === 1)!.current).toEqual({ kind: "value", value: number(16) });
  expect(request.context.cells.find(cell => cell.column === 3)).toMatchObject({ access: "read-only", value: number(1) });
  expect(request.context.cells.some(cell => cell.row === 1 && cell.column === 0)).toBe(true);
});

test("reverse references inside a defined name retain only the used name and its context", () => {
  const session = setup(scalarWorkbook([{ row: 0, column: 0, value: number(6) }, { row: 0, column: 1, value: number(999), formula: "=Chosen*2" }]));
  apply(session, [{ op: "setWorkbookDefinedName", name: { nameId: "chosen", name: "Chosen", expression: '=INDEX($A$1:$A$2,1,1)', origin: { row: 0, column: 0 } } }]);
  const request = build(session);
  expect(request.context.cells.find(cell => cell.column === 1)!.current).toEqual({ kind: "value", value: number(12) });
  expect(request.context.metadata.workbook.names?.map(name => name.nameId)).toEqual(["chosen"]);
});

test.each([range(0, 0), range(0, 1)])("nested names in unions retain exact native bindings in forward and reverse request context: %j", selected => {
  const session = setup(scalarWorkbook([
    { row: 0, column: 0, value: number(999), formula: "=SUM((Alias,$A$2))" }, { row: 0, column: 1, value: number(2) },
    { row: 1, column: 0, value: number(3) }, { row: 0, column: 2, value: { type: "text", value: "unused local private value" } },
  ], metadata => { metadata.workbook.names = [
    { nameId: "inner-global", name: "Inner", expression: "=Sheet1!$B$1", origin: { row: 0, column: 0 } },
    { nameId: "inner-local", name: "Inner", expression: "=Sheet1!$C$1", sheetId: "sheet-1", origin: { row: 0, column: 0 } },
    { nameId: "alias", name: "Alias", expression: "=Inner", origin: { row: 0, column: 0 } },
  ]; }), selected);
  const before = session.serialize(), request = build(session);
  expect(request.context.metadata.workbook.names?.map(name => name.nameId).sort()).toEqual(["alias", "inner-global"]);
  expect(request.context.cells.find(cell => cell.row === 0 && cell.column === 0)?.current).toEqual({ kind: "value", value: number(5) });
  expect(JSON.stringify(request)).not.toContain("unused local private value");
  expect(request.context.cells.some(cell => cell.column === 2)).toBe(false); expect(session.serialize()).toEqual(before);
});

test("outside validation targets and their relative thresholds are read-only dependencies of a threshold edit", () => {
  const session = setup(scalarWorkbook([{ row: 0, column: 0, value: number(4) }, { row: 0, column: 1, value: number(5) }], metadata => {
    metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "limit", ranges: [range(0, 0, 2, 0)], origin: { row: 0, column: 0 }, kind: "custom", expression: "=A1<=B1", allowBlank: true,
      showErrorMessage: true, showInputMessage: true, errorStyle: "stop", errorTitle: "", error: "too large", promptTitle: "", prompt: "" }];
  }), range(0, 1));
  const request = build(session);
  expect(request.context.cells.map(cell => [cell.row, cell.column, cell.access])).toEqual([[0, 0, "read-only"], [0, 1, "editable"]]);
  expect(request.context.metadata.workbook.sheets[0]!.dataValidations?.map(rule => rule.ruleId)).toEqual(["limit"]);
});

test("relative rules disclose dependencies for implicit blank coordinates without sampling the interior", () => {
  const session = setup(scalarWorkbook([{ row: 0, column: 1, value: number(5) }, { row: 1, column: 1, value: number(7) }, { row: 2, column: 1, value: number(9) }], metadata => {
    metadata.workbook.sheets[0]!.conditionalRules = [{ ruleId: "interior", ranges: [range(0, 0, 2, 0)], origin: { row: 0, column: 0 }, kind: "expression", expression: "=B1>0", priority: 1, stopIfTrue: false, cssClass: "cell-default" }];
  }), range(0, 0, 2, 0));
  const request = build(session);
  expect(request.context.cells.map(cell => [cell.row, cell.column, cell.access])).toEqual([[0, 1, "read-only"], [1, 1, "read-only"], [2, 1, "read-only"]]);
  expect(request.context.dependencies.some(entry => entry.range.from.row === 1 && entry.range.from.column === 1)).toBe(true);
});

test("a whole-column formatting rule yields while scanning implicit targets", () => {
  const session = setup(scalarWorkbook([{ row: 0, column: 1, value: number(5) }], metadata => {
    metadata.workbook.sheets[0]!.conditionalRules = [{ ruleId: "full-column", ranges: [range(0, 0, 1_048_575, 0)], origin: { row: 0, column: 0 }, kind: "expression", expression: "=A1<B1", priority: 1, stopIfTrue: false, cssClass: "cell-default" }];
  }), range(0, 1)), signal = { aborted: false }, before = session.serialize();
  const steps = buildWorkbookRequestSteps(captureWorkbookModel(session), { kind: "workbook", annotationIds: ["request"] }, { signal });
  for (;;) { const next = steps.next(); if (next.done) throw Error("Rule scan never yielded"); if (next.value.phase === "validation") break; }
  signal.aborted = true; expect(() => steps.next()).toThrowError(expect.objectContaining({ code: "OperationCancelled" })); expect(session.serialize()).toEqual(before);
});

test("byte, cell, node, dependency and cumulative evaluation limits fail without a partial request or source change", () => {
  const session = setup(scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=A2+A3" }, { row: 1, column: 0, value: number(2) }, { row: 2, column: 0, value: number(3) }])), before = session.serialize();
  for (const limits of [{ maxBytes: 100 }, { maxCells: 2 }, { maxNodes: 1 }, { maxDependencies: 1 }, { maxSteps: 4 }]) {
    expect(() => build(session, { limits })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  }
});

test("context cancellation after an actual 128-cell checkpoint leaves history and annotations unchanged", () => {
  const session = setup(scalarWorkbook(Array.from({ length: 300 }, (_, row) => ({ row, column: 0, value: number(row) }))), range(0, 0, 299, 0)), before = session.serialize(), signal = { aborted: false };
  const steps = buildWorkbookRequestSteps(captureWorkbookModel(session), { kind: "workbook", annotationIds: ["request"] }, { signal });
  for (;;) { const result = steps.next(); if (result.done) throw Error("No context checkpoint"); if (result.value.phase === "cells") break; }
  signal.aborted = true; expect(() => steps.next()).toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  expect(session.serialize()).toEqual(before); expect(session.getAnnotations()[0]!.status).toBe("open"); expect(session.history().undoSteps).toBe(0);
});
