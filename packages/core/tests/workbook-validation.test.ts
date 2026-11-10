import { expect, test } from "vitest";
import { createDocumentSession } from "../src/index.js";
import type { WorkbookDataValidation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const range = { from: { row: 0, column: 0 }, to: { row: 9, column: 0 } };
const base = { ruleId: "rule", ranges: [range], origin: range.from, allowBlank: true, showErrorMessage: true, showInputMessage: true, errorStyle: "stop" as const, errorTitle: "范围", error: "必须满足范围", promptTitle: "", prompt: "" };
const value = (row: number, column: number, amount: number) => ({ op: "setWorkbookCells" as const, sheetId: "sheet-1", cells: [{ row, column, source: { kind: "value" as const, value: { type: "number" as const, value: amount } } }] });
function session(rule: WorkbookDataValidation): WorkbookSession {
  return createDocumentSession(scalarWorkbook([{ row: 0, column: 1, value: { type: "number", value: 5 } }], metadata => { metadata.workbook.sheets[0]!.dataValidations = [rule]; }));
}
function transaction(s: WorkbookSession, operations: Parameters<WorkbookSession["prepare"]>[0]["operations"]) {
  return { protocolVersion: 1 as const, documentId: "scalar-kernel", baseRevision: s.snapshot().manifest.revision, operations };
}
test("validates the whole final candidate, and a stopped value leaves no preceding edits", () => {
  const s = session({ ...base, kind: "whole", operator: "between", formulas: ["=1", "=$B$1"] }), original = s.serialize();
  expect(() => s.prepare(transaction(s, [value(1, 1, 10), value(0, 0, 9)]))).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(original); expect(s.history().undoSteps).toBe(0);
  s.commit(s.prepare(transaction(s, [value(0, 0, 9), value(0, 1, 10)])));
  expect(s.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 9 }); expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.serialize().size).toBe(original.size);
});
test("unsupported stop formulas are never treated as a valid cached answer", () => {
  const s = session({ ...base, kind: "custom", expression: "=FUTUREFUNC(A1)" }), before = s.serialize();
  expect(() => s.prepare(transaction(s, [value(0, 0, 1)]), { validation: { acceptedRuleIds: ["rule"] } })).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(s.getCell("sheet-1", 0, 0)).toBeUndefined();
});
test("warning/information acceptance is explicit, while stop rules remain blocking", () => {
  const s = session({ ...base, errorStyle: "warning", kind: "decimal", operator: "lessThan", formulas: ["=5"] }), input = transaction(s, [value(0, 0, 9)]);
  expect(() => s.prepare(input)).toThrowError(expect.objectContaining({ code: "InvalidOperation", details: [expect.objectContaining({ code: "validation-confirmation-required" })] }));
  s.commit(s.prepare(input, { validation: { acceptedRuleIds: ["rule"] } })); expect(s.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 9 });
});
test("list rules bind the current final source and retain unavailable-member diagnostics", () => {
  const s = session({ ...base, kind: "list", expression: "=$B$1:$B$2", dropdown: true });
  s.commit(s.prepare(transaction(s, [value(0, 0, 10), value(1, 1, 10)])));
  expect(() => s.prepare(transaction(s, [value(0, 0, 99)]))).toThrow(); expect(s.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 10 });
});
test("edits to thresholds and calculated target values validate unchanged source targets", () => {
  const rule: WorkbookDataValidation = { ...base, kind: "whole", operator: "between", formulas: ["=1", "=$B$1"] };
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 4 } }, { row: 0, column: 1, value: { type: "number", value: 5 } }], metadata => { metadata.workbook.sheets[0]!.dataValidations = [rule]; })), before = s.serialize();
  expect(() => s.prepare(transaction(s, [value(0, 1, 3)]))).toThrowError(expect.objectContaining({ details: [expect.objectContaining({ path: ["sheet-1", 0, 0, "rule"] })] }));
  expect(s.serialize()).toEqual(before);
  const calculated = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 999 }, formula: "=B1" }, { row: 0, column: 1, value: { type: "number", value: 4 } }], metadata => { metadata.workbook.sheets[0]!.dataValidations = [{ ...base, kind: "whole", operator: "lessThan", formulas: ["=5"] }]; }));
  expect(() => calculated.prepare(transaction(calculated, [value(0, 1, 9)]))).toThrow(); expect(calculated.getCell("sheet-1", 0, 1)?.value).toEqual({ type: "number", value: 4 });
});
test("validation follows computed names and relative multi-row rule dependencies", () => {
  const s = createDocumentSession(scalarWorkbook([
    { row: 0, column: 0, value: { type: "number", value: 3 } }, { row: 1, column: 0, value: { type: "number", value: 4 } },
    { row: 0, column: 1, value: { type: "number", value: 5 } }, { row: 1, column: 1, value: { type: "number", value: 5 } },
  ], metadata => {
    metadata.workbook.names = [{ nameId: "limit", name: "Limit", expression: "=Sheet1!$B$1+0", origin: { row: 0, column: 0 } }];
    metadata.workbook.sheets[0]!.dataValidations = [{ ...base, kind: "custom", expression: "=AND(A1<=B1,A1<=Limit)" }];
  }));
  expect(() => s.prepare(transaction(s, [value(1, 1, 2)]))).toThrowError(expect.objectContaining({ details: [expect.objectContaining({ path: ["sheet-1", 1, 0, "rule"] })] }));
  expect(() => s.prepare(transaction(s, [value(0, 1, 2)]))).toThrow();
  s.commit(s.prepare(transaction(s, [value(1, 1, 8)]))); expect(s.getCell("sheet-1", 1, 1)?.value).toEqual({ type: "number", value: 8 });
});
test("changing a list member checks existing choices without editing those cells", () => {
  const s = createDocumentSession(scalarWorkbook([
    { row: 0, column: 0, value: { type: "text", value: "yes" } }, { row: 0, column: 1, value: { type: "text", value: "yes" } },
  ], metadata => { metadata.workbook.sheets[0]!.dataValidations = [{ ...base, kind: "list", expression: "=$B$1:$B$2", dropdown: true }]; }));
  expect(() => s.prepare(transaction(s, [value(0, 1, 1)]))).toThrow(); expect(s.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "yes" });
});
test("unrelated edits and formatting preserve pre-existing imported violations", () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 9 } }, { row: 0, column: 1, value: { type: "number", value: 5 } }], metadata => { metadata.workbook.sheets[0]!.dataValidations = [{ ...base, kind: "decimal", operator: "lessThan", formulas: ["=$B$1"] }]; }));
  s.commit(s.prepare(transaction(s, [value(4, 4, 1)])));
  s.commit(s.prepare(transaction(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, cssClass: "cell-default", source: { kind: "value", value: { type: "number", value: 9 } } }] }])));
  expect(() => s.prepare(transaction(s, [value(0, 1, 4)]))).toThrow(); expect(s.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 9 });
});
test("non-blank-permitting rules include implicit targets affected by a dependency edit", () => {
  const s = session({ ...base, ranges: [{ from: { row: 0, column: 0 }, to: { row: 2, column: 0 } }], allowBlank: false, kind: "custom", expression: "=A1<$B$1" });
  expect(() => s.prepare(transaction(s, [value(0, 1, -1)]))).toThrow(); expect(s.getCell("sheet-1", 0, 0)).toBeUndefined();
});
