import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument, WorkbookDataValidationSchema } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { createWorkbookRangeCopyTransaction } from "../src/workbook/copy-range.js";
const sessions: WorkbookSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
function apply(s: WorkbookSession, operations: readonly DocumentOperation[]) { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [...operations] })); }
function fixture(invalid = false, errorStyle: "stop" | "warning" = "stop") {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 8 } }, { row: 2, column: 0, value: { type: "number", value: 4 } }, { row: 0, column: 1, value: { type: "number", value: 10 } }], metadata => {
    const sheet = metadata.workbook.sheets[0]!;
    sheet.conditionalRules = [{ ruleId: "source-cf", kind: "expression", ranges: [range(0, 0, 2, 0)], origin: { row: 0, column: 0 }, expression: "=AND(A1>0,$B$1=10)", priority: 1, stopIfTrue: true }, { ruleId: "old-cf", kind: "expression", ranges: [range(0, 5, 5, 5)], origin: { row: 0, column: 5 }, expression: "=TRUE()", priority: 2147483647, stopIfTrue: false }];
    const common = { allowBlank: true, showInputMessage: true, showErrorMessage: true, errorStyle, errorTitle: "校验", error: "保留禰󠄀😀", promptTitle: "输入", prompt: "提示\r\n尾" };
    sheet.dataValidations = [{ ...common, ruleId: "source-dv", kind: "custom", ranges: [range(0, 0, 2, 0)], origin: { row: 0, column: 0 }, expression: invalid ? "=A1>100" : "=A1>=0" }, { ...common, ruleId: "old-dv", kind: "custom", ranges: [range(0, 5, 5, 5)], origin: { row: 0, column: 5 }, expression: "=TRUE()" }];
  })); sessions.push(s); return s;
}
const sheet = (s: WorkbookSession, id = "sheet-1") => s.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === id)!;
const covers = (ranges: readonly { readonly from: { readonly row: number; readonly column: number }; readonly to: { readonly row: number; readonly column: number } }[], row: number, column: number) => ranges.some(range => range.from.row <= row && range.to.row >= row && range.from.column <= column && range.to.column >= column);
test("ordinary copy translates rule templates, clips surrounding definitions and keeps one undo and an equivalent repeat", async () => {
  const s = fixture(), before = s.serialize(), old = sheet(s);
  const tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 2, 0), { row: 1, column: 5 });
  expect(tx.operations.filter(op => op.op === "replaceWorkbookRules")).toHaveLength(1); expect(s.serialize()).toEqual(before); s.commit(s.prepare(tx));
  const next = sheet(s), cf = next.conditionalRules!.find(rule => rule.ruleId.startsWith("copied-rule"))!, dv = next.dataValidations!.find(rule => rule.ruleId.startsWith("copied-rule"))!;
  expect(cf).toMatchObject({ kind: "expression", expression: "=AND((F2>0),($B$1=10))", origin: { row: 1, column: 5 }, ranges: [range(1, 5, 3, 5)], stopIfTrue: true });
  expect(dv).toMatchObject({ expression: "=(F2>=0)", prompt: "提示\r\n尾", error: "保留禰󠄀😀" });
  expect(next.conditionalRules!.find(rule => rule.ruleId === "old-cf")).toMatchObject({ priority: 2147483647, ranges: [range(0, 5), range(4, 5, 5, 5)] });
  expect(next.dataValidations!.find(rule => rule.ruleId === "old-dv")!.ranges).toEqual([range(0, 5), range(4, 5, 5, 5)]);
  expect(new Set(next.conditionalRules!.map(rule => rule.priority)).size).toBe(next.conditionalRules!.length);
  expect(s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 2, 0), { row: 1, column: 5 })))).toMatchObject({ changed: false });
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(sheet(s)).toEqual(old); s.redo(); expect(sheet(s)).toEqual(next);
  const reopened = createDocumentSession(openDocument(s.serialize(), { kind: "workbook" })); sessions.push(reopened); expect(sheet(reopened)).toEqual(next);
});
test("skip blanks keeps destination rules at skipped coordinates and copies only the actual footprint", async () => {
  const s = fixture(); s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 2, 0), { row: 1, column: 5 }, { skipBlanks: true })));
  for (const key of ["conditionalRules", "dataValidations"] as const) {
    const old = sheet(s)[key]!.find(rule => rule.ruleId.startsWith("old-"))!, copied = sheet(s)[key]!.find(rule => rule.ruleId.startsWith("copied-"))!;
    expect(covers(old.ranges, 2, 5)).toBe(true); expect(covers(copied.ranges, 2, 5)).toBe(false);
    expect(copied.ranges).toEqual([range(1, 5), range(3, 5)]); expect(covers(old.ranges, 1, 5)).toBe(false);
  }
});
test.each(["values", "formulas", "formats"] as const)("%s mode retains the appropriate destination rule categories", async mode => {
  const s = fixture(), before = sheet(s); s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0), { row: 1, column: 5 }, { mode })));
  expect(sheet(s).dataValidations).toEqual(before.dataValidations);
  if (mode === "formats") { expect(sheet(s).conditionalRules).not.toEqual(before.conditionalRules); expect(s.getCell("sheet-1", 1, 5)).toBeUndefined(); }
  else expect(sheet(s).conditionalRules).toEqual(before.conditionalRules);
});
test("cross-sheet rules copy local name dependencies even without source cell formulas", async () => {
  const s = fixture(), original = WorkbookDataValidationSchema.parse(sheet(s).dataValidations![0]!); if (original.kind !== "custom") throw new Error("Expected custom source validation");
  apply(s, [{ op: "setWorkbookDefinedName", name: { nameId: "limit", name: "Limit", sheetId: "sheet-1", origin: { row: 0, column: 0 }, expression: "=Sheet1!$B$1", hidden: false } }, { op: "setWorkbookDataValidation", sheetId: "sheet-1", rule: { ...original, expression: "=A1<Limit" } }, { op: "insertWorkbookSheet", name: "Other" }]);
  const targetSheetId = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
  apply(s, [{ op: "setWorkbookCells", sheetId: targetSheetId, cells: [{ row: 0, column: 1, source: { kind: "value", value: { type: "number", value: 10 } } }] }]); s.clearHistory();
  s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(2, 0), { row: 256, column: 32 }, { targetSheetId })));
  expect(sheet(s, targetSheetId).dataValidations![0]).toMatchObject({ kind: "custom", expression: "=(AG257<Limit)", origin: { row: 256, column: 32 } });
  expect(s.snapshot().metadata.workbook.names!.find(name => name.sheetId === targetSheetId)).toMatchObject({ name: "Limit", expression: "='Other'!$B$1" });
  expect(sheet(s, targetSheetId).conditionalRules![0]).toMatchObject({ expression: "=AND((AG257>0),($B$1=10))" });
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(sheet(s, targetSheetId).dataValidations).toBeUndefined(); expect(s.snapshot().metadata.workbook.names).toHaveLength(1);
});
test("overlapping rule copy reads original templates and blank sources clear target rules", async () => {
  const s = fixture(); s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 2, 0), { row: 1, column: 0 })));
  expect(sheet(s).dataValidations!.find(rule => covers(rule.ranges, 3, 0))).toMatchObject({ expression: "=(A2>=0)", ranges: [range(1, 0, 3, 0)] });
  s.undo(); s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(20, 0), { row: 1, column: 5 })));
  expect(sheet(s).conditionalRules!.some(rule => covers(rule.ranges, 1, 5))).toBe(false); expect(sheet(s).dataValidations!.some(rule => covers(rule.ranges, 1, 5))).toBe(false);
});
test("copied stop validation, budgets and cancellation cannot partially publish content or metadata", async () => {
  const s = fixture(true), before = s.serialize();
  const tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0), { row: 1, column: 5 }); expect(() => s.prepare(tx)).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before);
  await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0), { row: 1, column: 5 }, { maxSteps: 1 })).rejects.toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  const abort = new AbortController(); await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 299, 0), { row: 1, column: 5 }, { signal: abort.signal, yield: async () => { abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" })); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
test("copied warnings require the new rule identity and explicit acceptance remains one atomic undo", async () => {
  const s = fixture(true, "warning"), before = s.serialize(), tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0), { row: 1, column: 5 });
  const copied = tx.operations.find(operation => operation.op === "replaceWorkbookRules"); if (copied?.op !== "replaceWorkbookRules") throw Error("Missing copied rule batch");
  const ruleId = copied.dataValidations![0]!.ruleId;
  expect(() => s.prepare(tx, { validation: { acceptedRuleIds: ["source-dv"] } })).toThrowError(expect.objectContaining({ code: "InvalidOperation", details: [expect.objectContaining({ code: "validation-confirmation-required", path: ["sheet-1", 1, 5, ruleId] })] }));
  expect(s.serialize()).toEqual(before); s.commit(s.prepare(tx, { validation: { acceptedRuleIds: [ruleId] } })); expect(s.getCell("sheet-1", 1, 5)?.value).toEqual({ type: "number", value: 8 }); expect(s.history().undoSteps).toBe(1); s.undo(); expect(sheet(s).dataValidations!.map(rule => rule.ruleId)).toEqual(["source-dv", "old-dv"]);
});
