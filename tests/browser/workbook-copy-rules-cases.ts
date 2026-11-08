import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession } from "@doffi/core";
import type { DeepReadonly, WorkbookModel, WorkbookRange, WorkbookSheet } from "@doffi/core";
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
export function workbookCopyRulesSource(model: WorkbookModel) {
  const session = createDocumentSession(model), manifest = session.snapshot().manifest;
  session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [
    { op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 1, column: 16, source: { kind: "value", value: { type: "number", value: 8 } } }, { row: 3, column: 16, source: { kind: "value", value: { type: "number", value: 4 } } }] },
    { op: "setWorkbookConditionalRule", sheetId: "sheet-1", rule: { ruleId: "source-copy-cf", kind: "expression", ranges: [range(1, 16, 3, 16)], origin: { row: 1, column: 16 }, expression: "=Q2>5", priority: 100, stopIfTrue: true } },
    { op: "formatWorkbookConditionalRule", sheetId: "sheet-1", ruleId: "source-copy-cf", properties: { "background-color": "#112233" } },
    { op: "setWorkbookConditionalRule", sheetId: "sheet-1", rule: { ruleId: "target-copy-cf", kind: "expression", ranges: [range(0, 5, 5, 5)], origin: { row: 0, column: 5 }, expression: "=TRUE()", priority: 101, stopIfTrue: false } },
    { op: "setWorkbookDataValidation", sheetId: "sheet-1", rule: { ruleId: "source-copy-dv", kind: "custom", ranges: [range(1, 16, 3, 16)], origin: { row: 1, column: 16 }, expression: "=Q2>=0", allowBlank: true, showInputMessage: true, showErrorMessage: true, errorStyle: "stop", errorTitle: "验证", error: "保留禰󠄀😀", promptTitle: "输入", prompt: "提示\r\n尾" } },
    { op: "setWorkbookDataValidation", sheetId: "sheet-1", rule: { ruleId: "target-copy-dv", kind: "custom", ranges: [range(0, 5, 5, 5)], origin: { row: 0, column: 5 }, expression: "=TRUE()", allowBlank: true, showInputMessage: false, showErrorMessage: true, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" } },
    { op: "insertWorkbookSheet", name: "Other" },
  ] }));
  const source = session.serialize(), sheets = session.snapshot().metadata.workbook.sheets; session.dispose(); return { source, sheets };
}
export async function exerciseWorkbookCopyRules(page: Page, remote: boolean, original: readonly DeepReadonly<WorkbookSheet>[]): Promise<void> {
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { sheets: DeepReadonly<WorkbookSheet>[]; history: { undoSteps: number }; errors: string[] };
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const button = (name: string) => page.getByRole("button", { name, exact: true }), skip = page.getByLabel("区域粘贴时跳过空白", { exact: true });
  const select = async (range: string) => { const address = page.getByLabel("选择单元格或区域", { exact: true }); await address.fill(range); await address.press("Enter"); await settled(); };
  const covers = (ranges: readonly DeepReadonly<WorkbookRange>[], row: number, column: number) => ranges.some(range => range.from.row <= row && range.to.row >= row && range.from.column <= column && range.to.column >= column);
  const undo = async () => { await button("撤销").click(); await settled(); const current = await inspect(); expect(current.history.undoSteps).toBe(0); for (const key of ["conditionalRules", "dataValidations"] as const) expect(current.sheets[0]![key]).toEqual(original[0]![key]); };
  async function paste(mode = "contents", blanks = false) { await select("Q2:Q4"); await button("复制选区").click(); await page.getByLabel("区域粘贴方式", { exact: true }).selectOption(mode); await skip.setChecked(blanks); await select("F2"); await button("粘贴选区").click(); await settled(); }
  await paste(); let current = await inspect(); expect(current.history.undoSteps).toBe(1);
  const sourceRule = original[0]!.conditionalRules!.find(rule => rule.ruleId === "source-copy-cf")!;
  expect(current.sheets[0]!.conditionalRules!.find(rule => rule.ruleId.startsWith("copied-rule"))).toMatchObject({ kind: "expression", expression: "=(F2>5)", ...("cssClass" in sourceRule ? { cssClass: sourceRule.cssClass } : {}), ranges: [range(1, 5, 3, 5)] });
  expect(current.sheets[0]!.dataValidations!.find(rule => rule.ruleId.startsWith("copied-rule"))).toMatchObject({ kind: "custom", expression: "=(F2>=0)", prompt: "提示\r\n尾", error: "保留禰󠄀😀" });
  await expect.poll(() => page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCellAppearance(1, 5) : window.workbookEditor.displayedCellAppearance(1, 5), remote)).toMatchObject({ fill: "rgb(17, 34, 51)" }); await undo();
  await paste("contents", true); current = await inspect();
  for (const key of ["conditionalRules", "dataValidations"] as const) {
    const copied = current.sheets[0]![key]!.find(rule => rule.ruleId.startsWith("copied-rule"))!, old = current.sheets[0]![key]!.find(rule => rule.ruleId.startsWith("target-copy"))!;
    expect(copied.ranges).toEqual([range(1, 5), range(3, 5)]); expect(covers(old.ranges, 2, 5)).toBe(true); expect(covers(old.ranges, 1, 5)).toBe(false);
  }
  await undo(); await paste("formats"); current = await inspect(); expect(current.sheets[0]!.conditionalRules).not.toEqual(original[0]!.conditionalRules); expect(current.sheets[0]!.dataValidations).toEqual(original[0]!.dataValidations); await undo();
  await paste("values"); current = await inspect(); expect(current.sheets[0]!.conditionalRules).toEqual(original[0]!.conditionalRules); expect(current.sheets[0]!.dataValidations).toEqual(original[0]!.dataValidations); await undo();
  await select("Q2:Q4"); await button("复制选区").click(); await page.getByLabel("区域粘贴方式", { exact: true }).selectOption("contents"); await page.getByRole("tab", { name: "Other", exact: true }).click(); await settled(); await select("AG257"); await button("粘贴选区").click(); await settled();
  current = await inspect(); expect(current.sheets[1]!.conditionalRules![0]).toMatchObject({ expression: "=(AG257>5)", ranges: [range(256, 32, 258, 32)] }); expect(current.sheets[1]!.dataValidations![0]).toMatchObject({ expression: "=(AG257>=0)", prompt: "提示\r\n尾" }); expect(current.errors).toEqual([]); expect(current.history.undoSteps).toBe(1); await undo();
}
