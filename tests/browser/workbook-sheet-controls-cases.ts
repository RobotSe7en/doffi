import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import type { DeepReadonly, WorkbookSheet } from "@doffi/core";
import { scalarWorkbook } from "../../packages/core/tests/workbook-scalar-fixture.js";

export function workbookCopyNameConflictSource() {
  return scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 3 } }, { row: 0, column: 1, value: { type: "number", value: 999 }, formula: "=[0]!Tax+Tax" }], metadata => {
    metadata.workbook.names = [{ nameId: "global-tax", name: "Tax", expression: "=Sheet1!$A$1", origin: { row: 0, column: 0 } },
      { nameId: "local-tax", name: "Tax", expression: "=10", sheetId: "sheet-1", origin: { row: 0, column: 0 } }];
  });
}

export async function exerciseWorkbookCopyNameConflict(page: Page, remote: boolean): Promise<void> {
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  await page.getByText("工作表管理", { exact: true }).click(); await page.getByLabel("工作表操作", { exact: true }).selectOption("copy"); await page.getByLabel("工作表名称", { exact: true }).fill("Copied");
  const execute = page.getByRole("button", { name: "执行工作表操作", exact: true }); await execute.click(); await settled();
  expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, pending: true, sheets: [expect.objectContaining({ name: "Sheet1" })], errors: ["InvalidOperation"] });
  const conflict = page.getByLabel("复制名称冲突：Tax", { exact: true }); await expect(conflict).toBeVisible(); await conflict.fill("CopiedTax"); await execute.click(); await settled();
  expect(await inspect()).toMatchObject({ history: { undoSteps: 1 }, pending: false, sheets: [expect.objectContaining({ name: "Sheet1" }), expect.objectContaining({ name: "Copied" })], names: expect.arrayContaining([expect.objectContaining({ name: "CopiedTax" })]) });
  const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }); await address.fill("B1"); await address.press("Enter"); await settled();
  expect(await page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCell(0, 1) : window.workbookEditor.displayedCell(0, 1), remote)).toBe("13");
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, names: [expect.objectContaining({ nameId: "global-tax" }), expect.objectContaining({ nameId: "local-tax" })], sheets: [expect.objectContaining({ name: "Sheet1" })] });
}

export async function exerciseWorkbookSheetControls(page: Page, remote: boolean, original: readonly DeepReadonly<WorkbookSheet>[]): Promise<void> {
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { sheets: WorkbookSheet[]; history: { undoSteps: number }; selection: { anchor: { sheetId: string } }; pending: boolean };
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const execute = async () => { await page.getByRole("button", { name: "执行工作表操作", exact: true }).click(); await settled(); };
  const undo = async () => { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); };
  const sheet = page.getByLabel("管理工作表", { exact: true }), action = page.getByLabel("工作表操作", { exact: true }), name = page.getByLabel("工作表名称", { exact: true }), visibility = page.getByLabel("工作表可见性", { exact: true });
  const position = page.getByLabel("放在工作表之前", { exact: true }), save = page.getByRole("button", { name: "执行工作表操作", exact: true }), reload = page.getByRole("button", { name: "放弃当前工作表草稿并重新读取", exact: true });
  await page.getByText("工作表管理", { exact: true }).click(); await action.selectOption("copy"); await name.fill("Copied"); await position.selectOption(original[1]!.sheetId); expect(await inspect()).toMatchObject({ pending: true });
  await action.selectOption("properties"); await expect(name).toHaveValue(original[0]!.name); await action.selectOption("copy"); await expect(name).toHaveValue("Copied");
  await execute(); const copied = (await inspect()).sheets[1]!;
  expect((await inspect()).sheets.map(sheet => sheet.name)).toEqual([original[0]!.name, "Copied", original[1]!.name]); expect(await inspect()).toMatchObject({ pending: false, history: { undoSteps: 1 }, selection: { anchor: { sheetId: copied.sheetId } } });
  const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }), input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
  for (const [at, value] of [["A5", "99"], ["A6", "123"], ["H1", "88"]]) { await address.fill(at!); await address.press("Enter"); await settled(); await input.fill(value!); await input.press("Enter"); await settled(); }
  await address.fill("B5:B8"); await address.press("Enter"); await settled();
  expect(await page.evaluate(remote => [4, 5, 6, 7].map(row => remote ? window.workbookRemoteEditor.displayedCell(row, 1) : window.workbookEditor.displayedCell(row, 1)), remote)).toEqual(["187", "345", "88", "88"]);
  for (let index = 0; index < 4; index++) await undo(); expect((await inspect()).sheets).toEqual(original);
  await sheet.selectOption(original[1]!.sheetId); await action.selectOption("properties"); await name.fill("OtherPrime"); await visibility.selectOption("hidden"); await execute();
  expect((await inspect()).sheets[1]).toMatchObject({ sheetId: original[1]!.sheetId, name: "OtherPrime", visibility: "hidden" }); await expect(page.getByRole("tab", { name: "OtherPrime", exact: true })).toHaveCount(0);
  await visibility.selectOption("veryHidden"); await execute(); expect((await inspect()).sheets[1]!.visibility).toBe("veryHidden");
  await visibility.selectOption("visible"); await execute(); await expect(page.getByRole("tab", { name: "OtherPrime", exact: true })).toBeVisible();
  await action.selectOption("move"); await position.selectOption(original[0]!.sheetId); await execute(); expect((await inspect()).sheets[0]!.sheetId).toBe(original[1]!.sheetId);
  await action.selectOption("remove"); await execute(); expect((await inspect()).sheets).toHaveLength(1); for (let index = 0; index < 5; index++) await undo(); expect((await inspect()).sheets).toEqual(original);
  await sheet.selectOption(original[0]!.sheetId); await action.selectOption("insert"); await name.fill("NewSheet"); await position.selectOption(original[0]!.sheetId); await execute(); expect((await inspect()).sheets[0]!.name).toBe("NewSheet"); await undo();
  await sheet.selectOption(original[0]!.sheetId); await action.selectOption("properties");
  await name.evaluate(element => { const field = element as HTMLInputElement; field.focus(); field.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true })); field.value = "中文草稿"; field.dispatchEvent(new InputEvent("input", { bubbles: true, isComposing: true })); });
  await expect(save).toBeDisabled(); await expect(sheet).toBeDisabled(); await expect(name).toBeEnabled(); await name.dispatchEvent("compositionend");
  await page.evaluate(async remote => { if (remote) await window.workbookRemoteEditor.external("external"); else { window.workbookEditor.external("external"); await window.workbookEditor.settled(); } }, remote);
  await expect(name).toHaveValue("中文草稿"); await expect(save).toBeDisabled(); await expect(page.getByLabel("工作表编辑状态", { exact: true })).toContainText("草稿已保留");
  await action.selectOption("copy"); expect(await inspect()).toMatchObject({ pending: true }); await action.selectOption("properties"); await expect(name).toHaveValue("中文草稿"); await reload.click(); await settled(); expect(await inspect()).toMatchObject({ pending: false });
}
