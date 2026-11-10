import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import type { DeepReadonly, WorkbookDefinedName } from "@doffi/core";

export async function exerciseWorkbookNameControls(page: Page, remote: boolean, original: readonly DeepReadonly<WorkbookDefinedName>[], target: string): Promise<void> {
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const cell = (sheetId: string, row: number, column: number) => page.evaluate(({ remote, sheetId, row, column }) => remote ? window.workbookRemoteEditor.sourceCell(row, column, sheetId) : window.workbookEditor.cell(sheetId, row, column), { remote, sheetId, row, column });
  const names = page.getByLabel("选择定义名称", { exact: true }), name = page.getByLabel("定义名称", { exact: true }), scope = page.getByLabel("名称作用域", { exact: true }), expression = page.getByLabel("名称引用位置或公式", { exact: true });
  const save = page.getByRole("button", { name: "保存定义名称", exact: true }), reload = page.getByRole("button", { name: "放弃当前名称草稿并重新读取", exact: true });
  const global = original.find(name => name.name === "Tax" && name.sheetId === undefined)!, local = original.find(name => name.name === "Rate")!;
  await page.getByText("名称管理", { exact: true }).click(); await names.selectOption(global.nameId); await expect(scope).toBeDisabled();
  await name.fill("GrossTax"); expect(await inspect()).toMatchObject({ pending: true });
  await names.selectOption(local.nameId); await expect(name).toHaveValue("Rate"); await names.selectOption(global.nameId); await expect(name).toHaveValue("GrossTax");
  await save.click(); await settled();
  expect(await inspect()).toMatchObject({ pending: false, history: { undoSteps: 1 }, names: expect.arrayContaining([expect.objectContaining({ nameId: global.nameId, name: "GrossTax" }), expect.objectContaining({ name: "Tax", sheetId: target })]) });
  expect(await cell("sheet-1", 6, 1)).toMatchObject({ formula: { expression: "=GrossTax" } }); expect(await cell(target, 7, 1)).toMatchObject({ formula: { expression: "=[0]!GrossTax" } });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect(await inspect()).toMatchObject({ names: original, history: { undoSteps: 0 } });
  await names.selectOption(""); await expect(scope).toBeEnabled(); await name.fill("NewLocal"); await scope.selectOption(target); await expression.fill("=42");
  await page.getByLabel("名称相对引用基准单元格", { exact: true }).fill("B2"); await page.getByLabel("名称说明", { exact: true }).fill("中文禰󠄀😀"); await page.getByLabel("隐藏名称", { exact: true }).check();
  await save.click(); await settled(); expect(await inspect()).toMatchObject({ pending: false, history: { undoSteps: 1 }, names: expect.arrayContaining([expect.objectContaining({ name: "NewLocal", sheetId: target, expression: "=42", origin: { row: 1, column: 1 }, hidden: true, description: "中文禰󠄀😀" })]) });
  await page.getByRole("button", { name: "删除定义名称", exact: true }).click(); await settled(); expect(await inspect()).toMatchObject({ names: original, history: { undoSteps: 2 } });
  for (let index = 0; index < 2; index++) { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); }
  expect(await inspect()).toMatchObject({ names: original, history: { undoSteps: 0 } });
  await names.selectOption(global.nameId);
  // Native fill focuses and replaces text, which ends composition in Firefox.
  // Focus first, then deliver the same ordered events as an active IME draft.
  await name.evaluate(element => {
    const field = element as HTMLInputElement; field.focus(); field.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true })); field.value = "草稿";
    field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertCompositionText", data: "草稿", isComposing: true }));
  });
  await expect(save).toBeDisabled(); await expect(names).toBeDisabled(); await expect(name).toBeEnabled();
  await name.dispatchEvent("compositionend"); await name.fill("KeepDraft");
  await page.evaluate(async remote => { if (remote) await window.workbookRemoteEditor.external("external"); else { window.workbookEditor.external("external"); await window.workbookEditor.settled(); } }, remote);
  await expect(save).toBeDisabled(); await expect(name).toHaveValue("KeepDraft"); await expect(page.getByLabel("名称编辑状态", { exact: true })).toContainText("草稿已保留");
  await names.selectOption(local.nameId); expect(await inspect()).toMatchObject({ pending: true }); await names.selectOption(global.nameId); await expect(name).toHaveValue("KeepDraft");
  await reload.click(); await settled(); await expect(name).toHaveValue("Tax"); expect(await inspect()).toMatchObject({ pending: false, names: original });
  await name.fill("A1"); await save.click(); await settled(); expect(await inspect()).toMatchObject({ names: original, history: { undoSteps: 1 }, pending: true }); await reload.click(); await settled();
  expect(await inspect()).toMatchObject({ pending: false });
}
