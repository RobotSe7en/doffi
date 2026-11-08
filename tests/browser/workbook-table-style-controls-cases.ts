import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession } from "@doffi/core";
import type { DeepReadonly, WorkbookMetadata, WorkbookModel, WorkbookTableStyle } from "@doffi/core";

export function workbookTableStyleControlsSource(model: WorkbookModel) {
  const session = createDocumentSession(model), manifest = session.snapshot().manifest;
  session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [
    { op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 1, column: 6, source: { kind: "value", value: { type: "number", value: 4 } } }, { row: 1, column: 7, source: { kind: "value", value: { type: "number", value: 5 } } }] },
    { op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 0, column: 1, cssClass: null }] },
    { op: "insertWorkbookSheet", name: "Other" },
  ] }));
  const source = session.serialize(), workbook = session.snapshot().metadata.workbook; session.dispose(); return { source, workbook };
}
export async function exerciseWorkbookTableStyleControls(page: Page, remote: boolean, original: DeepReadonly<WorkbookMetadata>): Promise<void> {
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { history: { undoSteps: number }; pending: boolean; errors: string[]; tables: { sheetId: string; tables: { name: string }[] }[] };
  const source = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.tableStyleSource() : window.workbookEditor.tableStyleSource(), remote) as { tableStyles: DeepReadonly<WorkbookTableStyle>[]; defaultTableStyle?: DeepReadonly<WorkbookMetadata["defaultTableStyle"]> };
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const cellAppearance = (row: number, column: number) => page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.displayedCellAppearance(row, column) : window.workbookEditor.displayedCellAppearance(row, column), { remote, row, column });
  const label = (name: string) => page.getByLabel(name, { exact: true }), button = (name: string) => page.getByRole("button", { name, exact: true });
  const choice = label("选择自定义智能表样式"), name = label("自定义智能表样式名称"), region = label("选择智能表样式区域"), enabled = label("启用智能表样式区域"), makeDefault = label("新建智能表默认使用此样式"), save = button("保存自定义智能表样式"), discard = button("放弃表样式草稿并重新读取"), remove = button("删除自定义智能表样式");
  const definition = original.tableStyles![0]!, selected = async () => (await source()).tableStyles.find(style => style.styleId === definition.styleId)!;
  const undo = async () => { await button("撤销").click(); await settled(); };
  await page.locator("summary").filter({ hasText: /^自定义智能表样式$/ }).click(); await choice.selectOption(definition.styleId);
  await save.click(); await settled(); expect((await inspect()).history.undoSteps).toBe(0); expect((await source()).tableStyles).toEqual(original.tableStyles); await expect(remove).toBeDisabled();
  await region.selectOption("wholeTable"); await label("表样式区域修改填充颜色").check(); await label("表样式区域填充颜色").fill("#00aa77"); await expect(region).toBeDisabled(); await expect(choice).toBeDisabled(); await save.click(); await settled();
  const after = await selected(); expect(after.pivotCompatible).toBe(definition.pivotCompatible); expect(after.elements.filter(element => element.region !== "wholeTable")).toEqual(definition.elements.filter(element => element.region !== "wholeTable")); expect((await source()).tableStyles[1]).toEqual(original.tableStyles![1]);
  expect((await cellAppearance(3, 1)).fill).toBe("rgb(0, 170, 119)"); expect((await inspect()).history.undoSteps).toBe(1); await undo();
  const directHeader = await cellAppearance(0, 0);
  await region.selectOption("headerRow"); await label("表样式区域修改填充颜色").check(); await label("表样式区域填充颜色").fill("#3366bb"); await save.click(); await settled(); expect((await cellAppearance(0, 1)).fill).toBe("rgb(51, 102, 187)"); expect(await cellAppearance(0, 0)).toEqual(directHeader); await undo();
  await region.selectOption("firstRowStripe"); await label("智能表样式条纹大小").fill("3"); await save.click(); await settled(); expect((await selected()).elements.find(element => element.region === "firstRowStripe")!.size).toBe(3); await undo();
  await region.selectOption("lastTotalCell"); await expect(enabled).not.toBeChecked(); await enabled.check(); await label("智能表区域数字格式").selectOption("format:" + original.numberFormats[0]!.formatId); await save.click(); await settled();
  expect((await selected()).elements.find(element => element.region === "lastTotalCell")!.style?.numberFormatId).toBe(original.numberFormats[0]!.formatId); await undo();
  await choice.selectOption(original.tableStyles![1]!.styleId); await remove.click(); await settled(); expect((await source()).tableStyles).toHaveLength(1); await undo();
  await choice.selectOption(""); await region.selectOption("wholeTable"); await name.fill("新样式禰󠄀😀"); await enabled.check(); await label("表样式区域修改填充颜色").check(); await label("表样式区域填充颜色").fill("#cc33aa"); await makeDefault.check();
  await page.getByRole("tab", { name: "Other", exact: true }).click(); await settled(); await save.click(); await settled();
  const created = (await source()).tableStyles.find(style => style.name === "新样式禰󠄀😀")!; expect(created).toBeDefined(); expect((await source()).defaultTableStyle).toEqual({ kind: "custom", styleId: created.styleId }); expect((await inspect()).history.undoSteps).toBe(1);
  await page.getByRole("tab", { name: original.sheets[0]!.name, exact: true }).click(); await settled();
  await page.locator("summary").filter({ hasText: /^智能表管理$/ }).click(); const address = label("选择单元格或区域"); await address.fill("G1:H3"); await address.press("Enter"); await settled();
  await label("选择智能表").selectOption(""); await expect(label("智能表样式")).toHaveValue("custom:" + created.styleId); await label("智能表名称").fill("NewStyled"); await button("保存智能表").click(); await settled();
  expect((await inspect()).tables[0]!.tables).toHaveLength(2); expect((await cellAppearance(1, 6)).fill).toBe("rgb(204, 51, 170)"); await expect(remove).toBeDisabled(); await undo();
  await makeDefault.uncheck(); await remove.click(); await settled(); expect((await source()).tableStyles.some(style => style.styleId === created.styleId)).toBe(false); expect((await source()).defaultTableStyle).toBeUndefined(); await undo();
  expect((await source()).defaultTableStyle).toEqual({ kind: "custom", styleId: created.styleId }); await undo(); expect((await source()).tableStyles).toEqual(original.tableStyles); expect((await source()).defaultTableStyle).toEqual(original.defaultTableStyle); expect((await inspect()).history.undoSteps).toBe(0);
  await choice.selectOption(definition.styleId); await name.evaluate(element => { const input = element as HTMLTextAreaElement; input.focus(); input.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true })); input.value = "中文禰󠄀😀草稿"; input.dispatchEvent(new InputEvent("input", { bubbles: true, isComposing: true })); });
  await expect(save).toBeDisabled(); await expect(name).toBeEnabled();
  await page.evaluate(async remote => { if (remote) await window.workbookRemoteEditor.external("external", true, { row: 8, column: 0 }); else { window.workbookEditor.external("external", { row: 8, column: 0 }); await window.workbookEditor.settled(); } }, remote);
  await name.dispatchEvent("compositionend"); await expect(name).toHaveValue("中文禰󠄀😀草稿"); await expect(save).toBeDisabled(); expect((await inspect()).pending).toBe(true); await discard.click(); await settled(); await expect(save).toBeEnabled();
  expect(await inspect()).toMatchObject({ pending: false, history: { undoSteps: 1 }, errors: [] }); expect((await source()).tableStyles).toEqual(original.tableStyles);
}
