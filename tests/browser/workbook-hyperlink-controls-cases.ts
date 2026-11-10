import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import type { DeepReadonly, WorkbookSheet } from "@doffi/core";

export async function exerciseWorkbookHyperlinkControls(page: Page, remote: boolean, original: readonly DeepReadonly<WorkbookSheet>[]): Promise<void> {
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { sheets: DeepReadonly<WorkbookSheet>[]; history: { undoSteps: number }; pending: boolean; errors: string[] };
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }), select = async (at: string) => { await address.fill(at); await address.press("Enter"); await settled(); };
  const url = page.getByLabel("超链接外部地址", { exact: true }), location = page.getByLabel("超链接工作簿内位置", { exact: true }), tooltip = page.getByLabel("超链接屏幕提示", { exact: true });
  const save = page.getByRole("button", { name: "设置选区超链接", exact: true }), remove = page.getByRole("button", { name: "移除选区超链接", exact: true }), discard = page.getByRole("button", { name: "放弃超链接草稿并读取选区", exact: true });
  const links = async () => (await inspect()).sheets.find(sheet => sheet.sheetId === original[0]!.sheetId)!.hyperlinks;
  const point = async (row: number, column: number) => (await links())?.find(link => row >= link.range.from.row && row <= link.range.to.row && column >= link.range.from.column && column <= link.range.to.column);
  const undo = async () => { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); };
  await page.getByText("单元格超链接", { exact: true }).click(); await select("A1:B2"); await expect(url).toHaveValue(""); await expect(location).toHaveValue("");
  await url.fill("https://example.com/new-link"); await tooltip.fill("中文禰󠄀😀屏幕提示"); expect(await inspect()).toMatchObject({ pending: true, history: { undoSteps: 0 } });
  await save.click(); await settled(); expect(await inspect()).toMatchObject({ pending: false, history: { undoSteps: 1 }, errors: [] });
  expect(await point(0, 0)).toMatchObject({ range: { from: { row: 0, column: 0 }, to: { row: 1, column: 1 } }, url: "https://example.com/new-link", tooltip: "中文禰󠄀😀屏幕提示" });
  expect(await page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCell(0, 0) : window.workbookEditor.displayedCell(0, 0), remote)).toBe("Internal");
  await save.click(); await settled(); expect(await inspect()).toMatchObject({ history: { undoSteps: 1 }, pending: false });
  await select("A1"); await expect(url).toHaveValue("https://example.com/new-link"); await url.fill(""); await location.fill("Other!A1"); await save.click(); await settled();
  expect(await point(0, 0)).toMatchObject({ location: "Other!A1" }); expect((await point(0, 0))!.url).toBeUndefined(); expect(await point(1, 1)).toMatchObject({ url: "https://example.com/new-link" });
  await select("B1:B2"); await remove.click(); await settled(); expect(await point(0, 1)).toBeUndefined(); expect(await point(1, 1)).toBeUndefined(); expect(await point(1, 0)).toMatchObject({ url: "https://example.com/new-link" });
  for (let index = 0; index < 3; index++) await undo(); expect((await inspect()).sheets.map(sheet => sheet.hyperlinks)).toEqual(original.map(sheet => sheet.hyperlinks));
  await select("A1"); await location.fill(""); await url.fill("https://example.com/original-target");
  await page.getByRole("tab", { name: original[1]!.name, exact: true }).click(); await settled(); await expect(page.getByLabel("超链接编辑状态", { exact: true })).toContainText("原区域");
  await save.click(); await settled(); expect(await point(0, 0)).toMatchObject({ url: "https://example.com/original-target" }); expect((await inspect()).sheets[1]!.hyperlinks).toEqual(original[1]!.hyperlinks); await undo();
  await page.getByRole("tab", { name: original[0]!.name, exact: true }).click(); await settled(); await select("A1"); await url.fill("javascript:alert(1)"); await save.click(); await settled();
  expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, pending: true }); expect((await inspect()).errors).toHaveLength(1); expect(await links()).toEqual(original[0]!.hyperlinks);
  await url.evaluate(element => { const field = element as HTMLInputElement; field.focus(); field.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true })); field.value = "https://example.com/draft"; field.dispatchEvent(new InputEvent("input", { bubbles: true, isComposing: true })); });
  await expect(save).toBeDisabled(); await expect(url).toBeEnabled();
  await page.evaluate(async remote => { if (remote) await window.workbookRemoteEditor.external("external"); else { window.workbookEditor.external("external"); await window.workbookEditor.settled(); } }, remote);
  await url.dispatchEvent("compositionend"); await expect(url).toHaveValue("https://example.com/draft"); await expect(save).toBeDisabled(); expect(await inspect()).toMatchObject({ pending: true });
  await discard.click(); await settled(); await expect(url).toHaveValue(""); await expect(location).toHaveValue("Data!$A$5"); expect(await inspect()).toMatchObject({ pending: false, history: { undoSteps: 1 } }); expect(await links()).toEqual(original[0]!.hyperlinks);
}
