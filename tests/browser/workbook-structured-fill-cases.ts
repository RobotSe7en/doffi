import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";

export async function exerciseWorkbookStructuredFill(page: Page, remote: boolean): Promise<void> {
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const choose = async (value: string) => { const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }); await address.fill(value); await address.press("Enter"); await settled(); };
  const undo = async () => { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); };
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
  const write = async (formula: string) => { await choose("H260"); await input.fill(formula); await page.getByRole("button", { name: "确认输入", exact: true }).click(); await settled(); };
  const display = async (address: string, row: number, column: number) => { await choose(address); return page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.displayedCell(row, column) : window.workbookEditor.displayedCell(row, column), { remote, row, column }); };
  await page.getByRole("tab", { name: "Observer", exact: true }).click(); await settled();
  await write("=SUM(SalesTable[数量])"); await choose("H260:M260");
  await page.getByRole("button", { name: "向右填充", exact: true }).click(); await settled();
  for (const [address, column, value] of [["I260", 8, "100"], ["J260", 9, "240"], ["K260", 10, "0"], ["L260", 11, "10"], ["M260", 12, "100"]] as const) expect(await display(address, 259, column)).toBe(value);
  expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: 2 } }); await undo();
  await choose("H260"); await page.getByRole("button", { name: "复制选区", exact: true }).click(); await choose("I260");
  await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
  expect(await display("I260", 259, 8)).toBe("10"); await expect(input).toHaveValue("=SUM(SalesTable[数量])"); await undo();
  await choose("H260:H263"); await page.getByRole("button", { name: "向下填充", exact: true }).click(); await settled();
  expect(await display("H263", 262, 7)).toBe("10"); await expect(input).toHaveValue("=SUM(SalesTable[数量])"); await undo(); await undo();
  for (const [formula, value] of [["=SUM(SalesTable[[数量]:[单价]])", "110"], ["=SUM(SalesTable[[数量]:[数量]])", "10"]] as const) {
    await write(formula); await choose("H260:L260"); await page.getByRole("button", { name: "向右填充", exact: true }).click(); await settled();
    expect(await display("L260", 259, 11)).toBe(value); await expect(input).toHaveValue(formula); await undo(); await undo();
  }
  expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: 0 } });
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
}
