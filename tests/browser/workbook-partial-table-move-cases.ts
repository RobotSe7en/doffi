import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";

export async function exerciseWorkbookPartialTableMove(page: Page, remote: boolean): Promise<void> {
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const choose = async (value: string) => { const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }); await address.fill(value); await address.press("Enter"); await settled(); };
  const tab = async (name: string) => { await page.getByRole("tab", { name, exact: true }).click(); await settled(); };
  const display = async (address: string, row: number, column: number) => { await choose(address); return page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.displayedCell(row, column) : window.workbookEditor.displayedCell(row, column), { remote, row, column }); };
  const undo = async () => { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); };
  const before = (await inspect() as { tables: unknown }).tables;
  await choose("A2:D3"); await page.getByRole("button", { name: "剪切选区", exact: true }).click();
  await tab("Observer"); await choose("H260"); await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
  expect(await display("K260", 259, 10)).toBe("20.00"); expect(await display("K261", 260, 10)).toBe("60.00");
  expect(await inspect()).toMatchObject({ tables: before, errors: [], history: { undoSteps: 1 } });
  await choose("I260"); await page.getByRole("textbox", { name: "单元格或公式输入", exact: true }).fill("6");
  await page.getByRole("button", { name: "确认输入", exact: true }).click(); await settled();
  expect(await display("K260", 259, 10)).toBe("60.00"); expect(await display("C1", 0, 2)).toBe("160");
  await undo(); await undo(); expect(await inspect()).toMatchObject({ tables: before, errors: [], history: { undoSteps: 0 } });
  await tab("Sales"); await choose("D6"); await page.getByRole("button", { name: "剪切选区", exact: true }).click();
  await tab("Observer"); await choose("H260"); await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
  expect(await display("H260", 259, 7)).toBe("240.00"); expect(await display("B1", 0, 1)).toBe("240");
  expect(await inspect()).toMatchObject({ tables: before, errors: [], history: { undoSteps: 1 } }); await undo();
  await tab("Sales"); await choose("B2:C3"); await page.getByRole("button", { name: "剪切选区", exact: true }).click();
  await choose("B4"); await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ tables: before, history: { undoSteps: 0 } });
  await page.getByLabel("剪切时允许覆盖已有内容", { exact: true }).check();
  await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
  expect(await display("D2", 1, 3)).toBe("0.00"); expect(await display("D4", 3, 3)).toBe("20.00"); expect(await display("D5", 4, 3)).toBe("60.00");
  await tab("Observer"); expect(await display("C1", 0, 2)).toBe("80");
  expect(await inspect()).toMatchObject({ tables: before, history: { undoSteps: 1 } }); await undo();
  expect(await inspect()).toMatchObject({ tables: before, history: { undoSteps: 0, redoSteps: 1 } });
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
}
