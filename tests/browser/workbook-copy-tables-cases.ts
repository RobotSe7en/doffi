import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";

export async function exerciseWorkbookTableCopy(page: Page, remote: boolean): Promise<void> {
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const choose = async (value: string) => { const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }); await address.fill(value); await address.press("Enter"); await settled(); };
  const undo = async () => { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); };
  const display = async (address: string, row: number, column: number) => { await choose(address); return page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.displayedCell(row, column) : window.workbookEditor.displayedCell(row, column), { remote, row, column }); };
  const before = (await inspect() as { tables: unknown }).tables;
  await choose("A1:D6"); await page.getByRole("button", { name: "复制选区", exact: true }).click();
  await page.getByRole("tab", { name: "Observer", exact: true }).click(); await settled(); await choose("H260");
  await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: 1 }, tables: [{ tables: [{ name: "SalesTable" }] }, { tables: [{ name: "SalesTable_2" }] }] });
  expect(await display("K261", 260, 10)).toBe("20.00"); expect(await display("K265", 264, 10)).toBe("240.00");
  expect(await display("B1", 0, 1)).toBe("240");
  await choose("I261"); await page.getByRole("textbox", { name: "单元格或公式输入", exact: true }).fill("6");
  await page.getByRole("button", { name: "确认输入", exact: true }).click(); await settled();
  expect(await display("K261", 260, 10)).toBe("60.00"); expect(await display("K265", 264, 10)).toBe("280.00");
  expect(await display("B1", 0, 1)).toBe("240"); expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: 2 } });
  await undo(); await undo(); expect(await inspect()).toMatchObject({ tables: before, errors: [], history: { undoSteps: 0 } });
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
}
