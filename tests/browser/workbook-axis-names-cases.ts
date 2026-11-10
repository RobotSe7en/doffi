import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";

export async function exerciseWorkbookAxisNames(page: Page, remote: boolean): Promise<void> {
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const choose = async (address: string) => { const input = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }); await input.fill(address); await input.press("Enter"); await settled(); };
  const displayed = async (row: number, column: number) => {
    // The native fixture's wide columns put G/H outside the initial viewport.
    // Navigate to each value before checking the actual virtualized canvas.
    await choose(String.fromCharCode(65 + column) + (row + 1));
    return page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.displayedCell(row, column) : window.workbookEditor.displayedCell(row, column), { remote, row, column });
  };
  const tab = async (name: string) => { await page.getByRole("tab", { name, exact: true }).click(); await settled(); };
  const before = (await inspect() as { names: unknown }).names;
  await choose("A1"); await page.getByLabel("插入行列的格式", { exact: true }).selectOption("none");
  await page.getByRole("button", { name: "插入选区行", exact: true }).click(); await settled();
  await choose("C6:H6");
  expect(await displayed(5, 2)).toBe("15"); expect(await displayed(5, 3)).toBe("11"); expect(await displayed(5, 4)).toBe("11");
  expect(await displayed(5, 5)).toBe("0"); expect(await displayed(5, 6)).toBe("25"); expect(await displayed(5, 7)).toBe("48");
  await tab("Other"); await choose("C5:H5");
  expect(await displayed(4, 2)).toBe("14"); expect(await displayed(4, 5)).toBe("101"); expect(await displayed(4, 6)).toBe("24");
  expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: 1 } });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ names: before, history: { undoSteps: 0, redoSteps: 1 } });
  await tab("Data"); await choose("A1"); await page.getByRole("button", { name: "插入选区列", exact: true }).click(); await settled();
  await choose("D5:I5"); expect(await displayed(4, 3)).toBe("15"); expect(await displayed(4, 7)).toBe("25");
  await tab("Other"); await choose("C5:H5"); expect(await displayed(4, 2)).toBe("15"); expect(await displayed(4, 6)).toBe("25");
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ names: before, errors: [], history: { undoSteps: 0, redoSteps: 1 } });
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
}
