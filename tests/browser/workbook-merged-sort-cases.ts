import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import type { DeepReadonly, WorkbookRange } from "@doffi/core";

export async function exerciseWorkbookMergedSort(page: Page, remote: boolean, merges: readonly DeepReadonly<WorkbookRange>[]): Promise<void> {
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const display = (row: number, column: number) => page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.displayedCell(row, column) : window.workbookEditor.displayedCell(row, column), { remote, row, column });
  const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true });
  await address.fill("A1:D6"); await address.press("Enter"); await settled();
  await page.getByLabel("首行为表头", { exact: true }).uncheck();
  await page.getByRole("button", { name: "升序排序", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ errors: [], merges, history: { undoSteps: 1 } });
  expect(await Promise.all([0, 2, 4].map(row => display(row, 0)))).toEqual(["1", "1", "3"]);
  expect(await Promise.all([0, 2, 4].map(row => display(row, 2)))).toEqual(["20", "25", "30"]);
  expect(await display(0, 5)).toBe("1");
  await page.getByRole("button", { name: "降序排序", exact: true }).click(); await settled();
  expect(await Promise.all([0, 2, 4].map(row => display(row, 0)))).toEqual(["3", "1", "1"]);
  expect(await inspect()).toMatchObject({ errors: [], merges, history: { undoSteps: 2 } });
  for (let step = 0; step < 2; step++) { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); }
  expect(await inspect()).toMatchObject({ errors: [], merges, history: { undoSteps: 0, redoSteps: 2 } });
  expect(await Promise.all([0, 2, 4].map(row => display(row, 0)))).toEqual(["3", "1", "1"]);
}
