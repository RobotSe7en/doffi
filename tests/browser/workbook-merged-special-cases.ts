import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";

export async function exerciseWorkbookMergedSpecial(page: Page, remote: boolean): Promise<void> {
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const choose = async (value: string) => { const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }); await address.fill(value); await address.press("Enter"); await settled(); };
  const displayed = async (address: string, row: number, column: number) => { await choose(address); return page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.displayedCell(row, column) : window.workbookEditor.displayedCell(row, column), { remote, row, column }); };
  const cell = (row: number, column: number) => page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.sourceCell(row, column, "sheet-1") : window.workbookEditor.cell("sheet-1", row, column), { remote, row, column });
  const undo = async () => { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); };
  const before = (await inspect() as { merges: unknown[] }).merges;
  for (const mode of ["contents", "values", "formulas", "formats"]) for (const skip of [false, true]) {
    await choose("A256:B258"); await page.getByRole("button", { name: "复制选区", exact: true }).click();
    await page.getByLabel("区域粘贴方式", { exact: true }).selectOption(mode);
    await page.getByLabel("区域粘贴时跳过空白", { exact: true }).setChecked(skip);
    await choose("D256"); await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
    expect(await displayed("D256", 255, 3)).toBe(mode === "formats" ? "88" : "合并😀禰󠄀");
    if (mode !== "formats") expect(await displayed("D258", 257, 3)).toBe(mode === "values" ? "14" : "95");
    expect(await inspect()).toMatchObject({ merges: before, errors: [], history: { undoSteps: 1 } });
    await undo(); expect(await cell(255, 3)).toMatchObject({ value: { type: "number", value: 88 } });
    expect(await inspect()).toMatchObject({ merges: before, history: { undoSteps: 0, redoSteps: 1 } });
  }
  for (const format of ["source", "preserve"]) for (const direction of ["row", "column"]) {
    await page.getByLabel("填充格式", { exact: true }).selectOption(format);
    await choose(direction === "row" ? "L1:M10" : "O1:X2"); await page.getByRole("button", { name: "数列填充", exact: true }).click(); await settled();
    for (const [index, position] of (direction === "row" ? [[4, 11], [6, 11], [8, 11]] : [[0, 18], [0, 20], [0, 22]]).entries()) {
      const [row, column] = position;
      expect(await cell(row!, column!)).toMatchObject({ value: { type: "number", value: (index + 3) * 10 } });
    }
    expect(await displayed(direction === "row" ? "L9" : "W1", direction === "row" ? 8 : 0, direction === "row" ? 11 : 22)).toBe("50");
    expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: 1 } });
    expect((await inspect() as { merges: unknown[] }).merges).toHaveLength(before.length + (format === "source" ? 3 : 0));
    await undo(); expect(await inspect()).toMatchObject({ merges: before, history: { undoSteps: 0 } });
  }
  await page.getByLabel("填充格式", { exact: true }).selectOption("preserve");
  await choose("A256:B261"); await page.getByRole("button", { name: "向下填充", exact: true }).click(); await settled();
  expect(await displayed("A260", 259, 0)).toBe("合并😀禰󠄀");
  expect(await inspect()).toMatchObject({ merges: before, errors: [], history: { undoSteps: 1 } }); await undo();
  // Source-preserving paste cannot silently hide destination E1=98 under a
  // new merge when blanks or values are deliberately retained.
  for (const mode of ["contents", "formats"]) {
    await choose("A256:B258"); await page.getByRole("button", { name: "复制选区", exact: true }).click();
    await page.getByLabel("区域粘贴方式", { exact: true }).selectOption(mode); await page.getByLabel("区域粘贴时跳过空白", { exact: true }).check();
    await choose("D1"); await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
    expect(await cell(0, 3)).toMatchObject({ value: { type: "number", value: 88 } }); expect(await cell(0, 4)).toMatchObject({ value: { type: "number", value: 98 } });
    expect(await inspect()).toMatchObject({ merges: before, history: { undoSteps: 0 } });
  }
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
}
