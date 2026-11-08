import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession } from "@doffi/core";
import type { DeepReadonly, WorkbookHyperlink, WorkbookModel, WorkbookSheet } from "@doffi/core";

export function workbookCopyHyperlinksSource(model: WorkbookModel) {
  const session = createDocumentSession(model), manifest = session.snapshot().manifest, name = session.snapshot().metadata.workbook.sheets[0]!.name;
  const source = session.snapshot().metadata.workbook.sheets[0]!.hyperlinks![0]!, { range: originalRange, ...target } = source;
  if (originalRange.from.row !== 1) throw Error("Unexpected immutable hyperlink fixture");
  session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [
    { op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 1, column: 16, source: { kind: "value", value: { type: "text", value: "中文禰󠄀😀\r\n尾" } } }, { row: 3, column: 16, source: { kind: "value", value: { type: "number", value: 7 } } }] },
    { op: "setWorkbookHyperlink", sheetId: "sheet-1", range: { from: { row: 1, column: 16 }, to: { row: 2, column: 16 } }, target: { ...target, tooltip: "提示\r\n来源", display: "显示\r\n来源" } },
    { op: "setWorkbookHyperlink", sheetId: "sheet-1", range: { from: { row: 3, column: 16 }, to: { row: 3, column: 16 } }, target: { location: "'" + name.replaceAll("'", "''") + "'!$A$1" } },
    { op: "setWorkbookHyperlink", sheetId: "sheet-1", range: { from: { row: 0, column: 4 }, to: { row: 5, column: 6 } }, target: { url: "https://example.org/old", tooltip: "目标" } },
    { op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: { row: 1, column: 16 }, to: { row: 4, column: 16 } }], properties: { "font-weight": "bold" } },
    { op: "insertWorkbookSheet", name: "Other" },
  ] }));
  const files = session.serialize(), sheets = session.snapshot().metadata.workbook.sheets; session.dispose(); return { source: files, sheets };
}
export async function exerciseWorkbookCopyHyperlinks(page: Page, remote: boolean, original: readonly DeepReadonly<WorkbookSheet>[]): Promise<void> {
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { sheets: DeepReadonly<WorkbookSheet>[]; history: { undoSteps: number }; errors: string[] };
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const button = (name: string) => page.getByRole("button", { name, exact: true }), skip = page.getByLabel("区域粘贴时跳过空白", { exact: true });
  const select = async (range: string) => { const address = page.getByLabel("选择单元格或区域", { exact: true }); await address.fill(range); await address.press("Enter"); await settled(); };
  const linkAt = (links: readonly DeepReadonly<WorkbookHyperlink>[] | undefined, row: number, column: number) => links?.find(link => link.range.from.row <= row && link.range.to.row >= row && link.range.from.column <= column && link.range.to.column >= column);
  const source = linkAt(original[0]!.hyperlinks, 1, 16)!;
  async function paste(blank: boolean, mode = "contents") { await select("Q2:Q5"); await button("复制选区").click(); await skip.setChecked(blank); await page.getByLabel("区域粘贴方式", { exact: true }).selectOption(mode); await select("F2"); await button("粘贴选区").click(); await settled(); }
  const undo = async () => { await button("撤销").click(); await settled(); expect((await inspect()).history.undoSteps).toBe(0); expect((await inspect()).sheets[0]!.hyperlinks).toEqual(original[0]!.hyperlinks); };
  await paste(false); let current = await inspect(); expect(current.history.undoSteps).toBe(1);
  expect(linkAt(current.sheets[0]!.hyperlinks, 1, 5)).toMatchObject({ url: source.url, tooltip: source.tooltip, display: source.display }); expect(linkAt(current.sheets[0]!.hyperlinks, 2, 5)?.url).toBe(source.url);
  expect(linkAt(current.sheets[0]!.hyperlinks, 3, 5)?.location).toBe(linkAt(original[0]!.hyperlinks, 3, 16)!.location); expect(linkAt(current.sheets[0]!.hyperlinks, 4, 5)).toBeUndefined(); expect(linkAt(current.sheets[0]!.hyperlinks, 4, 4)?.url).toBe("https://example.org/old");
  await select("F2"); await expect(page.getByLabel("超链接外部地址", { exact: true })).toHaveValue(source.url!); await undo();
  await paste(true); current = await inspect(); expect(linkAt(current.sheets[0]!.hyperlinks, 1, 5)?.url).toBe(source.url); expect(linkAt(current.sheets[0]!.hyperlinks, 2, 5)?.url).toBe("https://example.org/old"); expect(linkAt(current.sheets[0]!.hyperlinks, 4, 5)?.url).toBe("https://example.org/old"); await undo();
  await paste(false, "values"); expect((await inspect()).sheets[0]!.hyperlinks).toEqual(original[0]!.hyperlinks); await undo();
  await paste(false, "formats"); expect((await inspect()).sheets[0]!.hyperlinks).toEqual(original[0]!.hyperlinks); await undo();
  await select("Q2:Q5"); await button("复制选区").click(); await page.getByLabel("区域粘贴方式", { exact: true }).selectOption("contents"); await page.getByRole("tab", { name: "Other", exact: true }).click(); await settled(); await select("AG257"); await button("粘贴选区").click(); await settled();
  current = await inspect(); expect(linkAt(current.sheets[1]!.hyperlinks, 256, 32)).toMatchObject({ url: source.url, tooltip: source.tooltip, display: source.display });
  const cell = await page.evaluate(({ remote, sheetId }) => remote ? window.workbookRemoteEditor.sourceCell(256, 32, sheetId) : window.workbookEditor.cell(sheetId, 256, 32), { remote, sheetId: original[1]!.sheetId }); expect(cell?.text).toBe("中文禰󠄀😀\r\n尾"); await undo();
  await page.getByRole("tab", { name: original[0]!.name, exact: true }).click(); await settled(); await select("Q2:Q5"); await skip.check(); await button("剪切选区").click(); await expect(skip).toBeDisabled(); await expect(skip).not.toBeChecked(); expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, errors: [] });
}
