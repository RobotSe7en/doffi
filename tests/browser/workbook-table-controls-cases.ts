import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession } from "@doffi/core";
import type { DeepReadonly, WorkbookModel, WorkbookSheet, WorkbookTable } from "@doffi/core";

export function workbookTableControlsSource(model: WorkbookModel) {
  const session = createDocumentSession(model), manifest = session.snapshot().manifest;
  session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [
    { op: "setWorkbookCells", sheetId: "sheet-1", cells: [
      { row: 2, column: 3, source: { kind: "value", value: { type: "number", value: 777 } } },
      { row: 0, column: 5, source: { kind: "formula", expression: "=SUM(SalesTable[金额])" } },
      { row: 1, column: 6, source: { kind: "value", value: { type: "number", value: 4 } } },
      { row: 2, column: 6, source: { kind: "value", value: { type: "number", value: 5 } } },
    ] }, { op: "insertWorkbookSheet", name: "Other" },
  ] }));
  const source = session.serialize(), sheets = session.snapshot().metadata.workbook.sheets; session.dispose(); return { source, sheets };
}
export async function exerciseWorkbookTableControls(page: Page, remote: boolean, original: readonly DeepReadonly<WorkbookSheet>[]): Promise<void> {
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { tables: { sheetId: string; tables: DeepReadonly<WorkbookTable>[] }[]; history: { undoSteps: number }; pending: boolean; errors: string[] };
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const displayed = (row: number, column: number) => page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.displayedCell(row, column) : window.workbookEditor.displayedCell(row, column), { remote, row, column });
  const cell = (row: number, column: number) => page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.sourceCell(row, column) : window.workbookEditor.cell("sheet-1", row, column), { remote, row, column });
  const definition = original[0]!.tables![0]!, table = async () => (await inspect()).tables[0]!.tables[0]!;
  const choice = page.getByLabel("选择智能表", { exact: true }), name = page.getByLabel("智能表名称", { exact: true }), range = page.getByLabel("智能表区域", { exact: true }), column = page.getByLabel("选择智能表列", { exact: true }), columnName = page.getByLabel("智能表列名", { exact: true });
  const calculated = page.getByLabel("智能表计算列公式", { exact: true }), save = page.getByRole("button", { name: "保存智能表", exact: true }), discard = page.getByRole("button", { name: "放弃智能表草稿并重新读取", exact: true });
  const undo = async () => { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); };
  const select = async (text: string) => { const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }); await address.fill(text); await address.press("Enter"); await settled(); };
  await page.getByText("智能表管理", { exact: true }).click(); await choice.selectOption(definition.tableId); await expect(name).toHaveValue("SalesTable");
  await name.fill("Revenue"); await column.selectOption(definition.columns[1]!.columnId); await columnName.fill("件数");
  await expect(choice).toBeDisabled(); expect(await inspect()).toMatchObject({ pending: true, history: { undoSteps: 0 } });
  await save.click(); await settled(); expect(await inspect()).toMatchObject({ pending: false, history: { undoSteps: 1 }, errors: [] });
  expect(await table()).toMatchObject({ tableId: definition.tableId, name: "Revenue", style: definition.style, columns: [definition.columns[0], { ...definition.columns[1], name: "件数" }, definition.columns[2], { ...definition.columns[3], calculatedFormula: "=(Revenue[[#This Row],[件数]]*Revenue[[#This Row],[单价]])" }] });
  expect(await cell(0, 1)).toMatchObject({ value: { type: "text", value: "件数" } }); expect(await cell(0, 5)).toMatchObject({ formula: { expression: "=SUM(Revenue[金额])" } }); expect(await cell(2, 3)).toMatchObject({ value: { type: "number", value: 777 } });
  expect(await displayed(0, 1)).toBe("件数"); expect(await displayed(2, 3)).toContain("777");
  await save.click(); await settled(); expect(await inspect()).toMatchObject({ history: { undoSteps: 1 } });
  await range.fill("A1:D8"); await save.click(); await settled(); expect(await table()).toMatchObject({ range: { to: { row: 7, column: 3 } }, autoFilter: { range: { to: { row: 6, column: 3 } } } });
  expect(await cell(2, 3)).toMatchObject({ value: { type: "number", value: 777 } }); expect(await cell(6, 3)).toMatchObject({ formula: { expression: "=(Revenue[[#This Row],[件数]]*Revenue[[#This Row],[单价]])" } }); expect(await cell(7, 3)).toMatchObject({ formula: { expression: "=SUBTOTAL(109,Revenue[金额])" } });
  await column.selectOption(definition.columns[3]!.columnId); await calculated.fill("=Revenue[[#This Row],[件数]]*2"); await save.click(); await settled(); expect((await cell(2, 3))!.formula?.expression).toBe("=(Revenue[[#This Row],[件数]]*2)"); await undo(); expect(await cell(2, 3)).toMatchObject({ value: { type: "number", value: 777 } });
  await column.selectOption(definition.columns[3]!.columnId); await page.getByLabel("智能表列汇总方式", { exact: true }).selectOption("custom"); await page.getByLabel("智能表列自定义汇总公式", { exact: true }).fill("=SUM(Revenue[金额])*2"); await save.click(); await settled(); expect(await cell(7, 3)).toMatchObject({ formula: { expression: "=SUM(Revenue[金额])*2" } }); await undo();
  await page.getByLabel("智能表样式", { exact: true }).selectOption("TableStyleDark3"); await page.getByLabel("智能表突出首列", { exact: true }).check(); await page.getByLabel("智能表行条纹", { exact: true }).uncheck(); await save.click(); await settled(); expect(await table()).toMatchObject({ style: { kind: "builtin", name: "TableStyleDark3" }, showFirstColumn: true, rowStripes: false }); await undo();
  await page.getByLabel("智能表对象名称", { exact: true }).fill("NativeObjectAlias"); await page.getByLabel("智能表说明", { exact: true }).fill("中文禰󠄀😀说明"); await save.click(); await settled(); expect(await table()).toMatchObject({ objectName: "NativeObjectAlias", description: "中文禰󠄀😀说明" });
  await page.getByLabel("智能表对象名称", { exact: true }).fill(""); await page.getByLabel("智能表说明", { exact: true }).fill(""); await save.click(); await settled(); expect((await table()).objectName).toBeUndefined(); expect((await table()).description).toBeUndefined(); await undo(); await undo();
  await page.getByRole("button", { name: "智能表转为普通区域", exact: true }).click(); await settled(); expect((await inspect()).tables[0]!.tables).toEqual([]); expect((await cell(0, 5))!.formula?.expression).not.toContain("Revenue["); expect(await cell(2, 3)).toMatchObject({ value: { type: "number", value: 777 } }); await undo(); await undo(); await undo();
  expect((await inspect()).tables[0]!.tables).toEqual(original[0]!.tables); expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, errors: [] });
  await choice.selectOption(""); await select("G1:H3"); await expect(range).toHaveValue("G1:H3"); await name.fill("Fresh");
  await column.selectOption({ index: 0 }); await columnName.fill("输入"); await column.selectOption({ index: 1 }); await columnName.fill("结果"); await calculated.fill("=G2*3");
  await page.getByRole("tab", { name: "Other", exact: true }).click(); await settled(); await save.click(); await settled();
  expect((await inspect()).tables[0]!.tables[1]).toMatchObject({ name: "Fresh", range: { from: { row: 0, column: 6 }, to: { row: 2, column: 7 } }, columns: [{ name: "输入" }, { name: "结果", calculatedFormula: "=G2*3" }] }); expect((await inspect()).tables[1]!.tables).toEqual([]); expect(await cell(2, 7)).toMatchObject({ formula: { expression: "=(G3*3)" } }); await undo();
  await page.getByRole("tab", { name: original[0]!.name, exact: true }).click(); await settled(); await choice.selectOption(definition.tableId); await range.fill("A1:B6");
  await page.getByRole("button", { name: "按智能表区域调整列", exact: true }).click(); await expect(column.locator("option")).toHaveCount(2); await discard.click(); await expect(column.locator("option")).toHaveCount(4);
  await columnName.evaluate(element => { const input = element as HTMLInputElement; input.focus(); input.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true })); input.value = "中文禰󠄀😀草稿"; input.dispatchEvent(new InputEvent("input", { bubbles: true, isComposing: true })); }); await expect(save).toBeDisabled(); await expect(columnName).toBeEnabled();
  await page.evaluate(async remote => { if (remote) await window.workbookRemoteEditor.external("external", true, { row: 8, column: 0 }); else { window.workbookEditor.external("external", { row: 8, column: 0 }); await window.workbookEditor.settled(); } }, remote);
  await columnName.dispatchEvent("compositionend"); await expect(columnName).toHaveValue("中文禰󠄀😀草稿"); await expect(save).toBeDisabled(); expect(await inspect()).toMatchObject({ pending: true }); await discard.click(); await settled(); await expect(save).toBeEnabled(); expect(await inspect()).toMatchObject({ pending: false, errors: [], history: { undoSteps: 1 } });
}
