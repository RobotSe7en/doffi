import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession } from "@doffi/core";
import type { DeepReadonly, WorkbookAutoFilter, WorkbookModel, WorkbookSheet, WorkbookTable } from "@doffi/core";

export function workbookFilterControlsSource(model: WorkbookModel) {
  const session = createDocumentSession(model), before = session.snapshot(), filter = JSON.parse(JSON.stringify(before.metadata.workbook.sheets[0]!.autoFilter)) as WorkbookAutoFilter;
  filter.columns[0]!.criterion = { kind: "values", blank: false, values: ["keep", ...Array.from({ length: 127 }, (_, index) => index === 0 ? "长".repeat(32767) : "未匹配" + index), "中文\r\n禰󠄀😀"], dates: [] };
  const table: WorkbookTable = { tableId: "filter-ui-table", name: "FilterTable", range: { from: { row: 8, column: 6 }, to: { row: 12, column: 7 } }, header: true, totals: true, totalsPreviouslyShown: true,
    columns: [{ columnId: "key", name: "状态", totalFunction: "none", totalLabel: "合计" }, { columnId: "amount", name: "数值", totalFunction: "sum" }], style: { kind: "builtin", name: "TableStyleMedium2" }, showFirstColumn: false, showLastColumn: false, rowStripes: true, columnStripes: false };
  session.commit(session.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: before.manifest.revision, operations: [
    { op: "setWorkbookAutoFilter", sheetId: "sheet-1", filter }, { op: "setWorkbookTable", sheetId: "sheet-1", table },
    { op: "setWorkbookCells", sheetId: "sheet-1", cells: ["keep", "skip", "keep"].flatMap((text, index) => [
      { row: index + 9, column: 6, source: { kind: "value" as const, value: { type: "text" as const, value: text } } },
      { row: index + 9, column: 7, source: { kind: "value" as const, value: { type: "number" as const, value: (index + 1) * 10 } } },
    ]) },
  ] }));
  const source = session.serialize(), sheets = session.snapshot().metadata.workbook.sheets; session.dispose(); return { source, sheets };
}
export async function exerciseWorkbookFilterControls(page: Page, remote: boolean, original: readonly DeepReadonly<WorkbookSheet>[]): Promise<void> {
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { sheets: DeepReadonly<WorkbookSheet>[]; history: { undoSteps: number }; pending: boolean; errors: string[] };
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const displayed = (row: number, column: number) => page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.displayedCell(row, column) : window.workbookEditor.displayedCell(row, column), { remote, row, column });
  const filter = async () => (await inspect()).sheets[0]!.autoFilter!, table = async () => (await inspect()).sheets[0]!.tables![0]!;
  const label = (name: string) => page.getByLabel(name, { exact: true }), button = (name: string) => page.getByRole("button", { name, exact: true });
  const save = button("应用筛选条件"), discard = button("放弃筛选草稿并重新读取"), kind = label("筛选条件类型");
  const undo = async () => { await button("撤销").click(); await settled(); }, tab = async (index: number) => { await page.getByRole("tab", { name: original[index]!.name, exact: true }).click(); await settled(); };
  const column = async (name: string) => { await label("筛选列").fill(name); await button("读取筛选列").click(); };
  await page.getByText("筛选管理", { exact: true }).click(); await expect(kind).toHaveValue("values"); await expect(label("筛选值列表项").locator("option")).toHaveCount(50); expect((await label("筛选值列表项").locator("option").allTextContents()).every(text => text.length < 60)).toBe(true);
  await button("筛选值下一页").click(); await button("筛选值下一页").click(); await expect(label("筛选值列表项").locator("option")).toHaveCount(29); await label("筛选值列表项").selectOption("128"); await expect(label("筛选值文字")).toHaveValue("中文\n禰󠄀😀"); expect(await inspect()).toMatchObject({ pending: false });
  await save.click(); await settled(); expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, pending: false, errors: [] }); expect(await filter()).toEqual(original[0]!.autoFilter);
  await button("筛选值下一页").click(); await button("筛选值下一页").click(); await label("筛选值列表项").selectOption("128"); await label("筛选值文字").fill("中文\n禰󠄀😀尾"); await save.click(); await settled();
  expect((await filter()).columns[0]!.criterion).toMatchObject({ kind: "values", values: [...(original[0]!.autoFilter!.columns[0]!.criterion as Extract<WorkbookAutoFilter["columns"][number]["criterion"], { kind: "values" }>).values.slice(0, -1), "中文\r\n禰󠄀😀尾"] }); await undo();
  for (const [index, expected, field, value] of [[1, "custom", "筛选第一个比较值", "2"], [2, "top", "筛选排名数量", "2"], [3, "top", "筛选排名方向", "bottom"], [4, "top", "筛选排名数量", "25"], [5, "dynamic", "筛选动态条件", "aboveAverage"], [7, "custom", "筛选第一个比较值", "a~*b"]] as const) {
    await tab(index); await expect(kind).toHaveValue(expected); await expect(label(field)).toHaveValue(value); expect(await inspect()).toMatchObject({ pending: false });
  }
  await tab(0); await column("B"); await kind.selectOption("custom"); await label("筛选第一个比较").selectOption("lessThan"); await label("筛选第一个比较值").fill("50"); await save.click(); await settled();
  expect(await displayed(0, 3)).toBe("1"); expect(await displayed(1, 3)).toBeUndefined(); expect((await inspect()).sheets[0]!.rowDimensions).toEqual(original[0]!.rowDimensions);
  await column("B"); await label("筛选使用第二个条件").check(); await label("筛选第二个比较").selectOption("greaterThan"); await label("筛选第二个比较值").fill("90"); await label("筛选条件组合").selectOption("or"); await save.click(); await settled(); expect(await displayed(0, 3)).toBe("101"); expect((await filter()).columns[1]!.criterion).toMatchObject({ kind: "custom", and: false, tests: [{ operator: "lessThan", value: "50" }, { operator: "greaterThan", value: "90" }] }); await undo();
  await column("B"); await kind.selectOption("top"); await label("筛选排名使用百分比").check(); await label("筛选排名数量").fill("150"); await save.click(); await settled(); expect(await inspect()).toMatchObject({ history: { undoSteps: 1 }, pending: true }); expect((await inspect()).errors).toHaveLength(1);
  await button("清除所有筛选条件").click(); await settled(); expect(await displayed(0, 3)).toBe("111"); expect(await displayed(1, 3)).toBeUndefined(); expect((await filter()).columns.every(column => column.criterion === undefined)).toBe(true); await undo(); await undo();
  await column("B"); await kind.selectOption("top"); await label("筛选排名数量").fill("1"); await save.click(); await settled(); expect(await displayed(0, 3)).toBe("100"); expect(await displayed(1, 3)).toBeUndefined(); await undo();
  await column("B"); await kind.selectOption("dynamic"); await label("筛选动态条件").selectOption("belowAverage"); await save.click(); await settled(); expect(await displayed(0, 3)).toBe("1"); await undo();
  await label("筛选包含空白").check(); await label("隐藏此列筛选按钮").check(); await label("显示此列筛选按钮").uncheck(); await save.click(); await settled(); expect((await filter()).columns[0]).toMatchObject({ column: 0, hiddenButton: true, showButton: false, criterion: { kind: "values", blank: true } }); await undo();
  await tab(1); await kind.selectOption("values"); await button("添加筛选日期").click(); await label("筛选日期精度").selectOption("day"); await label("筛选日期年").fill("1900"); await label("筛选日期月").fill("2"); await label("筛选日期日").fill("29"); await save.click(); await settled();
  expect((await inspect()).sheets[1]!.autoFilter!.columns[0]!.criterion).toMatchObject({ kind: "values", dates: [{ grouping: "day", year: 1900, month: 2, day: 29 }] }); await undo();
  await tab(0); await label("筛选所属区域").selectOption("filter-ui-table"); await expect(label("筛选区域")).toBeDisabled(); await expect(label("筛选列")).toHaveValue("G"); await kind.selectOption("values"); await button("添加筛选值").click(); await label("筛选值文字").fill("keep"); await save.click(); await settled();
  expect(await table()).toMatchObject({ ...original[0]!.tables![0], autoFilter: { range: { from: { row: 8, column: 6 }, to: { row: 11, column: 7 } }, columns: [{ column: 6, criterion: { kind: "values", values: ["keep"], dates: [] } }] } });
  const address = label("选择单元格或区域"); await address.fill("H13"); await address.press("Enter"); await settled(); expect(await displayed(12, 7)).toBe("40");
  await button("移除区域筛选").click(); await settled(); expect(await displayed(12, 7)).toBe("60"); await undo(); await undo(); expect(await table()).toEqual(original[0]!.tables![0]);
  await kind.selectOption("custom"); await label("筛选第一个比较值").fill("keep"); await tab(7); await save.click(); await settled(); expect((await table()).autoFilter?.columns[0]?.criterion).toMatchObject({ kind: "custom", tests: [{ operator: "equal", value: "keep" }] }); expect((await inspect()).sheets[7]!.autoFilter).toEqual(original[7]!.autoFilter); await undo();
  await kind.selectOption("custom");
  await label("筛选第一个比较值").evaluate(element => { const input = element as HTMLTextAreaElement; input.focus(); input.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true })); input.value = "中文禰󠄀😀"; input.dispatchEvent(new InputEvent("input", { bubbles: true, isComposing: true })); });
  await expect(save).toBeDisabled(); await expect(label("筛选第一个比较值")).toBeEnabled(); await page.evaluate(async remote => { if (remote) await window.workbookRemoteEditor.external("external", true, { row: 20, column: 0 }); else { window.workbookEditor.external("external", { row: 20, column: 0 }); await window.workbookEditor.settled(); } }, remote);
  await label("筛选第一个比较值").dispatchEvent("compositionend"); await expect(save).toBeDisabled(); expect(await inspect()).toMatchObject({ pending: true }); await discard.click(); await settled(); expect(await inspect()).toMatchObject({ pending: false, history: { undoSteps: 1 } });
}
