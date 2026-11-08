import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import { scalarWorkbook } from "../../packages/core/tests/workbook-scalar-fixture.js";
import type { WorkbookDataValidation } from "@doffi/core";

export const validationPreparation = { calculationContext: { now: { unixMilliseconds: Date.UTC(2026, 10, 9), utcOffsetMinutes: 480 }, locale: "zh-CN" } };
export function validationConfirmationSource() {
  return scalarWorkbook([0, 1].map(column => ({ row: 0, column, value: { type: "number", value: 9 } })), metadata => {
    metadata.workbook.sheets[0]!.dataValidations = [0, 1, 2, 3, 4].map(column => {
      const origin = { row: 0, column }, base = { ruleId: "rule-" + column, origin, ranges: [{ from: origin, to: { row: 9, column } }], allowBlank: true,
        errorStyle: column === 2 ? "stop" as const : column === 1 ? "information" as const : "warning" as const,
        showErrorMessage: true, showInputMessage: true, errorTitle: "范围", error: "必须小于五", promptTitle: "", prompt: "" };
      return (column === 3 ? { ...base, kind: "custom", expression: "=FUTUREFUNC(D1)" }
        : column === 4 ? { ...base, kind: "custom", expression: "=TODAY()>E1" }
        : { ...base, kind: "decimal", operator: "lessThan", formulas: ["=5"] }) satisfies WorkbookDataValidation;
    });
  }).serialize();
}
export async function exerciseValidationConfirmation(page: Page, remote: boolean): Promise<void> {
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { history: { undoSteps: number }; pending: boolean; input: { text: string; dirty: boolean }; errors: string[] };
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const cell = (row: number, column: number) => page.evaluate(({ remote, row, column }) => remote ? window.workbookRemoteEditor.sourceCell(row, column) : window.workbookEditor.cell("sheet-1", row, column), { remote, row, column });
  const button = (name: string) => page.getByRole("button", { name, exact: true }), input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
  const select = async (range: string) => { const address = page.getByLabel("选择单元格或区域", { exact: true }); await address.fill(range); await address.press("Enter"); await settled(); };
  const undo = async () => { await button("撤销").click(); await settled(); expect((await inspect()).history.undoSteps).toBe(0); };
  async function decide(action: () => Promise<unknown>, answers: boolean[]) {
    let count = 0;
    const onDialog = async (dialog: import("@playwright/test").Dialog) => { expect(dialog.type()).toBe("confirm"); const answer = answers[count++]; expect(answer).toBeDefined(); if (answer) await dialog.accept(); else await dialog.dismiss(); };
    page.on("dialog", onDialog);
    try { await action(); await settled(); expect(count).toBe(answers.length); } finally { page.off("dialog", onDialog); }
  }
  await select("A2"); await input.fill("9"); await decide(() => button("确认输入").click(), [false]);
  expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, pending: true, input: { text: "9", dirty: true } }); expect(await cell(1, 0)).toBeUndefined();
  await decide(() => button("确认输入").click(), [true]); expect(await inspect()).toMatchObject({ history: { undoSteps: 1 }, pending: false }); expect((await cell(1, 0))?.value).toEqual({ type: "number", value: 9 }); await undo();
  await select("A1:B3"); await decide(() => button("向下填充").click(), [true, false]);
  expect((await inspect()).history.undoSteps).toBe(0); expect(await cell(1, 0)).toBeUndefined(); expect(await cell(1, 1)).toBeUndefined();
  await decide(() => button("向下填充").click(), [true, true]); expect((await inspect()).history.undoSteps).toBe(1); expect((await cell(2, 1))?.value).toEqual({ type: "number", value: 9 }); await undo();
  await select("A1:B1"); await button("复制选区").click(); await page.getByLabel("区域粘贴方式", { exact: true }).selectOption("values"); await select("A2"); await decide(() => button("粘贴选区").click(), [true, true]);
  expect((await inspect()).history.undoSteps).toBe(1); expect((await cell(1, 1))?.value).toEqual({ type: "number", value: 9 }); await undo();
  await select("A2");
  await decide(() => input.evaluate(element => {
    const data = new DataTransfer(); data.setData("text/plain", "9\t9");
    const event = new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: data }); element.dispatchEvent(event);
  }), [true, true]);
  expect((await inspect()).history.undoSteps).toBe(1); expect((await cell(1, 1))?.value).toEqual({ type: "number", value: 9 }); await undo();
  for (const address of ["C2", "D2"]) {
    await select(address); await input.fill("9"); await decide(() => button("确认输入").click(), []);
    expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, pending: true, input: { text: "9", dirty: true } }); await button("取消输入").click(); await settled();
  }
  await select("E2"); await input.fill("1000000000"); await decide(() => button("确认输入").click(), [true]);
  expect((await inspect()).history.undoSteps).toBe(1); expect((await cell(1, 4))?.value).toEqual({ type: "number", value: 1000000000 }); await undo();
  // A custom synchronous host dialog can reenter the input. A bulk retry
  // must retain that newer draft and refuse to publish its old command.
  await select("A1:A3");
  await input.evaluate(element => {
    const original = window.confirm;
    window.confirm = () => { window.confirm = original; (element as HTMLTextAreaElement).value = "继续输入的草稿"; element.dispatchEvent(new Event("input", { bubbles: true })); return true; };
  });
  await button("向下填充").click(); await settled();
  expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, pending: true, input: { text: "继续输入的草稿", dirty: true } }); expect(await cell(1, 0)).toBeUndefined();
  await button("取消输入").click(); await settled();
}
