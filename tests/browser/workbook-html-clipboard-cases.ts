import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";

export async function exerciseWorkbookHtmlClipboard(page: Page, remote: boolean): Promise<void> {
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const state = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { revision: number; history: { undoSteps: number }; errors: unknown[]; merges: unknown[] };
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
  const paste = (html: string) => input.evaluate((field, html) => {
    const data = new DataTransfer(); data.setData("text/plain", "plain fallback must not override the HTML"); data.setData("text/html", html);
    const event = new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true });
    Object.defineProperty(event, "clipboardData", { value: data });
    return !field.dispatchEvent(event);
  }, html);
  const html = '<style>.amount{mso-number-format:"0.00";background-color:#ffeeaa}</style><table><tr><td><b>中文</b>😀禰󠄀</td><td class="amount" x:num="12.5">12.50</td></tr><tr><td colspan="2">合并</td></tr></table>';
  expect(await paste(html)).toBe(true); await settled();
  expect(await state()).toMatchObject({ history: { undoSteps: 1 }, errors: [], merges: [{ from: { row: 1, column: 0 }, to: { row: 1, column: 1 } }] });
  expect(await page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCell(0, 0) : window.workbookEditor.displayedCell(0, 0), remote)).toBe("中文😀禰󠄀");
  expect(await page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCell(0, 1) : window.workbookEditor.displayedCell(0, 1), remote)).toBe("12.50");
  if (!remote) expect((await page.evaluate(() => window.workbookEditor.displayedTextStyles(0, 0))).some(style => style.text === "中文" && style.weight === "700")).toBe(true);
  const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true });
  const choose = async (value: string) => { await address.fill(value); await address.press("Enter"); await settled(); };
  await choose("A1:B2");
  for (const kind of ["copy", "cut"] as const) {
    const before = await state();
    const copied = await input.evaluate((field, kind) => {
      const data = new DataTransfer(), event = new ClipboardEvent(kind, { clipboardData: data, bubbles: true, cancelable: true });
      Object.defineProperty(event, "clipboardData", { value: data });
      return { handled: !field.dispatchEvent(event), text: data.getData("text/plain"), html: data.getData("text/html") };
    }, kind);
    expect(copied.handled).toBe(true); expect(copied.text).toBe("中文😀禰󠄀\t12.50\r\n合并\t"); expect(copied.html).toContain('colspan="2"');
    expect(copied.html).not.toMatch(/data-node-id|data-sheet-id|data-cache-state/);
    expect((await state()).revision).toBe(before.revision); expect((await state()).history).toEqual(before.history);
    await choose("A5");
    expect(await input.evaluate((field, data) => {
      const transfer = new DataTransfer(); transfer.setData("text/plain", data.text); transfer.setData("text/html", data.html);
      const event = new ClipboardEvent("paste", { clipboardData: transfer, bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: transfer }); return !field.dispatchEvent(event);
    }, copied)).toBe(true); await settled();
    expect(await state()).toMatchObject({ history: { undoSteps: 2 }, errors: [] });
    expect(await page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCell(4, 0) : window.workbookEditor.displayedCell(4, 0), remote)).toBe("中文😀禰󠄀");
    await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); await choose("A1:B2");
    expect(await state()).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  }
  // Trusted keyboard events cross the browser's actual system clipboard.
  await input.press("Control+x"); await settled(); expect((await state()).history.undoSteps).toBe(1);
  await choose("A5"); await input.press("Control+v"); await settled();
  expect(await state()).toMatchObject({ history: { undoSteps: 2 }, errors: [] });
  expect(await page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCell(4, 0) : window.workbookEditor.displayedCell(4, 0), remote)).toBe("中文😀禰󠄀");
  await choose("A1");
  expect(await page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCell(0, 0) : window.workbookEditor.displayedCell(0, 0), remote)).toBe("");
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); await choose("A1:B2");
  // Dirty/IME input keeps the textarea's own paste path and never starts a grid transaction.
  await input.press("F2"); await input.fill("pending😀"); expect(await paste(html)).toBe(false); expect((await state()).history.undoSteps).toBe(1);
  await input.press("Escape"); await settled();
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect(await state()).toMatchObject({ history: { undoSteps: 0 }, errors: [], merges: [] });
  const revision = (await state()).revision;
  expect(await paste('<table><tr><td onclick="window.__clipboardExecuted=true">unsafe</td></tr></table>')).toBe(true); await settled();
  const failed = await state(); expect(failed).toMatchObject({ revision, history: { undoSteps: 0 }, merges: [] }); expect(failed.errors).toHaveLength(1);
  expect(await page.evaluate(() => (window as Window & { __clipboardExecuted?: boolean }).__clipboardExecuted)).toBeUndefined();
  const operations = [
    { op: "setWorkbookCells" as const, sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value" as const, value: { type: "text" as const, value: "ABCD" } } }] },
    { op: "formatWorkbookRanges" as const, sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], properties: { "text-decoration-line": "underline" } },
    { op: "formatWorkbookText" as const, sheetId: "sheet-1", at: { row: 0, column: 0 }, from: 1, to: 3, expectedText: "BC", properties: { "text-decoration-line": "none" } },
  ];
  await page.evaluate(async ({ remote, operations }) => { if (remote) await window.workbookRemoteEditor.perform(operations); else await window.workbookEditor.perform(operations); }, { remote, operations });
  const decorations = async (column: number) => (await page.evaluate(({ remote, column }) => remote ? window.workbookRemoteEditor.displayedCellDecorations(0, column) : window.workbookEditor.displayedCellDecorations(0, column), { remote, column })).flatMap(run => [...run.text].map(text => ({ text, lines: run.lines })));
  const expected = [{ text: "A", lines: ["underline"] }, { text: "B", lines: [] }, { text: "C", lines: [] }, { text: "D", lines: ["underline"] }];
  expect(await decorations(0)).toEqual(expected);
  await choose("A1");
  const data = await input.evaluate(field => {
    const transfer = new DataTransfer(), event = new ClipboardEvent("copy", { clipboardData: transfer, bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: transfer });
    field.dispatchEvent(event); return { text: transfer.getData("text/plain"), html: transfer.getData("text/html") };
  });
  await choose("D1");
  await input.evaluate((field, data) => {
    const transfer = new DataTransfer(); transfer.setData("text/plain", data.text); transfer.setData("text/html", data.html + "\n");
    const event = new ClipboardEvent("paste", { clipboardData: transfer, bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: transfer }); field.dispatchEvent(event);
  }, data); await settled(); expect(await decorations(3)).toEqual(expected);
  expect((await state()).errors).toHaveLength(1);
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled();
  expect((await state()).history.undoSteps).toBe(0);
}
