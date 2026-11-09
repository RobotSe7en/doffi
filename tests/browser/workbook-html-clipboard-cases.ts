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
  // Dirty/IME input keeps the textarea's own paste path and never starts a grid transaction.
  await input.press("F2"); await input.fill("pending😀"); expect(await paste(html)).toBe(false); expect((await state()).history.undoSteps).toBe(1);
  await input.press("Escape"); await settled();
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect(await state()).toMatchObject({ history: { undoSteps: 0 }, errors: [], merges: [] });
  const revision = (await state()).revision;
  expect(await paste('<table><tr><td onclick="window.__clipboardExecuted=true">unsafe</td></tr></table>')).toBe(true); await settled();
  const failed = await state(); expect(failed).toMatchObject({ revision, history: { undoSteps: 0 }, merges: [] }); expect(failed.errors).toHaveLength(1);
  expect(await page.evaluate(() => (window as Window & { __clipboardExecuted?: boolean }).__clipboardExecuted)).toBeUndefined();
}
