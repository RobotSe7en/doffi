import { fileURLToPath } from "node:url";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { createDocumentSession } from "@doffi/core";
import type { WorkbookSnapshot } from "@doffi/core";
import { writeDocumentZip } from "@doffi/office";
import { scalarWorkbook } from "../../packages/core/tests/workbook-scalar-fixture.js";
import { encodeTestPackage } from "./package-transport.js";
import { openSdkHost } from "./sdk-host.js";
import type * as directFixture from "./fixtures/workbook-editor.js";
import type * as remoteFixture from "./fixtures/workbook-remote-editor.js";
declare global { interface Window { workbookEditor: typeof directFixture; workbookRemoteEditor: typeof remoteFixture } }
const bundles = new Map<string, string>();
test.beforeEach(async ({ page }) => { await openSdkHost(page); });
test.beforeAll(async () => {
  for (const name of ["workbook-editor", "workbook-remote-editor", "workbook-session-worker"]) {
    const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/" + name + ".ts", import.meta.url)), formats: ["iife"], name: "WorkbookComment" } } });
    const output = Array.isArray(result) ? result.flatMap(part => part.output) : "output" in result ? result.output : [];
    bundles.set(name, output.filter(part => part.type === "chunk").map(part => part.code).join("\n"));
  }
});
interface Inspection { readonly annotations: WorkbookSnapshot["annotations"]; readonly comments: NonNullable<WorkbookSnapshot["metadata"]["workbook"]["comments"]>; readonly content: string; readonly errors: readonly string[] }
for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} comment controls preserve rich text/CRLF/mentions and manage classic boxes, replies, recovery and history`, async ({ page, browserName }) => {
  const s = createDocumentSession(scalarWorkbook([])), html = '<p><strong>A&#13;<br>禰󠄀😀</strong><em>Z</em></p>';
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0, operations: [{ op: "addWorkbookComment", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] }, html, author: { name: "王" } }] }));
  const root = s.getAnnotations()[0]!, originalBody = s.getNode(root.bodyNodeId!)!, authorId = s.snapshot().metadata.workbook.comments![0]!.authorId!;
  const strongId = s.getNode(originalBody.childNodeIds[0]!)!.childNodeIds.find(id => s.getNode(id)!.tagName === "strong")!;
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 1, operations: [{ op: "setWorkbookCommentBody", annotationId: root.annotationId, html, mentions: [{ mentionId: "retained", authorId, startIndex: 3, length: 5 }] }] }));
  const files = s.serialize(); s.dispose(); await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as Inspection;
  await page.getByText("备注与线程批注", { exact: true }).click();
  const threads = page.getByLabel("选择工作簿批注", { exact: true }), replies = page.getByLabel("选择批注消息", { exact: true }), text = page.getByLabel("批注正文", { exact: true });
  await threads.selectOption({ label: "王 · 1" }); await expect(text).toHaveValue("A\n禰󠄀😀Z");
  await text.fill("中文\n禰󠄀😀Z"); await page.getByRole("button", { name: "保存批注文字", exact: true }).click(); await settled();
  let result = await inspect(); expect(result.comments[0]).toMatchObject({ mentions: [{ startIndex: 4, length: 5 }] });
  expect(result.content).toContain('<strong data-node-id="' + strongId + '">');
  expect(result.content).toContain("中文&#13;<br"); expect(result.content).toContain("<em data-node-id=");
  if (browserName === "chromium") {
    const before = result.content, cdp = await page.context().newCDPSession(page);
    await text.focus(); await text.evaluate(element => (element as HTMLTextAreaElement).setSelectionRange(0, 0));
    await cdp.send("Input.imeSetComposition", { text: "预编辑", selectionStart: 3, selectionEnd: 3 }); await expect(page.getByRole("button", { name: "保存批注文字", exact: true })).toBeDisabled();
    expect((await inspect()).content).toBe(before);
    await cdp.send("Input.insertText", { text: "输入法禰󠄀😀" }); await expect(page.getByRole("button", { name: "保存批注文字", exact: true })).toBeEnabled();
    await page.getByRole("button", { name: "保存批注文字", exact: true }).click(); await settled(); expect((await inspect()).content).toContain("输入法禰󠄀😀中文&#13;");
    await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect((await inspect()).content).toBe(before); await cdp.detach();
  } else {
    await text.dispatchEvent("compositionstart", { data: "预编辑" }); await expect(page.getByRole("button", { name: "保存批注文字", exact: true })).toBeDisabled();
    await text.dispatchEvent("compositionend", { data: "" }); await expect(page.getByRole("button", { name: "保存批注文字", exact: true })).toBeEnabled();
  }
  const address = page.getByLabel("选择单元格或区域", { exact: true }); await address.fill("B3"); await address.press("Enter"); await settled();
  await page.getByLabel("批注格式", { exact: true }).selectOption("classic"); await page.getByLabel("批注作者", { exact: true }).fill("李"); await text.fill("经典备注禰󠄀😀");
  await page.getByRole("button", { name: "在选区添加批注", exact: true }).click(); await settled();
  await threads.selectOption({ label: "李 · 2" }); await page.getByLabel("显示经典备注框", { exact: true }).check(); await page.getByLabel("打印经典备注框", { exact: true }).check();
  await page.getByRole("button", { name: "应用经典备注框显示", exact: true }).click(); await settled();
  expect((await inspect()).comments[1]).toMatchObject({ format: "classic", box: { visible: true, printable: true } });
  await threads.selectOption({ label: "王 · 1" }); await text.fill("回复禰󠄀😀"); await page.getByLabel("批注作者", { exact: true }).fill("回复者");
  await page.getByRole("button", { name: "添加线程回复", exact: true }).click(); await settled();
  await replies.selectOption({ label: "回复者 · 1" }); await expect(text).toHaveValue("回复禰󠄀😀"); await text.fill("已复核禰󠄀😀");
  await page.getByRole("button", { name: "保存批注文字", exact: true }).click(); await settled();
  result = await inspect(); expect(result.annotations[0]!.replies).toHaveLength(1); expect(result.content).toContain("已复核禰󠄀😀");
  await page.getByRole("button", { name: "删除所选回复", exact: true }).click(); await settled(); expect((await inspect()).annotations[0]!.replies).toBeUndefined();
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect((await inspect()).annotations[0]!.replies).toHaveLength(1);
  await replies.selectOption(""); await text.fill("保留的批注草稿");
  await page.evaluate(remote => remote ? window.workbookRemoteEditor.external("external") : window.workbookEditor.external("external"), remote); await settled();
  await expect(page.getByLabel("批注编辑状态", { exact: true })).toContainText("批注草稿已保留"); await expect(text).toHaveValue("保留的批注草稿"); await expect(page.getByRole("button", { name: "保存批注文字", exact: true })).toBeDisabled();
  await page.getByRole("button", { name: "重新读取批注", exact: true }).click(); await expect(text).toHaveValue("中文\n禰󠄀😀Z");
  await page.getByLabel("批注状态", { exact: true }).selectOption("needs_reanchor"); await page.getByRole("button", { name: "更新批注作者时间与状态", exact: true }).click(); await settled();
  await expect(page.getByLabel("批注编辑状态", { exact: true })).toContainText("显式重新定位");
  await address.fill("D5:E6"); await address.press("Enter"); await settled(); await page.getByLabel("批注状态", { exact: true }).selectOption("open");
  await page.getByRole("button", { name: "将批注重新锚定到选区", exact: true }).click(); await settled();
  expect((await inspect()).annotations[0]).toMatchObject({ status: "open", anchor: { kind: "cells", ranges: [{ from: { row: 4, column: 3 }, to: { row: 5, column: 4 } }] } });
  await page.getByRole("button", { name: "删除整条批注", exact: true }).click(); await settled(); expect((await inspect()).annotations).toHaveLength(1);
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); result = await inspect(); expect(result.annotations).toHaveLength(2); expect(result.annotations[0]!.bodyNodeId).toBe(root.bodyNodeId); expect(result.errors).toEqual([]);
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
  else await page.evaluate(() => window.workbookEditor.dispose());
});

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} grid structure and clipboard controls move a source-owned thread and restore its exact body/history`, async ({ page }) => {
  const s = createDocumentSession(scalarWorkbook([]));
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0, operations: [{ op: "addWorkbookComment", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 1, column: 0 }, to: { row: 1, column: 0 } }] }, html: "<p><strong>结构禰󠄀😀</strong></p>", author: { name: "结构批注" } }] }));
  const root = s.getAnnotations()[0]!, body = s.getNode(root.bodyNodeId!)!.html, files = s.serialize(); s.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as Inspection;
  const address = page.getByLabel("选择单元格或区域", { exact: true }), choose = async (value: string) => { await address.fill(value); await address.press("Enter"); await settled(); };
  await choose("A2"); await page.getByRole("button", { name: "插入选区行", exact: true }).click(); await settled();
  expect((await inspect()).annotations[0]!.anchor).toMatchObject({ kind: "cells", ranges: [{ from: { row: 2, column: 0 }, to: { row: 2, column: 0 } }] });
  await choose("A3"); await page.getByRole("button", { name: "剪切选区", exact: true }).click(); await settled();
  await choose("AG257"); await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
  expect((await inspect()).annotations[0]!.anchor).toMatchObject({ kind: "cells", ranges: [{ from: { row: 256, column: 32 }, to: { row: 256, column: 32 } }] });
  await choose("AG257"); await page.getByRole("button", { name: "删除选区行", exact: true }).click(); await settled();
  expect((await inspect()).annotations[0]).toMatchObject({ status: "needs_reanchor", bodyNodeId: root.bodyNodeId });
  for (let index = 0; index < 3; index++) { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); }
  const restored = await inspect(); expect(restored.annotations[0]).toMatchObject({ status: "open", anchor: root.anchor }); expect(restored.content).toContain(body); expect(restored.errors).toEqual([]);
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
  else await page.evaluate(() => window.workbookEditor.dispose());
});
