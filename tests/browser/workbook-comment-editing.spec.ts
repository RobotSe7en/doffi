import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { WorkbookSnapshot } from "@doffi/core";
import { importXlsx, writeDocumentZip } from "@doffi/office";
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
for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} native dual note boxes follow sorted and cross-sheet cut cells through actual font measurement`, async ({ page }) => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync("packages/office/tests/fixtures/xlsx-dual-note-boxes-conflict.xlsx"), { documentId: "browser-note-transport" }), { kind: "workbook" }));
  const source = s.snapshot().metadata.workbook.sheets[0]!, original = s.snapshot().metadata.workbook.comments![0]!;
  if (original.format !== "classic") throw Error("Expected original native classic note");
  const initial = s.snapshot(), bodyId = initial.annotations[0]!.bodyNodeId!, files = s.serialize(); s.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes, true), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(files));
  // Prepare the sort keys and destination through the mounted production
  // measurer too: the imported alternate note box has a fixed size response.
  await page.evaluate(({ remote, sheetId, original }) => {
    const view = remote ? window.workbookRemoteEditor : window.workbookEditor;
    return view.perform([{ op: "insertWorkbookSheet", name: "Other" }, { op: "setWorkbookNoteBox", annotationId: original.annotationId, box: { ...original.box!, visible: true } },
      { op: "setWorkbookCells", sheetId, cells: [3, 1, 2].map((value, index) => ({ row: index + 1, column: 1, source: { kind: "value" as const, value: { type: "number" as const, value } } })) }]).then(() => {});
  }, { remote, sheetId: source.sheetId, original });
  const state = await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { sheets: WorkbookSnapshot["metadata"]["workbook"]["sheets"] };
  const otherId = state.sheets.at(-1)!.sheetId;
  await page.evaluate(({ remote, otherId }) => {
    const view = remote ? window.workbookRemoteEditor : window.workbookEditor;
    return view.perform([{ op: "updateWorkbookSheet", sheetId: otherId, changes: { defaultRowHeightPt: 30, defaultRowHeightCustom: true, defaultColumnWidth: 12 } }]).then(() => {});
  }, { remote, otherId });
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as Inspection & { history: { undoSteps: number } };
  const choose = async (value: string) => { const address = page.getByLabel("选择单元格或区域", { exact: true }); await address.fill(value); await address.press("Enter"); await settled(); };
  const geometry = async (row: number, column: number) => page.evaluate(({ remote, row, column, bodyId }) => {
    const view = remote ? window.workbookRemoteEditor : window.workbookEditor, box = view.displayedDrawingBounds().objects.find(object => object.nodeId === bodyId);
    if (!box) throw Error("The selected native note is not visible");
    const cell = view.displayedCellBounds(row, column); return { ...box, offsetX: box.xPt - cell.xPt, offsetY: box.yPt - cell.yPt };
  }, { remote, row, column, bodyId });
  const undo = async () => { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); };
  await choose("B2"); const before = await geometry(1, 1), beforeState = await inspect();
  await page.getByRole("button", { name: "剪切选区", exact: true }).click(); await page.getByRole("tab", { name: "Other", exact: true }).click(); await settled();
  await choose("H20"); await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: beforeState.history.undoSteps + 1 }, annotations: [{ anchor: { sheetId: otherId } }] });
  // The box starts to the right of H20; reveal it while keeping its owning
  // cell visible, rather than measuring an object outside the viewport.
  await choose("L24");
  const after = await geometry(19, 7);
  for (const key of ["offsetX", "offsetY", "widthPt", "heightPt"] as const) expect(after[key]).toBeCloseTo(before[key], 1);
  expect(await inspect()).toMatchObject({ errors: [], content: beforeState.content, history: { undoSteps: beforeState.history.undoSteps + 1 }, annotations: [{ bodyNodeId: bodyId, anchor: { sheetId: otherId, ranges: [{ from: { row: 19, column: 7 }, to: { row: 19, column: 7 } }] } }] });
  const current = (await inspect()).comments[0]!; if (current.format !== "classic") throw Error("Native binding changed kind");
  expect(current.commentPrBox).toMatchObject({ anchor: { kind: "two-cell", moveWithCells: true, sizeWithCells: false } });
  expect(current.boxSource).toBe("vml"); await undo(); expect(await inspect()).toMatchObject({ comments: beforeState.comments, history: { undoSteps: beforeState.history.undoSteps } });
  await page.getByRole("tab", { name: source.name, exact: true }).click(); await settled(); await choose("B2:C4");
  await page.getByRole("button", { name: "升序排序", exact: true }).click(); await settled(); await choose("B4");
  const sorted = await geometry(3, 1); for (const key of ["offsetX", "offsetY", "widthPt", "heightPt"] as const) expect(sorted[key]).toBeCloseTo(before[key], 1);
  expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: beforeState.history.undoSteps + 1 }, annotations: [{ bodyNodeId: bodyId, anchor: { ranges: [{ from: { row: 3, column: 1 }, to: { row: 3, column: 1 } }] } }] });
  await undo(); expect(await inspect()).toMatchObject({ comments: beforeState.comments, history: { undoSteps: beforeState.history.undoSteps }, errors: [] });
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, reopenedEqual: true }); else await page.evaluate(() => window.workbookEditor.dispose());
});
for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} comment drafts retain their message, unrelated fields, IME input and revision until explicit discard`, async ({ page }) => {
  const s = createDocumentSession(scalarWorkbook([])), manifest = s.snapshot().manifest, anchor = { kind: "cells" as const, sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] };
  s.commit(s.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [
    { op: "addWorkbookComment", anchor, html: "<p>原正文</p>", author: { name: "王" } }, { op: "addWorkbookComment", anchor, html: "<p>其他批注</p>", author: { name: "李" } },
  ] }));
  const [first, second] = s.snapshot().annotations;
  s.commit(s.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations: [{ op: "addWorkbookCommentReply", annotationId: first!.annotationId, html: "<p>原回复</p>", author: { name: "张" } }] }));
  const replyId = s.snapshot().annotations[0]!.replies![0]!.replyId, files = s.serialize(); s.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as Inspection & { pending: boolean; history: { undoSteps: number } };
  await page.getByText("备注与线程批注", { exact: true }).click(); const threads = page.getByLabel("选择工作簿批注", { exact: true }), replies = page.getByLabel("选择批注消息", { exact: true }), body = page.getByLabel("批注正文", { exact: true }), author = page.getByLabel("批注作者", { exact: true });
  const save = page.getByRole("button", { name: "更新批注作者时间与状态", exact: true }), discard = page.getByRole("button", { name: "放弃批注表单草稿", exact: true });
  await threads.selectOption(first!.annotationId); await body.fill("保存正文禰󠄀😀"); await author.fill("未保存作者"); expect(await inspect()).toMatchObject({ pending: true, history: { undoSteps: 0 } });
  await replies.selectOption(replyId); await expect(replies).toHaveValue(""); await expect(body).toHaveValue("保存正文禰󠄀😀");
  await threads.selectOption(second!.annotationId); await expect(threads).toHaveValue(first!.annotationId); await expect(author).toHaveValue("未保存作者");
  await page.getByRole("button", { name: "保存批注文字", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ pending: true, history: { undoSteps: 1 }, annotations: [expect.objectContaining({ author: { name: "王" } }), expect.anything()] });
  await expect(author).toHaveValue("未保存作者"); await expect(body).toHaveValue("保存正文禰󠄀😀"); await expect(save).toBeDisabled();
  await discard.click(); await settled(); expect(await inspect()).toMatchObject({ pending: false }); await expect(author).toHaveValue("王"); await expect(body).toHaveValue("保存正文禰󠄀😀");
  await author.evaluate(element => { const field = element as HTMLInputElement; field.focus(); field.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true })); field.value = "中文草稿"; field.dispatchEvent(new InputEvent("input", { bubbles: true, isComposing: true })); });
  await expect(author).toBeEnabled(); await expect(save).toBeDisabled(); await expect(threads).toBeDisabled(); await author.dispatchEvent("compositionend");
  await page.evaluate(async remote => { if (remote) await window.workbookRemoteEditor.external("external"); else { window.workbookEditor.external("external"); await window.workbookEditor.settled(); } }, remote);
  await expect(author).toHaveValue("中文草稿"); await expect(save).toBeDisabled(); expect(await inspect()).toMatchObject({ pending: true });
  await discard.click(); await settled(); expect(await inspect()).toMatchObject({ pending: false, errors: [] });
  await replies.selectOption(replyId); await expect(body).toHaveValue("原回复"); await expect(page.getByLabel("批注状态", { exact: true })).toBeDisabled();
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, reopenedEqual: true, endpointDisposed: true }); else await page.evaluate(() => window.workbookEditor.dispose());
});

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
