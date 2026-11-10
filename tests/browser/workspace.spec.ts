import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { openSdkHost } from './sdk-host.js';
import { openDocument } from "@doffi/core";
import { importPptx, readDocumentZip } from "@doffi/office";
import type {} from "./fixtures/interaction-host.js";

async function report(page: Page): Promise<void> {
  await page.addInitScript(() => { localStorage.setItem("doffi:library", "open"); localStorage.setItem("doffi:editing", "preview"); });
  await page.goto("/");
  await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]')).toBeVisible();
}
async function selectText(page: Page, from = 2, to = 8): Promise<void> {
  await page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]').evaluate((node, { from, to }) => {
    const range = document.createRange(); range.setStart(node.firstChild!, from); range.setEnd(node.firstChild!, to);
    const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range);
    document.dispatchEvent(new Event("selectionchange"));
  }, { from, to });
  if (await page.locator('.office-app').count()) {
    await page.getByRole('tab', { name: '审阅', exact: true }).click();
    await page.getByRole('button', { name: '新建批注', exact: true }).click();
  }
}

test("homepage drop opens a document and home navigation preserves its editing history", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "样例文档", exact: true })).toBeVisible();
  const bytes = [...readFileSync(new URL("../../examples/docx/01-text-numbering.docx", import.meta.url))];
  await page.locator(".upload-zone").evaluate((node, bytes) => {
    const transfer = new DataTransfer();
    transfer.items.add(new File([new Uint8Array(bytes)], "dropped.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    node.dispatchEvent(new DragEvent("drop", { dataTransfer: transfer, bubbles: true, cancelable: true }));
  }, bytes);
  await expect(page.getByRole("heading", { name: "dropped.docx", exact: true })).toBeVisible();
  await expect(page.getByRole("complementary", { name: "文档列表", exact: true })).toBeHidden();
  const heading = page.frameLocator('iframe[title="正文编辑"]').getByRole("heading", { name: "实施报告", exact: true });
  await heading.click();
  await heading.evaluate(node => {
    const range = document.createRange(); range.selectNodeContents(node); range.collapse(false);
    const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
    (document.querySelector("[data-office-editor]") as HTMLElement).focus();
    document.dispatchEvent(new Event("selectionchange"));
  });
  await page.keyboard.insertText("已编辑");
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await page.getByRole("button", { name: "返回首页", exact: true }).click();
  await expect(page.getByRole("heading", { name: "样例文档", exact: true })).toBeVisible();
  await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  await page.getByRole("button", { name: "显示文档列表", exact: true }).click();
  await page.getByRole("button", { name: "打开 dropped.docx", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await expect(page.frameLocator('iframe[title="正文编辑"]').getByRole("heading", { name: "实施报告已编辑", exact: true })).toBeVisible();
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.frameLocator('iframe[title="正文编辑"]').getByRole("heading", { name: "实施报告", exact: true })).toBeVisible();
});

test("Office ribbon keeps unavailable tools disabled and formats through the document history", async ({ page }) => {
  await page.addInitScript(() => localStorage.setItem("doffi:editing", "preview"));
  await page.goto("/"); await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  await expect(page.getByRole("tab", { name: "开始", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByRole("group", { name: "字体", exact: true })).toBeHidden();
  await expect(page.getByRole("tab", { name: "插入", exact: true })).toBeDisabled();
  await page.getByRole("tab", { name: "开始", exact: true }).click();
  await page.getByRole("button", { name: "编辑正文", exact: true }).click();
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="n-summary"]');
  await paragraph.evaluate(node => {
    const range = document.createRange(); range.selectNodeContents(node);
    const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
    (document.querySelector("[data-office-editor]") as HTMLElement).focus(); document.dispatchEvent(new Event("selectionchange"));
  });
  await page.getByRole("button", { name: "加粗", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await expect(paragraph.locator("span")).toHaveCSS("font-weight", "700");
  await page.getByRole("button", { name: "居中", exact: true }).click();
  await expect(paragraph).toHaveCSS("text-align", "center");
  await page.getByRole("button", { name: "返回预览与批注", exact: true }).click();
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "4");
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]')).toHaveText("项目计划于第三季度完成。");
});

test("document library filters and switches between isolated sessions while retaining history", async ({ page }) => {
  await report(page);
  await expect(page).toHaveTitle("doffi Playground");
  await expect(page.getByRole("link", { name: "doffi Playground", exact: true })).toBeVisible();
  const sidebar = page.getByRole("complementary", { name: "文档列表", exact: true });
  expect((await sidebar.boundingBox())!.x).toBeLessThan((await page.getByRole("region", { name: "文档工作区", exact: true }).boundingBox())!.x);
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "载入固定示例提案", exact: true }).click();
  await page.getByRole("button", { name: "整批应用提案", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await page.getByRole("group", { name: "筛选文档" }).getByRole("button", { name: "PPT", exact: true }).click();
  await expect(sidebar.getByRole("button", { name: "打开报告协作样例", exact: true })).toHaveCount(0);
  await page.getByRole("button", { name: "打开 PPT 协作样例", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await page.getByRole("group", { name: "筛选文档" }).getByRole("button", { name: "全部", exact: true }).click();
  await page.getByLabel("搜索文档", { exact: true }).fill("季度");
  await expect(sidebar.getByRole("navigation").getByRole("button")).toHaveCount(1);
  await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]')).toContainText("第三季度");
});

test("selection opens a bounded comment dialog; cancellation, IME and keyboard submission preserve transaction semantics", async ({ page }, info) => {
  await report(page); await selectText(page);
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true });
  await expect(dialog).toBeVisible(); await expect(dialog).toContainText("计划于第三季");
  const input = dialog.getByLabel("新批注内容", { exact: true });
  await input.fill("取消的草稿"); await input.press("Escape");
  await expect(dialog).toBeHidden(); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await selectText(page); await expect(input).toHaveValue(""); await input.fill("补充时间与负责人");
  await input.evaluate(node => node.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", ctrlKey: true, isComposing: true, bubbles: true, cancelable: true })));
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  const box = (await dialog.boundingBox())!, viewport = page.viewportSize()!;
  expect(box.x).toBeGreaterThanOrEqual(0); expect(box.y).toBeGreaterThanOrEqual(0);
  const toolbar = (await page.getByRole("toolbar", { name: "文件与导出", exact: true }).boundingBox())!;
  expect(box.y).toBeGreaterThanOrEqual(toolbar.y + toolbar.height);
  expect(box.x + box.width).toBeLessThanOrEqual(viewport.width); expect(box.y + box.height).toBeLessThanOrEqual(viewport.height);
  await page.screenshot({ path: info.outputPath("word-comment.png") });
  await input.press("Control+Enter"); await expect(dialog).toBeHidden();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  const pending = page.waitForEvent("download"); await page.getByRole("button", { name: "下载文档包", exact: true }).click();
  const path = info.outputPath("comment.zip"); await (await pending).saveAs(path);
  const model = openDocument(readDocumentZip(new Uint8Array(readFileSync(path))));
  expect(model.snapshot().annotations.at(-1)).toMatchObject({ instruction: "补充时间与负责人", anchor: { kind: "text", nodeId: "n-summary", from: 2, to: 8, quote: "计划于第三季" } });
  await selectText(page); await input.fill("过期草稿");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(dialog).toBeHidden(); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
  await page.getByRole("button", { name: "重做", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "3");
});

test("object and rectangle selections use the same floating composer", async ({ page }) => {
  await report(page);
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "对象选择", exact: true }).click();
  await page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-title"]').click();
  await page.getByRole("button", { name: "新建批注", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true });
  await expect(dialog).toContainText("对象：项目实施报告");
  const input = dialog.getByLabel("新批注内容");
  // A non-modal composer opens without focus; enter it as a user would. Firefox
  // keeps the iframe's native text selection until this explicit focus transfer.
  await input.click(); await expect(input).toBeFocused(); await input.fill("对象批注");
  await dialog.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await page.getByRole("button", { name: "矩形框选", exact: true }).click();
  const paragraph = await page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]').boundingBox();
  await page.mouse.move(paragraph!.x + 1, paragraph!.y + 1); await page.mouse.down();
  await page.mouse.move(paragraph!.x + 100, paragraph!.y + 20, { steps: 4 });
  await expect(dialog).toBeHidden();
  await page.mouse.up();
  await page.getByRole("button", { name: "新建批注", exact: true }).click();
  await expect(dialog).toBeVisible(); await input.click(); await input.fill("区域批注");
  await dialog.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
});

test("PPT comment popover preserves native comments and can be undone, exported and reopened", async ({ page }, info) => {
  await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto("/");
  await page.getByRole("button", { name: "打开 PPT 协作样例", exact: true }).click();
  const frame = page.frameLocator('iframe[title="幻灯片只读预览"]');
  await frame.locator('[data-node-id="heading-p"]').click();
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "新建批注", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true }); await expect(dialog).toBeVisible();
  await dialog.getByLabel("批注显示署名", { exact: true }).fill("审阅人");
  await dialog.getByLabel("新批注正文", { exact: true }).fill("确认演示标题");
  await dialog.getByRole("button", { name: "添加批注", exact: true }).click(); await expect(dialog).toBeHidden();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
  await page.getByRole("button", { name: "重做", exact: true }).click();
  await page.getByRole("tab", { name: "文件", exact: true }).click();
  const pending = page.waitForEvent("download"); await page.getByRole("button", { name: "导出 PPTX", exact: true }).click();
  const path = info.outputPath("comment.pptx"); await (await pending).saveAs(path);
  const model = openDocument(importPptx(new Uint8Array(readFileSync(path)), { documentId: "workspace-comment" }));
  expect(model.snapshot().annotations).toHaveLength(2);
  expect(model.snapshot().annotations.map(annotation => annotation.bodyNodeId ? model.getNode(annotation.bodyNodeId)?.text : annotation.instruction)).toContain("确认演示标题");
  expect(model.snapshot().annotations.at(-1)?.author?.name).toBe("审阅人");
});

test("switching documents discards comment drafts without changing either session", async ({ page }) => {
  await report(page); await selectText(page);
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true });
  await dialog.getByLabel("新批注内容", { exact: true }).fill("不能带到另一文档");
  await page.getByRole("button", { name: "打开 PPT 协作样例", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await expect(dialog).toBeHidden();
  await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await selectText(page);
  await expect(dialog.getByLabel("新批注内容", { exact: true })).toHaveValue("");
});

test("comment popover follows scrolling inside the canvas shadow tree", async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 1000 });
  await report(page); await selectText(page);
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true });
  await expect(dialog).toBeVisible();
  const before = (await dialog.boundingBox())!;
  const scrolled = await page.locator(".stage").evaluate(node => { node.scrollTop = 80; return node.scrollTop; });
  expect(scrolled).toBeGreaterThan(0);
  await expect.poll(async () => Math.abs((await dialog.boundingBox())!.y - before.y)).toBeGreaterThan(10);
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
});

test("workspace unmount releases views and floating controls while the host retains its session", async ({ page }) => {
  const built = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: {
    entry: fileURLToPath(new URL("./fixtures/interaction-host.ts", import.meta.url)), formats: ["es"], fileName: "interaction-host",
  } } });
  if (!("output" in built) && !Array.isArray(built)) throw new Error("Missing fixture output");
  const output = Array.isArray(built) ? built.flatMap(item => item.output) : built.output;
  await page.goto("/");
  await page.addScriptTag({ type: "module", content: output.flatMap(item => item.type === "chunk" ? [item.code] : []).join("\n") });
  await page.waitForFunction(() => Boolean(window.interactionHost));
  const files = ["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(path => [path, [...readFileSync(new URL("../../examples/report/" + path, import.meta.url))]] as [string, number[]]);
  for (let iteration = 0; iteration < 3; iteration++) {
    await page.evaluate(files => window.interactionHost.mount(files, false, "workspace"), files);
    await selectText(page);
    await expect(page.getByRole("dialog", { name: "新建批注", exact: true })).toBeVisible();
    await page.evaluate(() => { window.interactionHost.view!.dispose(); window.interactionHost.view!.dispose(); });
    await expect(page.locator("iframe")).toHaveCount(0);
    await expect(page.getByRole("dialog", { name: "新建批注", exact: true })).toHaveCount(0);
    expect(await page.evaluate(() => ({ revision: window.interactionHost.session!.snapshot().manifest.revision, text: window.interactionHost.session!.getNode("n-summary")!.text, errors: window.interactionHost.errors }))).toMatchObject({ revision: 0, errors: [], text: expect.stringContaining("第三季度") });
  }
  await page.evaluate(() => window.interactionHost.session!.dispose());
});

test("narrow workspace keeps page chrome within the viewport and comment actions above the canvas", async ({ page }, info) => {
  await page.setViewportSize({ width: 390, height: 844 }); await report(page);
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
  const toolbar = (await page.getByRole("toolbar", { name: "文件与导出", exact: true }).boundingBox())!;
  const canvas = (await page.locator('iframe[title="文档只读预览"]').boundingBox())!;
  expect(toolbar.y + toolbar.height).toBeLessThanOrEqual(canvas.y);
  await selectText(page); const dialog = page.getByRole("dialog", { name: "新建批注", exact: true }); await expect(dialog).toBeVisible();
  const box = (await dialog.boundingBox())!; expect(box.x + box.width).toBeLessThanOrEqual(390); expect(box.y + box.height).toBeLessThanOrEqual(844);
  await page.screenshot({ path: info.outputPath("mobile-workspace.png"), fullPage: true });
});

test("Office text toggles preserve the selected text, current state and shared history", async ({ page }) => {
  await page.goto('/'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="n-summary"]');
  await paragraph.evaluate(node => {
    (document.querySelector('[data-office-editor]') as HTMLElement).focus();
    const range = document.createRange(); range.selectNodeContents(node); const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
    document.dispatchEvent(new Event('selectionchange'));
  });
  const text = await paragraph.textContent();
  for (const name of ['加粗', '斜体', '下划线']) {
    const tool = page.getByRole('button', { name, exact: true });
    await tool.click(); await expect(tool).toHaveAttribute('aria-pressed', 'true');
    expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(text);
    await tool.click(); await expect(tool).toHaveAttribute('aria-pressed', 'false');
    expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(text);
  }
  await page.getByRole('group', { name: '字体', exact: true }).getByLabel('字号', { exact: true }).fill("18"); await page.getByRole('group', { name: '字体', exact: true }).getByLabel('字号', { exact: true }).press("Enter");
  await expect(page.getByRole('group', { name: '字体', exact: true }).getByLabel('字号', { exact: true })).toHaveValue('18');
  await page.getByRole('button', { name: '居中', exact: true }).click(); await expect(paragraph).toHaveCSS('text-align', 'center');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect(paragraph).not.toHaveCSS('text-align', 'center');
  await expect(page.getByRole('group', { name: '字体', exact: true }).getByLabel('字号', { exact: true })).toHaveValue('18');
  await page.getByRole('tab', { name: '审阅', exact: true }).click();
  await expect(page.getByRole('dialog', { name: '新建批注', exact: true })).toBeHidden();
  await page.getByRole('button', { name: '新建批注', exact: true }).click();
  await expect(page.getByRole('dialog', { name: '新建批注', exact: true })).toBeVisible(); await expect(paragraph).toBeVisible();
});

async function zoomGeometry(page: Page) {
  return page.locator('.stage').filter({ visible: true }).evaluate(stage => {
    const plane = stage.firstElementChild as HTMLElement, iframe = plane.querySelector('iframe')!;
    const transform = iframe.style.transform || (iframe.parentElement as HTMLElement).style.transform;
    const zoom = Number(transform.match(/[\d.]+/)?.[0] ?? 1), rect = stage.getBoundingClientRect(), content = plane.getBoundingClientRect();
    return { x: (rect.left + stage.clientLeft + stage.clientWidth / 2 - content.left - plane.offsetWidth / 2) / zoom,
      y: (rect.top + stage.clientTop + stage.clientHeight / 2 - content.top) / zoom, zoom };
  });
}
test("Office zoom keeps the document center and a panned viewport in both editing and preview", async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto('/');
  await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  await expect(page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="n-summary"]')).toBeVisible();
  for (const preview of [false, true]) {
    if (preview) await page.getByRole('button', { name: '返回预览与批注', exact: true }).click();
    const initial = await zoomGeometry(page); expect(Math.abs(initial.x)).toBeLessThan(1);
    await page.getByRole('button', { name: '放大', exact: true }).click();
    await expect.poll(async () => Math.abs((await zoomGeometry(page)).x)).toBeLessThan(1);
    await page.locator('.stage').filter({ visible: true }).evaluate(stage => { stage.scrollLeft += 70; stage.scrollTop = 180; });
    const panned = await zoomGeometry(page);
    await page.getByRole('button', { name: '放大', exact: true }).click();
    const after = await zoomGeometry(page); expect(Math.abs(after.x - panned.x)).toBeLessThan(1); expect(Math.abs(after.y - panned.y)).toBeLessThan(1);
    await page.getByRole('button', { name: '缩小', exact: true }).click();
    const back = await zoomGeometry(page); expect(Math.abs(back.x - panned.x)).toBeLessThan(1); expect(Math.abs(back.y - panned.y)).toBeLessThan(1);
    // Center explicitly before switching modes so both views verify the same point.
    await page.locator('.stage').filter({ visible: true }).evaluate(stage => { stage.scrollLeft = (stage.scrollWidth - stage.clientWidth) / 2; });
  }
  await expect(page.locator('.office-app')).toHaveAttribute("data-revision", "0");
});

test("PPT ribbon executes slide and text commands and zoom centers the slide", async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto('/');
  await page.getByRole('button', { name: '打开 PPT 协作样例', exact: true }).click();
  const text = page.frameLocator('iframe[title="幻灯片只读预览"]').locator('[data-node-id="heading-p"]'); await expect(text).toBeVisible();
  await expect.poll(async () => (await zoomGeometry(page)).zoom).toBe(1);
  await text.click(); const bold = page.getByRole('button', { name: '加粗', exact: true }); await expect(bold).toBeEnabled();
  const original = await bold.getAttribute('aria-pressed'); await bold.click(); await expect(bold).toHaveAttribute('aria-pressed', original === 'true' ? 'false' : 'true');
  await bold.click(); await expect(bold).toHaveAttribute('aria-pressed', original!);
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.getByRole('button', { name: '重做', exact: true }).click();
  await page.getByRole('button', { name: '新建幻灯片', exact: true }).click();
  await expect(page.locator('.office-app')).toHaveAttribute("data-revision", "5");
  const before = await zoomGeometry(page); await page.getByRole('button', { name: '放大', exact: true }).click();
  const after = await zoomGeometry(page); expect(Math.abs(after.x - before.x)).toBeLessThan(1);
  await page.getByRole('button', { name: '版式', exact: true }).click(); await expect(page.getByLabel('当前页布局', { exact: true })).toBeVisible();
});

test('partial underline removal preserves surrounding text through history and native DOCX reopen', async ({ page }, info) => {
  await page.goto('/'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="n-summary"]');
  await paragraph.evaluate(node => { (document.querySelector('[data-office-editor]') as HTMLElement).focus(); const range = document.createRange(); range.selectNodeContents(node); const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  await page.getByRole('button', { name: '下划线', exact: true }).click();
  await paragraph.evaluate(node => { const text = document.createTreeWalker(node, 4).nextNode()!; const range = document.createRange(); range.setStart(text, 2); range.setEnd(text, 6); const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  await page.getByRole('button', { name: '下划线', exact: true }).click();
  const underlines = async (node: import('@playwright/test').Locator) => node.evaluate(root => {
    const walker = document.createTreeWalker(root, 4), result: boolean[] = []; let text: Node | null;
    while ((text = walker.nextNode())) { let underlined = false; for (let element = text.parentElement; element && root.contains(element); element = element.parentElement) underlined ||= getComputedStyle(element).textDecorationLine.includes('underline'); result.push(...Array(text.textContent!.length).fill(underlined)); }
    return result;
  });
  const expected = [true, true, false, false, false, false, ...Array((await paragraph.textContent())!.length - 6).fill(true)];
  expect(await underlines(paragraph)).toEqual(expected);
  await page.getByRole('button', { name: '撤销', exact: true }).click(); expect((await underlines(paragraph)).every(Boolean)).toBe(true);
  await page.getByRole('button', { name: '重做', exact: true }).click(); expect(await underlines(paragraph)).toEqual(expected);
  await page.getByRole('tab', { name: '文件', exact: true }).click(); const pending = page.waitForEvent('download'); await page.getByRole('button', { name: '导出 DOCX', exact: true }).click();
  const path = info.outputPath('partial-underline.docx'); await (await pending).saveAs(path);
  await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles(path);
  const reopened = page.frameLocator('iframe[title="正文编辑"]').locator('p').filter({ hasText: /^项目计划于第三季度完成。$/ }); await expect(reopened).toBeVisible();
  expect(await underlines(reopened)).toEqual(expected);
});

test('Word rulers track the paper and the left selection bar selects wrapped visual lines', async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto('/');
  await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="n-summary"]');
  await paragraph.evaluate(node => { (document.querySelector('[data-office-editor]') as HTMLElement).focus(); const range = document.createRange(); range.selectNodeContents(node); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); });
  const text = '这是一个用于核对自动换行和左侧整行选择的段落，包含中文、Office、😀 和 12345。'.repeat(10);
  await page.keyboard.insertText(text); await expect(paragraph).toHaveText(text);
  let revision = await page.locator('.office-app').getAttribute('data-revision');
  for (const preview of [false, true]) {
    if (preview) await page.getByRole('button', { name: '返回预览与批注', exact: true }).click();
    const title = preview ? '文档只读预览' : '正文编辑', frame = page.frameLocator(`iframe[title="${title}"]`), paragraph = frame.locator('[data-node-id="n-summary"]');
    const horizontal = page.getByRole('img', { name: '水平标尺（厘米）', exact: true }), vertical = page.getByRole('img', { name: '垂直标尺（厘米）', exact: true });
    await expect(horizontal).toBeVisible(); await expect(vertical).toBeVisible();
    for (const zoom of [1, 1.25]) {
      if (zoom > 1) await page.getByRole('button', { name: '放大', exact: true }).click();
      const frameBox = (await page.locator(`iframe[title="${title}"]`).boundingBox())!;
      const localPaper = await frame.locator('.sdk-paper').first().evaluate(node => { const rect = node.getBoundingClientRect(); return { x: rect.x, y: rect.y, width: rect.width }; });
      const paper = { x: frameBox.x + localPaper.x * zoom, width: localPaper.width * zoom }, ruler = await horizontal.evaluate(node => { const rect = node.getBoundingClientRect(); return { x: rect.x, width: rect.width }; });
      expect(Math.abs(paper.x - ruler.x)).toBeLessThan(1); expect(Math.abs(paper.width - ruler.width)).toBeLessThan(1);
      const line = await paragraph.evaluate(root => {
        const all = document.createRange(); all.selectNodeContents(root); const rects = [...all.getClientRects()].filter(rect => rect.width && rect.height);
        const top = rects.map(rect => rect.top).filter((top, index, values) => values.findIndex(value => Math.abs(value - top) < 2) === index).sort((a, b) => a - b)[1]!;
        const band = rects.find(rect => Math.abs(rect.top - top) < 2)!, y = (band.top + band.bottom) / 2;
        let expected = '', node: Node | null; const walker = document.createTreeWalker(root, 4);
        while ((node = walker.nextNode())) { if (node.parentElement?.closest('.sdk-page-gap,.sdk-number-marker')) continue; let offset = 0; for (const character of node.textContent!) { const range = document.createRange(); range.setStart(node, offset); range.setEnd(node, offset + character.length); const rect = range.getBoundingClientRect(); if (rect.top < y && rect.bottom > y) expected += character; offset += character.length; } }
        const laterTop = rects.map(rect => rect.top).filter((top, index, values) => values.findIndex(value => Math.abs(value - top) < 2) === index).sort((a, b) => a - b)[3]!;
        const later = rects.find(rect => Math.abs(rect.top - laterTop) < 2)!;
        return { y, expected, laterY: (later.top + later.bottom) / 2 };
      });
      await page.mouse.click(paper.x + 12 * zoom, frameBox.y + line.y * zoom);
      expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(line.expected); expect(line.expected.length).toBeGreaterThan(5); expect(line.expected.length).toBeLessThan(text.length);
      await expect(page.locator('.office-app')).toHaveAttribute('data-revision', revision!);
      if (zoom === 1) {
        await page.mouse.move(paper.x + 12, frameBox.y + line.y); await page.mouse.down(); await page.mouse.move(paper.x + 12, frameBox.y + line.laterY, { steps: 4 }); await page.mouse.up();
        const selected = await paragraph.evaluate(() => document.getSelection()!.toString()); expect(selected).toContain(line.expected); expect(selected.length).toBeGreaterThan(line.expected.length); expect(selected.length).toBeLessThan(text.length);
      }
      if (!preview && zoom > 1) {
        await page.getByRole('button', { name: '加粗', exact: true }).click();
        expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(line.expected);
        await page.getByRole('button', { name: '撤销', exact: true }).click(); revision = await page.locator('.office-app').getAttribute('data-revision');
        await page.mouse.dblclick(paper.x + 12 * zoom, frameBox.y + line.y * zoom); expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(text);
        await page.mouse.click(paper.x + 12 * zoom, frameBox.y + line.y * zoom, { clickCount: 3 }); expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toContain('项目实施报告');
      }
      if (zoom > 1) await page.getByRole('button', { name: '缩小', exact: true }).click();
    }
  }
  await page.getByRole('tab', { name: '视图', exact: true }).click(); await page.getByRole('checkbox', { name: '标尺', exact: true }).uncheck();
  await expect(page.getByRole('img', { name: '水平标尺（厘米）', exact: true })).toBeHidden();
  await page.getByRole('checkbox', { name: '标尺', exact: true }).check(); await expect(page.getByRole('img', { name: '水平标尺（厘米）', exact: true })).toBeVisible();
  await page.getByRole('tabpanel', { name: '视图', exact: true }).getByRole('button', { name: '连续视图', exact: true }).click();
  await expect(page.getByRole('img', { name: '水平标尺（厘米）', exact: true })).toBeVisible(); await expect(page.getByRole('img', { name: '垂直标尺（厘米）', exact: true })).toBeHidden();
});

test('Office audit: native DOCX fonts change visible glyphs and survive export/reopen', async ({ page }, info) => {
  await page.goto('/'); await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles('examples/docx/09-chinese-fonts.docx');
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('p').filter({ hasText: 'ABCxyz' }).first(); await expect(paragraph).toBeVisible();
  const font = page.getByRole('group', { name: '字体', exact: true }).getByRole('combobox', { name: '字体', exact: true });
  const size = page.getByRole('group', { name: '字体', exact: true }).getByRole('combobox', { name: '字号', exact: true });
  await paragraph.evaluate(root => {
    const walker = document.createTreeWalker(root, 4); let node: Node | null; while ((node = walker.nextNode())) if (node.textContent!.includes('ABCxyz')) { const range = document.createRange(), start = node.textContent!.indexOf('ABCxyz'); range.setStart(node, start); range.setEnd(node, start + 6); root.closest<HTMLElement>('[contenteditable]')!.focus(); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); break; }
  });
  await expect(font).toHaveValue('Liberation Serif');
  await paragraph.evaluate(root => { root.closest<HTMLElement>('[contenteditable]')!.focus(); const range = document.createRange(); range.selectNodeContents(root); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  const selected = await paragraph.textContent(), before = await paragraph.screenshot();
  await page.getByRole('button', { name: '展开字体', exact: true }).click(); await page.getByRole('option', { name: 'FandolKai', exact: true }).click();
  await expect(font).toHaveValue('FandolKai'); expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(selected);
  const applied = await paragraph.evaluate(root => { const walker = document.createTreeWalker(root, 4), values: string[] = []; let node: Node | null; while ((node = walker.nextNode())) if (node.textContent?.trim()) values.push(getComputedStyle(node.parentElement!).fontFamily); return values; });
  expect(applied.every(font => font.includes('FandolKai'))).toBe(true); expect((await paragraph.screenshot()).equals(before)).toBe(false);
  const spanCount = await paragraph.locator('span').count();
  for (const family of ['Noto Serif CJK SC', 'Calibri', 'Aptos', 'FandolKai', 'FandolSong', 'FandolKai']) { await font.fill(family); await font.press('Enter'); await expect(font).toHaveValue(family); }
  expect(await paragraph.locator('span').count()).toBeLessThanOrEqual(spanCount);
  await size.fill('10.5'); await size.press('Enter'); await expect(size).toHaveValue('10.5');
  await page.getByRole('tab', { name: '文件', exact: true }).click(); const pending = page.waitForEvent('download'); await page.getByRole('button', { name: '导出 DOCX', exact: true }).click();
  const path = info.outputPath('font-audit.docx'); await (await pending).saveAs(path); await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles(path);
  await expect(page.getByRole('heading', { name: 'font-audit.docx', exact: true })).toBeVisible();
  const reopened = page.frameLocator('iframe[title="正文编辑"]').locator('p').filter({ hasText: 'ABCxyz' }).first();
  await reopened.evaluate(root => { root.closest<HTMLElement>('[contenteditable]')!.focus(); const range = document.createRange(); range.selectNodeContents(root); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  await expect(font).toHaveValue('FandolKai'); await expect(size).toHaveValue('10.5');
  await font.fill('Calibri'); await font.press('Enter'); await expect(page.getByRole('status', { name: /本机没有.*Calibri/ })).toBeVisible();
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect(font).toHaveValue('FandolKai');
});

test('Office audit: pending typing format follows its caret and Word keyboard commands use shared history', async ({ page }) => {
  await page.goto('/'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  const frame = page.frameLocator('iframe[title="正文编辑"]'), paragraph = frame.locator('[data-node-id="n-summary"]'), heading = frame.locator('[data-node-id="n-title"]');
  const font = page.getByRole('group', { name: '字体', exact: true }).getByRole('combobox', { name: '字体', exact: true });
  await paragraph.evaluate(root => { root.closest<HTMLElement>('[contenteditable]')!.focus(); const range = document.createRange(); range.selectNodeContents(root); range.collapse(false); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  await font.fill('FandolKai'); await font.press('Enter'); await expect(font).toHaveValue('FandolKai'); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '0');
  await heading.evaluate(root => { const range = document.createRange(); range.selectNodeContents(root); range.collapse(false); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  await expect(font).not.toHaveValue('FandolKai'); await page.keyboard.insertText('测试'); await expect(heading).toContainText('测试');
  expect(await heading.evaluate(root => [...root.querySelectorAll('span')].some(node => node.style.fontFamily.includes('FandolKai')))).toBe(false);
  await paragraph.evaluate(root => { root.closest<HTMLElement>('[contenteditable]')!.focus(); const range = document.createRange(); range.selectNodeContents(root); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  await page.keyboard.press('Control+b'); await expect(page.getByRole('button', { name: '加粗', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('Control+z'); await expect(page.getByRole('button', { name: '加粗', exact: true })).toHaveAttribute('aria-pressed', 'false');
  await page.keyboard.press('Control+y'); await expect(page.getByRole('button', { name: '加粗', exact: true })).toHaveAttribute('aria-pressed', 'true');
  for (const [key, alignment] of [['e', 'center'], ['l', 'left'], ['r', 'right']]) { await page.keyboard.press('Control+' + key); await expect(paragraph).toHaveCSS('text-align', alignment!); }
});

test('Office audit: ruler rails stay outside the paper and selection cursor reverses only in the margin', async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto('/'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  const frame = page.frameLocator('iframe[title="正文编辑"]'), paragraph = frame.locator('[data-node-id="n-summary"]'); await expect(paragraph).toBeVisible();
  const geometry = () => page.locator('iframe[title="正文编辑"]').evaluate(iframe => { const root = iframe.getRootNode() as ShadowRoot, stage = root.querySelector<HTMLElement>('.stage')!, h = root.querySelector<SVGSVGElement>('[data-doffi-ruler=horizontal]')!, v = root.querySelector<SVGSVGElement>('[data-doffi-ruler=vertical]')!, vr = v.parentElement!.getBoundingClientRect(), hr = h.parentElement!.getBoundingClientRect(), s = stage.getBoundingClientRect(), horizontal = h.getBoundingClientRect(); return { inside: stage.contains(h) || stage.contains(v), verticalX: vr.x, stageX: s.x, horizontalY: hr.y, stageY: s.y, horizontalX: horizontal.x }; });
  const initial = await geometry(); expect(initial.inside).toBe(false); expect(initial.verticalX).toBeLessThan(initial.stageX); expect(initial.horizontalY).toBeLessThan(initial.stageY);
  await page.getByRole('button', { name: '放大', exact: true }).click(); const before = await geometry();
  await page.locator('.stage').filter({ visible: true }).evaluate(stage => { stage.scrollTop = 120; stage.scrollLeft += 50; });
  await expect.poll(async () => (await geometry()).horizontalX).toBeCloseTo(before.horizontalX - 50, 0);
  const after = await geometry(); expect(after.verticalX).toBe(before.verticalX); expect(after.horizontalY).toBe(before.horizontalY);
  const coords = await paragraph.evaluate(root => { const rect = root.getBoundingClientRect(), paper = document.querySelector('.sdk-paper')!.getBoundingClientRect(); return { y: rect.y + rect.height / 2, x: paper.left + 12, textX: rect.left + 20 }; });
  const box = (await page.locator('iframe[title="正文编辑"]').boundingBox())!;
  await page.mouse.move(box.x + coords.x * 1.25, box.y + coords.y * 1.25);
  expect(await paragraph.evaluate((_, at) => getComputedStyle(document.elementFromPoint(at.x, at.y)!).cursor, { x: coords.x, y: coords.y })).toContain('M18');
  await page.mouse.move(box.x + coords.textX * 1.25, box.y + coords.y * 1.25); expect(await paragraph.evaluate(() => document.documentElement.hasAttribute('data-doffi-line-cursor'))).toBe(false);
  await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '0');
});

test('Office audit: editable font fields validate size and their list fits a narrow viewport', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 }); await page.goto('/'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="n-summary"]');
  await paragraph.evaluate(root => { root.closest<HTMLElement>('[contenteditable]')!.focus(); const range = document.createRange(); range.selectNodeContents(root); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  await page.getByRole('button', { name: '展开字体', exact: true }).click(); const list = page.getByRole('listbox', { name: '字体选项', exact: true }); await expect(list).toBeVisible();
  const box = (await list.boundingBox())!; expect(box.x).toBeGreaterThanOrEqual(0); expect(box.x + box.width).toBeLessThanOrEqual(390); expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
  await page.getByRole('option', { name: 'FandolSong', exact: true }).click();
  const size = page.getByRole('group', { name: '字体', exact: true }).getByRole('combobox', { name: '字号', exact: true });
  const revision = await page.locator('.office-app').getAttribute('data-revision'); await size.fill('10.3'); await size.press('Enter'); await expect(size).toHaveAttribute('aria-invalid', 'true'); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', revision!);
  await size.press('Escape'); await expect(size).toHaveAttribute('aria-invalid', 'false'); await size.fill('10.5'); await size.press('Enter'); await expect(size).toHaveValue('10.5');
});

test('Office audit: font list keyboard navigation and Tab preserve the document selection', async ({ page }) => {
  await page.goto('/'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="n-summary"]');
  await paragraph.evaluate(root => { root.closest<HTMLElement>('[contenteditable]')!.focus(); const range = document.createRange(); range.selectNodeContents(root); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  const font = page.getByRole('combobox', { name: '字体', exact: true }), size = page.getByRole('combobox', { name: '字号', exact: true });
  await font.fill('FandolK'); await font.press('ArrowDown'); await font.press('Enter'); await expect(font).toHaveValue('FandolKai');
  await font.fill('FandolSong'); await font.press('Tab'); await expect(size).toBeFocused(); await expect(font).toHaveValue('FandolSong'); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '2');
  await size.fill('15.5'); await size.press('Shift+Tab'); await expect(font).toBeFocused(); await expect(size).toHaveValue('15.5'); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '3');
  expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe('项目计划于第三季度完成。');
});

test('Office audit: the presentation font field formats the selected text box through history', async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto('/'); await page.getByRole('button', { name: '打开 PPT 协作样例', exact: true }).click();
  const text = page.frameLocator('iframe[title="幻灯片只读预览"]').locator('[data-node-id="heading-p"]'); await text.click();
  const font = page.getByRole('group', { name: '字体', exact: true }).getByRole('combobox', { name: '字体', exact: true }); await expect(font).toBeEnabled();
  const before = await text.screenshot(); await font.fill('FandolKai'); await font.press('Enter'); await expect(font).toHaveValue('FandolKai'); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '1');
  expect((await text.screenshot()).equals(before)).toBe(false); expect(await text.textContent()).toBe('项目进展');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '2'); await expect(font).not.toHaveValue('FandolKai');
  await page.getByRole('button', { name: '重做', exact: true }).click(); await expect(font).toHaveValue('FandolKai');
});

for (const file of ['04-styles-sections.docx', '05-long-page-break.docx']) {
  test('Word preview keeps separate paper geometry after editing and repeated mode switches: ' + file, async ({ page }) => {
    const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
    await page.goto('/'); await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles('examples/docx/' + file);
    async function papers(title: string) {
      const frame = page.frameLocator(`iframe[title="${title}"]`); await expect(frame.locator('.sdk-paper')).toHaveCount(file.startsWith('04') ? 2 : 3);
      return frame.locator('.sdk-paper').first().evaluate(paper => {
        const root = paper.ownerDocument.querySelector('[data-office-kind="document"]')!.getBoundingClientRect();
        return [...paper.ownerDocument.querySelectorAll('.sdk-paper')].map(node => { const rect = node.getBoundingClientRect(); return { x: rect.x - root.x, y: rect.y - root.y, width: rect.width, height: rect.height }; });
      });
    }
    const expected = await papers('正文编辑');
    expect(expected[1]!.y - expected[0]!.y - expected[0]!.height).toBeGreaterThan(20);
    for (let cycle = 0; cycle < 3; cycle++) {
      if (cycle) {
        await page.getByRole('button', { name: '编辑正文', exact: true }).click(); await expect.poll(() => papers('正文编辑')).toEqual(expected);
        const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-office-editor] p').first();
        await paragraph.evaluate(node => { const root = node.closest<HTMLElement>('[contenteditable]')!; root.focus(); const range = document.createRange(); range.selectNodeContents(node); range.collapse(false); const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range); });
        await page.keyboard.insertText('测'); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', String(cycle));
      }
      await page.getByRole('button', { name: cycle % 2 ? '缩小' : '放大', exact: true }).click();
      await page.getByRole('button', { name: '返回预览与批注', exact: true }).click();
      await expect.poll(() => papers('文档只读预览')).toEqual(expected);
      await expect(page.getByRole('alert')).toHaveCount(0);
    }
    expect(errors).toEqual([]);
  });
}

test('the SDK Office ribbon owns its styles, preview behavior and independent instance navigation', async ({ page }) => {
  const bundle = await build({ configFile: false, logLevel: 'silent', define: { 'process.env.NODE_ENV': '"production"' }, resolve: { alias: { react: fileURLToPath(new URL('../../apps/playground/node_modules/react', import.meta.url)), 'react-dom': fileURLToPath(new URL('../../apps/playground/node_modules/react-dom', import.meta.url)) } }, build: { write: false, target: 'es2022', lib: { entry: fileURLToPath(new URL('./fixtures/office-ribbon-host.ts', import.meta.url)), formats: ['es'] } } });
  const output = Array.isArray(bundle) ? bundle[0]!.output : 'output' in bundle ? bundle.output : [];
  await openSdkHost(page); await page.addScriptTag({ type: 'module', content: output.flatMap(item => item.type === 'chunk' ? [item.code] : []).join('\n') });
  const first = page.getByRole('region', { name: '文档一', exact: true }), second = page.getByRole('region', { name: '文档二', exact: true });
  const firstFont = first.getByRole('combobox', { name: '字体', exact: true }), secondFont = second.getByRole('combobox', { name: '字体', exact: true });
  await expect(firstFont).toBeVisible(); await expect(first.getByRole('button', { name: '加粗', exact: true })).toHaveCSS('font-size', '11px');
  await expect(page.locator('#outside')).toHaveCSS('background-color', 'rgb(255, 0, 0)'); await expect(page.locator('#outside')).toHaveCSS('font-size', '21px');
  expect(await firstFont.getAttribute('aria-controls')).not.toBe(await secondFont.getAttribute('aria-controls'));
  await firstFont.fill('FandolKai'); await firstFont.press('Enter'); await expect(firstFont).toHaveValue('FandolKai'); await expect(secondFont).toHaveValue('Arial');
  await first.getByRole('button', { name: '切换预览', exact: true }).click(); await expect(first.getByRole('group', { name: '字体', exact: true })).toBeHidden(); await expect(first.getByRole('tab', { name: '插入', exact: true })).toBeDisabled();
  await first.getByRole('tab', { name: '文件', exact: true }).click(); await first.getByRole('button', { name: '打开', exact: true }).click(); await expect(first.getByLabel('打开次数')).toHaveText('1'); await expect(second.getByLabel('打开次数')).toHaveText('0');
  await second.getByRole('tab', { name: '开始', exact: true }).press('End'); await expect(second.getByRole('tab', { name: '帮助', exact: true })).toBeFocused(); await expect(first.getByRole('tab', { name: '文件', exact: true })).toHaveAttribute('aria-selected', 'true');
  await first.getByRole('button', { name: '切换编辑', exact: true }).click(); await expect(firstFont).toBeVisible(); await expect(firstFont).toHaveValue('FandolKai');
});
