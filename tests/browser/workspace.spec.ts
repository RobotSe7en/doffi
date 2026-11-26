import { existsSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { openSdkHost } from './sdk-host.js';
import { openDocument } from "@doffi/core";
import { importPptx, readDocumentZip } from "@doffi/office";
import { body as nativeBody, mutateDocx } from "../../packages/office/tests/regression-fixtures.js";
import type {} from "./fixtures/interaction-host.js";
import { wordProducerDrawingFixture } from "../../packages/office/tests/word-producer-drawing-fixture.js";

test("Word shape selection exposes its contextual format tab and edits text on the page", async ({ page }) => {
  await page.addInitScript(() => localStorage.setItem("doffi:editing", "preview"));
  await page.goto("/");
  await page.getByLabel("打开文件", { exact: true }).setInputFiles({
    name: "contextual-shape.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from(wordProducerDrawingFixture({ scale: 90 })),
  });
  await page.getByRole("button", { name: "编辑正文", exact: true }).click();
  const frame = page.frameLocator('iframe[title="正文编辑"]'), object = frame.locator('.sdk-word-floating-layer [data-word-drawing="textBox"]');
  await expect(page.getByRole("tab", { name: "形状格式", exact: true })).toHaveCount(0);
  await object.click({ position: { x: 15, y: 15 } });
  await expect(page.getByRole("region", { name: "绘图设置", exact: true })).toHaveCount(0);
  await page.getByRole("tab", { name: "形状格式", exact: true }).click();
  await expect(page.getByRole("tabpanel", { name: "形状格式", exact: true })).toBeVisible();
  const paper = await frame.locator(".sdk-paper").first().boundingBox();
  await page.getByRole("button", { name: "大小和位置", exact: true }).click();
  await expect(page.getByRole("dialog", { name: "形状格式", exact: true })).toBeVisible();
  expect((await frame.locator(".sdk-paper").first().boundingBox())!.y).toBeCloseTo(paper!.y, 1);
  await page.getByRole("button", { name: "关闭绘图设置", exact: true }).click();
  await page.getByRole("button", { name: "编辑文字", exact: true }).click();
  const editor = object.locator("[data-office-editor]");
  await expect(editor).toBeVisible();
  await expect(page.getByText("正在编辑辅助内容", { exact: false })).toHaveCount(0);
  await expect(page.getByRole("tab", { name: "形状格式", exact: true })).toBeVisible();
  await editor.press("Escape");
  await frame.locator('td p').filter({ hasText: "中文排版" }).first().click({ position: { x: 4, y: 4 } });
  await expect(page.getByRole("tab", { name: "形状格式", exact: true })).toHaveCount(0);
  await expect(page.getByRole("tab", { name: "开始", exact: true })).toHaveAttribute("aria-selected", "true");
});

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

test("Word preview exposes find, heading navigation and statistics without entering editing", async ({ page }) => {
  await report(page);
  const before = await page.locator(".office-app").getAttribute("data-revision");
  await page.getByRole("button", { name: "查找", exact: true }).click();
  const find = page.getByRole("region", { name: "全文查找", exact: true });
  await find.getByRole("searchbox", { name: "查找内容" }).fill("第三季度");
  await find.getByRole("button", { name: "下一处", exact: true }).click();
  await expect(find).toContainText("处匹配 · 第 1 处");
  await expect(page.locator('iframe[title="正文编辑"]')).toHaveCount(0);
  await find.getByRole("button", { name: "关闭导航", exact: true }).click();
  await page.getByRole("tab", { name: "视图", exact: true }).click();
  await page.getByRole("button", { name: "导航窗格", exact: true }).click();
  await expect(page.getByRole("region", { name: "文档导航", exact: true })).toBeVisible();
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "字数统计", exact: true }).click();
  await expect(page.getByRole("region", { name: "字数统计", exact: true })).toContainText("字符数");
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", before!);
  await page.getByRole("button", { name: "编辑正文", exact: true }).click();
  await expect(page.locator('iframe[title="正文编辑"]')).toBeVisible();
  await expect(page.getByRole("region", { name: "字数统计", exact: true })).toBeHidden();
});

test("Word preview and editing load Playground web fonts and keep the native font names", async ({ page }, info) => {
  const manifest = new URL("../../fonts/local/sources.json", import.meta.url);
  const internal = existsSync(manifest);
  const families = internal ? [
    ["宋体", "SimSun"],
    ["黑体", "SimHei"],
    ["仿宋", "FangSong"],
    ["楷体", "KaiTi"],
  ] as const : [
    ["宋体", "Doffi Serif SC"],
    ["黑体", "Doffi Sans SC"],
    ["仿宋_GB2312", "Doffi Fangsong SC"],
    ["楷体", "Doffi Kai SC"],
  ] as const;
  const maxRequests = 4 + (internal ? JSON.parse(readFileSync(manifest, "utf8")).fonts.length : 0);
  const fontRequests: string[] = [];
  page.on("request", (request) => {
    if (request.resourceType() === "fetch" && /\.woff2?$/.test(new URL(request.url()).pathname))
      fontRequests.push(request.url());
  });
  await page.addInitScript(() => localStorage.setItem("doffi:editing", "preview"));
  await page.goto("/");
  expect(fontRequests).toHaveLength(0);
  const name = "word-web-fonts.docx";
  const bytes = mutateDocx("01-text-numbering.docx", {
    "word/document.xml": (source) => nativeBody(source, families.map(([family], index) =>
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="' + family + '"/><w:i/></w:rPr>' +
      '<w:t>汉字测试' + index + 'ABC</w:t></w:r></w:p>',
    ).join("")),
  });
  await page.getByLabel("打开文件", { exact: true }).setInputFiles({
    name, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from(bytes),
  });
  async function checkFonts(title: string): Promise<void> {
    const frame = page.frameLocator(`iframe[title="${title}"]`);
    await expect(frame.locator("[data-page-id]").first()).toBeAttached();
    for (const [index, [requested, family]] of families.entries()) {
      const paragraph = frame.locator("p").filter({ hasText: new RegExp("汉字测试\\s*" + index + "ABC") }).first();
      const cjk = paragraph.getByText("汉字测试", { exact: true });
      const selection = await cjk.evaluate((node) => ({
        family: getComputedStyle(node).fontFamily.split(",")[0]!.replace(/^["']|["']$/g, ""),
        loaded: [...node.ownerDocument.fonts].filter((face) => face.status === "loaded")
          .map((face) => [face.family, face.style]),
      }));
      if (internal || selection.family !== requested) {
        expect(selection.family).toBe(family);
        expect(selection.loaded).toContainEqual([family, "normal"]);
      }
      if (internal) {
        await expect(paragraph.getByText(index + "ABC", { exact: true })).toHaveCSS("font-family", /Times New Roman/);
        expect(selection.loaded).toContainEqual(["Times New Roman", "italic"]);
      }
    }
  }
  await checkFonts("文档只读预览");
  const loadedCount = fontRequests.length;
  expect(loadedCount).toBeGreaterThan(0);
  expect(loadedCount).toBeLessThanOrEqual(maxRequests);
  expect(fontRequests.every((url) => new URL(url).origin === new URL(page.url()).origin)).toBe(true);
  await page.getByRole("button", { name: "编辑正文", exact: true }).click();
  await checkFonts("正文编辑");
  expect(fontRequests).toHaveLength(loadedCount);
  const target = page.frameLocator('iframe[title="正文编辑"]').locator("p").filter({ hasText: /汉字测试\s*0ABC/ }).first();
  await target.click();
  await target.evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); range.collapse(false);
    const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
    document.dispatchEvent(new Event("selectionchange"));
  });
  await page.keyboard.insertText("增");
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  const pending = page.waitForEvent("download");
  await page.getByRole("button", { name: "文件菜单", exact: true }).click(); await page.getByRole("menuitem", { name: "文件和导出…", exact: true }).click();
  await page.getByRole("button", { name: "下载文档包", exact: true }).click();
  const saved = info.outputPath("web-font-edit.zip");
  await (await pending).saveAs(saved);
  const snapshot = openDocument(readDocumentZip(new Uint8Array(readFileSync(saved)))).snapshot();
  if (!('html' in snapshot)) throw new Error("Expected a Word document.");
  for (const [family] of families) expect(snapshot.html).toContain(family);
  expect(snapshot.html).toContain("增");
  expect(snapshot.html).not.toMatch(/Doffi (Serif|Sans|Fangsong|Kai) SC/);
});

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
  await page.getByRole("button", { name: "文件菜单", exact: true }).click(); await page.getByRole("menuitem", { name: "显示文档列表", exact: true }).click();
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
  const toolbar = (await page.locator(".office-ribbon").boundingBox())!;
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
  await page.getByRole("button", { name: "选择", exact: true }).click(); await page.getByRole("menuitem", { name: "对象选择", exact: true }).click();
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
  await page.getByRole("button", { name: "选择", exact: true }).click(); await page.getByRole("menuitem", { name: "矩形框选", exact: true }).click();
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
  await expect(frame.locator('[data-node-id="heading-p"]')).toBeVisible();
  await expect(page.locator(".comments-pane")).toBeAttached();
  await expect(page.locator(".comments-pane")).toBeHidden();
  await expect(page.getByRole("heading", { name: "当前页对象", exact: true })).toHaveCount(0);
  await expect(page.locator(".object-list")).toHaveCount(0);
  await frame.locator('[data-node-id="heading-p"]').click();
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "显示批注", exact: true }).click();
  await expect(page.locator(".comments-pane")).toBeVisible();
  await page.getByRole("button", { name: "显示批注", exact: true }).click();
  await expect(page.locator(".comments-pane")).toBeHidden();
  await page.getByRole("button", { name: "新建批注", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true }); await expect(dialog).toBeVisible();
  await dialog.getByLabel("批注显示署名", { exact: true }).fill("审阅人");
  await dialog.getByLabel("新批注正文", { exact: true }).fill("确认演示标题");
  await dialog.getByRole("button", { name: "添加批注", exact: true }).click(); await expect(dialog).toBeHidden();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
  await page.getByRole("button", { name: "重做", exact: true }).click();
  await page.getByRole("button", { name: "文件菜单", exact: true }).click(); await page.getByRole("menuitem", { name: "文件和导出…", exact: true }).click();
  const pending = page.waitForEvent("download"); await page.getByRole("button", { name: "导出 PPTX", exact: true }).click();
  const path = info.outputPath("comment.pptx"); await (await pending).saveAs(path);
  const model = openDocument(importPptx(new Uint8Array(readFileSync(path)), { documentId: "workspace-comment" }));
  expect(model.snapshot().annotations).toHaveLength(2);
  expect(model.snapshot().annotations.map(annotation => annotation.bodyNodeId ? model.getNode(annotation.bodyNodeId)?.text : annotation.instruction)).toContain("确认演示标题");
  expect(model.snapshot().annotations.at(-1)?.author?.name).toBe("审阅人");
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "选择", exact: true }).click(); await page.getByRole("menuitem", { name: "矩形框选", exact: true }).click();
  const slide = (await frame.locator('[data-node-id="slide-root-1"]').boundingBox())!;
  await page.mouse.move(slide.x + 35, slide.y + 25); await page.mouse.down();
  await page.mouse.move(slide.x + 175, slide.y + 80, { steps: 4 }); await page.mouse.up();
  await page.getByRole("button", { name: "新建批注", exact: true }).click();
  await expect(dialog).toContainText("所选幻灯片区域");
  await dialog.getByRole("button", { name: "取消批注", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "3");
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
  const toolbar = (await page.locator(".office-ribbon").boundingBox())!;
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
  await expect(page.getByRole('group', { name: '字体', exact: true }).getByLabel('字号', { exact: true })).toHaveValue('小二');
  await page.getByRole('button', { name: '居中', exact: true }).click(); await expect(paragraph).toHaveCSS('text-align', 'center');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect(paragraph).not.toHaveCSS('text-align', 'center');
  await expect(page.getByRole('group', { name: '字体', exact: true }).getByLabel('字号', { exact: true })).toHaveValue('小二');
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
  await expect(page.locator('.slide-list button[aria-pressed=true]')).toHaveCSS('border-color', 'rgb(183, 71, 42)');
  await expect(page.locator('.stage')).toHaveCSS('height', '780px');
  await expect.poll(async () => { const g = await zoomGeometry(page); return g.zoom >= 0.5 && g.zoom <= 1 && Math.abs(g.x) < 1; }).toBe(true);
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
  await page.getByRole('button', { name: '文件菜单', exact: true }).click(); await page.getByRole('menuitem', { name: '文件和导出…', exact: true }).click(); const pending = page.waitForEvent('download'); await page.getByRole('button', { name: '导出 DOCX', exact: true }).click();
  const path = info.outputPath('partial-underline.docx'); await (await pending).saveAs(path);
  await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles(path);
  const reopened = page.frameLocator('iframe[title="正文编辑"]').locator('p').filter({ hasText: /^项目计划于第三季度完成。$/ }); await expect(reopened).toBeVisible();
  expect(await underlines(reopened)).toEqual(expected);
});

test('Word left margin selects wrapped visual lines in editing and preview', async ({ page }) => {
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
    await page.getByRole('tab', { name: '视图', exact: true }).click();
    await expect(page.getByRole('checkbox', { name: '标尺', exact: true })).toHaveCount(0);
    await expect(page.locator('[data-doffi-ruler-rail], [data-doffi-ruler]')).toHaveCount(0);
    if (!preview) await page.getByRole('tab', { name: '开始', exact: true }).click();
    for (const zoom of [1, 1.25]) {
      if (zoom > 1) await page.getByRole('button', { name: '放大', exact: true }).click();
      const frameBox = (await page.locator(`iframe[title="${title}"]`).boundingBox())!;
      const paperX = frameBox.x + await frame.locator('.sdk-paper').first().evaluate(node => node.getBoundingClientRect().left) * zoom;
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
      await page.mouse.click(paperX + 12 * zoom, frameBox.y + line.y * zoom);
      expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(line.expected); expect(line.expected.length).toBeGreaterThan(5); expect(line.expected.length).toBeLessThan(text.length);
      await expect(page.locator('.office-app')).toHaveAttribute('data-revision', revision!);
      if (zoom === 1) {
        await page.mouse.move(paperX + 12, frameBox.y + line.y); await page.mouse.down(); await page.mouse.move(paperX + 12, frameBox.y + line.laterY, { steps: 4 }); await page.mouse.up();
        const selected = await paragraph.evaluate(() => document.getSelection()!.toString()); expect(selected).toContain(line.expected); expect(selected.length).toBeGreaterThan(line.expected.length); expect(selected.length).toBeLessThan(text.length);
      }
      if (!preview && zoom > 1) {
        await page.getByRole('button', { name: '加粗', exact: true }).click();
        expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(line.expected);
        await page.getByRole('button', { name: '撤销', exact: true }).click(); revision = await page.locator('.office-app').getAttribute('data-revision');
        await page.mouse.dblclick(paperX + 12 * zoom, frameBox.y + line.y * zoom); expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(text);
        await page.mouse.click(paperX + 12 * zoom, frameBox.y + line.y * zoom, { clickCount: 3 }); expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toContain('项目实施报告');
      }
      if (zoom > 1) await page.getByRole('button', { name: '缩小', exact: true }).click();
    }
  }
  await page.getByRole('tabpanel', { name: '视图', exact: true }).getByRole('button', { name: 'Web 版式', exact: true }).click();
  const continuous = page.frameLocator('iframe[title="文档只读预览"]');
  await expect(continuous.locator('[data-node-id="n-summary"]')).toHaveText(text);
  await expect(continuous.locator('.sdk-paper')).toHaveCount(0);
});

test('Office audit: native DOCX fonts change visible glyphs and survive export/reopen', async ({ page }, info) => {
  await page.goto('/'); await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles('examples/docx/09-chinese-fonts.docx');
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-office-editor] p').filter({ hasText: 'ABCxyz' }).first(); await expect(paragraph).toBeVisible();
  const font = page.getByRole('group', { name: '字体', exact: true }).getByRole('combobox', { name: '字体', exact: true });
  const size = page.getByRole('group', { name: '字体', exact: true }).getByRole('combobox', { name: '字号', exact: true });
  await paragraph.evaluate(root => {
    const walker = document.createTreeWalker(root, 4); let node: Node | null; while ((node = walker.nextNode())) if (node.textContent!.includes('ABCxyz')) { const range = document.createRange(), start = node.textContent!.indexOf('ABCxyz'); range.setStart(node, start); range.setEnd(node, start + 6); root.closest<HTMLElement>('[contenteditable]')!.focus(); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); break; }
  });
  await expect(font).toHaveValue('Liberation Serif');
  await paragraph.evaluate(root => { root.closest<HTMLElement>('[contenteditable]')!.focus(); const range = document.createRange(); range.selectNodeContents(root); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  const selected = await paragraph.textContent(), before = await paragraph.screenshot();
  const chosenFont = 'Noto Sans CJK SC';
  await page.getByRole('button', { name: '展开字体', exact: true }).click(); await page.getByRole('option', { name: chosenFont, exact: true }).click();
  await expect(font).toHaveValue(chosenFont); expect(await paragraph.evaluate(() => document.getSelection()!.toString())).toBe(selected);
  const applied = await paragraph.evaluate(root => { const walker = document.createTreeWalker(root, 4), values: string[] = []; let node: Node | null; while ((node = walker.nextNode())) if (node.textContent?.trim()) values.push(getComputedStyle(node.parentElement!).fontFamily); return values; });
  expect(applied.every(font => font.includes(chosenFont))).toBe(true); expect((await paragraph.screenshot()).equals(before)).toBe(false);
  const spanCount = await paragraph.locator('span').count();
  for (const family of ['Noto Serif CJK SC', 'Calibri', 'Aptos', 'FandolKai', 'FandolSong', chosenFont]) { await font.fill(family); await font.press('Enter'); await expect(font).toHaveValue(family); }
  expect(await paragraph.locator('span').count()).toBeLessThanOrEqual(spanCount);
  await size.fill('10.5'); await size.press('Enter'); await expect(size).toHaveValue('五号');
  await page.getByRole('button', { name: '文件菜单', exact: true }).click(); await page.getByRole('menuitem', { name: '文件和导出…', exact: true }).click(); const pending = page.waitForEvent('download'); await page.getByRole('button', { name: '导出 DOCX', exact: true }).click();
  const path = info.outputPath('font-audit.docx'); await (await pending).saveAs(path); await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles(path);
  await expect(page.getByRole('heading', { name: 'font-audit.docx', exact: true })).toBeVisible();
  const reopened = page.frameLocator('iframe[title="正文编辑"]').locator('[data-office-editor] p').filter({ hasText: 'ABCxyz' }).first();
  await reopened.evaluate(root => { root.closest<HTMLElement>('[contenteditable]')!.focus(); const range = document.createRange(); range.selectNodeContents(root); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  await expect(font).toHaveValue(chosenFont); await expect(size).toHaveValue('五号');
  await font.fill('Calibri'); await font.press('Enter'); await expect(page.locator('.office-font-hint')).toHaveCount(0);
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect(font).toHaveValue(chosenFont);
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

test('Office audit: selection cursor reverses only in the paper margin', async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto('/'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  const frame = page.frameLocator('iframe[title="正文编辑"]'), paragraph = frame.locator('[data-node-id="n-summary"]'); await expect(paragraph).toBeVisible();
  await page.getByRole('button', { name: '放大', exact: true }).click();
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
  await size.press('Escape'); await expect(size).toHaveAttribute('aria-invalid', 'false'); await size.fill('10.5'); await size.press('Enter'); await expect(size).toHaveValue('五号');
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
    // Runtime host fonts can change automatic overflow; fixed-font pagination
    // fixtures assert exact counts. This case verifies stable geometry across modes.
    let paperCount: number | undefined;
    async function papers(title: string) {
      const frame = page.frameLocator(`iframe[title="${title}"]`);
      if (paperCount === undefined) { await expect.poll(() => frame.locator('.sdk-paper').count()).toBeGreaterThanOrEqual(file.startsWith('04') ? 2 : 3); paperCount = await frame.locator('.sdk-paper').count(); }
      await expect(frame.locator('.sdk-paper')).toHaveCount(paperCount);
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

test("Word page navigation scrolls to measured paper in preview and editing", async ({ page }) => {
  await page.addInitScript(() => localStorage.setItem("doffi:editing", "preview"));
  await page.goto("/");
  await page.getByLabel("打开文件", { exact: true }).filter({ visible: false }).setInputFiles("examples/docx/05-long-page-break.docx");
  async function jump(title: string) {
    const paper = page.frameLocator(`iframe[title="${title}"]`).locator(".sdk-paper").nth(2);
    const viewport = page.locator(".preview .stage").filter({ visible: true }).first();
    await expect(paper).toBeAttached();
    const target = page.getByRole("spinbutton", { name: "目标页码" });
    await target.fill("3");
    await target.press("Enter");
    await expect.poll(async () => Math.abs((await paper.boundingBox())!.y - (await viewport.boundingBox())!.y)).toBeLessThan(24);
  }
  await jump("文档只读预览");
  await page.getByRole("button", { name: "编辑正文", exact: true }).click();
  await jump("正文编辑");
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await expect(page.getByRole("button", { name: "撤销", exact: true })).toBeDisabled();
});

test("PowerPoint starts a read-only slide show from the current slide or the beginning", async ({ page }) => {
  await page.goto("/");
  await page.getByLabel("打开文件", { exact: true }).filter({ visible: false }).setInputFiles("examples/pptx/13-multi-slide.pptx");
  const second = page.getByRole("button", { name: "2. Slide 2", exact: true });
  await second.click();
  await page.getByLabel("文档缩放", { exact: true }).fill("200");
  await page.getByRole("button", { name: "幻灯片适应窗口" }).click();
  await expect(page.getByLabel("文档缩放", { exact: true })).not.toHaveValue("200");
  await expect(second).toHaveAttribute("aria-pressed", "true");
  const dialog = page.getByRole("dialog", { name: "幻灯片放映" });
  const visible = dialog.frameLocator('iframe[title="放映画布"]').locator('article>[data-office-kind="slide"]:visible');
  const launch = page.getByRole("button", { name: "从当前幻灯片放映", exact: true });
  await launch.click();
  await expect(visible).toHaveCount(1);
  await expect(visible).toHaveAttribute("data-node-id", "node-22");
  await expect(dialog.getByRole("button", { name: "下一张" })).toBeDisabled();
  await page.keyboard.press("ArrowLeft");
  await expect(visible).toHaveAttribute("data-node-id", "node-8");
  await page.keyboard.press("Escape");
  await expect(dialog).toBeHidden();
  await expect(launch).toBeFocused();
  await expect(second).toHaveAttribute("aria-pressed", "true");
  await page.getByRole("tab", { name: "幻灯片放映" }).click();
  await page.getByRole("button", { name: "从头播放" }).click();
  await expect(dialog.locator(".slide-show-surface")).toHaveAttribute("data-ready", "true");
  await expect(visible).toHaveAttribute("data-node-id", "node-8");
  await page.keyboard.press("PageDown");
  await expect(visible).toHaveAttribute("data-node-id", "node-22");
  await dialog.getByRole("button", { name: "退出放映" }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
});

test('the SDK Office ribbon owns its styles, preview behavior and independent instance navigation', async ({ page }) => {
  await page.setViewportSize({ width: 1920, height: 1080 });
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
  await first.getByRole('button', { name: '文件菜单', exact: true }).click(); await first.getByRole('menuitem', { name: '文件和导出…', exact: true }).click(); await first.getByRole('button', { name: '打开', exact: true }).click(); await expect(first.getByLabel('打开次数')).toHaveText('1'); await expect(second.getByLabel('打开次数')).toHaveText('0');
  await second.getByRole('tab', { name: '开始', exact: true }).press('End'); await expect(second.getByRole('tab', { name: '视图', exact: true })).toBeFocused(); await expect(first.getByRole('tabpanel', { name: '文件', exact: true })).toBeVisible();
  await first.getByRole('button', { name: '切换编辑', exact: true }).click(); await expect(firstFont).toBeVisible(); await expect(firstFont).toHaveValue('FandolKai');
  await expect(first.getByRole('tab', { name: '文件', exact: true })).toHaveCount(0);
  for (const name of ['文档一', '演示文稿', '工作簿']) {
    const region = page.getByRole('region', { name, exact: true });
    await expect(region.locator('.doffi-office-titlebar')).toHaveCSS('height', '40px');
    await expect(region.locator('.office-tabs')).toHaveCSS('height', '32px');
    await expect(region.locator('.office-ribbon-panel')).toHaveCSS('height', '74px');
    await expect(region.getByRole('button', { name: '粘贴', exact: true }).locator('svg')).toHaveAttribute('width', '32');
  }
  await expect(first.getByRole('group', { name: '字体', exact: true })).toHaveCSS('width', '375px');
  await expect(first.getByRole('group', { name: '段落', exact: true })).toHaveCSS('width', '338px');
  await first.getByRole('button', { name: '字体颜色选项', exact: true }).click();
  const colorMenu = first.getByRole('menu', { name: '字体颜色选项', exact: true }); await expect(colorMenu).toBeVisible();
  const triggerBox = (await first.getByRole('button', { name: '字体颜色选项', exact: true }).boundingBox())!, menuBox = (await colorMenu.boundingBox())!;
  expect(menuBox.y - triggerBox.y - triggerBox.height).toBeGreaterThanOrEqual(0); expect(menuBox.y - triggerBox.y - triggerBox.height).toBeLessThan(8);
  await first.getByRole('button', { name: '字体颜色选项', exact: true }).press('Escape');
  await expect(first.getByRole('menu')).toHaveCount(0);
  await page.setViewportSize({ width: 1100, height: 900 });
  await expect(first.locator('.style-gallery')).toBeHidden();
  await expect(first.getByRole('button', { name: '更多样式', exact: true }).locator('svg')).toBeVisible();

});

for (const [label, offset] of [['实施报告', 0], ['实施报告', 4], ['第一项', 0], ['第一项', 3], ['子项', 2], ['第二项', 3], ['从四开始', 4]] as const) {
  test(`native DOCX Enter keeps its new paragraph editable: ${label} at ${offset}`, async ({ page, browserName }, info) => {
    const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
    await page.goto('/'); await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles('examples/docx/01-text-numbering.docx');
    const frame = page.frameLocator('iframe[title="正文编辑"]'), original = frame.locator('[data-word-script]').filter({ hasText: label }).first();
    await expect(original.locator('xpath=ancestor::*[@contenteditable="true"]')).toBeVisible();
    const id = await original.evaluate((node, offset) => {
      node.closest<HTMLElement>('[contenteditable]')!.focus(); const text = document.createTreeWalker(node, 4).nextNode()!, range = document.createRange(); range.setStart(text, offset); range.collapse(true);
      const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range); return node.closest<HTMLElement>('p,h1')!.dataset.nodeId!;
    }, offset);
    await page.keyboard.press('Enter'); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '1');
    const previous = frame.locator(`[data-node-id="${id}"]`), nextId = await previous.evaluate(node => { let next = node.nextElementSibling; while (next && !next.hasAttribute('data-node-id')) next = next.nextElementSibling; return next!.getAttribute('data-node-id')!; });
    const next = frame.locator(`[data-node-id="${nextId}"]`), logical = async (node: typeof next) => node.evaluate(element => { const marker = element.querySelector('.sdk-number-marker'); return element.textContent!.slice(marker?.textContent?.length ?? 0); });
    await next.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
    expect((await next.boundingBox())!.height).toBeGreaterThan(0);
    if (browserName === 'chromium') {
      const cdp = await page.context().newCDPSession(page);
      await cdp.send('Input.imeSetComposition', { text: 'xinzeng', selectionStart: 7, selectionEnd: 7 });
      await cdp.send('Input.imeSetComposition', { text: '新增中文', selectionStart: 4, selectionEnd: 4 });
      await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '1');
      await cdp.send('Input.insertText', { text: '新增中文😀' }); await cdp.detach();
    } else await page.keyboard.insertText('新增中文😀');
    await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '2');
    await expect.poll(() => logical(previous)).toBe(label.slice(0, offset)); await expect.poll(() => logical(next)).toBe('新增中文😀' + label.slice(offset));
    await expect(page.getByRole('alert')).toHaveCount(0);
    await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect.poll(() => logical(next)).toBe(label.slice(offset));
    await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect(next).toHaveCount(0); await expect.poll(() => logical(previous)).toBe(label);
    await page.getByRole('button', { name: '重做', exact: true }).click(); await page.getByRole('button', { name: '重做', exact: true }).click(); await expect.poll(() => logical(next)).toBe('新增中文😀' + label.slice(offset));
    if (offset === label.length && (label === '实施报告' || label === '第一项')) {
      await page.getByRole('button', { name: '文件菜单', exact: true }).click(); await page.getByRole('menuitem', { name: '文件和导出…', exact: true }).click(); const download = page.waitForEvent('download'); await page.getByRole('button', { name: '导出 DOCX', exact: true }).click();
      const path = info.outputPath('enter-reopen.docx'); await (await download).saveAs(path); await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles(path);
      await expect(page.getByRole('heading', { name: 'enter-reopen.docx', exact: true })).toBeVisible();
      await expect(frame.locator('[data-office-editor]')).toContainText('新增中文😀'); await expect(page.getByRole('alert')).toHaveCount(0);
    }
    expect(errors).toEqual([]);
  });
}


test('native paragraph dialog shows current values, cancels without an edit and converts character and line units in one undo', async ({ page }) => {
  await page.goto('/'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="n-summary"]');
  await paragraph.evaluate(node => { node.closest<HTMLElement>('[contenteditable]')!.focus(); const range = document.createRange(); range.selectNodeContents(node); const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range); document.dispatchEvent(new Event('selectionchange')); });
  const font = await paragraph.evaluate(node => parseFloat(getComputedStyle(node).fontSize) * .75);
  const before = await paragraph.evaluate(node => ({ marginLeft: getComputedStyle(node).marginLeft, marginBottom: getComputedStyle(node).marginBottom }));
  const open = async () => { await page.getByRole('button', { name: '行和段落间距', exact: true }).click(); await page.getByRole('menuitem', { name: '行距选项…', exact: true }).click(); };
  await open(); let dialog = page.getByRole('dialog', { name: '段落', exact: true });
  await expect(dialog).toHaveCSS('width', '580px'); await dialog.getByLabel('左侧（字符）', { exact: true }).fill('2'); await dialog.press('Escape');
  await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '0');
  await open(); dialog = page.getByRole('dialog', { name: '段落', exact: true });
  await dialog.getByLabel('左侧（字符）', { exact: true }).fill('1.5'); await dialog.getByLabel('段后（行）', { exact: true }).fill('1.2'); await dialog.getByRole('button', { name: '确定', exact: true }).click();
  await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '1');
  await expect.poll(() => paragraph.evaluate(node => parseFloat(getComputedStyle(node).marginLeft) * .75)).toBeCloseTo(font * 1.5, 1);
  await expect.poll(() => paragraph.evaluate(node => parseFloat(getComputedStyle(node).marginBottom) * .75)).toBeCloseTo(font * 1.2, 1);
  await page.getByRole('button', { name: '撤销', exact: true }).click();
  expect(await paragraph.evaluate(node => ({ marginLeft: getComputedStyle(node).marginLeft, marginBottom: getComputedStyle(node).marginBottom }))).toEqual(before);
});

test('Excel workspace opens its native modal above hidden SDK panels and formula buttons use the shared history', async ({ page }) => {
  await page.addInitScript(() => { const attach = Element.prototype.attachShadow; Element.prototype.attachShadow = function(init) { return attach.call(this, { ...init, mode: 'open' }); }; });
  await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto('/');
  await page.getByLabel('打开文件', { exact: true }).filter({ visible: false }).setInputFiles('examples/xlsx/01-values-text.xlsx');
  const formats = page.getByRole('combobox', { name: '数字格式', exact: true }); await expect(formats).toBeEnabled();
  await page.getByRole('button', { name: '格式', exact: true }).click(); await page.getByRole('menuitem', { name: '设置单元格格式…', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: '设置单元格格式', exact: true });
  await expect.poll(() => page.locator('dialog[aria-label="设置单元格格式"]').evaluateAll(nodes => nodes.map(node => { const el = node as HTMLDialogElement; return { open: el.open, display: getComputedStyle(el).display, parent: getComputedStyle(el.parentElement!).display }; }))).toEqual([{ open: true, display: 'block', parent: 'contents' }]); await expect(dialog).toBeVisible();
  expect(await dialog.getByRole('tab').allTextContents()).toEqual(['数字', '对齐', '字体', '边框', '填充', '保护']);
  await dialog.getByRole('tab', { name: '字体', exact: true }).click(); await expect(dialog.getByLabel('西文字体', { exact: true })).toBeVisible();
  await dialog.getByRole('button', { name: '取消', exact: true }).click(); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '0');
  await formats.fill('货币'); await formats.press('Enter'); await expect(formats).toHaveValue('货币'); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '1');
  const input = page.getByRole('textbox', { name: '单元格或公式输入', exact: true }); await input.fill('42');
  await page.getByRole('button', { name: '确认单元格输入', exact: true }).click(); await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '2'); await expect(input).toHaveValue('42');
  await input.fill('cancelled'); await page.getByRole('button', { name: '取消单元格输入', exact: true }).click(); await expect(input).toHaveValue('42');
  await expect(page.locator('.office-app')).toHaveAttribute('data-revision', '2');
  const footer = (await page.locator('.doffi-office-statusbar').boundingBox())!; expect(footer.y + footer.height).toBeLessThanOrEqual(1001);
});
