import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { openDocument } from "@doffi/core";
import { importDocx, readDocumentZip } from "@doffi/office";
async function report(page: Page): Promise<void> {
  await page.addInitScript(() => { localStorage.setItem("doffi:library", "open"); localStorage.setItem("doffi:editing", "preview"); });
  await page.goto("/"); await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
}
async function typeTitle(page: Page): Promise<void> {
  await page.getByRole("button", { name: "编辑正文", exact: true }).click();
  const frame = page.frameLocator('iframe[title="正文编辑"]');
  await expect(frame.locator('[data-node-id="n-title"]')).toBeVisible();
  await frame.locator('[data-node-id="n-title"]').evaluate((element) => {
    const range = document.createRange(); range.selectNodeContents(element); range.collapse(false);
    const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
    (document.querySelector('[data-office-editor]') as HTMLElement).focus(); document.dispatchEvent(new Event("selectionchange"));
  });
  await page.keyboard.insertText("（人工）");
  await expect(frame.locator('[data-node-id="n-title"]')).toHaveText("项目实施报告（人工）");
  await page.getByRole("button", { name: "返回预览与批注", exact: true }).click();
}
function mainText(model: ReturnType<typeof openDocument>): string {
  const story = model.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!; return model.getNode(story.rootNodeId)!.text;
}
test("Playground closes HTML/manual input/proposal/DOCX/PDF/package/reopen through public UI", async ({ page }, info) => {
  const errors: string[] = []; page.on("pageerror", (error) => errors.push(error.message));
  await report(page); await typeTitle(page);
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "统一提交批注", exact: true }).click();
  await page.getByRole("button", { name: "载入固定示例提案", exact: true }).click();
  await expect(page.getByLabel("提案差异", { exact: true })).toBeVisible();
  await page.getByRole("button", { name: "整批应用提案", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
  await page.getByRole("tab", { name: "文件", exact: true }).click();
  let download = page.waitForEvent("download"); await page.getByRole("button", { name: "导出 DOCX", exact: true }).click();
  const docx = info.outputPath("edited.docx"); await (await download).saveAs(docx);
  const reopened = openDocument(importDocx(new Uint8Array(readFileSync(docx)), { documentId: "ui-docx" }));
  expect(mainText(reopened)).toContain("项目实施报告（人工）"); expect(mainText(reopened)).toContain("第四季度");
  expect(reopened.snapshot().annotations).toHaveLength(1); expect(reopened.snapshot().annotations[0]!.status).toBe("resolved");
  await expect(page.getByLabel("DOCX 导出报告")).toContainText("a-summary：需要重新定位");
  download = page.waitForEvent("download"); await page.getByRole("button", { name: "下载导出报告", exact: true }).click();
  const reportPath = info.outputPath("report.json"); await (await download).saveAs(reportPath);
  const exportedReport = JSON.parse(readFileSync(reportPath, "utf8"));
  expect(exportedReport.revision).toBe(2); expect(exportedReport.unexportedAnnotations).toEqual([{ annotationId: "a-summary", reason: "needs_reanchor" }]);
  download = page.waitForEvent("download"); await page.getByRole("button", { name: "导出 PDF", exact: true }).click();
  const pdf = info.outputPath("edited.pdf"); await (await download).saveAs(pdf);
  const text = execFileSync("pdftotext", [pdf, "-"], { encoding: "utf8" }).replace(/\s/g, "");
  expect(text).toContain("项目实施报告（人工）"); expect(text).toContain("第四季度");
  download = page.waitForEvent("download"); await page.getByRole("button", { name: "下载文档包", exact: true }).click();
  const archive = info.outputPath("edited.zip"); await (await download).saveAs(archive);
  const saved = openDocument(readDocumentZip(new Uint8Array(readFileSync(archive))));
  expect(saved.snapshot().manifest.revision).toBe(2); expect(saved.getNode("n-title")!.text).toBe("项目实施报告（人工）");
  expect(saved.snapshot().annotations.map((annotation) => annotation.status)).toEqual(["needs_reanchor", "resolved"]);
  await page.getByLabel("打开文件", { exact: true }).setInputFiles({ name: "bad.docx", mimeType: "application/octet-stream", buffer: Buffer.from([1, 2, 3]) });
  await expect(page.getByRole("alert")).toBeVisible(); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "3");
  await page.getByRole("button", { name: "重做", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "4");
  await page.getByLabel("打开文件", { exact: true }).setInputFiles(docx);
  await expect(page.locator(".office-app")).toHaveAttribute("data-annotation-count", "1");
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await expect(page.getByLabel("DOCX 导出报告")).toHaveCount(0);
  await page.getByLabel("打开文件", { exact: true }).setInputFiles(archive);
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
  await expect(page.locator(".office-app")).toHaveAttribute("data-annotation-count", "2");
  expect(errors).toEqual([]);
});
test("switching documents terminates delayed DOCX work and never downloads a stale revision", async ({ page }) => {
  await page.addInitScript(() => {
    const prototype = Worker.prototype as unknown as { postMessage(message: unknown, transfer?: Transferable[]): void };
    const original = prototype.postMessage, state = window as unknown as { latePosts: number }; state.latePosts = 0;
    prototype.postMessage = function (message, transfer): void { setTimeout(() => { state.latePosts++; original.call(this, message, transfer); }, 1200); };
  });
  await report(page); const downloads: string[] = []; page.on("download", (download) => downloads.push(download.suggestedFilename()));
  await page.getByRole("tab", { name: "文件", exact: true }).click();
  await page.getByRole("button", { name: "导出 DOCX", exact: true }).click();
  await expect(page.getByRole("status").filter({ hasText: "正在导出 DOCX" })).toBeVisible();
  await page.getByRole("button", { name: "打开 DOCX 样例", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-annotation-count", "0");
  await page.waitForFunction(() => (window as unknown as { latePosts: number }).latePosts > 0);
  expect(downloads).toEqual([]); await expect(page.getByLabel("DOCX 导出报告")).toHaveCount(0);
  await page.getByRole("tab", { name: "文件", exact: true }).click();
  await expect(page.getByRole("button", { name: "导出 DOCX", exact: true })).toBeEnabled();
});
for (const action of ["cancel", "edit"] as const) test(`${action} stops DOCX export without downloading an outdated result`, async ({ page }) => {
  await page.addInitScript(() => {
    const prototype = Worker.prototype as unknown as { postMessage(message: unknown, transfer?: Transferable[]): void };
    const original = prototype.postMessage;
    const terminate = Worker.prototype.terminate, state = window as unknown as { terminatedWorkers: number }; state.terminatedWorkers = 0;
    Worker.prototype.terminate = function (): void { state.terminatedWorkers++; terminate.call(this); };
    prototype.postMessage = function (message, transfer): void { setTimeout(() => original.call(this, message, transfer), 5000); };
  });
  await report(page); const downloads: string[] = []; page.on("download", (value) => downloads.push(value.suggestedFilename()));
  await page.getByRole("tab", { name: "文件", exact: true }).click();
  await page.getByRole("button", { name: "导出 DOCX", exact: true }).click();
  await expect(page.getByRole("status").filter({ hasText: "正在导出 DOCX" })).toBeVisible();
  if (action === "cancel") await page.getByRole("button", { name: "取消导出", exact: true }).click();
  else await typeTitle(page);
  await expect(page.getByRole("alert")).toContainText(action === "cancel" ? "已取消 DOCX 导出" : "导出期间文档已变化");
  await page.getByRole("tab", { name: "文件", exact: true }).click();
  await expect(page.getByRole("button", { name: "导出 DOCX", exact: true })).toBeEnabled();
  expect(await page.evaluate(() => (window as unknown as { terminatedWorkers: number }).terminatedWorkers)).toBe(1);
  await expect(page.getByLabel("DOCX 导出报告")).toHaveCount(0); expect(downloads).toEqual([]);
});
