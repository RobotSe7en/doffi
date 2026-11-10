import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { openDocument } from "@doffi/core";
import { importPptx, readDocumentZip } from "@doffi/office";
async function example(page: Page): Promise<void> { await page.addInitScript(() => localStorage.setItem("doffi:library", "open")); await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto("/"); await page.getByRole("button", { name: "打开 PPT 协作样例", exact: true }).click(); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0"); await expect(page.frameLocator('iframe[title="幻灯片只读预览"]').locator('[data-node-id="heading-p"]')).toHaveText("项目进展"); }
async function typeHeading(page: Page): Promise<void> { await page.getByRole("tab", { name: "开始", exact: true }).click(); const frame = page.frameLocator('iframe[title="幻灯片只读预览"]'); await frame.locator('[data-node-id="heading-p"]').click(); await page.getByRole("button", { name: "编辑文字", exact: true }).click(); const input = frame.getByRole("textbox", { name: "幻灯片文字输入" }); await input.press("End"); await page.keyboard.insertText("（人工）"); await expect(input).toContainText("项目进展（人工）"); await page.getByRole("button", { name: "完成输入", exact: true }).click(); }
test("Playground PPT input/proposal/native/PDF/PNG/package/reopen uses the current revision", async ({ page }, info) => {
    test.setTimeout(60000);
    const errors: string[] = [];
    page.on("pageerror", error => errors.push(error.message));
    await example(page);
    await typeHeading(page);
    await page.getByRole("tab", { name: "审阅", exact: true }).click();
    await page.getByRole("checkbox", { name: "补充上线风险及应对措施。", exact: true }).check();
    await page.getByRole("button", { name: "生成 PPT 请求", exact: true }).click();
    await page.getByRole("button", { name: "载入固定示例提案", exact: true }).click();
    await expect(page.getByLabel("提案差异", { exact: true })).toBeVisible();
    await page.getByRole("button", { name: "应用 PPT 提案", exact: true }).click();
    await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
    let pending = page.waitForEvent("download");
    await page.getByRole("tab", { name: "文件", exact: true }).click();
    await page.getByRole("button", { name: "导出 PPTX", exact: true }).click();
    const path = info.outputPath("edited.pptx");
    await (await pending).saveAs(path);
    const native = openDocument(importPptx(new Uint8Array(readFileSync(path)), { documentId: "ui-pptx" })), p = native.snapshot().metadata.presentation!;
    expect(native.getNode(p.slides[0]!.rootNodeId)!.text).toContain("项目进展（人工）");
    expect(native.getNode(p.slides[0]!.rootNodeId)!.text).toContain("11 月正式上线");
    expect(native.snapshot().annotations).toHaveLength(1);
    expect(native.snapshot().annotations[0]!.replies).toHaveLength(1);
    await expect(page.getByLabel("PPTX 导出报告")).toBeVisible();
    pending = page.waitForEvent("download");
    await page.getByRole("button", { name: "下载导出报告", exact: true }).click();
    const reportPath = info.outputPath("report.json");
    await (await pending).saveAs(reportPath);
    const report = JSON.parse(readFileSync(reportPath, "utf8"));
    expect(report.revision).toBe(2);
    expect(report.unexportedAnnotations).toEqual([]);
    for (const format of ["PDF", "PNG"] as const) {
        pending = page.waitForEvent("download");
        await page.getByRole("button", { name: "导出 " + format, exact: true }).click();
        const output = info.outputPath("edited." + format.toLowerCase());
        await (await pending).saveAs(output);
        if (format === "PDF") {
            const text = execFileSync("pdftotext", ["-raw", output, "-"], { encoding: "utf8" }).replace(/\s/g, "");
            expect(text).toContain("项目进展（人工）");
            expect(text).toContain("11月正式上线");
        }
        else
            expect([...readFileSync(output).subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    }
    pending = page.waitForEvent("download");
    await page.getByRole("button", { name: "下载文档包", exact: true }).click();
    const archive = info.outputPath("edited.zip");
    await (await pending).saveAs(archive);
    expect(openDocument(readDocumentZip(new Uint8Array(readFileSync(archive)))).snapshot().manifest.revision).toBe(2);
    await page.getByRole("button", { name: "撤销", exact: true }).click();
    await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "3");
    await expect(page.getByLabel("PPTX 导出报告")).toHaveCount(0);
    await page.getByLabel("打开文件", { exact: true }).setInputFiles(path);
    await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
    await page.getByLabel("打开文件", { exact: true }).setInputFiles(archive);
    await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "2");
    expect(errors).toEqual([]);
});
for (const action of ["cancel", "edit", "switch"] as const)
    test("PPTX worker " + action + " terminates work and releases stale output", async ({ page }) => {
        await page.addInitScript(() => {
            const prototype = Worker.prototype as unknown as {
                postMessage(message: unknown, transfer?: Transferable[]): void;
            };
            const original = prototype.postMessage, terminate = Worker.prototype.terminate, state = window as unknown as {
                terminatedWorkers: number;
                latePosts: number;
            };
            state.terminatedWorkers = 0;
            state.latePosts = 0;
            Worker.prototype.terminate = function () { state.terminatedWorkers++; terminate.call(this); };
            prototype.postMessage = function (message: unknown, transfer?: Transferable[]) { setTimeout(() => { state.latePosts++; original.call(this, message, transfer!); }, 1200); };
        });
        await example(page);
        const downloads: string[] = [];
        page.on("download", download => downloads.push(download.suggestedFilename()));
        await page.getByRole("tab", { name: "文件", exact: true }).click();
        await page.getByRole("button", { name: "导出 PPTX", exact: true }).click();
        await expect(page.getByRole("status").filter({ hasText: "正在导出 PPTX" })).toBeVisible();
        if (action === "cancel")
            await page.getByRole("button", { name: "取消导出", exact: true }).click();
        else if (action === "edit")
            await typeHeading(page);
        else
            await page.getByRole("button", { name: "打开 DOCX 样例", exact: true }).click();
        await expect.poll(() => page.evaluate(() => (window as unknown as {
            terminatedWorkers: number;
        }).terminatedWorkers)).toBe(1);
        await page.waitForFunction(() => (window as unknown as {
            latePosts: number;
        }).latePosts > 0);
        expect(downloads).toEqual([]);
        await expect(page.getByLabel("PPTX 导出报告")).toHaveCount(0);
        if (action !== "switch") { await page.getByRole("tab", { name: "文件", exact: true }).click();
            await expect(page.getByRole("button", { name: "导出 PPTX", exact: true })).toBeEnabled(); }
    });
test("new presentation and native PPTX sample open independently of Word", async ({ page }) => { await page.addInitScript(() => localStorage.setItem("doffi:library", "open")); await page.goto("/"); await page.getByRole("button", { name: "新建 PPT", exact: true }).click(); await page.getByRole("tab", { name: "文件", exact: true }).click(); await expect(page.getByRole("button", { name: "导出 PPTX", exact: true })).toBeEnabled(); await page.getByRole("button", { name: "打开 PPTX 样例", exact: true }).click(); await expect(page.locator(".office-app")).toHaveAttribute("data-annotation-count", "2"); await expect(page.getByRole("button", { name: "导出 DOCX", exact: true })).toHaveCount(0); });
