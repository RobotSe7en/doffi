import { readFileSync, readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { writeDocumentZip } from "@doffi/office";
import { openSdkHost } from "./sdk-host.js";
import type * as fixture from "./fixtures/workbook-workbench.js";
declare global { interface Window { workbookWorkbench: typeof fixture } }
let code = "", worker = "";
const root = new URL("../../examples/workbook/", import.meta.url), encoded = Buffer.from(writeDocumentZip(new Map(readdirSync(root).filter(name => /\.(?:json|html|css)$/.test(name)).map(name => [name, new Uint8Array(readFileSync(new URL(name, root)))])))).toString("base64");
test.beforeAll(async () => {
  for (const name of ["workbook-workbench", "workbook-session-worker"]) {
    const result = await build({ configFile: false, logLevel: "silent", define: { "process.env.NODE_ENV": JSON.stringify("production") }, resolve: { alias: { react: fileURLToPath(new URL("../../apps/playground/node_modules/react", import.meta.url)), "react-dom": fileURLToPath(new URL("../../apps/playground/node_modules/react-dom", import.meta.url)) } }, build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL(`./fixtures/${name}.ts`, import.meta.url)), formats: ["iife"], name: "WorkbookWorkbenchFixture" } } });
    const parts = Array.isArray(result) ? result.flatMap(part => part.output) : "output" in result ? result.output : [];
    const text = parts.flatMap(part => part.type === "chunk" ? [part.code] : []).join("\n"); if (name === "workbook-workbench") code = text; else worker = text;
  }
});
for (const remote of [false, true]) {
  const name = remote ? "unique Worker" : "direct source";
  test.describe(name, () => {
    test.beforeEach(async ({ page }) => { await openSdkHost(page); await page.addScriptTag({ content: code }); await page.evaluate(({ encoded, worker }) => window.workbookWorkbench.mount(encoded, worker || undefined), { encoded, worker: remote ? worker : "" }); });
    test.afterEach(async ({ page }) => { await page.evaluate(() => window.workbookWorkbench.finish()); });
    test("reviews isolated source and consumer results, then applies once and undoes the instruction and content", async ({ page }) => {
      await page.getByRole("button", { name: "生成工作簿请求", exact: true }).click(); await expect(page.getByRole("status", { name: "工作簿协作状态", exact: true })).toContainText("请求已生成");
      const original = await page.evaluate(() => window.workbookWorkbench.inspect()); expect(original.request!.context.cells[1]).toMatchObject({ access: "read-only", value: { value: 999 }, current: { value: { value: 5 } } });
      await page.evaluate(() => window.workbookWorkbench.load());
      await page.getByText("当前结果与只读消费者重算（2）", { exact: true }).click(); await expect(page.getByLabel("工作簿提案差异")).toContainText("只读源重算");
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ revision: 0, cells: [{ value: { value: 4 } }, { value: { value: 999 } }], annotations: [{ status: "open" }] });
      await page.getByRole("button", { name: "整批应用工作簿提案", exact: true }).click(); await expect.poll(() => page.evaluate(() => window.workbookWorkbench.revision())).toBe(1);
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ history: { undoSteps: 1 }, cells: [{ value: { value: 10 } }, { formula: { expression: "=A1+1" }, value: { value: 999 } }], annotations: [{ status: "resolved" }] });
      await page.getByRole("button", { name: "撤销", exact: true }).click(); await expect.poll(() => page.evaluate(() => window.workbookWorkbench.revision())).toBe(2); await page.evaluate(() => window.workbookWorkbench.settled()); expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ revision: 2, annotations: [{ status: "open" }], cells: [{ value: { value: 4 } }, { formula: { expression: "=A1+1" }, value: { value: 999 } }] });
    });
    test("rejects no-op writes to a read-only source, releases failed candidates and preserves a pending input", async ({ page }) => {
      await page.evaluate(() => window.workbookWorkbench.build()); const problem = await page.evaluate(() => window.workbookWorkbench.loadNoop().then(() => "accepted", (error: { code: string }) => error.code)); expect(problem).toBe("InvalidProposal");
      await page.evaluate(() => window.workbookWorkbench.load()); await page.getByRole("button", { name: "放弃工作簿提案", exact: true }).click(); await expect(page.getByLabel("工作簿提案差异")).toBeHidden();
      await page.evaluate(() => window.workbookWorkbench.focusInput()); const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await input.fill("draft😀");
      await expect(page.getByRole("button", { name: "生成工作簿请求", exact: true })).toBeDisabled(); expect(await page.evaluate(() => window.workbookWorkbench.build().then(() => "accepted", (error: { code: string }) => error.code))).toBe("InvalidOperation"); await expect(input).toHaveValue("draft😀"); await page.evaluate(() => window.workbookWorkbench.cancelInput());
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, annotations: [{ status: "open" }] });
      await page.getByText("名称管理", { exact: true }).click(); const name = page.getByLabel("定义名称", { exact: true }); await name.fill("UnsavedName");
      await expect(page.getByRole("button", { name: "生成工作簿请求", exact: true })).toBeDisabled();
      expect(await page.evaluate(() => window.workbookWorkbench.build().then(() => "accepted", (error: { code: string }) => error.code))).toBe("InvalidOperation"); await expect(name).toHaveValue("UnsavedName");
      await page.getByRole("button", { name: "放弃当前名称草稿并重新读取", exact: true }).click(); await expect(page.getByRole("button", { name: "生成工作簿请求", exact: true })).toBeEnabled();
    });
    test("host-returned proposals stay candidates and late host results cannot overwrite a newer revision", async ({ page }) => {
      await page.getByRole("button", { name: "提交工作簿批注", exact: true }).click(); await expect(page.getByRole("button", { name: "整批应用工作簿提案", exact: true })).toBeEnabled();
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ revision: 0, annotations: [{ status: "open" }] }); await page.getByRole("button", { name: "放弃工作簿提案", exact: true }).click();
      await page.evaluate(() => window.workbookWorkbench.mode("hold")); await page.getByRole("button", { name: "提交工作簿批注", exact: true }).click(); await expect.poll(() => page.evaluate(async () => (await window.workbookWorkbench.inspect()).held)).toBe(true);
      await page.evaluate(() => window.workbookWorkbench.external(7)); await page.evaluate(() => window.workbookWorkbench.deliver(20)); await expect(page.getByLabel("工作簿请求与提案").getByRole("alert")).toContainText("工作簿已变化");
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ revision: 1, cells: [{ value: { value: 7 } }, { formula: { expression: "=A1+1" }, value: { value: 999 } }], annotations: [{ status: "open" }] }); await expect(page.getByRole("button", { name: "整批应用工作簿提案", exact: true })).toBeDisabled();
    });
    test("cancels a held host callback and releases the view while the owning session remains usable", async ({ page }) => {
      await page.evaluate(() => window.workbookWorkbench.mode("hold")); await page.getByRole("button", { name: "提交工作簿批注", exact: true }).click(); await expect.poll(() => page.evaluate(async () => (await window.workbookWorkbench.inspect()).held)).toBe(true);
      await page.getByRole("button", { name: "取消协作操作", exact: true }).click(); await expect(page.getByRole("button", { name: "生成工作簿请求", exact: true })).toBeEnabled(); await page.evaluate(() => window.workbookWorkbench.deliver());
      expect(await page.evaluate(() => window.workbookWorkbench.disposeView())).toMatchObject({ disposed: true, revision: 1, cells: [{ value: { value: 42 } }, { formula: { expression: "=A1+1" }, value: { value: 999 } }], annotations: [{ status: "open" }] });
    });
    test("downloads exact frozen JSON and imports formulas only after explicit read expansion", async ({ page }) => {
      const request = await page.evaluate(() => window.workbookWorkbench.build());
      const [download] = await Promise.all([page.waitForEvent("download"), page.getByRole("button", { name: "下载工作簿请求", exact: true }).click()]);
      expect(JSON.parse(readFileSync((await download.path())!, "utf8"))).toEqual(request);
      const packet = (value: typeof request, operations: unknown[]) => ({ protocolVersion: 1, requestId: value.requestId, documentId: value.documentId, baseRevision: value.baseRevision, proposalId: "file", annotationIds: ["review-1"], summary: "使用额外只读输入", operations });
      const formula = { op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=A20+1" } }] };
      const input = page.getByLabel("导入工作簿提案 JSON", { exact: true });
      await input.setInputFiles({ name: "formula.json", mimeType: "application/json", buffer: Buffer.from(JSON.stringify(packet(request, [formula]))) });
      await expect.poll(() => page.evaluate(async () => (await window.workbookWorkbench.inspect()).errors.at(-1))).toBe("InvalidProposal");
      await page.getByLabel("额外只读输入区域", { exact: true }).fill("A20");
      const expanded = await page.evaluate(() => window.workbookWorkbench.build()); expect(expanded.context.writableRanges).toEqual(request.context.writableRanges); expect(expanded.context.requestedReadRanges).toEqual([{ sheetId: "sheet-1", range: { from: { row: 19, column: 0 }, to: { row: 19, column: 0 } } }]);
      await input.setInputFiles({ name: "formula.json", mimeType: "application/json", buffer: Buffer.from(JSON.stringify(packet(expanded, [formula]))) });
      await expect(page.getByRole("button", { name: "整批应用工作簿提案", exact: true })).toBeEnabled(); expect(await page.evaluate(() => window.workbookWorkbench.revision())).toBe(0);
      await page.getByRole("button", { name: "放弃工作簿提案", exact: true }).click(); await expect(page.getByLabel("工作簿提案差异")).toBeHidden();
      const beforeErrors = (await page.evaluate(() => window.workbookWorkbench.inspect())).errors.length;
      await input.setInputFiles({ name: "readonly.json", mimeType: "application/json", buffer: Buffer.from(JSON.stringify(packet(expanded, [{ op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 19, column: 0 }] }]))) });
      await expect.poll(() => page.evaluate(async () => (await window.workbookWorkbench.inspect()).errors.length)).toBe(beforeErrors + 1); expect(await page.evaluate(() => window.workbookWorkbench.revision())).toBe(0);
    });
    test("pages every direct source difference without retaining all rows in the DOM", async ({ page }) => {
      await page.evaluate(() => window.workbookWorkbench.selectRows(150)); await page.getByLabel("请求写入范围", { exact: true }).selectOption("selection"); await page.evaluate(() => window.workbookWorkbench.build()); await page.evaluate(() => window.workbookWorkbench.formatRows(150));
      await page.getByText("直接修改的值、公式与格式", { exact: true }).click(); const region = page.getByLabel("工作簿来源差异分页", { exact: true });
      await expect(region).toContainText("1–50 / 150 个单元格"); await expect(region.locator("details")).toHaveCount(50);
      await page.getByRole("button", { name: "下一页源差异", exact: true }).click(); await expect(region).toContainText("51–100 / 150 个单元格");
      await page.getByRole("button", { name: "下一页源差异", exact: true }).click(); await expect(region).toContainText("101–150 / 150 个单元格"); await expect(region.locator("details")).toHaveCount(50);
      await page.getByRole("button", { name: "放弃工作簿提案", exact: true }).click(); await expect(region).toHaveCount(0); expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, annotations: [{ status: "open" }] });
    });
    test("disposing after apply starts preserves the authoritative commit receipt", async ({ page }) => {
      await page.evaluate(() => window.workbookWorkbench.build()); await page.evaluate(() => window.workbookWorkbench.load());
      expect(await page.evaluate(() => window.workbookWorkbench.applyAndDispose())).toMatchObject({ changed: true, revision: 1 });
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ disposed: true, revision: 1, history: { undoSteps: 1 }, cells: [{ value: { value: 10 } }, { formula: { expression: "=A1+1" } }], annotations: [{ status: "resolved" }] });
    });
    test("drawing source review shows exact old and new rich bodies while rejection preserves the source", async ({ page }) => {
      const nodeId = await page.evaluate(() => window.workbookWorkbench.addObject()); await page.getByLabel("允许修改的绘图", { exact: true }).selectOption(nodeId);
      await page.evaluate(() => window.workbookWorkbench.build()); await page.evaluate(nodeId => window.workbookWorkbench.replaceObject(nodeId), nodeId);
      await page.getByText("结构、规则、绘图、批注锚点与样式", { exact: true }).click();
      await expect(page.getByLabel("工作簿提案差异")).toContainText("原对象"); await expect(page.getByLabel("工作簿提案差异")).toContainText("新对象😀");
      expect((await page.evaluate(() => window.workbookWorkbench.inspect())).content).not.toContain("新对象😀");
      await page.getByRole("button", { name: "放弃工作簿提案", exact: true }).click(); await expect(page.getByLabel("工作簿提案差异")).toBeHidden();
      expect((await page.evaluate(() => window.workbookWorkbench.inspect())).content).toContain("原对象"); expect(await page.evaluate(() => window.workbookWorkbench.revision())).toBe(1);
    });
  });
}

for (const remote of [false, true]) test(`React workbook wrapper owns view cleanup over ${remote ? "Worker" : "direct"} source`, async ({ page }) => {
  await openSdkHost(page); await page.addScriptTag({ content: code });
  await page.evaluate(({ encoded, worker }) => window.workbookWorkbench.mount(encoded, worker || undefined, true), { encoded, worker: remote ? worker : "" });
  await page.getByRole("button", { name: "提交工作簿批注", exact: true }).click(); await expect(page.getByRole("button", { name: "整批应用工作簿提案", exact: true })).toBeEnabled();
  expect(await page.evaluate(() => window.workbookWorkbench.disposeView())).toMatchObject({ disposed: true, revision: 1, cells: [{ value: { value: 42 } }, { formula: { expression: "=A1+1" }, value: { value: 999 } }], annotations: [{ status: "open" }] });
  const audit = await page.evaluate(() => window.workbookWorkbench.finish()); if (remote) expect(audit).toMatchObject({ sessionsCreated: 1, endpointDisposed: true });
});
