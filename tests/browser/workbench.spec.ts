import { openSdkHost } from './sdk-host.js';
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { build } from "vite";
import type { DocumentInteractionRequest, EditProposal } from "@doffi/core";
import type { WorkbenchHost } from "./fixtures/workbench-host.js";

declare global { interface Window { workbenchHost: WorkbenchHost } }
let hostCode: string;
const report = ["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [
  path, [...readFileSync(new URL("../../examples/report/" + path, import.meta.url))],
] as [string, number[]]);
const exampleProposal = JSON.parse(readFileSync(new URL("../../examples/report.proposal.json", import.meta.url), "utf8")) as EditProposal;
async function mount(page: Page): Promise<void> {
  await openSdkHost(page); await page.addScriptTag({ type: "module", content: hostCode });
  await page.waitForFunction(() => !!window.workbenchHost); await page.evaluate((files) => window.workbenchHost.mount(files), report);
}
async function generated(page: Page): Promise<DocumentInteractionRequest> {
  await page.getByRole("button", { name: "生成请求", exact: true }).click();
  return page.evaluate(() => window.workbenchHost.view!.request()!);
}
function proposal(request: DocumentInteractionRequest, changes: Partial<EditProposal> = {}): EditProposal {
  return { ...exampleProposal, requestId: request.requestId, documentId: request.documentId, baseRevision: request.baseRevision, ...changes };
}
async function importProposal(page: Page, value: EditProposal): Promise<void> {
  await page.getByLabel("导入提案 JSON", { exact: true }).setInputFiles({ name: "proposal.json", mimeType: "application/json", buffer: Buffer.from(JSON.stringify(value)) });
}
test.beforeAll(async () => {
  const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: {
    entry: fileURLToPath(new URL("./fixtures/workbench-host.ts", import.meta.url)), formats: ["es"], fileName: "workbench-host",
  } } });
  if (!("output" in result) && !Array.isArray(result)) throw new Error("Expected browser fixture build output");
  const output = Array.isArray(result) ? result.flatMap((entry) => entry.output) : result.output;
  hostCode = output.flatMap((item) => item.type === "chunk" ? [item.code] : []).join("\n");
});

test("plain JS host exports a frozen two-annotation request, reviews a proposal, applies and undoes atomically", async ({ page }, info) => {
  await mount(page);
  const request = await generated(page);
  expect(request.annotations).toHaveLength(2);
  expect(request.context.blocks.map((block) => block.nodeId)).toEqual(["n-title", "n-summary"]);
  const download = page.waitForEvent("download"); await page.getByRole("button", { name: "下载请求 JSON" }).click();
  const path = info.outputPath("request.json"); await (await download).saveAs(path);
  expect(JSON.parse(readFileSync(path, "utf8"))).toEqual(request);
  await importProposal(page, proposal(request));
  await expect(page.getByLabel("提案差异", { exact: true })).toBeVisible();
  await expect(page.getByLabel("修改后").filter({ hasText: "第四季度" })).toHaveCount(1);
  expect(await page.evaluate(() => window.workbenchHost.session!.snapshot().manifest.revision)).toBe(0);
  await page.getByRole("button", { name: "整批应用提案", exact: true }).click();
  await expect(page.getByLabel("会话状态")).toContainText("修订 1");
  expect(await page.evaluate(() => ({ text: window.workbenchHost.session!.getNode("n-summary")!.text,
    statuses: window.workbenchHost.session!.getAnnotations().map((a) => a.status), history: window.workbenchHost.session!.history().undoSteps })))
    .toEqual({ text: exampleProposal.operations[0]!.op === "replaceText" ? exampleProposal.operations[0]!.text : "", statuses: ["needs_reanchor", "resolved"], history: 1 });
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  expect(await page.evaluate(() => window.workbenchHost.session!.getAnnotations().map((a) => a.status))).toEqual(["open", "open"]);
  await page.getByRole("button", { name: "重做", exact: true }).click();
  expect(await page.evaluate(() => window.workbenchHost.session!.snapshot().manifest.revision)).toBe(3);
  expect(await page.evaluate(() => window.workbenchHost.errors)).toEqual([]);
});

test("choosing an annotation subset bounds the request, and rejecting a candidate changes nothing", async ({ page }) => {
  await mount(page);
  await page.getByLabel("提交批注：改为第四季度完成，交付内容说明包含系统上线和用户培训。", { exact: true }).uncheck();
  const request = await generated(page);
  expect(request.annotations.map((a) => a.annotationId)).toEqual(["a-heading"]);
  await importProposal(page, proposal(request, { annotationIds: ["a-heading"], operations: [exampleProposal.operations[1]!] }));
  await expect(page.getByRole("button", { name: "整批应用提案" })).toBeEnabled();
  await page.getByRole("button", { name: "拒绝提案", exact: true }).click();
  await expect(page.getByLabel("提案差异", { exact: true })).toBeHidden();
  expect(await page.evaluate(() => ({ revision: window.workbenchHost.session!.snapshot().manifest.revision, history: window.workbenchHost.session!.history().undoSteps })))
    .toEqual({ revision: 0, history: 0 });
});

test("submission hands data to the host and host failure preserves document and annotations", async ({ page }) => {
  await mount(page);
  await page.evaluate(() => { window.workbenchHost.failSubmission = true; });
  await page.getByRole("button", { name: "统一提交批注" }).click();
  await expect(page.getByRole("alert")).toContainText("host could not accept");
  expect(await page.evaluate(() => window.workbenchHost.session!.snapshot().manifest.revision)).toBe(0);
  await page.evaluate(() => { window.workbenchHost.failSubmission = false; });
  await page.getByRole("button", { name: "统一提交批注" }).click();
  await expect(page.getByLabel("协作状态")).toContainText("请求已交给宿主");
  expect(await page.evaluate(() => ({ requests: window.workbenchHost.requests.length, event: window.workbenchHost.events.at(-1),
    frozen: Object.isFrozen(window.workbenchHost.requests[0]!.context.blocks[0]) })))
    .toEqual({ requests: 1, event: "annotationsSubmitted", frozen: true });
});

test("an external edit disables stale application while retaining proposal differences", async ({ page }) => {
  await mount(page); const request = await generated(page); await importProposal(page, proposal(request));
  await expect(page.getByRole("button", { name: "整批应用提案" })).toBeEnabled();
  await page.evaluate(() => {
    const session = window.workbenchHost.session!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "report-demo", baseRevision: 0,
      operations: [{ op: "setStyle", nodeId: "n-title", property: "color", value: "red" }] }));
  });
  await expect(page.getByRole("button", { name: "整批应用提案" })).toBeDisabled();
  await expect(page.getByLabel("提案差异", { exact: true })).toBeVisible();
  await expect(page.getByLabel("协作状态")).toContainText("请求已过期");
  await page.getByRole("button", { name: "拒绝提案", exact: true }).click();
  await importProposal(page, proposal(request));
  await expect(page.getByRole("alert")).toContainText("different document revision");
  expect(await page.evaluate(() => window.workbenchHost.session!.snapshot().manifest.revision)).toBe(1);
});

test("one-candidate budgets support sequential reviews without hidden replacement capacity", async ({ page }) => {
  await mount(page);
  await page.evaluate((files) => window.workbenchHost.mount(files, 1), report);
  const request = await generated(page), value = proposal(request);
  await importProposal(page, value);
  await expect(page.getByRole("button", { name: "整批应用提案" })).toBeEnabled();
  await expect(page.getByLabel("导入提案 JSON", { exact: true })).toBeDisabled();
  const code = await page.evaluate((value) => {
    try { window.workbenchHost.view!.loadProposal(window.workbenchHost.view!.request()!, value); return "success"; }
    catch (error) { return (error as { code: string }).code; }
  }, value);
  expect(code).toBe("InvalidOperation");
  await expect(page.getByRole("button", { name: "整批应用提案" })).toBeEnabled();
  await page.getByRole("button", { name: "拒绝提案", exact: true }).click();
  await importProposal(page, { ...value, proposalId: "second" });
  await expect(page.getByRole("button", { name: "整批应用提案" })).toBeEnabled();
  expect(await page.evaluate(() => window.workbenchHost.session!.history().undoSteps)).toBe(0);
});

for (const deferred of [false, true]) {
  test("host callback can return a proposal before submission completes: deferred=" + deferred, async ({ page }) => {
    await mount(page);
    await page.evaluate(({ value, deferred }) => {
      window.workbenchHost.submissionProposal = value; window.workbenchHost.deferSubmission = deferred;
    }, { value: exampleProposal, deferred });
    await page.getByRole("button", { name: "统一提交批注" }).click();
    await expect(page.getByRole("button", { name: "整批应用提案" })).toBeEnabled();
    if (deferred) await page.evaluate(() => window.workbenchHost.release!());
    await expect.poll(() => page.evaluate(() => window.workbenchHost.events.filter((type) => type === "annotationsSubmitted").length)).toBe(1);
    await expect(page.getByLabel("协作状态")).toContainText("提案已准备");
    expect(await page.evaluate(() => window.workbenchHost.events.filter((type) => type === "proposalReady").length)).toBe(1);
    expect(await page.evaluate(() => window.workbenchHost.session!.snapshot().manifest.revision)).toBe(0);
  });
}

for (const action of ["整批应用提案", "拒绝提案"]) {
  test("late handoff completion preserves a proposal's final status after " + action, async ({ page }) => {
    await mount(page);
    await page.evaluate((value) => { window.workbenchHost.submissionProposal = value; window.workbenchHost.deferSubmission = true; }, exampleProposal);
    await page.getByRole("button", { name: "统一提交批注" }).click();
    await expect(page.getByRole("button", { name: action, exact: true })).toBeEnabled();
    await page.getByRole("button", { name: action, exact: true }).click();
    const before = await page.getByLabel("协作状态").textContent();
    await page.evaluate(() => window.workbenchHost.release!());
    await expect.poll(() => page.evaluate(() => window.workbenchHost.events.filter((type) => type === "annotationsSubmitted").length)).toBe(1);
    await expect(page.getByLabel("协作状态")).toHaveText(before!);
  });
}

test("malicious, mismatched and out-of-scope proposal files leave state intact", async ({ page }) => {
  await mount(page); const request = await generated(page);
  for (const value of [proposal(request, { requestId: "wrong" }), proposal(request, { operations: [{ op: "removeNode", nodeId: "n-item" }] }),
    proposal(request, { operations: [{ op: "replaceFragment", nodeId: "n-summary", html: '<img src="https://example.com/private" onerror="bad()">' }] })]) {
    await importProposal(page, value); await expect(page.getByRole("alert")).toBeVisible();
    expect(await page.evaluate(() => window.workbenchHost.session!.snapshot().manifest.revision)).toBe(0);
  }
  await expect(page.locator('img[src^="https:"]')).toHaveCount(0);
});

test("unmount ignores late submission results and leaves the external session usable", async ({ page }) => {
  await mount(page);
  await page.evaluate(() => { window.workbenchHost.deferSubmission = true; });
  await page.getByRole("button", { name: "统一提交批注" }).click();
  await page.waitForFunction(() => !!window.workbenchHost.release);
  expect(await page.evaluate(async () => {
    const host = window.workbenchHost; host.view!.dispose(); host.release!(); await Promise.resolve(); await Promise.resolve();
    return { submitted: host.events.includes("annotationsSubmitted"), revision: host.session!.snapshot().manifest.revision };
  })).toEqual({ submitted: false, revision: 0 });
  await expect(page.locator("iframe")).toHaveCount(0);
});

test("React Playground submits two annotations, loads a fixed proposal and downloads a reopenable result", async ({ page }, info) => {
  await page.goto("/"); await page.getByRole("button", { name: "打开报告协作样例" }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-annotation-count", "2");
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "统一提交批注" }).click();
  await expect(page.getByText("Playground 已接收批注请求", { exact: false })).toBeVisible();
  await page.getByRole("button", { name: "载入固定示例提案" }).click();
  await expect(page.getByRole("button", { name: "整批应用提案" })).toBeEnabled();
  await page.getByRole("button", { name: "整批应用提案" }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  const download = page.waitForEvent("download"); await page.getByRole("button", { name: "下载文档包" }).click();
  const path = info.outputPath("modified.zip"); await (await download).saveAs(path);
  await page.getByLabel("打开文件").setInputFiles(path);
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]')).toContainText("第四季度");
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-title"]')).toHaveCSS("color", "rgb(31, 75, 153)");
});
