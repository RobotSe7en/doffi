import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { writeDocumentZip } from "@doffi/office";
import type { WorkbookRenderReport } from "@doffi/render";
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
    const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/" + name + ".ts", import.meta.url)), formats: ["iife"], name: "WorkbookDrawing" } } });
    const output = Array.isArray(result) ? result.flatMap(part => part.output) : "output" in result ? result.output : [];
    bundles.set(name, output.filter(part => part.type === "chunk").map(part => part.code).join("\n"));
  }
});
for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} drawing controls edit typed source, local images and current chart references with history and stale-form protection`, async ({ page }) => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "甲" } }, { row: 1, column: 0, value: { type: "text", value: "乙" } },
    { row: 0, column: 1, value: { type: "number", value: 7 } }, { row: 1, column: 1, value: { type: "number", value: 9 } }]);
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(model.serialize())).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(model.serialize()));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  await page.getByText("图片、形状与图表", { exact: true }).click();
  await page.getByLabel("对象名称", { exact: true }).fill("说明"); await page.getByLabel("对象文字", { exact: true }).fill("中文禰󠄀😀\n第二行");
  await page.getByLabel("形状预设", { exact: true }).selectOption("ellipse"); await page.getByRole("button", { name: "插入基础形状", exact: true }).click();
  expect((await settled() as WorkbookRenderReport).objectFragmentCount).toBeGreaterThan(0); expect(await inspect()).toMatchObject({ drawings: [{ kind: "geometry", preset: "ellipse", name: "说明" }], history: { undoSteps: 1 }, errors: [] });
  await page.getByLabel("选择绘图对象", { exact: true }).selectOption({ label: "说明" });
  await page.getByLabel("对象宽度（磅）", { exact: true }).fill("180"); await page.getByLabel("对象旋转角度", { exact: true }).fill("30");
  await page.getByRole("button", { name: "应用对象位置与属性", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ drawings: [{ anchor: { kind: "one-cell", widthPt: 180 } }], history: { undoSteps: 2 }, errors: [] });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled();
  await expect(page.getByLabel("对象宽度（磅）", { exact: true })).toHaveValue("240");
  await page.getByLabel("对象名称", { exact: true }).fill("未提交对象草稿");
  await page.evaluate(remote => remote ? window.workbookRemoteEditor.external("external") : window.workbookEditor.external("external"), remote); await settled();
  await expect(page.getByLabel("绘图编辑状态", { exact: true })).toContainText("表单草稿已保留");
  await expect(page.getByRole("button", { name: "应用对象位置与属性", exact: true })).toBeDisabled(); await expect(page.getByLabel("对象名称", { exact: true })).toHaveValue("未提交对象草稿");
  await page.getByRole("button", { name: "重新读取对象", exact: true }).click(); await expect(page.getByLabel("对象名称", { exact: true })).toHaveValue("说明");
  await page.getByLabel("选择本地图片", { exact: true }).setInputFiles({ name: "sample.png", mimeType: "image/png", buffer: readFileSync("examples/presentation/assets/sample.png") });
  await page.getByLabel("图片替代文字", { exact: true }).fill("授权图片"); await page.getByRole("button", { name: "插入本地图片", exact: true }).click(); await settled();
  await page.getByLabel("选择绘图对象", { exact: true }).selectOption({ label: "图片 2" }); await page.getByLabel("图片左侧裁剪比例", { exact: true }).fill("0.2");
  await page.getByRole("button", { name: "应用图片裁剪", exact: true }).click(); await settled();
  expect((await inspect() as { content: string }).content).toContain('data-office-crop-left="0.2"');
  const address = page.getByLabel("选择单元格或区域", { exact: true }); await address.fill("A1:B2"); await address.press("Enter"); await settled();
  await page.getByLabel("图表标题", { exact: true }).fill("当前数据"); await page.getByRole("button", { name: "从选区插入图表", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ charts: [{ kind: "column", categories: { expression: "='Sheet1'!A1:A2" }, series: [{ values: { expression: "='Sheet1'!B1:B2" } }] }], errors: [] });
  await page.getByLabel("选择绘图对象", { exact: true }).selectOption({ label: "图表 3" }); await page.getByRole("button", { name: "复制绘图对象", exact: true }).click(); await settled();
  expect((await inspect() as { drawings: unknown[] }).drawings).toHaveLength(4);
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect((await inspect() as { drawings: unknown[] }).drawings).toHaveLength(3);
  const cellInput = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await cellInput.fill("pending cell draft"); await expect(page.getByRole("button", { name: "删除绘图对象", exact: true })).toBeDisabled();
  await cellInput.press("Escape"); await settled();
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
  else await page.evaluate(() => window.workbookEditor.dispose());
});
