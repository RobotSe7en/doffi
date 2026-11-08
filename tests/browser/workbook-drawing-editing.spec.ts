import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { writeDocumentZip } from "@doffi/office";
import { importXlsx } from "@doffi/office";
import { createDocumentSession, openDocument, WorkbookChartSchema } from "@doffi/core";
import type { WorkbookChart } from "@doffi/core";
import type { WorkbookRenderReport } from "@doffi/render";
import { scalarWorkbook } from "../../packages/core/tests/workbook-scalar-fixture.js";
import { encodeTestPackage } from "./package-transport.js";
import { openSdkHost } from "./sdk-host.js";
import type * as directFixture from "./fixtures/workbook-editor.js";
import type * as remoteFixture from "./fixtures/workbook-remote-editor.js";
declare global { interface Window { workbookEditor: typeof directFixture; workbookRemoteEditor: typeof remoteFixture } }
declare global { interface Window { workbookImageReadGate: { readonly started: boolean; release(): void; restore(): void } } }
const bundles = new Map<string, string>();
test.beforeEach(async ({ page }) => { await openSdkHost(page); });
test.beforeAll(async () => {
  for (const name of ["workbook-editor", "workbook-remote-editor", "workbook-session-worker"]) {
    const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/" + name + ".ts", import.meta.url)), formats: ["iife"], name: "WorkbookDrawing" } } });
    const output = Array.isArray(result) ? result.flatMap(part => part.output) : "output" in result ? result.output : [];
    bundles.set(name, output.filter(part => part.type === "chunk").map(part => part.code).join("\n"));
  }
});

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} chart source controls retain sparse literals and native caches while series, axes and titles share guarded history`, async ({ page, browserName }) => {
  const session = createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/07-three-charts.xlsx"), { documentId: "chart-controls" }), { kind: "workbook" }));
  const original = WorkbookChartSchema.parse(session.snapshot().metadata.workbook.charts![0]!), sheet = session.snapshot().metadata.workbook.sheets[0]!;
  const categories = original.categories!, series = original.series;
  session.commit(session.prepare({ protocolVersion: 1, documentId: "chart-controls", baseRevision: 0, operations: [{ op: "setWorkbookChartData", nodeId: original.nodeId,
    data: { categories, series: series.map((item, index) => ({ seriesId: item.seriesId, ...(item.name ? { name: item.name } : {}), values: index ? item.values : {
      kind: "literal", valueType: "number", count: 205, numberFormat: "0.00", points: [...Array.from({ length: 101 }, (_, index) => ({ index, value: index + 1, ...(index === 0 ? { numberFormat: "0%" } : {}) })), { index: 204, value: null }],
    } })) } }] }));
  const files = session.serialize(), before = session.snapshot().metadata.workbook.charts!; session.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { charts: WorkbookChart[]; content: string; revision: number; history: { undoSteps: number }; errors: string[] };
  await page.getByText("图片、形状与图表", { exact: true }).click(); await page.getByLabel("选择绘图对象", { exact: true }).selectOption(original.nodeId);
  await page.getByRole("button", { name: "应用图表数据", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ revision: 1, history: { undoSteps: 0 }, charts: before, errors: [] });
  await page.getByRole("button", { name: "系列数值下一页", exact: true }).click();
  await page.getByLabel("系列数值数据点 101内容", { exact: true }).fill("-44.5");
  await page.getByRole("button", { name: "系列数值上一页", exact: true }).click();
  await expect(page.getByLabel("系列数值数据点 1数字格式", { exact: true })).toHaveValue("0%");
  await page.getByRole("button", { name: "系列数值下一页", exact: true }).click(); await expect(page.getByLabel("系列数值数据点 101内容", { exact: true })).toHaveValue("-44.5");
  await page.getByLabel("系列名称来源", { exact: true }).selectOption("text"); await page.getByLabel("系列名称或引用", { exact: true }).fill("常量禰󠄀😀");
  await page.getByRole("button", { name: "新增图表系列", exact: true }).click();
  await page.getByRole("button", { name: "系列数值新增数据点", exact: true }).click(); await page.getByLabel("系列数值数据点 1内容", { exact: true }).fill("12");
  await page.getByRole("button", { name: "图表系列上移", exact: true }).click();
  await page.getByRole("button", { name: "应用图表数据", exact: true }).click(); await settled();
  let result = await inspect(), updated = result.charts[0]!;
  expect(result).toMatchObject({ revision: 2, history: { undoSteps: 1 }, errors: [] }); expect(updated.series).toHaveLength(3);
  expect(updated.series[0]).toMatchObject({ seriesId: series[0]!.seriesId, name: { kind: "text", text: "常量禰󠄀😀" }, values: { kind: "literal", count: 205, numberFormat: "0.00" } });
  const points = "kind" in updated.series[0]!.values ? updated.series[0]!.values.points : []; expect(points[100]).toEqual({ index: 100, value: -44.5 }); expect(points.at(-1)).toEqual({ index: 204, value: null });
  expect(updated.series[2]).toEqual(before[0]!.series[1]); expect(updated.categories).toEqual(categories); expect(result.charts.slice(1)).toEqual(before.slice(1));
  await page.getByLabel("数值轴最小值", { exact: true }).fill("-60"); await page.getByLabel("数值轴最大值", { exact: true }).fill("240"); await page.getByLabel("数值轴主刻度", { exact: true }).fill("30");
  await page.getByLabel("分类轴逆序", { exact: true }).check(); await page.getByLabel("数据标签分类", { exact: true }).check();
  await page.getByLabel("系列名称来源", { exact: true }).selectOption("text"); await page.getByLabel("系列名称或引用", { exact: true }).fill("未提交的系列草稿");
  await page.getByRole("button", { name: "应用坐标轴与标签", exact: true }).click(); await settled();
  expect((await inspect()).charts[0]).toMatchObject({ minimum: -60, maximum: 240, majorUnit: 30, reverseCategories: true, showLabels: true, labelContents: expect.arrayContaining(["category"]) });
  await expect(page.getByLabel("系列名称或引用", { exact: true })).toHaveValue("未提交的系列草稿"); await expect(page.getByRole("button", { name: "应用图表数据", exact: true })).toBeDisabled();
  expect((await inspect()).charts[0]!.series[1]!.name).toEqual({ kind: "text", text: "新系列" });
  await page.getByRole("button", { name: "重新读取对象", exact: true }).click();
  const reference = "='" + sheet.name.replace(/'/g, "''") + "'!$A$1";
  await page.getByLabel("数值轴标题来源", { exact: true }).selectOption("reference"); await page.getByLabel("数值轴标题文字或引用", { exact: true }).fill(reference);
  await page.getByRole("button", { name: "应用数值轴标题", exact: true }).click(); await settled();
  expect((await inspect()).charts[0]!.valueAxisTitle).toMatchObject({ reference: { expression: reference, valueType: "text" } });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled();
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled();
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect((await inspect()).charts).toEqual(before);
  await page.getByLabel("主标题来源", { exact: true }).selectOption("text"); const title = page.getByLabel("主标题文字或引用", { exact: true }); await title.fill("标题");
  if (browserName === "chromium") {
    const cdp = await page.context().newCDPSession(page); await title.focus(); await cdp.send("Input.imeSetComposition", { text: "预编辑", selectionStart: 3, selectionEnd: 3 });
    await expect(page.getByRole("button", { name: "应用主标题", exact: true })).toBeDisabled(); await expect(title).toBeEnabled();
    await cdp.send("Input.insertText", { text: "禰󠄀😀" }); await cdp.detach();
  } else { await title.dispatchEvent("compositionstart"); await expect(page.getByRole("button", { name: "应用主标题", exact: true })).toBeDisabled(); await title.dispatchEvent("compositionend"); }
  await title.fill("保留的标题草稿禰󠄀😀");
  await page.evaluate(remote => remote ? window.workbookRemoteEditor.external("external") : window.workbookEditor.external("external"), remote); await settled();
  await expect(page.getByRole("button", { name: "应用主标题", exact: true })).toBeDisabled(); await expect(title).toHaveValue("保留的标题草稿禰󠄀😀");
  expect((await inspect()).charts).toEqual(before);
  await page.getByRole("button", { name: "重新读取对象", exact: true }).click(); await expect(title).not.toHaveValue("保留的标题草稿禰󠄀😀");
  expect((await inspect()).errors).toEqual([]);
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
  else await page.evaluate(() => window.workbookEditor.dispose());
});

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} delayed local image input keeps its captured source revision across file reading`, async ({ page }) => {
  const files = scalarWorkbook([]).serialize();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(files));
  await page.getByText("图片、形状与图表", { exact: true }).click();
  await page.getByLabel("选择本地图片", { exact: true }).setInputFiles({ name: "sample.png", mimeType: "image/png", buffer: readFileSync("examples/presentation/assets/sample.png") });
  await page.evaluate(() => {
    const original = File.prototype.arrayBuffer; let release!: () => void, started = false;
    const gate = new Promise<void>(resolve => { release = resolve; });
    File.prototype.arrayBuffer = async function() { started = true; const bytes = await original.call(this); await gate; return bytes; };
    window.workbookImageReadGate = { get started() { return started; }, release, restore() { File.prototype.arrayBuffer = original; } };
  });
  try {
    await page.getByRole("button", { name: "插入本地图片", exact: true }).click(); await expect.poll(() => page.evaluate(() => window.workbookImageReadGate.started)).toBe(true);
    await page.evaluate(remote => remote ? window.workbookRemoteEditor.external("external", false) : window.workbookEditor.external("external"), remote);
    await page.evaluate(() => window.workbookImageReadGate.release());
    await page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
    expect(await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote)).toMatchObject({ revision: 1, history: { undoSteps: 1 }, drawings: [], errors: ["RevisionConflict"] });
    if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
    else await page.evaluate(() => window.workbookEditor.dispose());
  } finally { await page.evaluate(() => { window.workbookImageReadGate.release(); window.workbookImageReadGate.restore(); }); }
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
