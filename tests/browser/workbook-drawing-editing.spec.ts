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

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} object paragraph input keeps rich runs, exact text, per-paragraph drafts and owning history`, async ({ page, browserName }) => {
  const session = createDocumentSession(scalarWorkbook([]));
  session.commit(session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: 0, operations: [{ op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "absolute", xPt: 30, yPt: 30, widthPt: 240, heightPt: 120 }, name: "富文本框",
    object: { kind: "text", html: '<p style="text-align:right">A&#13;<br><strong>禰󠄀😀</strong>Z</p><p><em>第二段</em></p>', text: { wrap: true } } },
    { op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "absolute", xPt: 300, yPt: 30, widthPt: 120, heightPt: 60 }, name: "另一对象", object: { kind: "text", html: "<p>其他</p>" } }] }));
  const object = session.snapshot().metadata.workbook.objects![0]!, first = session.getNode(object.nodeId)!.childNodeIds[0]!, second = session.getNode(object.nodeId)!.childNodeIds[1]!, strong = session.getNode(first)!.childNodeIds.find(id => session.getNode(id)!.tagName === "strong")!, before = session.snapshot().content!.html, files = session.serialize(); session.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { content: string; revision: number; history: { undoSteps: number }; drawings: { text?: { wrap?: boolean; autoFit?: string; fontScale?: number } }[]; errors: string[] };
  await page.getByText("图片、形状与图表", { exact: true }).click(); await page.getByLabel("选择绘图对象", { exact: true }).selectOption(object.nodeId);
  const paragraphs = page.getByLabel("对象文字段落", { exact: true }), text = page.getByLabel("所选对象段落文字", { exact: true }), save = page.getByRole("button", { name: "保存对象段落文字", exact: true });
  await expect(text).toHaveValue("A\n禰󠄀😀Z");
  if (browserName === "chromium") {
    const cdp = await page.context().newCDPSession(page); await text.focus(); await text.evaluate(element => (element as HTMLTextAreaElement).setSelectionRange(0, 0));
    await cdp.send("Input.imeSetComposition", { text: "预编辑", selectionStart: 3, selectionEnd: 3 }); await expect(save).toBeDisabled(); await expect(text).toBeEnabled(); expect((await inspect()).content).toBe(before);
    await cdp.send("Input.insertText", { text: "中文" }); await cdp.detach();
  } else { await text.dispatchEvent("compositionstart"); await expect(save).toBeDisabled(); await text.dispatchEvent("compositionend"); await text.fill("中文A\n禰󠄀😀Z"); }
  await save.click(); await settled(); expect(await inspect()).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  expect((await inspect()).content).toContain("中文A&#13;<br"); expect((await inspect()).content).toContain('<strong data-node-id="' + strong + '">');
  await page.getByLabel("对象段落西文字体", { exact: true }).fill("Liberation Sans"); await page.getByLabel("对象段落中文字体", { exact: true }).fill("Noto Sans CJK SC");
  await page.getByLabel("对象段落文字颜色", { exact: true }).fill("#a02030"); await page.getByLabel("对象段落字号（磅）", { exact: true }).fill("18");
  await text.focus(); await text.evaluate(element => { const field = element as HTMLTextAreaElement, from = field.value.indexOf("禰"); field.setSelectionRange(from, from + "禰󠄀😀".length); });
  await page.getByRole("button", { name: "应用对象选中文字字体", exact: true }).click(); await settled();
  expect((await inspect()).content).toContain("--office-text-source:direct"); expect((await inspect()).content).toContain("--office-font-east-asia"); expect((await inspect()).content).toContain('<strong data-node-id="' + strong + '">');
  await text.fill("未提交的第一段草稿");
  await page.getByLabel("选择绘图对象", { exact: true }).selectOption({ label: "另一对象" }); await expect(page.getByLabel("选择绘图对象", { exact: true })).toHaveValue(object.nodeId); await expect(text).toHaveValue("未提交的第一段草稿");
  await paragraphs.selectOption(second); await text.fill("已提交的第二段"); await paragraphs.selectOption(first); await expect(text).toHaveValue("未提交的第一段草稿");
  await paragraphs.selectOption(second); await save.click(); await settled(); await expect(save).toBeDisabled();
  await paragraphs.selectOption(first); await expect(text).toHaveValue("未提交的第一段草稿"); expect((await inspect()).content).toContain("已提交的第二段"); expect((await inspect()).content).not.toContain("未提交的第一段草稿");
  await page.getByRole("button", { name: "重新读取对象", exact: true }).click(); await expect(text).toHaveValue("中文A\n禰󠄀😀Z");
  await page.getByRole("button", { name: "新增空对象段落", exact: true }).click(); await settled(); await expect(text).toHaveValue(""); await expect(paragraphs.locator("option")).toHaveCount(3);
  await text.fill("新增段落禰󠄀😀"); await save.click(); await settled(); await page.getByRole("button", { name: "删除当前对象段落", exact: true }).click(); await settled(); await expect(paragraphs.locator("option")).toHaveCount(2);
  for (let index = 0; index < 6; index++) { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); }
  expect(await inspect()).toMatchObject({ content: before, history: { undoSteps: 0 }, errors: [] });
  await page.getByLabel("对象文字自动换行", { exact: true }).selectOption("false"); await page.getByLabel("对象文字自适应", { exact: true }).selectOption("normal"); await page.getByLabel("对象文字缩放比例", { exact: true }).fill("0.8");
  await page.getByRole("button", { name: "应用对象文字换行与自适应", exact: true }).click(); await settled(); expect((await inspect()).drawings[0]).toMatchObject({ text: { wrap: false, autoFit: "normal", fontScale: .8 } });
  await page.getByLabel("对象段落对齐", { exact: true }).selectOption("center"); await page.getByRole("button", { name: "应用对象段落排版", exact: true }).click(); await settled(); expect((await inspect()).content).toContain("text-align:center");
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect((await inspect()).content).toBe(before);
  expect((await inspect()).errors).toEqual([]);
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
  else await page.evaluate(() => window.workbookEditor.dispose());
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
  await page.getByLabel("编辑图表系列", { exact: true }).selectOption("0"); await expect(page.getByRole("button", { name: "应用图表数据", exact: true })).toBeDisabled();
  await page.getByLabel("编辑图表系列", { exact: true }).selectOption("1"); await expect(page.getByLabel("系列名称或引用", { exact: true })).toHaveValue("未提交的系列草稿");
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


for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} chart detail controls update crossings, current label formats and empty separators atomically`, async ({ page }) => {
  const session = createDocumentSession(scalarWorkbook([]));
  const commit = (operations: Parameters<typeof session.prepare>[0]["operations"]) => session.commit(session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: session.snapshot().manifest.revision, operations }));
  commit([{ op: "insertWorkbookChart", sheetId: "sheet-1", kind: "line", anchor: { kind: "absolute", xPt: 20, yPt: 20, widthPt: 300, heightPt: 180 }, data: {
    categories: { kind: "literal", valueType: "text", count: 2, points: [{ index: 0, value: "A" }, { index: 1, value: "B" }] }, series: [{ values: { kind: "literal", valueType: "number", count: 2, points: [{ index: 0, value: 2 }, { index: 1, value: 4 }] } }],
  } }]);
  const nodeId = session.snapshot().metadata.workbook.charts![0]!.nodeId;
  commit([{ op: "setWorkbookChart", nodeId, changes: { showLabels: true, labelContents: ["category", "value"] } }]);
  const initial = session.snapshot().metadata.workbook.charts![0]!, files = session.serialize(); session.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes, true), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { charts: WorkbookChart[]; history: { undoSteps: number }; errors: string[] };
  await page.getByText("图片、形状与图表", { exact: true }).click(); await page.getByLabel("选择绘图对象", { exact: true }).selectOption(nodeId);
  const target = page.getByLabel("图表样式部件", { exact: true }), apply = page.getByRole("button", { name: "应用当前图表部件样式", exact: true });
  await target.selectOption("categoryAxis"); await page.getByLabel("坐标轴交叉方式", { exact: true }).selectOption("at");
  const crossing = page.getByLabel("坐标轴在另一轴上的交叉值", { exact: true }); await expect(crossing).toBeEnabled(); await crossing.fill("3.25");
  await page.getByLabel("坐标轴主刻度线", { exact: true }).selectOption("out"); await page.getByLabel("分类轴标签间隔（0 自动）", { exact: true }).fill("2");
  await apply.click(); await settled(); expect((await inspect()).history.undoSteps).toBe(1);
  const paint = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedChart() : window.workbookEditor.displayedChart(), remote);
  expect((await paint()).categories).toEqual(["A"]);
  const axes = async () => (await paint()).axes;
  const first = await axes(); expect(first.y).toBeGreaterThan(first.top); expect(first.y).toBeLessThan(first.bottom);
  await target.selectOption("label"); await expect(page.getByLabel("坐标轴交叉方式", { exact: true })).toBeDisabled();
  await page.getByLabel("当前轴或标签数值格式来源", { exact: true }).selectOption("custom"); const format = page.getByLabel("当前轴或标签自定义数值格式", { exact: true }); await expect(format).toBeEnabled(); await format.fill("0.00");
  await format.dispatchEvent("compositionstart"); await expect(format).toBeEnabled(); await expect(apply).toBeDisabled(); await format.dispatchEvent("compositionend"); await expect(format).toHaveValue("0.00");
  await page.getByLabel("数据标签分隔符设置", { exact: true }).selectOption("custom"); await expect(page.getByLabel("数据标签分隔文字", { exact: true })).toBeEnabled();
  await page.getByLabel("数据标签位置", { exact: true }).selectOption("r");
  await page.getByLabel("数据标签引导线", { exact: true }).selectOption("true"); await page.getByLabel("数据标签引导线颜色", { exact: true }).fill("#207040"); await page.getByLabel("数据标签引导线宽度（磅）", { exact: true }).fill("2");
  await apply.click(); await settled();
  expect((await paint()).labels).toEqual(["A2.00", "B4.00"]);
  expect((await paint()).leaders).toEqual(Array(2).fill({ color: 'rgba(32,112,64,1)', width: 2 }));
  await target.selectOption("categoryAxis"); await page.getByLabel("坐标轴交叉方式", { exact: true }).selectOption("max"); await apply.click(); await settled();
  const last = await axes(); expect(last.y).toBeCloseTo(last.top, 4); // The explicit old crossesAt must be cleared.
  for (let count = 0; count < 3; count++) { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); }
  expect(await inspect()).toMatchObject({ charts: [initial], history: { undoSteps: 0 }, errors: [] });
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, reopenedEqual: true }); else await page.evaluate(() => window.workbookEditor.dispose());
});

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} chart style controls retain part drafts, marker overrides and owning history`, async ({ page }) => {
  const session = createDocumentSession(scalarWorkbook([]));
  session.commit(session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: 0, operations: [{ op: "insertWorkbookChart", sheetId: "sheet-1", kind: "line", anchor: { kind: "absolute", xPt: 20, yPt: 20, widthPt: 300, heightPt: 180 }, data: { series: [{ name: { kind: "text", text: "样式系列" }, values: { kind: "literal", valueType: "number", count: 2, points: [{ index: 0, value: 2 }, { index: 1, value: 4 }] } }] } }] }));
  const chart = session.snapshot().metadata.workbook.charts![0]!, initial = chart, files = session.serialize(); session.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { revision: number; charts: WorkbookChart[]; errors: string[]; history: { undoSteps: number } };
  await page.getByText("图片、形状与图表", { exact: true }).click(); await page.getByLabel("选择绘图对象", { exact: true }).selectOption(chart.nodeId);
  const target = page.getByLabel("图表样式部件", { exact: true }), fill = page.getByLabel("图表部件填充颜色", { exact: true }), apply = page.getByRole("button", { name: "应用当前图表部件样式", exact: true });
  await target.selectOption("point"); await fill.fill("#c02030"); await page.getByLabel("图表数据标记", { exact: true }).selectOption("diamond"); await page.getByLabel("图表标记大小（2 至 72 磅，0 恢复默认）", { exact: true }).fill("12");
  await expect(page.getByLabel("图表标记大小（2 至 72 磅，0 恢复默认）", { exact: true })).toHaveValue("12"); await fill.focus();
  await fill.dispatchEvent("compositionstart"); await expect(fill).toBeEnabled(); await expect(apply).toBeDisabled(); await fill.dispatchEvent("compositionend"); await expect(page.getByLabel("图表标记大小（2 至 72 磅，0 恢复默认）", { exact: true })).toHaveValue("12");
  await apply.click(); await settled(); const first = (await inspect()).charts[0]!.pointStyles![0]!; expect(first).toMatchObject({ index: 0, marker: "diamond", markerSize: 12 });
  await target.selectOption("series"); await fill.fill("#2050c0"); await page.getByLabel("图表数据标记", { exact: true }).selectOption("auto"); await apply.click(); await settled();
  expect((await inspect()).charts[0]!.pointStyles![0]).toEqual(first); expect((await inspect()).charts[0]!.seriesStyles![0]).toMatchObject({ marker: "auto" });
  await fill.fill("#111111"); await target.selectOption("point"); await page.getByLabel("图表样式原始数据位置（从 1 开始）", { exact: true }).fill("2"); await fill.fill("#20a040"); await target.selectOption("series"); await expect(fill).toHaveValue("#111111");
  await apply.click(); await settled(); await expect(apply).toBeDisabled(); await target.selectOption("point"); await expect(fill).toHaveValue("#20a040"); await expect(apply).toBeDisabled();
  await page.getByRole("button", { name: "重新读取对象", exact: true }).click(); await expect(fill).toHaveValue("");
  for (let i = 0; i < 3; i++) { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); }
  expect(await inspect()).toMatchObject({ charts: [initial], history: { undoSteps: 0 }, errors: [] });
  await target.selectOption("point"); await expect(page.getByLabel("图表部件文字装饰", { exact: true })).toBeDisabled();
  await target.selectOption("pointLabel"); await expect(page.getByLabel("图表部件文字装饰", { exact: true })).toBeEnabled(); await page.getByLabel("图表部件文字装饰", { exact: true }).selectOption("none");
  await apply.click(); await settled(); expect((await inspect()).charts[0]!.pointStyles![0]!.labelCssClass).toBeTruthy();
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect((await inspect()).charts).toEqual([initial]);
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, reopenedEqual: true }); else await page.evaluate(() => window.workbookEditor.dispose());
});
