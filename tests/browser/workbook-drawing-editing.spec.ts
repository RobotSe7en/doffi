import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { writeDocumentZip } from "@doffi/office";
import { importXlsx } from "@doffi/office";
import { createDocumentSession, openDocument, WorkbookChartSchema } from "@doffi/core";
import type { WorkbookChart, WorkbookNoteBox } from "@doffi/core";
import type { WorkbookRenderReport } from "@doffi/render";
import { scalarWorkbook } from "../../packages/core/tests/workbook-scalar-fixture.js";
import { encodeTestPackage } from "./package-transport.js";
import { openSdkHost } from "./sdk-host.js";
import type * as directFixture from "./fixtures/workbook-editor.js";
import type * as remoteFixture from "./fixtures/workbook-remote-editor.js";
declare global { interface Window { workbookEditor: typeof directFixture; workbookRemoteEditor: typeof remoteFixture } }
declare global { interface Window { workbookImageReadGate: { readonly started: boolean; release(): void; restore(): void } } }
const bundles = new Map<string, string>();

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} object text clears ancestor underline while preserving strike, fonts and undo`, async ({ page }) => {
  const s = createDocumentSession(scalarWorkbook([]));
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0, operations: [{ op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "absolute", xPt: 20, yPt: 20, widthPt: 240, heightPt: 100 }, object: { kind: "text", html: '<p style="font-size:18pt;text-decoration:underline line-through double red;--office-font-ascii:&quot;Liberation Sans&quot;;--office-font-east-asia:&quot;Noto Sans CJK SC&quot;"><strong>A禰󠄀😀Z</strong></p>' } }] }));
  const object = s.snapshot().metadata.workbook.objects![0]!, paragraph = s.getNode(object.nodeId)!.childNodeIds[0]!, strong = s.getNode(paragraph)!.childNodeIds[0]!, before = s.snapshot().content!.html, source = s.serialize(); s.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes, true), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(source)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(source));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { content: string; history: { undoSteps: number }; errors: string[] };
  const paint = (id: string) => page.evaluate(({ remote, id }) => remote ? window.workbookRemoteEditor.displayedDrawingText(id) : window.workbookEditor.displayedDrawingText(id), { remote, id });
  await page.getByText("图片、形状与图表", { exact: true }).click(); await page.getByLabel("选择绘图对象", { exact: true }).selectOption(object.nodeId);
  const text = page.getByLabel("所选对象段落文字", { exact: true }); await expect(text).toHaveValue("A禰󠄀😀Z"); await page.getByLabel("对象段落下划线", { exact: true }).selectOption("none");
  await text.focus(); await text.evaluate(element => (element as HTMLTextAreaElement).setSelectionRange(1, 6)); await page.getByRole("button", { name: "应用对象选中文字字体", exact: true }).click(); await settled();
  const current = await inspect(); expect(current.history.undoSteps).toBe(1); expect(current.errors).toEqual([]);
  const wrapper = await page.evaluate(({ content, strong }) => {
    const root = new DOMParser().parseFromString(content, "text/html"), parent = [...root.querySelectorAll("[data-node-id]")].find(element => element.getAttribute("data-node-id") === strong)!;
    return parent.querySelector('[style*="--office-font-decoration-reset"]')!.getAttribute("data-node-id")!;
  }, { content: current.content, strong });
  await expect.poll(() => paint(wrapper)).toEqual([expect.objectContaining({ text: "禰󠄀😀", decoration: "line-through", decorationStyle: "double", weight: "700" })]);
  await expect.poll(() => paint(strong)).toEqual([expect.objectContaining({ text: "A禰󠄀😀Z", decoration: "underline line-through" })]);
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect((await inspect()).content).toBe(before); expect((await inspect()).history.undoSteps).toBe(0);
  await page.getByRole("button", { name: "重做", exact: true }).click(); await settled(); expect((await inspect()).content).toBe(current.content);
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, reopenedEqual: true }); else await page.evaluate(() => window.workbookEditor.dispose());
});
for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} partial object formatting keeps unrequested fonts against existing important span rules`, async ({ page }) => {
  const s = createDocumentSession(scalarWorkbook([]));
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0, operations: [{ op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "absolute", xPt: 20, yPt: 20, widthPt: 240, heightPt: 100 }, object: { kind: "text", html: '<p style="font-family:&quot;Liberation Sans&quot;;font-size:18pt;--office-font-ascii:&quot;Liberation Sans&quot;"><strong style="background-color:#ddeeff">GlyphAB</strong></p>' } }] }));
  const object = s.snapshot().metadata.workbook.objects![0]!, paragraph = s.getNode(object.nodeId)!.childNodeIds[0]!, strong = s.getNode(paragraph)!.childNodeIds[0]!, source = new Map(s.serialize()); s.dispose();
  source.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(source.get("styles.css")!) + 'span{font-family:"Liberation Serif"!important;font-size:40pt!important;font-weight:normal!important;font-style:italic!important;vertical-align:super!important;background-color:yellow!important;--office-font-ascii:"Liberation Serif"!important}'));
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes, true), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(source)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(source));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { content: string; history: { undoSteps: number }; errors: string[] };
  const paint = (id: string) => page.evaluate(({ remote, id }) => remote ? window.workbookRemoteEditor.displayedDrawingText(id, true) : window.workbookEditor.displayedDrawingText(id, true), { remote, id });
  const before = await inspect(), original = await paint(strong); expect(original).toMatchObject([{ size: 24, weight: "700", italic: "normal" }]);
  await page.getByText("图片、形状与图表", { exact: true }).click(); await page.getByLabel("选择绘图对象", { exact: true }).selectOption(object.nodeId);
  await page.getByLabel("对象段落文字颜色", { exact: true }).fill("#a02030"); const text = page.getByLabel("所选对象段落文字", { exact: true }); await text.focus(); await text.evaluate(element => (element as HTMLTextAreaElement).setSelectionRange(1, 4));
  await page.getByRole("button", { name: "应用对象选中文字字体", exact: true }).click(); await settled(); const current = await inspect();
  const wrapper = await page.evaluate(({ content, strong }) => { const doc = new DOMParser().parseFromString(content, "text/html"); return [...doc.querySelectorAll("[data-node-id]")].find(element => element.getAttribute("data-node-id") === strong)!.querySelector("span")!.getAttribute("data-node-id")!; }, { content: current.content, strong });
  await expect.poll(() => paint(wrapper)).toMatchObject([{ text: "lyp", color: "rgb(160, 32, 48)", size: 24, weight: "700", italic: "normal", font: original[0]!.font }]);
  const selected = (await paint(wrapper))[0]!; expect(selected.topPt).toBeCloseTo(original[0]!.topPt!, 3); expect(selected.background).toBe("rgba(0, 0, 0, 0)"); expect((await paint(strong))[0]!.background).toBe(original[0]!.background);
  expect(current.history.undoSteps).toBe(1); expect(current.errors).toEqual([]);
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect((await inspect()).content).toBe(before.content);
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, reopenedEqual: true }); else await page.evaluate(() => window.workbookEditor.dispose());
});
test.beforeEach(async ({ page }) => { await openSdkHost(page); });
test.beforeAll(async () => {
  for (const name of ["workbook-editor", "workbook-remote-editor", "workbook-session-worker"]) {
    const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/" + name + ".ts", import.meta.url)), formats: ["iife"], name: "WorkbookDrawing" } } });
    const output = Array.isArray(result) ? result.flatMap(part => part.output) : "output" in result ? result.output : [];
    bundles.set(name, output.filter(part => part.type === "chunk").map(part => part.code).join("\n"));
  }
});

test("measured fixed and mixed objects and note boxes follow actual font and wrapped-row layout without geometry drift", async ({ page }) => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "A" } }, { row: 0, column: 1, value: { type: "text", value: "B" } },
    { row: 3, column: 0, value: { type: "text", value: "Wrapped text with several words on multiple lines" } }]));
  const at = (row: number, column: number) => ({ row, column, offsetXPt: 2, offsetYPt: 3 });
  const anchor = (moveWithCells: boolean, sizeWithCells: boolean) => ({ kind: "two-cell" as const, from: at(2, 1), to: at(6, 3), moveWithCells, sizeWithCells });
  const box: WorkbookNoteBox = { anchor: anchor(false, false), visible: true, printable: true, locked: true, lockText: false, defaultSize: false,
    disabled: false, autoFill: false, autoLine: false, autoScale: false, justifyLastLine: false };
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0, operations: [
    ...[false, true].flatMap(move => [false, true].map(size => ({ op: "insertWorkbookObject" as const, sheetId: "sheet-1", anchor: anchor(move, size), name: `${move}/${size}`, object: { kind: "geometry" as const, preset: "rect" as const, html: "<p>形</p>" } }))),
    { op: "addWorkbookComment", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 1, column: 1 }, to: { row: 1, column: 1 } }] }, format: "classic", html: "<p>备注</p>", box },
  ] }));
  const initial = s.snapshot(), ids = [...initial.metadata.workbook.objects!.map(object => object.nodeId), initial.annotations[0]!.bodyNodeId!], files = s.serialize(); s.dispose();
  await page.addScriptTag({ content: bundles.get("workbook-editor")! }); await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(files));
  const measure = () => page.evaluate(() => window.workbookEditor.displayedDrawingBounds());
  const bounds = (value: Awaited<ReturnType<typeof measure>>, index: number) => value.objects.find(object => object.nodeId === ids[index])!;
  const before = await measure(); expect(before.objects).toHaveLength(5);
  await page.evaluate(() => window.workbookEditor.measuredApply([{ op: "setWorkbookStyleClass", cssClass: "cell-default", properties: { "font-size": "20pt", "white-space": "pre-wrap" } }]));
  const after = await measure(); expect(after.objects).toHaveLength(5); expect(after.columnWidthPt!).toBeGreaterThan(before.columnWidthPt!);
  for (const index of [0, 4]) for (const field of ["xPt", "yPt", "widthPt", "heightPt"] as const) expect(bounds(after, index)[field]).toBeCloseTo(bounds(before, index)[field], 1);
  for (const field of ["xPt", "yPt"] as const) expect(bounds(after, 1)[field]).toBeCloseTo(bounds(before, 1)[field], 1);
  for (const field of ["widthPt", "heightPt"] as const) expect(bounds(after, 2)[field]).toBeCloseTo(bounds(before, 2)[field], 1);
  for (const index of [1, 3]) for (const field of ["widthPt", "heightPt"] as const) expect(bounds(after, index)[field]).toBeGreaterThan(bounds(before, index)[field]);
  for (const index of [2, 3]) for (const field of ["xPt", "yPt"] as const) expect(bounds(after, index)[field]).toBeGreaterThan(bounds(before, index)[field]);
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  expect(await page.locator("iframe").count()).toBe(1);
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ drawings: initial.metadata.workbook.objects, comments: initial.metadata.workbook.comments, history: { undoSteps: 0 }, errors: [] });
  await page.evaluate(() => window.workbookEditor.dispose());
});

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} axis edits move fixed one-cell drawings and resize both-marker drawings and notes in the actual canvas`, async ({ page }) => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "A" } }, { row: 0, column: 1, value: { type: "text", value: "B" } }], metadata => {
    metadata.workbook.sheets[0]!.defaultRowHeightPt = 20; metadata.workbook.sheets[0]!.defaultRowHeightCustom = true;
  }));
  const at = (row: number, column: number) => ({ row, column, offsetXPt: 2, offsetYPt: 3 });
  const note: WorkbookNoteBox = { anchor: { kind: "two-cell", from: at(3, 1), to: at(5, 3), moveWithCells: true, sizeWithCells: true }, visible: true, printable: true,
    locked: true, lockText: false, defaultSize: false, disabled: false, autoFill: false, autoLine: false, autoScale: false, justifyLastLine: false };
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations: [
    { op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "one-cell", from: at(2, 1), widthPt: 80, heightPt: 30 }, name: "Fixed extent", object: { kind: "text", html: "<p>固定</p>" } },
    { op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "two-cell", from: at(2, 2), to: at(6, 4), moveWithCells: true, sizeWithCells: true }, name: "Resizing", object: { kind: "geometry", preset: "rect", html: "<p>缩放</p>" } },
    { op: "addWorkbookComment", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 1, column: 1 }, to: { row: 1, column: 1 } }] }, html: "<p>备注</p>", format: "classic", box: note },
  ] }));
  const initial = s.snapshot(), ids = [...initial.metadata.workbook.objects!.map(object => object.nodeId), initial.annotations[0]!.bodyNodeId!], files = s.serialize(); s.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes, true), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const measure = async () => (await page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedDrawingBounds() : window.workbookEditor.displayedDrawingBounds(), remote));
  const bounds = (value: Awaited<ReturnType<typeof measure>>, index: number) => { const result = value.objects.find(object => object.nodeId === ids[index]); if (!result) throw Error("The expected source object is not visible."); return result; };
  const address = page.getByLabel("选择单元格或区域", { exact: true }), select = async (range: string) => { await address.fill(range); await address.press("Enter"); await settled(); };
  const before = await measure(); expect(before.objects).toHaveLength(3); expect(before.columnWidthPt).toBeGreaterThan(0);
  await select("A2:A3"); await page.getByRole("button", { name: "插入选区行", exact: true }).click(); await settled(); const moved = await measure();
  for (let index = 0; index < 3; index++) { expect(bounds(moved, index).yPt - bounds(before, index).yPt).toBeCloseTo(40, 1); expect(bounds(moved, index).heightPt).toBeCloseTo(bounds(before, index).heightPt, 1); }
  await select("A7"); await page.getByRole("button", { name: "插入选区行", exact: true }).click(); await settled(); const resized = await measure();
  expect(bounds(resized, 0).heightPt).toBeCloseTo(30, 1); expect(bounds(resized, 0).yPt).toBeCloseTo(bounds(moved, 0).yPt, 1);
  for (const index of [1, 2]) { expect(bounds(resized, index).heightPt - bounds(moved, index).heightPt).toBeCloseTo(20, 1); expect(bounds(resized, index).yPt).toBeCloseTo(bounds(moved, index).yPt, 1); }
  await select("B1:C1"); await page.getByRole("button", { name: "插入选区列", exact: true }).click(); await settled(); const shifted = await measure();
  for (let index = 0; index < 3; index++) { expect(bounds(shifted, index).xPt - bounds(resized, index).xPt).toBeCloseTo(2 * before.columnWidthPt!, 1); expect(bounds(shifted, index).widthPt).toBeCloseTo(bounds(resized, index).widthPt, 1); }
  for (let index = 0; index < 3; index++) { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); }
  const restored = await measure(); for (let index = 0; index < 3; index++) for (const field of ["xPt", "yPt", "widthPt", "heightPt"] as const) expect(bounds(restored, index)[field]).toBeCloseTo(bounds(before, index)[field], 1);
  expect(await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote)).toMatchObject({ history: { undoSteps: 0 }, drawings: initial.metadata.workbook.objects, comments: initial.metadata.workbook.comments, errors: [] });
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, reopenedEqual: true, hostShadowsRestored: true }); else await page.evaluate(() => window.workbookEditor.dispose());
});

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} new drawing composition binds its starting revision before the first input event`, async ({ page }) => {
  const files = scalarWorkbook([]).serialize(); await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  await page.getByText("图片、形状与图表", { exact: true }).click(); const name = page.getByLabel("对象名称", { exact: true }), insert = page.getByRole("button", { name: "插入文本框", exact: true });
  await name.focus(); await name.dispatchEvent("compositionstart"); expect(await inspect()).toMatchObject({ pending: true });
  await page.evaluate(async remote => { if (remote) await window.workbookRemoteEditor.external("external"); else { window.workbookEditor.external("external"); await window.workbookEditor.settled(); } }, remote);
  await name.evaluate(element => { const field = element as HTMLInputElement; field.value = "中文草稿"; field.dispatchEvent(new InputEvent("input", { bubbles: true, isComposing: true })); });
  await name.dispatchEvent("compositionend"); await expect(name).toHaveValue("中文草稿"); await expect(insert).toBeDisabled(); expect(await inspect()).toMatchObject({ pending: true, history: { undoSteps: 1 }, drawings: [] });
  await page.getByRole("button", { name: "放弃绘图表单草稿", exact: true }).click(); await expect(name).toHaveValue(""); expect(await inspect()).toMatchObject({ pending: false });
  await name.fill("新对象"); await insert.click(); await settled(); expect(await inspect()).toMatchObject({ pending: false, history: { undoSteps: 2 }, drawings: [expect.objectContaining({ name: "新对象", kind: "text" })], errors: [] });
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, reopenedEqual: true, endpointDisposed: true }); else await page.evaluate(() => window.workbookEditor.dispose());
});

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} whole worksheet copies display their own current chart references and preserve literal vectors`, async ({ page }) => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "A" } }, { row: 1, column: 0, value: { type: "text", value: "B" } },
    { row: 0, column: 1, value: { type: "number", value: 2 } }, { row: 1, column: 1, value: { type: "number", value: 4 } }]));
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0, operations: [{ op: "insertWorkbookChart", sheetId: "sheet-1", kind: "line", anchor: { kind: "absolute", xPt: 15, yPt: 20, widthPt: 300, heightPt: 180 },
    data: { categories: { expression: "=A1:A2", valueType: "text" }, series: [{ values: { expression: "=B1:B2", valueType: "number" } },
      { values: { kind: "literal", valueType: "number", count: 2, points: [{ index: 0, value: 7 }, { index: 1, value: 8 }] } }] } }] }));
  const nodeId = s.snapshot().metadata.workbook.objects![0]!.nodeId;
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations: [{ op: "setWorkbookChart", nodeId, changes: { showLabels: true } }] }));
  const initial = s.snapshot().metadata.workbook.charts!, files = s.serialize(); s.dispose();
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes, true), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const paint = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedChart() : window.workbookEditor.displayedChart(), remote);
  expect((await paint()).labels).toEqual(["2", "4", "7", "8"]);
  await page.getByText("工作表管理", { exact: true }).click(); await page.getByLabel("工作表操作", { exact: true }).selectOption("copy"); await page.getByLabel("工作表名称", { exact: true }).fill("Copied");
  await page.getByRole("button", { name: "执行工作表操作", exact: true }).click(); await settled();
  expect(await inspect()).toMatchObject({ charts: expect.arrayContaining([...initial]), history: { undoSteps: 1 }, errors: [] });
  const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }), input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
  for (const [at, value] of [["B1", "9"], ["A1", "新类"]]) { await address.fill(at!); await address.press("Enter"); await settled(); await input.fill(value!); await input.press("Enter"); await settled(); }
  expect(await paint()).toMatchObject({ categories: ["新类", "B"], labels: ["9", "4", "7", "8"] });
  await page.getByRole("tab", { name: "Sheet1", exact: true }).click(); await settled(); expect(await paint()).toMatchObject({ categories: ["A", "B"], labels: ["2", "4", "7", "8"] });
  for (let index = 0; index < 3; index++) { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); }
  expect(await inspect()).toMatchObject({ charts: initial, history: { undoSteps: 0 }, errors: [] });
  await expect(page.getByRole("tab", { name: "Copied", exact: true })).toHaveCount(0);
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, reopenedEqual: true, hostShadowsRestored: true }); else await page.evaluate(() => window.workbookEditor.dispose());
});

for (const remote of [false, true]) test(`${remote ? "Worker" : "direct"} copied object text retains node-ID styles and the original cascade in the actual view`, async ({ page }) => {
  const source = new Map(scalarWorkbook([]).serialize()); source.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(source.get("styles.css")!) + ".drawing-scope{}.drawing-accent{}"));
  const s = createDocumentSession(openDocument(source, { kind: "workbook" }));
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0, operations: [{ op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "absolute", xPt: 20, yPt: 20, widthPt: 180, heightPt: 80 }, name: "复制外观", object: { kind: "text", html: '<p class="drawing-scope"><strong class="drawing-accent">Glyph</strong></p>' } }] }));
  const object = s.snapshot().metadata.workbook.objects![0]!, paragraph = s.getNode(object.nodeId)!.childNodeIds[0]!, run = s.getNode(paragraph)!.childNodeIds[0]!, files = new Map(s.serialize()); s.dispose();
  files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + '[data-node-id="' + paragraph + '"] .drawing-accent{color:red}[data-node-id="' + paragraph + '"] [data-node-id="' + run + '"]{font-size:20pt}.drawing-scope .drawing-accent{color:#334455}'));
  await page.addScriptTag({ content: bundles.get(remote ? "workbook-remote-editor" : "workbook-editor")! });
  if (remote) await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes, true), { code: bundles.get("workbook-session-worker")!, bytes: Buffer.from(writeDocumentZip(files)).toString("base64") });
  else await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(files));
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const styles = (id: string) => page.evaluate(({ remote, id }) => remote ? window.workbookRemoteEditor.displayedDrawingText(id) : window.workbookEditor.displayedDrawingText(id), { remote, id });
  const before = await styles(run); expect(before).toHaveLength(1); expect(before[0]).toMatchObject({ text: "Glyph", color: "rgb(51, 68, 85)" }); expect(before[0]!.size).toBeCloseTo(20 / .75, 1);
  await page.getByText("图片、形状与图表", { exact: true }).click(); await page.getByLabel("选择绘图对象", { exact: true }).selectOption(object.nodeId);
  await page.getByRole("button", { name: "复制绘图对象", exact: true }).click(); await settled();
  const result = await page.evaluate(({ remote, original }) => {
    return Promise.resolve(remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect()).then(value => {
      const data = value as { drawings: { nodeId: string }[]; content: string; history: { undoSteps: number }; errors: string[] };
      const copy = data.drawings.find(item => item.nodeId !== original)!;
      const template = document.createElement("template"); template.innerHTML = data.content;
      const root = [...template.content.querySelectorAll("[data-node-id]")].find(element => element.getAttribute("data-node-id") === copy.nodeId)!;
      return { run: root.querySelector("strong")!.getAttribute("data-node-id")!, history: data.history, errors: data.errors };
    });
  }, { remote, original: object.nodeId });
  expect(result).toMatchObject({ history: { undoSteps: 1 }, errors: [] }); expect(result.run).not.toBe(run); expect(await styles(result.run)).toEqual(before); expect(await styles(run)).toEqual(before);
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); expect(await styles(result.run)).toEqual([]); expect(await styles(run)).toEqual(before);
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true, hostShadowsRestored: true });
  else await page.evaluate(() => window.workbookEditor.dispose());
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
  await page.getByLabel("图表部件填充颜色", { exact: true }).fill("#f0e0c0"); await page.getByLabel("图表部件线条颜色", { exact: true }).fill("#704090"); await page.getByLabel("图表部件线宽（磅）", { exact: true }).fill("1.5");
  await apply.click(); await settled();
  expect((await paint()).labels).toEqual(["A2.00", "B4.00"]);
  expect((await paint()).leaders).toEqual(Array(2).fill({ color: 'rgba(32,112,64,1)', width: 2 }));
  expect((await paint()).frames).toEqual(Array(2).fill({ fill: 'rgba(240,224,192,1)', color: 'rgba(112,64,144,1)', width: 1.5 }));
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
