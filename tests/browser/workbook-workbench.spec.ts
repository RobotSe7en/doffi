import { readFileSync, readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { openDocument } from "@doffi/core";
import { mixedEditingOperations } from "../workbook-combination-operations.js";
import type { WorkbookDocumentMetadata } from "@doffi/core";
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
    test("mixed native proposal freezes full context and reviews table, rules, drawings and thread changes before one apply", async ({ page }) => {
      const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/workbook-editing-combinations.xlsx'), { documentId: 'mixed-workbench' }), { kind: 'workbook' });
      const operations = mixedEditingOperations(model.snapshot()), bytes = Buffer.from(writeDocumentZip(model.serialize())).toString('base64');
      await page.evaluate(({ encoded, worker }) => window.workbookWorkbench.mount(encoded, worker || undefined), { encoded: bytes, worker: remote ? worker : '' });
      await page.getByLabel('请求写入范围', { exact: true }).selectOption('global');
      await page.getByRole('button', { name: '生成工作簿请求', exact: true }).click(); await expect(page.getByRole('status', { name: '工作簿协作状态', exact: true })).toContainText('请求已生成');
      const before = await page.evaluate(() => window.workbookWorkbench.inspect()), request = before.request!;
      expect(request.context.scope.kind).toBe('global'); expect(request.context.metadata.workbook.objects).toHaveLength(4); expect(request.context.metadata.workbook.sheets[0]!.tables).toHaveLength(1);
      expect(request.context.cells.find(cell => cell.sheetId === 'sheet-1' && cell.row === 1 && cell.column === 3)).toMatchObject({ value: { value: 999 }, current: { value: { value: 20 } } });
      await page.evaluate(operations => window.workbookWorkbench.loadOperations(operations), operations);
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, request });
      const details = page.getByText('结构、规则、绘图、批注锚点与样式', { exact: true }); await details.click();
      const diff = JSON.parse((await details.locator('..').locator('pre').textContent())!) as { after: { metadata: WorkbookDocumentMetadata } };
      expect(diff.after.metadata.workbook.names).toEqual(expect.arrayContaining([expect.objectContaining({ name: 'CurrentAmounts' })]));
      expect(diff.after.metadata.workbook.sheets[0]!.tables![0]!.autoFilter!.columns).toHaveLength(2);
      await page.getByRole('button', { name: '整批应用工作簿提案', exact: true }).click(); await expect.poll(() => page.evaluate(() => window.workbookWorkbench.revision())).toBe(1); await page.evaluate(() => window.workbookWorkbench.settled());
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
      await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect.poll(() => page.evaluate(() => window.workbookWorkbench.revision())).toBe(2); await page.evaluate(() => window.workbookWorkbench.settled());
      const undone = await page.evaluate(() => window.workbookWorkbench.inspect()); expect(undone).toMatchObject({ drawings: before.drawings, history: { undoSteps: 0 }, errors: [] });
      expect(undone.annotations.map(root => [root.status, root.anchor, root.replies])).toEqual(before.annotations.map(root => [root.status, root.anchor, root.replies]));
    });
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
      for (const form of [
        { panel: "名称管理", field: "定义名称", discard: "放弃当前名称草稿并重新读取" },
        { panel: "工作表管理", field: "工作表名称", discard: "放弃当前工作表草稿并重新读取" },
        { panel: "图片、形状与图表", field: "对象名称", discard: "放弃绘图表单草稿" },
        { panel: "备注与线程批注", field: "批注正文", discard: "放弃批注表单草稿" },
        { panel: "单元格超链接", field: "超链接外部地址", discard: "放弃超链接草稿并读取选区" },
        { panel: "智能表管理", field: "智能表名称", discard: "放弃智能表草稿并重新读取" },
        { panel: "自定义智能表样式", field: "自定义智能表样式名称", discard: "放弃表样式草稿并重新读取" },
        { panel: "筛选管理", field: "筛选区域", discard: "放弃筛选草稿并重新读取" },
        { panel: "条件格式与数据验证", field: "规则公式一", discard: "放弃规则草稿并重新读取" },
      ]) {
        await page.getByText(form.panel, { exact: true }).click(); const field = page.getByLabel(form.field, { exact: true }); await field.fill("UnsavedDraft");
        await expect(page.getByRole("button", { name: "生成工作簿请求", exact: true })).toBeDisabled();
        expect(await page.evaluate(() => window.workbookWorkbench.build().then(() => "accepted", (error: { code: string }) => error.code))).toBe("InvalidOperation"); await expect(field).toHaveValue("UnsavedDraft");
        await page.getByRole("button", { name: form.discard, exact: true }).click(); await expect(page.getByRole("button", { name: "生成工作簿请求", exact: true })).toBeEnabled();
      }
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
    test("layout compensation requires explicit object scope and preserves fixed drawing bounds through candidate apply and undo", async ({ page }) => {
      await page.evaluate(({ encoded, worker }) => window.workbookWorkbench.mount(encoded, worker || undefined, false, false, true), { encoded, worker: remote ? worker : "" });
      const nodeId = await page.evaluate(() => window.workbookWorkbench.addFixedObject()), before = await page.evaluate(() => window.workbookWorkbench.inspect());
      const bounds = () => page.evaluate(() => window.workbookWorkbench.drawingBounds()); const original = (await bounds()).objects.find(object => object.nodeId === nodeId)!; expect(original).toBeDefined();
      await page.evaluate(() => window.workbookWorkbench.build());
      expect(await page.evaluate(() => window.workbookWorkbench.formatFont().then(() => "accepted", (error: { code: string }) => error.code))).toBe("InvalidProposal");
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ revision: before.revision, drawings: before.drawings, history: before.history });
      await page.getByLabel("允许修改的绘图", { exact: true }).selectOption(nodeId); await page.evaluate(() => window.workbookWorkbench.build());
      await page.evaluate(() => window.workbookWorkbench.formatFont());
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ revision: before.revision, drawings: before.drawings });
      const title = page.getByText("结构、规则、绘图、批注锚点与样式", { exact: true }); await title.click();
      const details = JSON.parse((await title.locator("..").locator("pre").textContent())!) as { after: { metadata: WorkbookDocumentMetadata } };
      const candidate = details.after.metadata.workbook.objects!.find(object => object.nodeId === nodeId)!.anchor;
      expect(candidate.kind).toBe("two-cell"); if (candidate.kind !== "two-cell") throw Error("Missing cell anchor.");
      expect(candidate.from.offsetYPt).toBeLessThan(3);
      await page.getByRole("button", { name: "整批应用工作簿提案", exact: true }).click(); await expect.poll(() => page.evaluate(() => window.workbookWorkbench.revision())).toBe(before.revision + 1); await page.evaluate(() => window.workbookWorkbench.settled());
      const after = (await bounds()).objects.find(object => object.nodeId === nodeId)!;
      for (const field of ["xPt", "yPt", "widthPt", "heightPt"] as const) expect(after[field]).toBeCloseTo(original[field], 1);
      expect(await page.evaluate(() => window.workbookWorkbench.inspect())).toMatchObject({ drawings: [expect.objectContaining({ nodeId, anchor: candidate })], history: { undoSteps: before.history.undoSteps + 1 } });
      await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookWorkbench.settled());
      const restored = await page.evaluate(() => window.workbookWorkbench.inspect()); expect(restored.drawings).toEqual(before.drawings);
      expect(restored.annotations).toEqual(before.annotations.map(annotation => ({ ...annotation, baseRevision: restored.revision })));
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

function warningSource(): string {
  const files = new Map(readDocumentZip(Buffer.from(encoded, 'base64'))), metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: 'warning', kind: 'decimal', operator: 'lessThan', formulas: ['=5'], origin: { row: 0, column: 0 },
    ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], allowBlank: true, errorStyle: 'warning', showErrorMessage: true, showInputMessage: true,
    errorTitle: '范围', error: '必须小于五', promptTitle: '', prompt: '' }];
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata))); return Buffer.from(writeDocumentZip(files)).toString('base64');
}
for (const mode of ['direct', 'worker', 'react'] as const) test(`${mode} workbench forwards warning decisions to human input and isolated proposals`, async ({ page }) => {
  await openSdkHost(page); await page.addScriptTag({ content: code });
  await page.evaluate(({ source, worker, react }) => window.workbookWorkbench.mount(source, worker || undefined, react, true), { source: warningSource(), worker: mode === 'direct' ? '' : worker, react: mode === 'react' });
  try {
    const settled = () => page.evaluate(() => window.workbookWorkbench.settled());
    const inspect = async () => { await settled(); return page.evaluate(() => window.workbookWorkbench.inspect()); }, original = await inspect();
    const afterChange = async (previous: number) => { await expect.poll(() => page.evaluate(() => window.workbookWorkbench.revision())).toBe(previous + 1); await settled(); };
    const button = (name: string) => page.getByRole('button', { name, exact: true }), input = page.getByLabel('单元格或公式输入', { exact: true });
    await input.fill('9'); page.once('dialog', dialog => dialog.accept()); await button('确认输入').click();
    await afterChange(original.revision); expect((await inspect()).history.undoSteps).toBe(1); expect((await inspect()).cells[0]?.value).toEqual({ type: 'number', value: 9 });
    await button('撤销').click(); await afterChange(original.revision + 1); expect((await inspect()).history.undoSteps).toBe(0); expect((await inspect()).cells[0]?.value).toEqual(original.cells[0]?.value);
    page.once('dialog', dialog => dialog.dismiss()); await button('提交工作簿批注').click(); await expect(button('生成工作簿请求')).toBeEnabled();
    expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, cells: [{ value: original.cells[0]!.value }, {}], annotations: [{ status: 'open' }] }); await expect(button('整批应用工作簿提案')).toBeDisabled();
    page.once('dialog', dialog => dialog.accept()); await button('提交工作簿批注').click(); await expect(button('整批应用工作簿提案')).toBeEnabled();
    expect(await inspect()).toMatchObject({ history: { undoSteps: 0 }, cells: [{ value: original.cells[0]!.value }, {}], annotations: [{ status: 'open' }] });
    await button('整批应用工作簿提案').click(); await afterChange(original.revision + 2); expect((await inspect()).history.undoSteps).toBe(1);
    expect(await inspect()).toMatchObject({ cells: [{ value: { type: 'number', value: 10 } }, {}], annotations: [{ status: 'resolved' }] });
    await button('撤销').click(); await afterChange(original.revision + 3); expect((await inspect()).history.undoSteps).toBe(0); expect((await inspect()).cells[0]?.value).toEqual(original.cells[0]?.value);
  } finally { await page.evaluate(() => window.workbookWorkbench.finish()); }
});
