import { expect, test } from '@playwright/test';
import type { Page, Route } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { openDocument } from '@doffi/core';
import type { ConversionRequestV1 } from '@doffi/node';
import { importXlsx, readDocumentZip, writeDocumentZip } from '@doffi/office';

let workbookZip: Buffer, reportZip: Buffer;
declare global {
  interface Window {
    workbookHostProbe: { created: number; terminated: number; roots: ShadowRoot[]; failLayout: boolean; deferReplies: boolean; replies: Array<() => void> };
  }
}
test.beforeAll(async () => {
  workbookZip = Buffer.from(writeDocumentZip(await importXlsx(readFileSync('examples/xlsx/09-classic-notes.xlsx'), { documentId: 'workspace-workbook-package' })));
  reportZip = Buffer.from(writeDocumentZip(new Map(['manifest.json', 'content.html', 'styles.css', 'metadata.json', 'annotations.json'].map(path => [path, new Uint8Array(readFileSync('examples/report/' + path))]))));
});
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    localStorage.setItem("doffi:library", "open"); localStorage.setItem("doffi:editing", "preview");
    window.workbookHostProbe = { created: 0, terminated: 0, roots: [], failLayout: false, deferReplies: false, replies: [] };
    const attach = Element.prototype.attachShadow;
    Element.prototype.attachShadow = function (options) { const root = attach.call(this, { ...options, mode: 'open' }); window.workbookHostProbe.roots.push(root); return root; };
    const add = EventTarget.prototype.addEventListener, guardedPorts = new WeakSet<MessagePort>();
    EventTarget.prototype.addEventListener = function (type, listener, options) {
      if (this instanceof MessagePort && type === 'message' && !guardedPorts.has(this)) {
        guardedPorts.add(this); const port = this;
        add.call(port, 'message', event => { if (!window.workbookHostProbe.deferReplies) return; event.stopImmediatePropagation(); const message = event as MessageEvent<unknown>; window.workbookHostProbe.replies.push(() => port.dispatchEvent(new MessageEvent('message', { data: message.data, ports: [...message.ports] }))); }, { capture: true });
      }
      return add.call(this, type, listener, options);
    };
    const NativeWorker = window.Worker;
    window.Worker = class extends NativeWorker {
      constructor(url: string | URL, options?: WorkerOptions) { super(url, options); window.workbookHostProbe.created++; }
      override terminate(): void { window.workbookHostProbe.terminated++; super.terminate(); }
    };
    const computed = window.getComputedStyle;
    window.getComputedStyle = function (element, pseudo) {
      if (window.parent.workbookHostProbe?.failLayout && element.closest('office-workbook-source')) throw Error('Injected source measurement failure');
      return computed.call(this, element, pseudo);
    };
  });
  await page.goto('/');
});
const status = (page: Page) => page.locator('.office-app');
const upload = (page: Page) => page.getByLabel('打开文件', { exact: true }).filter({ visible: false });
const formulaInput = (page: Page) => page.getByRole('textbox', { name: '单元格或公式输入', exact: true });

test('Playground Excel zoom resizes the grid and keeps pointer selection and formula editing aligned', async ({ page }) => {
  await upload(page).setInputFiles('examples/xlsx/01-values-text.xlsx');
  await expect(status(page)).toHaveAttribute('data-revision', '0');
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  const grid = page.frameLocator('iframe[title="工作簿只读预览"]');
  const cells = grid.getByRole('gridcell');
  const second = cells.nth(1);
  const column = page.getByRole('columnheader', { name: 'B', exact: true });
  const row = page.getByRole('rowheader', { name: '1', exact: true });
  const zoomedCellWidth = async () => (await column.boundingBox())!.width;
  await expect(second).toBeVisible();
  const original = await zoomedCellWidth();
  await page.getByRole('button', { name: '放大', exact: true }).click();
  await expect.poll(async () => (await zoomedCellWidth()) / original).toBeGreaterThan(1.2);
  const x = (await column.boundingBox())!, y = (await row.boundingBox())!;
  await page.mouse.click(x.x + x.width / 2, y.y + y.height / 2);
  await expect(address).toHaveValue('B1');
  await expect(status(page)).toHaveAttribute('data-revision', '0');
  await formulaInput(page).fill('=2*9');
  await formulaInput(page).press('Control+Enter');
  await expect(status(page)).toHaveAttribute('data-revision', '1');
  await expect(second).toHaveText('18');
  await page.getByRole('button', { name: '缩小', exact: true }).click();
  await expect.poll(zoomedCellWidth).toBeCloseTo(original, 0);
  await expect(address).toHaveValue('B1');
});

test('Playground exports the current revision as editable XLSX and a reopenable package from one Worker', async ({ page }) => {
  await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx');
  await expect(status(page)).toHaveAttribute('data-revision', '0', { timeout: 30_000 });
  const input = formulaInput(page), address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  await address.fill('AG257'); await address.press('Enter');
  await expect(input).toHaveValue('');
  await input.fill('=7*3');
  await page.getByRole('button', { name: '文件菜单', exact: true }).click(); await page.getByRole('menuitem', { name: '文件和导出…', exact: true }).click();
  const exportNative = page.getByRole('button', { name: '导出 XLSX', exact: true });
  await expect(exportNative).toBeDisabled();
  await input.press('Control+Enter');
  await expect(status(page)).toHaveAttribute('data-revision', '1');
  await expect(exportNative).toBeEnabled();
  const xlsx = page.waitForEvent('download');
  await exportNative.click();
  const native = await xlsx;
  expect(native.suggestedFilename()).toBe('09-classic-notes.xlsx');
  const reopened = openDocument(await importXlsx(readFileSync(await native.path()), { documentId: 'native-reopen' }), { kind: 'workbook' });
  const sheetId = reopened.snapshot().metadata.workbook.sheets[0]!.sheetId;
  expect(reopened.getCell(sheetId, 256, 32)).toMatchObject({ value: { type: 'number', value: 21 }, formula: { expression: '=7*3' } });
  await expect(page.getByRole('region', { name: '工作簿导出报告' })).toContainText('原生批注');

  const zip = page.waitForEvent('download');
  await page.getByRole('button', { name: '另存为文档包', exact: true }).click();
  const saved = await zip;
  expect(saved.suggestedFilename()).toBe('09-classic-notes.zip');
  const packageBytes = readFileSync(await saved.path());
  const source = openDocument(readDocumentZip(packageBytes), { kind: 'workbook' });
  expect(source.snapshot().manifest.revision).toBe(1);
  expect(source.getCell(source.snapshot().metadata.workbook.sheets[0]!.sheetId, 256, 32)?.formula?.expression).toBe('=7*3');
  expect(await page.evaluate(() => ({ created: window.workbookHostProbe.created, terminated: window.workbookHostProbe.terminated }))).toEqual({ created: 1, terminated: 0 });
  await expect(status(page)).toHaveAttribute('data-revision', '1');

  await upload(page).setInputFiles({ name: 'reopened-workbook.zip', mimeType: 'application/zip', buffer: packageBytes });
  await expect(page.getByRole('heading', { name: 'reopened-workbook.zip' })).toBeVisible();
  await expect(status(page)).toHaveAttribute('data-revision', '1');
  await expect(page.getByRole('button', { name: '撤销', exact: true })).toBeDisabled();
});

test('Playground cancels and retries native export of the ten-thousand-row sample', { tag: '@scale' }, async ({ page }) => {
  test.setTimeout(180_000);
  await page.getByRole('button', { name: '打开一万行 XLSX 样例', exact: true }).click();
  await expect(status(page)).toHaveAttribute('data-revision', '0', { timeout: 60_000 });
  await page.getByRole('button', { name: '文件菜单', exact: true }).click(); await page.getByRole('menuitem', { name: '文件和导出…', exact: true }).click();
  const action = page.getByRole('button', { name: '导出 XLSX', exact: true });
  await action.click();
  await expect(page.getByRole('status').filter({ hasText: '正在导出 XLSX' })).toContainText('计算公式', { timeout: 15_000 });
  await page.getByRole('button', { name: '取消导出', exact: true }).click();
  await expect(page.getByRole('alert')).toHaveText('已取消 XLSX 导出。');
  await expect(action).toBeEnabled();
  const download = page.waitForEvent('download', { timeout: 75_000 });
  await action.click();
  const output = await download;
  expect(output.suggestedFilename()).toBe('一万行工作簿.xlsx');
  const model = openDocument(await importXlsx(readFileSync(await output.path()), { documentId: 'large-output' }), { kind: 'workbook' });
  const { manifest, metadata } = model.snapshot();
  expect(manifest.chunks).toHaveLength(40);
  expect(model.getCell(metadata.workbook.sheets[0]!.sheetId, 9999, 19)?.value).toEqual({ type: 'number', value: 3_400_152 });
  await expect(status(page)).toHaveAttribute('data-revision', '0');
  expect(await page.evaluate(() => ({ created: window.workbookHostProbe.created, terminated: window.workbookHostProbe.terminated }))).toEqual({ created: 1, terminated: 0 });
});

interface PrintCall { readonly route: Route; readonly request: ConversionRequestV1 }
async function capturePrint(page: Page): Promise<PrintCall[]> {
  const calls: PrintCall[] = [];
  await page.context().route('**/__doffi/convert', route => {
    calls.push({ route, request: route.request().postDataJSON() as ConversionRequestV1 });
  });
  return calls;
}
function replyPrint(call: PrintCall, revision: number): Promise<void> {
  const { request, route } = call;
  return route.fulfill({ json: { protocolVersion: 1, taskId: request.taskId, ok: true, output: {
    mimeType: request.format === 'pdf' ? 'application/pdf' : 'image/png',
    bytesBase64: Buffer.from('print result').toString('base64'),
    report: { kind: 'workbook', documentId: request.input.documentId, revision, pages: [{}, {}],
      formulaValues: request.options?.workbook?.formulaValues },
  } } });
}
async function openPrint(page: Page, format: 'PDF' | 'PNG'): Promise<void> {
  await page.getByRole('button', { name: '文件菜单', exact: true }).click(); await page.getByRole('menuitem', { name: '文件和导出…', exact: true }).click();
  await page.getByRole('button', { name: '导出 ' + format, exact: true }).click();
}

test('Playground prints the selected cells with explicit saved values and a PNG page', async ({ page }) => {
  const calls = await capturePrint(page);
  await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx');
  await expect(status(page)).toHaveAttribute('data-revision', '0', { timeout: 30_000 });
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  await address.fill('B2:D4'); await address.press('Enter');
  await expect(address).toHaveValue('B2:D4');
  await openPrint(page, 'PNG');
  await page.getByLabel('打印范围').selectOption('selection');
  await page.getByLabel('打印页码').fill('2');
  await page.getByLabel('使用已保存的公式值').check();
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: '生成 PNG', exact: true }).click();
  await expect.poll(() => calls.length).toBe(1);
  const call = calls[0]!, source = openDocument(readDocumentZip(Buffer.from(call.request.input.bytesBase64, 'base64')), { kind: 'workbook' });
  expect(call.request).toMatchObject({ format: 'png', input: { format: 'document-package-zip', documentId: source.snapshot().manifest.documentId },
    options: { pageIndex: 1, workbook: { fileName: '09-classic-notes.xlsx', formulaValues: 'cached',
      sheets: [{ sheetId: source.snapshot().metadata.workbook.sheets[0]!.sheetId, areas: [{ from: { row: 1, column: 1 }, to: { row: 3, column: 3 } }] }] } } });
  await replyPrint(call, 0);
  expect((await download).suggestedFilename()).toBe('09-classic-notes.png');
  await expect(page.getByRole('region', { name: '工作簿导出报告' })).toContainText('已使用保存的公式值');
});

test('Playground cancels print on demand, on revision changes and on switching documents', async ({ page }) => {
  const calls = await capturePrint(page), downloads: string[] = [];
  page.on('download', result => downloads.push(result.suggestedFilename()));
  await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx');
  await expect(status(page)).toHaveAttribute('data-revision', '0', { timeout: 30_000 });
  const start = async (count: number) => {
    await openPrint(page, 'PDF');
    await page.getByRole('button', { name: '生成 PDF', exact: true }).click();
    await expect.poll(() => calls.length).toBe(count);
  };
  await start(1);
  await page.getByRole('button', { name: '取消导出', exact: true }).click();
  await expect(page.getByRole('alert')).toHaveText('已取消 PDF 导出。');
  await replyPrint(calls[0]!, 0).catch(() => {});
  await start(2);
  const input = formulaInput(page);
  await input.fill('=20+22'); await input.press('Control+Enter');
  await expect(status(page)).toHaveAttribute('data-revision', '1');
  await expect(page.getByRole('alert')).toContainText('工作簿已变化');
  await replyPrint(calls[1]!, 0).catch(() => {});
  await start(3);
  await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  await expect(page.getByRole('heading', { name: '季度项目报告', exact: true })).toBeVisible();
  await replyPrint(calls[2]!, 1).catch(() => {});
  await expect(page.getByRole('button', { name: '取消导出', exact: true })).toHaveCount(0);
  expect(downloads).toEqual([]);

  await page.getByRole('button', { name: '打开 09-classic-notes.xlsx', exact: true }).click();
  await expect(status(page)).toHaveAttribute('data-revision', '1');
  await start(4);
  const returned = page.waitForEvent('download');
  await replyPrint(calls[3]!, 1);
  expect((await returned).suggestedFilename()).toBe('09-classic-notes.pdf');
  await start(5);
  await page.getByRole('button', { name: '返回首页', exact: true }).click();
  await expect(status(page)).toHaveClass(/is-home/);
  await replyPrint(calls[4]!, 1).catch(() => {});
  expect(downloads).toEqual(['09-classic-notes.pdf']);
  expect(await page.evaluate(() => ({ created: window.workbookHostProbe.created, terminated: window.workbookHostProbe.terminated }))).toEqual({ created: 1, terminated: 0 });
});

test('Playground edits through its Worker, retains drafts on file switches and reopens the same history', async ({ page }) => {
  await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx'); await expect(status(page)).toHaveAttribute("data-revision", "0", { timeout: 30_000 });
  const input = page.getByRole('textbox', { name: '单元格或公式输入', exact: true }), address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  await address.fill('AG257'); await address.press('Enter'); await expect(input).toHaveValue('');
  await input.fill('=1+1'); await input.press('Control+Enter'); await expect(status(page)).toHaveAttribute("data-revision", "1"); await expect(input).toHaveValue('=1+1');
  await input.fill('draft pending'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('请先确认或取消'); await expect(input).toHaveValue('draft pending');
  await expect(page.getByRole('heading', { name: '09-classic-notes.xlsx', exact: true })).toBeVisible();
  await page.getByRole('button', { name: '新建 PPT', exact: true }).click(); await expect(input).toHaveValue('draft pending');
  await expect(page.getByRole('heading', { name: '09-classic-notes.xlsx', exact: true })).toBeVisible();
  await input.press('Escape'); await expect(input).toHaveValue('=1+1'); await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await page.getByRole('button', { name: '打开 09-classic-notes.xlsx', exact: true }).click(); await expect(status(page)).toHaveAttribute("data-revision", "1"); await expect(input).toHaveValue('=1+1');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect(status(page)).toHaveAttribute("data-revision", "2"); await expect(input).toHaveValue('');
  expect(await page.evaluate(() => ({ created: window.workbookHostProbe.created, terminated: window.workbookHostProbe.terminated }))).toEqual({ created: 1, terminated: 0 });
});
async function workbookState(page: Page) {
  const frame = page.frames().find(frame => frame.url() === 'about:srcdoc');
  if (!frame) throw Error('Workbook frame is missing');
  return frame.evaluate(() => {
    const roots = window.workbookHostProbe.roots.filter(root => root.host.isConnected);
    const grids = roots.flatMap(root => [...root.querySelectorAll('[role="grid"]')]);
    return { grids: grids.length, cells: grids.reduce((sum, grid) => sum + grid.querySelectorAll('[role="gridcell"]').length, 0),
      text: grids.map(grid => grid.textContent).join(''), readonly: grids.every(grid => grid.getAttribute('aria-readonly') === 'true') };
  });
}

for (const form of [
  { kind: 'name', tab: '公式', panel: '名称管理器', field: '定义名称', discard: '放弃当前名称草稿并重新读取' },
  { kind: 'worksheet', tab: '数据', panel: '工作表管理', field: '工作表名称', discard: '放弃当前工作表草稿并重新读取' },
  { kind: 'drawing', tab: '插入', panel: '图片', field: '对象名称', discard: '放弃绘图表单草稿' },
  { kind: 'comment', tab: '审阅', panel: '显示批注', field: '批注正文', discard: '放弃批注表单草稿' },
  { kind: 'hyperlink', tab: '插入', panel: '链接', field: '超链接外部地址', discard: '放弃超链接草稿并读取选区' },
  { kind: 'table', tab: '数据', panel: '智能表管理', field: '智能表名称', discard: '放弃智能表草稿并重新读取' },
  { kind: 'filter', tab: '数据', panel: '筛选', field: '筛选区域', discard: '放弃筛选草稿并重新读取' },
  { kind: 'rule', tab: '数据', panel: '数据验证', field: '规则公式一', discard: '放弃规则草稿并重新读取' },
]) test(`Playground retains an unsaved ${form.kind} draft on file switches until explicit discard`, async ({ page }) => {
  await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx'); await expect(status(page)).toHaveAttribute("data-revision", "0", { timeout: 30_000 });
  await page.getByRole('tab', { name: form.tab, exact: true }).click(); await page.getByRole('button', { name: form.panel, exact: true }).click(); const name = page.getByLabel(form.field, { exact: true }), original = await name.inputValue(); await name.fill('UnsavedDraft');
  await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('请先确认或取消'); await expect(name).toHaveValue('UnsavedDraft');
  await page.getByRole('button', { name: '新建 PPT', exact: true }).click(); await expect(name).toHaveValue('UnsavedDraft');
  await expect(page.getByRole('heading', { name: '09-classic-notes.xlsx', exact: true })).toBeVisible();
  await page.getByRole('button', { name: form.discard, exact: true }).click(); await expect(name).toHaveValue(original);
  await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click(); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await page.getByRole('button', { name: '打开 09-classic-notes.xlsx', exact: true }).click(); await expect(status(page)).toHaveAttribute("data-revision", "0");
  await expect(page.getByRole('button', { name: '撤销', exact: true })).toBeDisabled();
  expect(await page.evaluate(() => ({ created: window.workbookHostProbe.created, terminated: window.workbookHostProbe.terminated }))).toEqual({ created: 1, terminated: 0 });
});

test('Playground opens a small workbook and switches between isolated Word, PPT and workbook views', async ({ page }, info) => {
  test.setTimeout(120_000);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  const reopen = page.getByRole('button', { name: '打开 09-classic-notes.xlsx', exact: true });
  await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx');
  await expect(status(page)).toHaveAttribute("data-sheet-count", "1", { timeout: 60_000 });
  await expect(page.getByRole('heading', { name: '09-classic-notes.xlsx', exact: true })).toBeVisible();
  await expect(page.getByLabel('Excel 工作簿编辑', { exact: true })).toBeVisible();
  const first = await workbookState(page);
  expect(first.grids).toBe(1); expect(first.readonly).toBe(true); expect(first.cells).toBeGreaterThan(0); expect(first.cells).toBeLessThan(2000);
  await expect(page.getByRole('toolbar', { name: '文件与导出' })).toBeVisible();
  await page.getByRole('button', { name: '文件菜单', exact: true }).click(); await page.getByRole('menuitem', { name: '文件和导出…', exact: true }).click();
  await expect(page.getByRole('button', { name: '导出 XLSX', exact: true })).toBeEnabled();
  await expect(page.getByRole('button', { name: '导出 DOCX', exact: true })).toHaveCount(0);
  await page.screenshot({ path: info.outputPath('workbook-playground.png') });
  await page.setViewportSize({ width: 1000, height: 800 });
  await expect.poll(() => page.locator('.workbook-pane').evaluate(element => element.scrollWidth <= element.clientWidth)).toBe(true);
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]')).toBeVisible();
  await page.getByRole('tab', { name: '审阅', exact: true }).click();
  await page.getByRole('button', { name: '载入固定示例提案', exact: true }).click();
  const apply = page.getByRole('button', { name: '整批应用提案', exact: true });
  await expect(apply).toBeVisible(); await apply.click();
  await page.getByRole('button', { name: '打开 PPT 协作样例', exact: true }).click();
  await expect(page.locator('iframe[title="幻灯片只读预览"]')).toBeVisible();
  await reopen.click();
  await expect(status(page)).toHaveAttribute("data-sheet-count", "1");
  await expect.poll(() => workbookState(page)).toEqual(first);
  await expect(page.getByRole('alert')).toHaveCount(0);
  await reopen.click();
  await expect(page.getByRole('button', { name: '取消读取' })).toHaveCount(0);
  expect(await page.evaluate(() => ({ created: window.workbookHostProbe.created, terminated: window.workbookHostProbe.terminated }))).toEqual({ created: 1, terminated: 0 });
  await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  expect(errors).toEqual([]);
});

test('cancelling an actual import terminates its Worker, preserves the active document and allows retry', { tag: '@scale' }, async ({ page }) => {
  test.setTimeout(120_000);
  await page.getByRole('button', { name: '打开报告协作样例', exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await upload(page).setInputFiles('examples/xlsx/15-ten-thousand-rows.xlsx');
  await expect(page.getByRole('status').filter({ hasText: /读取工作表|校验工作簿/ })).toBeVisible({ timeout: 30_000 });
  await page.getByRole('button', { name: '取消读取', exact: true }).click();
  await expect(page.getByRole('alert')).toHaveText('已取消读取。');
  await expect(page.getByRole('heading', { name: '季度项目报告', exact: true })).toBeVisible();
  await expect(page.locator('iframe[title="文档只读预览"]')).toBeVisible();
  await expect.poll(() => page.evaluate(() => window.workbookHostProbe.terminated)).toBe(1);
  await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx');
  await expect(status(page)).toHaveAttribute("data-annotation-count", "1", { timeout: 30_000 });
  await expect(page.getByRole('heading', { name: '09-classic-notes.xlsx', exact: true })).toBeVisible();
  expect(await workbookState(page)).toMatchObject({ grids: 1, readonly: true });
});

test('workbook ZIP and native failures retain the accepted canvas and content ZIPs still reopen through the Word session', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await upload(page).setInputFiles({ name: 'notes.zip', mimeType: 'application/zip', buffer: workbookZip });
  await expect(status(page)).toHaveAttribute("data-annotation-count", "1");
  const accepted = await workbookState(page);
  await page.evaluate(() => { window.workbookHostProbe.failLayout = true; });
  await upload(page).setInputFiles('examples/xlsx/08-images-shapes.xlsx');
  await expect(page.getByRole('alert')).toBeVisible();
  await expect(page.getByRole('button', { name: '取消读取' })).toHaveCount(0);
  await expect(page.getByRole('heading', { name: 'notes.zip', exact: true })).toBeVisible();
  await page.evaluate(() => { window.workbookHostProbe.failLayout = false; });
  expect(await workbookState(page)).toEqual(accepted);
  await upload(page).setInputFiles({ name: 'broken.xlsx', mimeType: 'application/zip', buffer: Buffer.from('not a ZIP') });
  await expect(page.getByRole('alert')).toBeVisible();
  await expect(page.getByRole('button', { name: '取消读取' })).toHaveCount(0);
  expect(await workbookState(page)).toEqual(accepted);
  await expect.poll(() => page.evaluate(() => window.workbookHostProbe.terminated)).toBe(2);
  await upload(page).setInputFiles({ name: 'report.zip', mimeType: 'application/zip', buffer: reportZip }); await expect(page.getByRole('heading', { name: 'report.zip', exact: true })).toBeVisible();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  await page.getByRole('button', { name: '文件菜单', exact: true }).click(); await page.getByRole('menuitem', { name: '文件和导出…', exact: true }).click();
  await expect(page.getByRole('button', { name: '导出 DOCX', exact: true })).toBeEnabled();
  await expect.poll(() => page.evaluate(() => window.workbookHostProbe.terminated)).toBe(3);
  expect(errors).toEqual([]);
});

test('Excel font fields commit through the Worker before moving keyboard focus', async ({ page }) => {
  await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx'); await expect(status(page)).toHaveAttribute('data-revision', '0');
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true }), font = page.getByRole('combobox', { name: '字体', exact: true }), size = page.getByRole('combobox', { name: '字号', exact: true });
  await expect(font).toBeEnabled(); await page.evaluate(() => { window.workbookHostProbe.deferReplies = true; });
  await address.fill('C3'); await address.press('Enter'); await expect.poll(() => page.evaluate(() => window.workbookHostProbe.replies.length)).toBeGreaterThan(0); await expect(font).toBeDisabled();
  await page.evaluate(() => { window.workbookHostProbe.deferReplies = false; for (const reply of window.workbookHostProbe.replies.splice(0)) reply(); });
  await expect(font).toBeEnabled(); await font.fill('FandolKai'); await font.press('Tab');
  await expect(size).toBeFocused(); await expect(font).toHaveValue('FandolKai'); await expect(status(page)).toHaveAttribute('data-revision', '1');
  await size.fill('12.5'); await size.press('Shift+Tab'); await expect(font).toBeFocused(); await expect(size).toHaveValue('12.5'); await expect(status(page)).toHaveAttribute('data-revision', '2');
});

test('Excel ribbon formats through its Worker and double click edits while Escape cancels', async ({ page }) => {
  await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx'); await expect(status(page)).toHaveAttribute("data-revision", "0");
  const input = page.getByRole('textbox', { name: '单元格或公式输入', exact: true }), address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  await address.fill('C3'); await address.press('Enter');
  await expect(input).toBeFocused(); await expect(input).toHaveValue('');
  await input.fill('原值'); await input.press('Control+Enter'); await expect(status(page)).toHaveAttribute("data-revision", "1");
  const cell = page.frameLocator('iframe').getByRole('gridcell').filter({ hasText: '原值' });
  await expect(cell).toHaveAttribute('aria-colindex', '3');
  await expect(cell.locator('..')).toHaveAttribute('aria-rowindex', '3');
  const bold = page.getByRole('button', { name: '加粗', exact: true }); await expect(bold).toBeEnabled();
  await bold.click(); await expect(status(page)).toHaveAttribute("data-revision", "2"); await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await bold.click(); await expect(status(page)).toHaveAttribute("data-revision", "3"); await expect(bold).toHaveAttribute('aria-pressed', 'false');
  await page.getByLabel('字号', { exact: true }).fill("18"); await page.getByLabel('字号', { exact: true }).press("Enter"); await expect(status(page)).toHaveAttribute("data-revision", "4"); await expect(page.getByLabel('字号', { exact: true })).toHaveValue('18');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect(status(page)).toHaveAttribute("data-revision", "5"); await expect(page.getByLabel('字号', { exact: true })).toHaveValue('11');
  await cell.dblclick(); await expect(page.getByRole('button', { name: '加粗', exact: true })).toBeDisabled(); await expect(input).toBeFocused(); await input.press('End'); await page.keyboard.insertText('追加'); await expect(input).toHaveValue('原值追加');
  await input.press('Escape'); await expect(input).toHaveValue('原值'); await expect(status(page)).toHaveAttribute("data-revision", "5");
  await input.press('F2'); await input.press('End'); await page.keyboard.insertText('追加'); await input.press('Enter');
  await expect(status(page)).toHaveAttribute("data-revision", "6");
  await address.fill('C3'); await address.press('Enter'); await expect(input).toHaveValue('原值追加');
  expect(await page.evaluate(() => window.workbookHostProbe.created)).toBe(1);
});
