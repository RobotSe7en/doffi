import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { importXlsx, writeDocumentZip } from '@doffi/office';

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

for (const large of [false, true]) test(`Playground opens a ${large ? 'ten-thousand-row' : 'small'} workbook and switches between isolated Word, PPT and workbook views`, { tag: large ? '@scale' : [] }, async ({ page }, info) => {
  test.setTimeout(120_000);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  const filename = large ? '一万行工作簿.xlsx' : '09-classic-notes.xlsx';
  const reopen = page.getByRole('button', { name: large ? '打开一万行 XLSX 样例' : '打开 09-classic-notes.xlsx', exact: true });
  if (large) await reopen.click(); else await upload(page).setInputFiles('examples/xlsx/09-classic-notes.xlsx');
  await expect(status(page)).toHaveAttribute("data-sheet-count", "1", { timeout: 60_000 });
  await expect(page.getByRole('heading', { name: filename, exact: true })).toBeVisible();
  await expect(page.getByLabel('Excel 工作簿编辑', { exact: true })).toBeVisible();
  const first = await workbookState(page);
  expect(first.grids).toBe(1); expect(first.readonly).toBe(true); expect(first.cells).toBeGreaterThan(0); expect(first.cells).toBeLessThan(2000);
  await expect(page.getByRole('toolbar', { name: '文件与导出' })).toBeVisible();
  await page.getByRole('tab', { name: '文件', exact: true }).click();
  await expect(page.getByRole('button', { name: '导出 XLSX', exact: true })).toBeDisabled();
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
  await page.getByRole('tab', { name: '文件', exact: true }).click();
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
  await address.fill('C3'); await address.press('Enter'); await input.fill('原值'); await input.press('Control+Enter'); await expect(status(page)).toHaveAttribute("data-revision", "1");
  const bold = page.getByRole('button', { name: '加粗', exact: true }); await expect(bold).toBeEnabled();
  await bold.click(); await expect(status(page)).toHaveAttribute("data-revision", "2"); await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await bold.click(); await expect(status(page)).toHaveAttribute("data-revision", "3"); await expect(bold).toHaveAttribute('aria-pressed', 'false');
  await page.getByLabel('字号', { exact: true }).fill("18"); await page.getByLabel('字号', { exact: true }).press("Enter"); await expect(status(page)).toHaveAttribute("data-revision", "4"); await expect(page.getByLabel('字号', { exact: true })).toHaveValue('18');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await expect(status(page)).toHaveAttribute("data-revision", "5"); await expect(page.getByLabel('字号', { exact: true })).toHaveValue('11');
  const cell = page.frameLocator('iframe').getByRole('gridcell').filter({ hasText: '原值' });
  await cell.dblclick(); await expect(page.getByRole('button', { name: '加粗', exact: true })).toBeDisabled(); await expect(input).toBeFocused(); await input.press('End'); await page.keyboard.insertText('追加'); await expect(input).toHaveValue('原值追加');
  await input.press('Escape'); await expect(input).toHaveValue('原值'); await expect(status(page)).toHaveAttribute("data-revision", "5");
  await input.press('F2'); await input.press('End'); await page.keyboard.insertText('追加'); await input.press('Enter');
  await expect(status(page)).toHaveAttribute("data-revision", "6");
  await address.fill('C3'); await address.press('Enter'); await expect(input).toHaveValue('原值追加');
  expect(await page.evaluate(() => window.workbookHostProbe.created)).toBe(1);
});
