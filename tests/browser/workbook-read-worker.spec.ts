import { openSdkHost } from './sdk-host.js';
import { expect, test } from '@playwright/test';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { createHash } from 'node:crypto';
import { writeDocumentZip } from '@doffi/office';
import { scalarWorkbook } from '../../packages/core/tests/workbook-scalar-fixture.js';
import { createDocumentSession } from '@doffi/core';
import type * as editorFixture from './fixtures/workbook-remote-editor.js';
import type * as fixture from './fixtures/workbook-read-host.js';
import type * as sessionFixture from './fixtures/workbook-session-host.js';
import type * as performanceFixture from './fixtures/workbook-editing-performance.js';
import type * as mountedPerformanceFixture from './fixtures/workbook-editor-performance.js';
declare global { interface Window { workbookReadHost: typeof fixture } }
declare global { interface Window { workbookSessionHost: typeof sessionFixture } }
declare global { interface Window { workbookEditingPerformance: typeof performanceFixture } }
declare global { interface Window { workbookEditorPerformance: typeof mountedPerformanceFixture } }
declare global { interface Window { workbookRemoteEditor: typeof editorFixture } }
let workerCode = '', hostCode = '', sessionWorkerCode = '', sessionHostCode = '';
let performanceHostCode = '';
let editorHostCode = '';
async function bundle(name: string): Promise<string> {
  const result = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL(`./fixtures/${name}.ts`, import.meta.url)), formats: ['iife'], name: 'WorkbookReadProbe' } } });
  if (!Array.isArray(result) && !('output' in result)) throw Error('Unexpected read-worker bundle');
  const output = Array.isArray(result) ? result.flatMap(part => part.output) : result.output;
  return output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
}
test.beforeAll(async () => {
  workerCode = await bundle('workbook-read-worker'); hostCode = await bundle('workbook-read-host');
  sessionWorkerCode = await bundle('workbook-session-worker'); sessionHostCode = await bundle('workbook-session-host');
  performanceHostCode = await bundle('workbook-editing-performance');
  editorHostCode = await bundle('workbook-remote-editor');
});
test('mounted remote workbook editor shares Worker input, address selection and history and leaves its host usable', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const bytes = Buffer.from(writeDocumentZip(scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 1 } },
    { row: 0, column: 1, value: { type: 'number', value: 999 }, formula: '=A1+1' }]).serialize())).toString('base64');
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  const input = page.getByRole('textbox', { name: '单元格或公式输入', exact: true }); await expect(input).toHaveValue('1');
  await input.fill('5'); await input.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, source: { text: '5' }, current: { kind: 'value', value: { type: 'number', value: 6 } }, errors: [] });
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true }); await address.fill('A1'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled()); await expect(input).toHaveValue('5');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled()); await expect(input).toHaveValue('1');
  await page.getByRole('button', { name: '重做', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled()); await expect(input).toHaveValue('5');
  await input.fill('='); await page.evaluate(() => window.workbookRemoteEditor.external('external')); await input.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ source: { text: 'external' }, input: { text: '=', dirty: true }, errors: ['RevisionConflict'] });
  await input.press('Escape'); await page.evaluate(() => window.workbookRemoteEditor.settled()); await expect(input).toHaveValue('external');
  expect(await page.evaluate(() => window.workbookRemoteEditor.disposeEditor())).toBe(true);
  expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, hostUsable: { type: 'number', value: 42 } });
});

test('mounted remote private clipboard repeats copies, takes current values and rejects a stale cut without removing its source', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const bytes = Buffer.from(writeDocumentZip(scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 7 } },
    { row: 0, column: 1, value: { type: 'number', value: 999 }, formula: '=A1+1' }]).serialize())).toString('base64');
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  const choose = async (value: string) => { await address.fill(value); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled()); };
  await choose('A1:B1'); await page.getByRole('button', { name: '复制选区', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  for (const row of [3, 5]) {
    await choose('A' + row); await page.getByRole('button', { name: '粘贴选区', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
    expect(await page.evaluate(row => window.workbookRemoteEditor.value(row - 1, 1), row)).toBe('=(A' + row + '+1)');
  }
  await choose('A7'); await page.getByLabel('区域粘贴方式', { exact: true }).selectOption('values');
  await page.getByRole('button', { name: '粘贴选区', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.value(6, 1))).toBe('8');
  await choose('A1:B1'); await page.getByRole('button', { name: '剪切选区', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.evaluate(() => window.workbookRemoteEditor.external('external', false)); await choose('A9');
  await page.getByRole('button', { name: '粘贴选区', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 4, history: { undoSteps: 4 }, errors: ['RevisionConflict'] });
  expect(await page.evaluate(() => Promise.all([window.workbookRemoteEditor.value(0, 0), window.workbookRemoteEditor.value(8, 0)]))).toEqual(['external', '']);
  expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, hostUsable: { type: 'number', value: 42 } });
});

test('mounted remote range appearance rejects unsafe colors atomically, then formats current numeric display and restores one Worker history step', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const bytes = Buffer.from(writeDocumentZip(scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 12.3 } },
    { row: 0, column: 1, value: { type: 'number', value: 999 }, formula: '=A1+1' }]).serialize())).toString('base64');
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  await page.getByText('设置单元格格式', { exact: true }).click(); await page.getByLabel('文字颜色', { exact: true }).fill('url(https://private.invalid)');
  await page.getByRole('button', { name: '应用单元格格式', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, errors: ['InvalidOperation'] });
  await page.getByLabel('文字颜色', { exact: true }).fill('#112233'); await page.getByLabel('数字格式代码', { exact: true }).fill('0.00');
  await page.getByLabel('字重', { exact: true }).selectOption('bold');
  await page.getByRole('button', { name: '应用单元格格式', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, source: { text: '12.3' }, current: { kind: 'value', value: { type: 'number', value: 13.3 } } });
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 2, history: { undoSteps: 0, redoSteps: 1 }, source: { text: '12.3' } });
  expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true });
});

test('mounted remote selected text fonts resolve native offsets inside the Worker and preserve rich source, draft protection and history', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode }); const s = createDocumentSession(scalarWorkbook([])), manifest = s.snapshot().manifest;
  s.commit(s.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [{ op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [
    { row: 0, column: 0, source: { kind: 'richText', html: '<span class="cell-default">=A&#13;<br>禰󠄀😀Z</span>' } },
  ] }] })); const bytes = Buffer.from(writeDocumentZip(s.serialize())).toString('base64'); s.dispose();
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  const before = await page.evaluate(() => window.workbookRemoteEditor.sourceCell(0, 0)), input = page.getByRole('textbox', { name: '单元格或公式输入', exact: true });
  await page.getByText('设置单元格格式', { exact: true }).click(); await page.getByLabel('字重', { exact: true }).selectOption('bold');
  await input.evaluate(element => { const field = element as HTMLTextAreaElement, from = field.value.indexOf('禰'); field.setSelectionRange(from, from + 5); field.dispatchEvent(new Event('select', { bubbles: true })); });
  await page.getByRole('button', { name: '应用选中文字字体', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  const cell = await page.evaluate(() => window.workbookRemoteEditor.sourceCell(0, 0)); expect(cell?.text).toBe('=A\r\n禰󠄀😀Z'); expect(cell?.html).toContain('workbook-run-');
  expect(await input.evaluate(element => { const field = element as HTMLTextAreaElement; return field.value.slice(field.selectionStart, field.selectionEnd); })).toBe('禰󠄀😀');
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.sourceCell(0, 0))).toEqual(before);
  await input.fill('draft'); await expect(page.getByRole('button', { name: '应用选中文字字体', exact: true })).toBeDisabled();
  await input.press('Escape'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true });
});

test('mounted remote sorting uses one Worker current batch and keeps headers, hidden rows, history and drafts', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const model = scalarWorkbook([
    { row: 0, column: 0, value: { type: 'text', value: 'Key' } }, { row: 0, column: 1, value: { type: 'text', value: 'Label' } },
    ...[3, 100, 1, 2].map((value, index) => ({ row: index + 1, column: 0, value: { type: 'number' as const, value: 999 }, formula: '=' + value })),
    ...['three', 'hidden', 'one', 'two'].map((value, index) => ({ row: index + 1, column: 1, value: { type: 'text' as const, value } })),
    { row: 5, column: 3, value: { type: 'number', value: 999 }, formula: '=A2' },
  ], metadata => { const sheet = metadata.workbook.sheets[0]!; sheet.defaultRowHeightCustom = true; sheet.rowDimensions = [{ from: 2, to: 2, hidden: true }]; });
  const bytes = Buffer.from(writeDocumentZip(model.serialize())).toString('base64');
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true }); await address.fill('A1:B5'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.getByLabel('首行为表头', { exact: true }).check(); await page.getByRole('button', { name: '升序排序', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => Promise.all([1, 2, 3, 4].map(row => window.workbookRemoteEditor.value(row, 1))))).toEqual(['one', 'hidden', 'two', 'three']);
  expect(await page.evaluate(() => window.workbookRemoteEditor.value(5, 3))).toBe('=A2');
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => Promise.all([1, 2, 3, 4].map(row => window.workbookRemoteEditor.value(row, 1))))).toEqual(['three', 'hidden', 'one', 'two']);
  const input = page.getByRole('textbox', { name: '单元格或公式输入', exact: true }); await input.fill('draft'); await expect(page.getByRole('button', { name: '降序排序', exact: true })).toBeDisabled();
  await input.press('Escape'); await page.evaluate(() => window.workbookRemoteEditor.settled()); await page.getByRole('button', { name: '降序排序', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => Promise.all([1, 2, 3, 4].map(row => window.workbookRemoteEditor.value(row, 1))))).toEqual(['three', 'hidden', 'two', 'one']);
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ history: { undoSteps: 1, redoSteps: 0 }, errors: [] });
  expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, hostUsable: { type: 'number', value: 42 } });
});
test('mounted remote workbook fill and scoped single/all replacements remain atomic Worker commands', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const bytes = Buffer.from(writeDocumentZip(scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 1 } }, { row: 1, column: 0, value: { type: 'number', value: 3 } },
    { row: 0, column: 2, value: { type: 'text', value: '甲猫😀猫' } }, { row: 1, column: 2, value: { type: 'text', value: '猫' } }, { row: 2, column: 2, value: { type: 'text', value: 'outside猫' } }]).serialize())).toString('base64');
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  await address.fill('A1:A5'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.getByRole('button', { name: '数列填充', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.value(4, 0))).toBe('9');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled()); expect(await page.evaluate(() => window.workbookRemoteEditor.value(4, 0))).toBe('');
  await address.fill('C1:C2'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.getByLabel('查找范围', { exact: true }).selectOption('selection'); await page.getByLabel('查找内容', { exact: true }).fill('猫');
  await page.getByRole('button', { name: '查找下一个', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await expect(page.getByRole('status', { name: '工作簿编辑状态', exact: true })).toContainText('匹配 1 / 3');
  await page.getByLabel('替换为', { exact: true }).fill('$&'); await page.getByRole('button', { name: '替换当前项', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.value(0, 2))).toBe('甲$&😀猫');
  await page.getByRole('button', { name: '查找下一个', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await expect(page.getByRole('status', { name: '工作簿编辑状态', exact: true })).toContainText('匹配 1 / 2');
  await page.getByRole('button', { name: '全部替换', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => Promise.all([window.workbookRemoteEditor.value(0, 2), window.workbookRemoteEditor.value(1, 2), window.workbookRemoteEditor.value(2, 2)]))).toEqual(['甲$&😀$&', '$&', 'outside猫']);
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ history: { undoSteps: 2 }, errors: [] });
  await page.evaluate(() => window.workbookRemoteEditor.finish());
});

test('mounted remote editor rejects a delayed address result after an external source revision', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const bytes = Buffer.from(writeDocumentZip(scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 1 } }]).serialize())).toString('base64');
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  await page.evaluate(() => window.workbookRemoteEditor.holdReference());
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true }); await address.fill('C5'); await address.press('Enter');
  await expect.poll(() => page.evaluate(() => window.workbookRemoteEditor.referenceHeld())).toBe(true);
  await page.evaluate(() => window.workbookRemoteEditor.external('external', false));
  await page.evaluate(() => window.workbookRemoteEditor.releaseReference()); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, source: { text: 'external' },
    selection: { baseRevision: 1, anchor: { ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] } }, errors: ['RevisionConflict'] });
  await expect(address).toHaveValue('A1'); await page.evaluate(() => window.workbookRemoteEditor.finish());
});

test('mounted remote merge requires explicit covered-content clearing and preserves both undo steps in its Worker', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const bytes = Buffer.from(writeDocumentZip(scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 7 } }, { row: 0, column: 1, value: { type: 'number', value: 3 } }]).serialize())).toString('base64');
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true }); await address.fill('A1:B1'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  const area = { from: { row: 0, column: 0 }, to: { row: 0, column: 1 } };
  await page.getByRole('button', { name: '合并单元格', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await expect(page.getByRole('alert')).toContainText('合并范围包含其他内容');
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, merges: [], errors: ['InvalidOperation'] });
  await page.getByRole('button', { name: '合并并清除覆盖内容', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, merges: [area], source: { text: '7' } });
  expect(await page.evaluate(() => window.workbookRemoteEditor.value(0, 1))).toBe('');
  await page.getByRole('button', { name: '取消合并', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 2, merges: [] });
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ merges: [area] });
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 4, merges: [], history: { undoSteps: 0, redoSteps: 2 }, source: { text: '7' } });
  expect(await page.evaluate(() => window.workbookRemoteEditor.value(0, 1))).toBe('3');
  expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true });
});

test('mounted remote freeze controls use the selected source address and the Worker history', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const bytes = Buffer.from(writeDocumentZip(scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 7 } }]).serialize())).toString('base64');
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  await address.fill('ALM10001'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.getByRole('button', { name: '冻结选区之前', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await expect(page.getByRole('alert')).toContainText('冻结区域将占满当前视口');
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 0, frozen: { rows: 0, columns: 0 }, history: { undoSteps: 0 }, errors: ['InvalidOperation'] });
  await address.fill('B2'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.getByRole('button', { name: '冻结选区之前', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 1, frozen: { rows: 1, columns: 1 }, history: { undoSteps: 1 }, source: { text: '7' }, errors: ['InvalidOperation'] });
  await page.getByRole('button', { name: '取消冻结', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 2, frozen: { rows: 0, columns: 0 } });
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ frozen: { rows: 1, columns: 1 } });
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 4, frozen: { rows: 0, columns: 0 }, history: { undoSteps: 0, redoSteps: 2 }, source: { text: '7' }, errors: ['InvalidOperation'] });
  expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true });
});

test('mounted remote axis controls move cross-chunk cells and formulas through one Worker with draft protection and four undo steps', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const bytes = Buffer.from(writeDocumentZip(scalarWorkbook([{ row: 1, column: 31, value: { type: 'number', value: 7 } }, { row: 2, column: 32, value: { type: 'number', value: 8 } },
    { row: 0, column: 1, value: { type: 'number', value: 999 }, formula: '=SUM(AF2:AG3)' }]).serialize())).toString('base64');
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  const input = page.getByRole('textbox', { name: '单元格或公式输入', exact: true }), address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  await input.fill('草稿'); await expect(page.getByRole('button', { name: '插入选区行', exact: true })).toBeDisabled();
  await input.press('Escape'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.getByLabel('插入行列的格式', { exact: true }).selectOption('none');
  await address.fill('AF2:AG3'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.getByRole('button', { name: '插入选区行', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await expect(input).toBeFocused();
  expect(await page.evaluate(() => Promise.all([window.workbookRemoteEditor.value(1, 31), window.workbookRemoteEditor.value(3, 31), window.workbookRemoteEditor.value(4, 32), window.workbookRemoteEditor.value(0, 1)])))
    .toEqual(['', '7', '8', '=SUM(AF4:AG5)']);
  await address.fill('AF4:AG5'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.getByRole('button', { name: '插入选区列', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await expect(input).toBeFocused();
  expect(await page.evaluate(() => Promise.all([window.workbookRemoteEditor.value(3, 33), window.workbookRemoteEditor.value(4, 34), window.workbookRemoteEditor.value(0, 1)])))
    .toEqual(['7', '8', '=SUM(AH4:AI5)']);
  await page.getByRole('button', { name: '删除选区列', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await address.fill('AF2:AG3'); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  await page.getByRole('button', { name: '删除选区行', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 4, history: { undoSteps: 4 }, current: { kind: 'value', value: { type: 'number', value: 15 } }, errors: [] });
  expect(await page.evaluate(() => Promise.all([window.workbookRemoteEditor.value(1, 31), window.workbookRemoteEditor.value(2, 32), window.workbookRemoteEditor.value(0, 1)])))
    .toEqual(['7', '8', '=SUM(AF2:AG3)']);
  for (let index = 0; index < 4; index++) { await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled()); }
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ revision: 8, history: { undoSteps: 0, redoSteps: 4 }, errors: [] });
  expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
});

test('mounted table row insertion fills only the new row, preserves a manual exception and restores native table source on undo', async ({ page }) => {
  await page.addScriptTag({ content: editorHostCode });
  const s = createDocumentSession(scalarWorkbook([{ row: 1, column: 0, value: { type: 'number', value: 3 } }, { row: 2, column: 0, value: { type: 'number', value: 4 } }]));
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0, operations: [
    { op: 'setWorkbookTable', sheetId: 'sheet-1', table: { tableId: 'sales', name: 'Sales', range: { from: { row: 0, column: 0 }, to: { row: 3, column: 1 } },
      header: true, totals: true, totalsPreviouslyShown: true, showFirstColumn: false, showLastColumn: false, rowStripes: true, columnStripes: false,
      columns: [{ columnId: 'amount', name: 'Amount', totalFunction: 'none' }, { columnId: 'double', name: 'Double', calculatedFormula: '=A2*2', totalFunction: 'sum' }] } },
    { op: 'setWorkbookCells', sheetId: 'sheet-1', cells: [{ row: 2, column: 1, source: { kind: 'value', value: { type: 'number', value: 99 } } }] },
  ] }));
  const bytes = Buffer.from(writeDocumentZip(s.serialize())).toString('base64'); s.dispose();
  await page.evaluate(({ code, bytes }) => window.workbookRemoteEditor.mount(code, bytes), { code: sessionWorkerCode, bytes });
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true });
  const choose = async (value: string) => { await address.fill(value); await address.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled()); };
  await choose('A3'); await page.getByRole('button', { name: '插入选区行', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => Promise.all([window.workbookRemoteEditor.value(2, 1), window.workbookRemoteEditor.value(3, 1)]))).toEqual(['=(A3*2)', '99']);
  await choose('A3'); const input = page.getByRole('textbox', { name: '单元格或公式输入', exact: true }); await input.fill('5'); await input.press('Enter'); await page.evaluate(() => window.workbookRemoteEditor.settled());
  expect(await page.evaluate(() => window.workbookRemoteEditor.sourceCell(2, 1))).toMatchObject({ formula: { expression: '=(A3*2)' } });
  for (let i = 0; i < 2; i++) { await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookRemoteEditor.settled()); }
  expect(await page.evaluate(() => Promise.all([window.workbookRemoteEditor.value(1, 1), window.workbookRemoteEditor.value(2, 1)]))).toEqual(['=A2*2', '99']);
  expect(await page.evaluate(() => window.workbookRemoteEditor.inspect())).toMatchObject({ history: { undoSteps: 0, redoSteps: 2 }, errors: [] });
  expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
});

test('100 ordinary edits and complete undo/redo retain bounded current previews on 200,000 actual cells', { tag: '@scale' }, async ({ page }, info) => {
  test.setTimeout(360_000); await page.addScriptTag({ content: performanceHostCode });
  const result = await page.evaluate(({ code, bytes }) => window.workbookEditingPerformance.ordinaryWorkerEdits(code, bytes), {
    code: sessionWorkerCode, bytes: readFileSync('examples/xlsx/15-ten-thousand-rows.xlsx').toString('base64') });
  const sha256 = (input: string | Uint8Array) => createHash('sha256').update(input).digest('hex');
  const report = JSON.stringify({ ...result, sourceSha256: sha256(readFileSync('examples/xlsx/15-ten-thousand-rows.xlsx')),
    workerBundleSha256: sha256(sessionWorkerCode), hostBundleSha256: sha256(performanceHostCode) }, null, 2); writeFileSync(info.outputPath('workbook-ordinary-editing.json'), report + '\n');
  await info.attach('workbook-ordinary-editing.json', { body: report, contentType: 'application/json' });
  expect(result.samples).toHaveLength(100); expect(result.maximumReadRow).toBeLessThan(1000);
  expect(result).toMatchObject({ sameFrame: true, originalChunksRestored: true, restoredInput: '22', failures: [], fullHistory: { undoSteps: 100, redoSteps: 0 },
    undoneHistory: { undoSteps: 0, redoSteps: 100 }, redoneHistory: { undoSteps: 100, redoSteps: 0 },
    audit: { cells: 200_000, formulas: 20_000, sessionsCreated: 1, revision: 400, endpointDisposed: true, reopenedEqual: true, history: { undoSteps: 0, redoSteps: 100 } } });
  expect(result.fullHistory.retainedBytes).toBeLessThan(512 * 1024 * 1024); expect(result.largestReplyUnits).toBeLessThan(2 * 1024 * 1024);
  // Record the complete edit→calculation→preview distribution. A non-pinned
  // host must not silently certify the 4-CPU/8-GiB stage's 500ms p95 gate.
});
test('mounted editor displays 100 current dependent results and full history on 200,000 actual cells', { tag: '@scale' }, async ({ page }, info) => {
  test.setTimeout(360_000); const mountedPerformanceHostCode = await bundle('workbook-editor-performance'); await page.addScriptTag({ content: mountedPerformanceHostCode });
  const result = await page.evaluate(({ code, bytes }) => window.workbookEditorPerformance.ordinaryMountedWorkerEdits(code, bytes), {
    code: sessionWorkerCode, bytes: readFileSync('examples/xlsx/15-ten-thousand-rows.xlsx').toString('base64') });
  const sha256 = (input: string | Uint8Array) => createHash('sha256').update(input).digest('hex');
  const report = JSON.stringify({ ...result, sourceSha256: sha256(readFileSync('examples/xlsx/15-ten-thousand-rows.xlsx')),
    workerBundleSha256: sha256(sessionWorkerCode), hostBundleSha256: sha256(mountedPerformanceHostCode) }, null, 2); writeFileSync(info.outputPath('workbook-mounted-editing.json'), report + '\n');
  await info.attach('workbook-mounted-editing.json', { body: report, contentType: 'application/json' });
  expect(result.samples).toHaveLength(100); expect(result.maximumReadRow).toBeLessThan(1000);
  expect(result).toMatchObject({ sameFrame: true, originalChunksRestored: true, restoredInput: '22', undoneText: '492', redoneText: '669', restoredText: '492', finalRevision: 400, failures: [],
    fullHistory: { undoSteps: 100, redoSteps: 0 }, undoneHistory: { undoSteps: 0, redoSteps: 100 }, redoneHistory: { undoSteps: 100, redoSteps: 0 },
    audit: { cells: 200_000, formulas: 20_000, sessionsCreated: 1, revision: 400, endpointDisposed: true, reopenedEqual: true, history: { undoSteps: 0, redoSteps: 100 } } });
  expect(result.fullHistory.retainedBytes).toBeLessThan(512 * 1024 * 1024); expect(result.largestReplyUnits).toBeLessThan(2 * 1024 * 1024);
  // This host is not pinned. Record every edit→calculation→visible formula
  // sample; final four-engine CPU/memory/p95 gates remain mandatory.
});

test.beforeEach(async ({ page }) => { await openSdkHost(page); await page.addScriptTag({ content: hostCode }); });
test('native XLSX pictures and charts stay in a host Worker and feed the public preview through bounded reads', async ({ page }) => {
  for (const name of ['07-three-charts', '08-images-shapes']) {
    const result = await page.evaluate(({ code, bytes, name }) => window.workbookReadHost.nativeWorkerRead(code, bytes, name, false), { code: workerCode, bytes: readFileSync(`examples/xlsx/${name}.xlsx`).toString('base64'), name });
    expect(result).toMatchObject({ frozen: true, retained: true, detachedInput: true, failures: [] });
    expect(['InvalidOperation', 'OperationCancelled']).toContain(result.closed); expect(result.progress).toBeGreaterThan(0);
    expect(result.first.objectFragmentCount).toBeGreaterThan(0); expect(result.first.documentId).toBe(name); expect(result.largestReplyUnits).toBeLessThan(1024 * 1024);
  }
});
test('native conditional and validation states reach the public preview through the host Worker', async ({ page }) => {
  const result = await page.evaluate(({ code, bytes }) => window.workbookReadHost.nativeWorkerRead(code, bytes, 'worker-rules', false), { code: workerCode, bytes: readFileSync('examples/xlsx/06-rules-validation.xlsx').toString('base64') });
  expect(result).toMatchObject({ frozen: true, retained: true, failures: [], first: { diagnostics: [expect.objectContaining({ code: 'validation-invalid', row: 2, column: 4 })] } });
  expect(result.ruleStates!.cells.map(cell => cell.validation?.state)).toEqual(['valid', 'invalid', 'valid', 'valid', 'valid', 'valid', 'valid']);
  expect(result.largestReplyUnits).toBeLessThan(1024 * 1024);
});
test('10,000 actual XLSX rows open and scroll through the production read transport without moving full cell chunks to the main thread', { tag: '@scale' }, async ({ page, browserName }, info) => {
  test.setTimeout(120_000);
  const bytes = readFileSync('examples/xlsx/15-ten-thousand-rows.xlsx').toString('base64');
  const profile = process.env.DOFFI_WORKBOOK_CPU_PROFILE && browserName === 'chromium' ? await page.context().newCDPSession(page) : undefined;
  if (profile) { await profile.send('Profiler.enable'); await profile.send('Profiler.start'); }
  const result = await page.evaluate(({ code, bytes }) => window.workbookReadHost.nativeWorkerRead(code, bytes, 'ten-thousand-native', true), { code: workerCode, bytes });
  if (profile) { writeFileSync(info.outputPath('workbook-main.cpuprofile'), JSON.stringify((await profile.send('Profiler.stop')).profile)); await profile.detach(); }
  expect(result.counts).toEqual({ kind: 'audit', cells: 200_000, formulas: 20_000, chunks: 40 });
  expect(result.last).toHaveLength(20); expect(result.last.at(-1)?.value).toEqual({ type: 'number', value: 3_400_152 });
  expect(result).toMatchObject({ frozen: true, retained: true, detachedInput: true, failures: [] });
  expect(result.last.filter(cell => cell.formula?.cacheState === 'imported')).toHaveLength(2);
  expect(result.afterScroll.formulaCaches.imported).toBeGreaterThan(0); expect(result.afterScroll.viewport.scrollTopPt).toBe(149_800);
  expect(result.first.storedCellCount).toBeLessThan(5000); expect(result.afterScroll.storedCellCount).toBeLessThan(5000);
  expect(result.snapshotUnits).toBeLessThan(100_000); expect(result.largestReplyUnits).toBeLessThan(1024 * 1024); expect(result.ticks).toBeGreaterThan(10);
  const cancelled = await page.evaluate(({ code, bytes }) => window.workbookReadHost.cancelNativeWorker(code, bytes), { code: workerCode, bytes });
  expect(cancelled.error).toBe('OperationCancelled'); expect(cancelled.progress).toBeGreaterThan(0);
  // Measurements are recorded; this host is not the final pinned 4-vCPU/8-GiB
  // acceptance environment, so a passing test does not certify that SLA.
  const report = JSON.stringify({ ...result, cancelled }, null, 2);
  writeFileSync(info.outputPath('native-workbook-worker.json'), report + '\n');
  await info.attach('native-workbook-worker.json', { body: report, contentType: 'application/json' });
});

test('native workbook edits, cross-sheet copy, current values, cancellation and history stay in one host Worker session', async ({ page }) => {
  await page.addScriptTag({ content: sessionHostCode });
  for (const name of ['01-values-text', '04-formulas-names']) {
    const result = await page.evaluate(({ code, bytes, name }) => window.workbookSessionHost.editWorker(code, bytes, name, false, true), {
      code: sessionWorkerCode, bytes: readFileSync(`examples/xlsx/${name}.xlsx`).toString('base64'), name });
    expect(result).toMatchObject({ isolated: true, cancelled: 'OperationCancelled', cancellationPoint: 256, copyUndo: true, originalChunksRestored: true,
      detachedInput: true, failures: [], cancellationState: { revision: 1, history: { undoSteps: 1 } },
      current: { kind: 'value', value: { type: 'number', value: 2 } }, audit: { sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true,
        revision: 4, history: { undoSteps: 0, redoSteps: 2 }, hostUsable: { type: 'number', value: 42 } } });
    expect(result.largestControlReplyUnits).toBeLessThan(64 * 1024); expect(result.largestViewReplyUnits).toBeLessThan(2 * 1024 * 1024);
    expect(result.states).toEqual([1, 2, 3, 4]); expect(result.selection?.anchor).toMatchObject({ kind: 'cells', sheetId: result.targetSheetId });
    expect(result.shown.sheetId).toBe(result.targetSheetId); expect(result.shown.sheetId).not.toBe(result.first.sheetId);
  }
});

test('200,000 actual cells copy and cancel through the editing Worker without transferring full source chunks', { tag: '@scale' }, async ({ page }, info) => {
  test.setTimeout(180_000); await page.addScriptTag({ content: sessionHostCode });
  const result = await page.evaluate(({ code, bytes }) => window.workbookSessionHost.editWorker(code, bytes, 'capacity-editing-worker', true), {
    code: sessionWorkerCode, bytes: readFileSync('examples/xlsx/15-ten-thousand-rows.xlsx').toString('base64') });
  expect(result).toMatchObject({ isolated: true, cancelled: 'OperationCancelled', cancellationPoint: 256, copyUndo: true, originalChunksRestored: true,
    detachedInput: true, failures: [], current: { kind: 'value', value: { type: 'number', value: 3_400_152 } },
    audit: { cells: 200_000, formulas: 20_000, sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true, revision: 4, history: { undoSteps: 0, redoSteps: 2 } } });
  expect(result.audit.chunks).toHaveLength(40); expect(result.copied).toBeGreaterThan(700); expect(result.prepared).toBeGreaterThan(800);
  expect(result.largestControlReplyUnits).toBeLessThan(2 * 1024 * 1024); expect(result.largestViewReplyUnits).toBeLessThan(2 * 1024 * 1024); expect(result.ticks).toBeGreaterThan(100);
  // This capacity integration keeps the real source counts and browser event
  // loop; final fixed-resource editing latency/history gates remain separate.
  const report = JSON.stringify(result, null, 2); writeFileSync(info.outputPath('workbook-editing-worker.json'), report + '\n');
  await info.attach('workbook-editing-worker.json', { body: report, contentType: 'application/json' });
});
