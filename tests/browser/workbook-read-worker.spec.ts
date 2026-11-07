import { openSdkHost } from './sdk-host.js';
import { expect, test } from '@playwright/test';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { createHash } from 'node:crypto';
import type * as fixture from './fixtures/workbook-read-host.js';
import type * as sessionFixture from './fixtures/workbook-session-host.js';
import type * as performanceFixture from './fixtures/workbook-editing-performance.js';
declare global { interface Window { workbookReadHost: typeof fixture } }
declare global { interface Window { workbookSessionHost: typeof sessionFixture } }
declare global { interface Window { workbookEditingPerformance: typeof performanceFixture } }
let workerCode = '', hostCode = '', sessionWorkerCode = '', sessionHostCode = '';
let performanceHostCode = '';
test.beforeAll(async () => {
  async function bundle(name: string): Promise<string> {
    const result = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL(`./fixtures/${name}.ts`, import.meta.url)), formats: ['iife'], name: 'WorkbookReadProbe' } } });
    if (!Array.isArray(result) && !('output' in result)) throw Error('Unexpected read-worker bundle');
    const output = Array.isArray(result) ? result.flatMap(part => part.output) : result.output;
    return output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  }
  workerCode = await bundle('workbook-read-worker'); hostCode = await bundle('workbook-read-host');
  sessionWorkerCode = await bundle('workbook-session-worker'); sessionHostCode = await bundle('workbook-session-host');
  performanceHostCode = await bundle('workbook-editing-performance');
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

test('native workbook edits, current values, cancellation and history stay in one host Worker session', async ({ page }) => {
  await page.addScriptTag({ content: sessionHostCode });
  for (const name of ['01-values-text', '04-formulas-names']) {
    const result = await page.evaluate(({ code, bytes, name }) => window.workbookSessionHost.editWorker(code, bytes, name), {
      code: sessionWorkerCode, bytes: readFileSync(`examples/xlsx/${name}.xlsx`).toString('base64'), name });
    expect(result).toMatchObject({ isolated: true, cancelled: 'OperationCancelled', cancellationPoint: 256, copyUndo: true, originalChunksRestored: true,
      detachedInput: true, failures: [], cancellationState: { revision: 1, history: { undoSteps: 1 } },
      current: { kind: 'value', value: { type: 'number', value: 2 } }, audit: { sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true,
        revision: 4, history: { undoSteps: 0, redoSteps: 2 }, hostUsable: { type: 'number', value: 42 } } });
    expect(result.largestControlReplyUnits).toBeLessThan(64 * 1024); expect(result.largestViewReplyUnits).toBeLessThan(2 * 1024 * 1024);
    expect(result.states).toEqual([1, 2, 3, 4]); expect(result.selection?.anchor.kind).toBe('cells');
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
