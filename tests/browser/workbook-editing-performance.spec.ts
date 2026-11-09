import { expect, test } from '@playwright/test';
import { build } from 'vite';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type * as fixture from './fixtures/workbook-editor-performance.js';
declare global { interface Window { workbookEditorPerformance: typeof fixture } }

test.describe.configure({ retries: 0 });
test.use({ trace: 'off', screenshot: 'off', video: 'off' });
const editingProfile = process.env.DOFFI_WORKBOOK_CPU_PROFILE === 'editing';
const diagnostic = process.env.DOFFI_WORKBOOK_EDIT_PROBE === '1', iterations = diagnostic ? editingProfile ? 20 : 3 : 100;
test('pinned-resource mounted editor calculates 100 edits and retains full undo/redo on 200,000 actual cells', async ({ page, browserName }, info) => {
  test.skip(process.env.DOFFI_WORKBOOK_BUDGET !== '1', 'Run through the explicit 4-vCPU/8-GiB acceptance wrapper.');
  test.setTimeout(360_000);
  if (diagnostic) page.on('console', async value => console.log(value.text(), ...(await Promise.all(value.args().filter(arg => arg.toString() === 'JSHandle@object').map(arg => arg.jsonValue())))));
  const bundle = async (name: string) => {
    const result = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/' + name + '.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookEditingBudget' } } });
    if (!Array.isArray(result) && !('output' in result)) throw Error('Unexpected editing acceptance fixture bundle.');
    const output = Array.isArray(result) ? result.flatMap(part => part.output) : result.output;
    return output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  };
  const worker = await bundle('workbook-session-worker'), host = await bundle('workbook-editor-performance');
  const profile = process.env.DOFFI_WORKBOOK_CPU_PROFILE && browserName === 'chromium' ? await page.context().newCDPSession(page) : undefined;
  let workerSession: string | undefined, sequence = 0, pageProfile: unknown, workerProfile: unknown;
  const commands = new Map<number, { resolve(value: Record<string, unknown>): void; reject(error: Error): void }>();
  async function workerCommand(method: string): Promise<Record<string, unknown>> {
    const id = ++sequence, response = new Promise<Record<string, unknown>>((resolve, reject) => commands.set(id, { resolve, reject }));
    await profile!.send('Target.sendMessageToTarget', { sessionId: workerSession!, message: JSON.stringify({ id, method }) }); return response;
  }
  if (profile) {
    await profile.send('Profiler.enable'); if (!editingProfile) await profile.send('Profiler.start');
    profile.on('Target.receivedMessageFromTarget', event => {
      const message = JSON.parse(event.message), pending = commands.get(message.id);
      if (!pending) return; commands.delete(message.id);
      if (message.error) pending.reject(Error(JSON.stringify(message.error))); else pending.resolve(message.result);
    });
    page.on('worker', async () => {
      const targets = await profile.send('Target.getTargets'), target = targets.targetInfos.find(target => target.type === 'worker' && target.url.startsWith('blob:'));
      if (!target) throw Error('The editing Worker profiling target is missing.');
      workerSession = (await profile.send('Target.attachToTarget', { targetId: target.targetId, flatten: false })).sessionId;
      await workerCommand('Profiler.enable'); if (!editingProfile) await workerCommand('Profiler.start');
    });
    page.on('console', async value => {
      if (editingProfile && value.text().startsWith('Editing diagnostic sample 0 ')) {
        await profile.send('Profiler.start'); await workerCommand('Profiler.start'); return;
      }
      if (editingProfile ? !value.text().startsWith('Editing diagnostic edits complete') : !value.text().startsWith('Editing diagnostic ready')) return;
      pageProfile = (await profile.send('Profiler.stop')).profile; workerProfile = (await workerCommand('Profiler.stop')).profile;
    });
  }
  await page.route('**/workbook-editing-budget.html', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><html><head><meta charset="utf-8"><title>Workbook editing budget</title></head><body></body></html>' }));
  await page.goto('/workbook-editing-budget.html'); await page.addScriptTag({ content: host });
  const bytes = readFileSync('packages/office/tests/fixtures/workbook-editing-scale.xlsx');
  const running = page.evaluate(({ code, bytes, diagnostic, iterations }) => window.workbookEditorPerformance.ordinaryMountedWorkerEdits(code, bytes, { performance: true, cancelBulk: true, nativeFeedback: !diagnostic, ...(diagnostic ? { iterations } : {}) }), { code: worker, bytes: bytes.toString('base64'), diagnostic, iterations });
  void running.catch(() => {});
  if (!diagnostic) try {
    await Promise.race([expect.poll(() => page.evaluate(() => window.workbookEditorPerformance.feedbackReady()), { timeout: 90000 }).toBe(true), running.then(() => { throw Error('Capacity ended before native feedback.'); })]);
    for (let index = 0; index < 30; index++) {
      const point = await page.evaluate(index => window.workbookEditorPerformance.feedbackPoint(index), index); await page.mouse.click(point.x, point.y);
      await page.evaluate(index => window.workbookEditorPerformance.feedbackSample('selection', index), index);
    }
    await page.evaluate(() => window.workbookEditorPerformance.feedbackFocus());
    for (let index = 0; index < 30; index++) {
      await page.keyboard.press('ControlOrMeta+A'); await page.keyboard.press('Digit1');
      await page.evaluate(index => window.workbookEditorPerformance.feedbackSample('input', index), index);
    }
    await page.keyboard.press('Escape'); await page.evaluate(() => window.workbookEditorPerformance.feedbackComplete());
  } catch (error) { writeFileSync(info.outputPath('capacity-failure.json'), JSON.stringify(await page.evaluate(() => window.workbookEditorPerformance.capacityStatus()), null, 2)); await page.screenshot({ path: info.outputPath('native-feedback-failure.png') }); await page.evaluate(() => window.workbookEditorPerformance.capacityCancel()); await running.catch(() => {}); throw error; }
  const result = await running;
  if (profile) {
    const phase = editingProfile ? 'editing' : 'opening';
    writeFileSync(info.outputPath('workbook-' + phase + '-main.cpuprofile'), JSON.stringify(pageProfile)); writeFileSync(info.outputPath('workbook-' + phase + '-worker.cpuprofile'), JSON.stringify(workerProfile));
    await profile.detach();
  }
  const sha = (input: string | Uint8Array) => createHash('sha256').update(input).digest('hex');
  const report = { ...result, sourceSha256: sha(bytes), workerBundleSha256: sha(worker), hostBundleSha256: sha(host), viewport: { width: 1280, height: 800 },
    instrumentation: { host: 'mounted public Worker editor', trace: 'off', screenshots: 'off', video: 'off', cpuProfiler: Boolean(profile) || process.env.MOZ_PROFILER_STARTUP === '1', diagnostic } };
  writeFileSync(info.outputPath('workbook-editing-budget.json'), JSON.stringify(report, null, 2) + '\n');
  await info.attach('workbook-editing-budget.json', { body: JSON.stringify(report), contentType: 'application/json' });
  expect(result.samples).toHaveLength(iterations); expect(result.maximumReadRow).toBeLessThan(1000);
  expect(result).toMatchObject({ sameFrame: true, originalChunksRestored: true, restoredInput: '22', undoneText: '492', redoneText: String(569 + iterations), restoredText: '492', finalRevision: 4 * iterations, failures: [],
    fullHistory: { undoSteps: iterations, redoSteps: 0 }, undoneHistory: { undoSteps: 0, redoSteps: iterations }, redoneHistory: { undoSteps: iterations, redoSteps: 0 },
    audit: { cells: 200_000, formulas: 20_000, objects: 3, sessionsCreated: 1, revision: 4 * iterations, endpointDisposed: true, reopenedEqual: true,
      hostUsable: { type: 'number', value: 42 }, history: { undoSteps: 0, redoSteps: iterations } } });
  expect(result.fullHistory.retainedBytes).toBeLessThan(512 * 1024 * 1024); expect(result.largestReplyUnits).toBeLessThan(2 * 1024 * 1024);
  expect(result.firstScreenMs).toBeLessThanOrEqual(30_000); expect(result.editAndPreviewP95Ms).toBeLessThanOrEqual(500);
  expect(result.opening).toMatchObject({ calculated: 20000, formulaResults: 20000 }); expect(result.opening!.calculationMs).toBeLessThanOrEqual(15000);
  expect(result.cancellations).toEqual([{ command: 'sort', error: 'OperationCancelled', progress: expect.any(Number), unchanged: true }, { command: 'transaction', error: 'OperationCancelled', progress: expect.any(Number), unchanged: true }]);
  expect(result.cancellations.every(value => value.progress >= 256)).toBe(true);
  expect(result.maximumFragments).toBeGreaterThan(0); expect(result.maximumFragments).toBeLessThanOrEqual(5000);
  if (!diagnostic) {
    expect(result.nativeFeedback?.restored).toBe(true);
    for (const samples of [result.nativeFeedback!.selections, result.nativeFeedback!.inputs]) {
      expect(samples).toHaveLength(30); expect([...samples].sort((a, b) => a - b)[28]).toBeLessThanOrEqual(100);
    }
  }
});
