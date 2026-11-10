import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { parseArgs } from 'node:util';
import { build } from 'vite';

const { values } = parseArgs({ options: { output: { type: 'string' }, record: { type: 'boolean' } } });
if (Boolean(values.output) === Boolean(values.record)) throw Error('Use exactly one of --record or --output <file>.');
const sha = input => createHash('sha256').update(input).digest('hex');
const digest = path => sha(readFileSync(path));
async function bundle(name) {
  const result = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: resolve('tests/browser/fixtures/' + name + '.ts'), formats: ['iife'], name: 'WorkbookEditingBudget' } } });
  return sha((Array.isArray(result) ? result.flatMap(part => part.output) : result.output).filter(part => part.type === 'chunk').map(part => part.code).join('\n'));
}
const worker = await bundle('workbook-session-worker'), host = await bundle('workbook-editor-performance'), runs = [];
const p95 = samples => [...samples].sort((a, b) => a - b)[Math.ceil(samples.length * .95) - 1];
const source = 'packages/office/tests/fixtures/workbook-editing-scale.xlsx';
for (const browser of ['chromium', 'firefox', 'webkit', 'legacy-chromium']) {
  const directory = resolve('.cache/workbook-editing-budget', browser), output = join(directory, 'browser');
  const paths = readdirSync(output, { recursive: true }).filter(path => String(path).endsWith('/workbook-editing-budget.json'));
  assert.equal(paths.length, 1, 'Expected exactly one latest editing measurement.');
  const measurement = JSON.parse(readFileSync(join(output, paths[0]), 'utf8')), resources = JSON.parse(readFileSync(join(directory, 'resources.json'), 'utf8'));
  assert.equal(resources.browser, browser); assert.equal(resources.mode, 'editing'); assert.equal(resources.childExitCode, 0);
  assert.deepEqual(resources.invalidAffinity, []); assert.equal(resources.memoryMax, 8 * 1024 ** 3);
  assert.equal(resources.swapMax, '0'); assert.equal(resources.swapCurrent, '0');
  const [quota, period] = resources.cpuMax.split(' ').map(Number); assert.equal(quota / period, 4);
  assert.equal(resources.cpuSet, resources.affinity); assert.ok(resources.processes.length && resources.threadChecks > 0);
  assert.equal(resources.cpuSet.split(',').reduce((sum, span) => { const [from, to = from] = span.split('-').map(Number); return sum + to - from + 1; }, 0), 4);
  assert.ok(resources.browserSummedProcessHighWater > 0 && resources.browserSummedProcessHighWater <= 2 * 1024 ** 3);
  assert.equal(measurement.sourceSha256, digest(source)); assert.equal(measurement.workerBundleSha256, worker); assert.equal(measurement.hostBundleSha256, host);
  assert.deepEqual(measurement.viewport, { width: 1280, height: 800 });
  assert.equal(measurement.instrumentation.host, 'mounted public Worker editor'); assert.equal(measurement.instrumentation.cpuProfiler, false);
  assert.equal(measurement.instrumentation.diagnostic, false);
  for (const name of ['trace', 'screenshots', 'video']) assert.equal(measurement.instrumentation[name], 'off');
  assert.equal(measurement.samples.length, 100); assert.deepEqual(measurement.failures, []);
  assert.ok(measurement.samples.every(sample => ['inputMs', 'previewMs', 'totalMs'].every(key => Number.isFinite(sample[key]) && sample[key] >= 0)));
  assert.equal(measurement.editAndPreviewP95Ms, p95(measurement.samples.map(sample => sample.totalMs)));
  assert.ok(measurement.firstScreenMs > 0 && measurement.firstScreenMs <= 30_000); assert.ok(measurement.editAndPreviewP95Ms <= 500);
  assert.equal(measurement.opening.calculated, 20000); assert.equal(measurement.opening.formulaResults, 20000);
  assert.ok(measurement.opening.calculationMs > 0 && measurement.opening.calculationMs <= 15000);
  assert.ok(measurement.maximumFragments > 0 && measurement.maximumFragments <= 5000);
  assert.equal(measurement.nativeFeedback.restored, true);
  for (const samples of [measurement.nativeFeedback.selections, measurement.nativeFeedback.inputs]) {
    assert.equal(samples.length, 30); assert.ok(samples.every(value => Number.isFinite(value) && value >= 0)); assert.ok(p95(samples) <= 100);
  }
  assert.equal(measurement.cancellations.length, 2);
  assert.deepEqual(measurement.cancellations.map(value => value.command), ['sort', 'transaction']);
  for (const value of measurement.cancellations) { assert.equal(value.error, 'OperationCancelled'); assert.ok(value.progress >= 256); assert.equal(value.unchanged, true); }
  assert.equal(measurement.sameFrame, true); assert.equal(measurement.originalChunksRestored, true);
  assert.equal(measurement.restoredInput, '22'); assert.equal(measurement.undoneText, '492'); assert.equal(measurement.redoneText, '669'); assert.equal(measurement.restoredText, '492');
  assert.equal(measurement.finalRevision, 400); assert.ok(measurement.maximumReadRow < 1000 && measurement.largestReplyUnits < 2 * 1024 ** 2);
  for (const [key, undo, redo] of [['fullHistory', 100, 0], ['undoneHistory', 0, 100], ['redoneHistory', 100, 0]]) {
    assert.equal(measurement[key].undoSteps, undo); assert.equal(measurement[key].redoSteps, redo);
  }
  assert.ok(measurement.fullHistory.retainedBytes < 512 * 1024 ** 2);
  for (const [key, expected] of Object.entries({ cells: 200000, formulas: 20000, objects: 3, sessionsCreated: 1, revision: 400, endpointDisposed: true, reopenedEqual: true })) assert.equal(measurement.audit[key], expected);
  assert.deepEqual(measurement.audit.hostUsable, { type: 'number', value: 42 });
  runs.push({ browser, measurement, resources });
}
for (const run of runs) {
  assert.deepEqual(run.resources.fontFiles, runs[0].resources.fontFiles); assert.equal(run.resources.fontListSha256, runs[0].resources.fontListSha256);
  assert.equal(run.resources.cpuSet, runs[0].resources.cpuSet); assert.deepEqual(run.resources.os.reportedCpuTopology, runs[0].resources.os.reportedCpuTopology);
}
const report = { scope: 'R3-09–16 fixed-resource mounted editing, current dependent results and 100-step history subset; other stage gates are separate.',
  source: { path: source, sha256: digest(source), generator: 'scripts/generate-workbook-editing-scale.py', generatorSha256: digest('scripts/generate-workbook-editing-scale.py'),
    inputs: ['examples/xlsx/15-ten-thousand-rows.xlsx', 'examples/xlsx/08-images-shapes.xlsx'].map(path => ({ path, sha256: digest(path) })) },
  scripts: Object.fromEntries(['scripts/workbook-reading-budget.mjs', 'scripts/collect-workbook-editing-budget.mjs', 'tests/browser/workbook-editing-performance.spec.ts'].map(path => [path, digest(path)])),
  fontFiles: runs[0].resources.fontFiles, runs: runs.map(({ browser, measurement, resources: { fontFiles: _files, ...resources } }) => ({ browser, measurement, resources })) };
writeFileSync(resolve(values.output ?? 'docs/evidence/workbook-editing-budget.json'), JSON.stringify(report, null, 2) + '\n');
process.stdout.write('Recorded four consistent fixed-resource editing runs.\n');
