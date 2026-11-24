import { parseArgs } from 'node:util';
import { resolve as outputResolve } from 'node:path';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { workbookCapacityCpuTopology } from './workbook-capacity-resources.mjs';
const { values: outputOptions } = parseArgs({ options: { record: { type: 'boolean' }, output: { type: 'string' } } });
if (Boolean(outputOptions.record) === Boolean(outputOptions.output)) throw Error('Use exactly one of --record or --output <file> to save verified evidence.');
const root = fileURLToPath(new URL('../', import.meta.url));
const sha = value => createHash('sha256').update(value).digest('hex');
const digest = path => sha(readFileSync(join(root, path)));
const p95 = samples => [...samples].sort((a, b) => a - b)[Math.ceil(samples.length * .95) - 1];
async function bundle(name) {
  const result = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: join(root, 'tests/browser/fixtures/' + name + '.ts'), formats: ['iife'], name: 'WorkbookBudget' } } });
  if (!Array.isArray(result) && !('output' in result)) throw Error('Unexpected workbook acceptance bundle.');
  const output = Array.isArray(result) ? result.flatMap(part => part.output) : result.output;
  return sha(output.filter(part => part.type === 'chunk').map(part => part.code).join('\n'));
}
const worker = await bundle('workbook-read-worker'), host = await bundle('workbook-reading-budget');
const sourcePath = 'examples/xlsx/15-ten-thousand-rows.xlsx', runs = [];
for (const browser of ['chromium', 'firefox', 'webkit', 'legacy-chromium']) {
  const directory = join(root, '.cache/workbook-reading-budget', browser), output = join(directory, 'browser');
  const paths = readdirSync(output, { recursive: true }).filter(path => String(path).endsWith('/workbook-reading-budget.json'));
  assert.equal(paths.length, 1, 'Expected exactly one latest browser measurement.');
  const measurement = JSON.parse(readFileSync(join(output, paths[0]), 'utf8')), resources = JSON.parse(readFileSync(join(directory, 'resources.json'), 'utf8'));
  assert.equal(resources.browser, browser); assert.equal(resources.childExitCode, 0); assert.deepEqual(resources.invalidAffinity, []);
  assert.equal(resources.memoryMax, 8 * 1024 ** 3); assert.equal(resources.swapMax, '0'); assert.equal(resources.swapCurrent, '0');
  const [quota, period] = resources.cpuMax.split(' ').map(Number); assert.equal(quota / period, 4);
  assert.equal(resources.cpuSet, resources.affinity); assert.ok(resources.processes.length && resources.threadChecks > 0);
  assert.equal(resources.cpuSet.split(',').reduce((sum, span) => { const [from, to = from] = span.split('-').map(Number); return sum + to - from + 1; }, 0), 4);
  assert.ok(resources.browserSummedProcessHighWater > 0 && resources.browserSummedProcessHighWater <= 2 * 1024 ** 3);
  assert.equal(measurement.sourceSha256, digest(sourcePath)); assert.equal(measurement.workerBundleSha256, worker); assert.equal(measurement.hostBundleSha256, host);
  assert.deepEqual(measurement.counts, { kind: 'audit', cells: 200_000, formulas: 20_000, chunks: 40 }); assert.deepEqual(measurement.errors, []);
  assert.equal(measurement.instrumentation.host, 'standalone public-SDK preview'); assert.equal(measurement.instrumentation.cpuProfiler, false);
  for (const mode of ['trace', 'screenshots', 'video']) assert.equal(measurement.instrumentation[mode], 'off');
  assert.deepEqual(measurement.viewport, { width: 1280, height: 800 }); assert.deepEqual(measurement.grid, { widthPt: 880, heightPt: 450 });
  assert.equal(measurement.selections.length, 30); assert.equal(measurement.wheels.length, 40);
  assert.ok([...measurement.selections, ...measurement.wheels].every(value => Number.isFinite(value) && value >= 0));
  assert.equal(measurement.selectionP95Ms, p95(measurement.selections)); assert.equal(measurement.wheelP95Ms, p95(measurement.wheels));
  assert.ok(measurement.firstScreenMs > 0 && measurement.firstScreenMs <= 30_000);
  assert.ok(measurement.selectionP95Ms <= 100 && measurement.wheelP95Ms <= 120);
  assert.ok(measurement.initialFragments > 0 && measurement.maxFragments <= 5000);
  assert.equal(measurement.cancellation.error, 'OperationCancelled'); assert.equal(measurement.cancellation.retained, true);
  runs.push({ browser, measurement, resources });
}
const fonts = runs[0].resources.fontFiles;
const topology = workbookCapacityCpuTopology(runs[0].resources);
for (const run of runs) {
  assert.deepEqual(run.resources.fontFiles, fonts); assert.equal(run.resources.fontListSha256, runs[0].resources.fontListSha256);
  assert.deepEqual(workbookCapacityCpuTopology(run.resources), topology, run.browser + ' must use an equivalent CPU topology.');
}
const report = { scope: 'R3-08 fixed-resource reading performance subset; not the complete XLSX reading gate.',
  method: 'docs/workbook-reading-budget.md', source: { path: sourcePath, sha256: digest(sourcePath), generator: 'scripts/generate-xlsx-fixtures.py', generatorSha256: digest('scripts/generate-xlsx-fixtures.py') },
  scripts: Object.fromEntries(['scripts/workbook-reading-budget.mjs', 'scripts/collect-workbook-reading-budget.mjs', 'tests/browser/workbook-reading-performance.spec.ts'].map(path => [path, digest(path)])),
  fontFiles: fonts, runs: runs.map(({ browser, measurement, resources: { fontFiles: _fonts, ...resources } }) => ({ browser, measurement, resources })) };
writeFileSync(outputOptions.output ? outputResolve(outputOptions.output) : join(root, 'docs/evidence/workbook-reading-budget.json'), JSON.stringify(report, null, 2) + '\n');
process.stdout.write('Recorded four consistent, fully passing fixed-resource reading runs.\n');
