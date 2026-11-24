/** Stage-only capacity acceptance. Daily CI keeps the independent small cases. */
import assert from 'node:assert/strict';
import { spawnSync, execFileSync } from 'node:child_process';
import { globSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseArgs } from 'node:util';
import { runWorkbookCapacityBudget, workbookCapacityBrowsers } from './workbook-capacity-resources.mjs';

const { values: { part, browser } } = parseArgs({
  options: { part: { type: 'string', default: 'all' }, browser: { type: 'string' } },
});
if (!['all', 'native', 'browser', 'summary'].includes(part) ||
  (part === 'browser' ? !workbookCapacityBrowsers.includes(browser) : browser !== undefined))
  throw Error('Select all, native, browser with an engine, or summary.');
const pnpm = process.env.npm_execpath;
if (!pnpm) throw Error('Run capacity acceptance through pnpm.');
if (process.env.DOFFI_CI_MODE === 'daily') throw Error('Stage capacity acceptance cannot run in daily mode.');
const root = resolve(import.meta.dirname, '..'), directory = resolve(root, '.cache/workbook-scale-acceptance');
const sourceHead = process.env.DOFFI_SOURCE_HEAD ?? execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim();
const env = { ...process.env, DOFFI_SOURCE_HEAD: sourceHead };
mkdirSync(directory, { recursive: true });
const acceptance = resolve(directory, 'acceptance.json');
rmSync(acceptance, { force: true });
const writeReport = (file, data) => writeFileSync(resolve(directory, file), JSON.stringify({ sourceHead, ...data }, null, 2) + '\n');
const readReport = file => JSON.parse(readFileSync(resolve(directory, file), 'utf8'));
function run(command, args, environment = env) {
  const result = spawnSync(command, args, { cwd: root, stdio: 'inherit', env: environment });
  if (result.error) throw result.error;
  if (result.status !== 0) throw Error('Scale acceptance failed: ' + command + ' ' + args.join(' '));
}

function nativeCapacity() {
  rmSync(resolve(directory, 'native.json'), { force: true });
  for (const script of ['check:xlsx-fixtures', 'check:xlsx-reference', 'check:xlsx-export'])
    run(process.execPath, [pnpm, 'run', script]);
  run(process.execPath, ['scripts/workbook-performance-acceptance.mjs', '--prepare-only']);
  const unitReport = resolve(directory, 'unit-scale.json');
  const testFiles = globSync(['tests/**/*.test.ts', 'packages/*/tests/**/*.test.ts'], { cwd: root })
    .filter(file => readFileSync(resolve(root, file), 'utf8').includes('@scale')).sort();
  assert.ok(testFiles.length > 0, 'Capacity unit tests must be discovered.');
  run(process.execPath, [pnpm, 'exec', 'vitest', 'run', ...testFiles, '--testNamePattern=@scale',
    '--reporter=default', '--reporter=json', '--outputFile=' + unitReport]);
  const unit = readReport('unit-scale.json');
  assert.equal(unit.numFailedTests, 0);
  assert.ok(unit.numPassedTests >= 9, 'All current native and editing capacity unit cases must actually execute.');
  writeReport('native.json', { complete: true, unitTests: unit.numPassedTests });
}

function browserCapacity(engine) {
  const report = 'browser-' + engine + '.json';
  rmSync(resolve(directory, report), { force: true });
  if (engine === 'firefox') run(process.execPath, ['scripts/check-acceptance-fonts.mjs']);
  const files = ['workbook-storage', 'workbook-preview', 'workbook-read-worker', 'workbook-playground'].map(name => name + '.spec.ts');
  const legacy = engine === 'legacy-chromium', project = legacy ? 'chromium' : engine;
  run(process.execPath, ['scripts/playwright.mjs', 'test', ...files, '--grep=@scale|completes the public reading matrix',
    '--project=' + project, '--reporter=list', '--output=.cache/workbook-reading-matrix/' + engine],
  { ...env, ...(legacy ? { DOFFI_BROWSER_EXECUTABLE: resolve(root, '.cache/legacy-browser/chromium-1041/chrome-linux/chrome') } : {}) });
  const logicalCpus = runWorkbookCapacityBudget(engine, 'reading');
  assert.deepEqual(runWorkbookCapacityBudget(engine, 'editing'), logicalCpus);
  writeReport(report, { browser: engine, complete: true, logicalCpus });
}

function summarize() {
  const native = readReport('native.json');
  assert.equal(native.sourceHead, sourceHead);
  assert.equal(native.complete, true);
  assert.ok(native.unitTests >= 9);
  assert.equal(readReport('font-environment.json').sourceHead, sourceHead);
  const logicalCpus = {};
  for (const engine of workbookCapacityBrowsers) {
    const shard = readReport('browser-' + engine + '.json');
    assert.equal(shard.sourceHead, sourceHead);
    assert.equal(shard.browser, engine);
    assert.equal(shard.complete, true);
    logicalCpus[engine] = shard.logicalCpus;
  }
  for (const report of ['reading-matrix', 'reading-budget', 'editing-budget'])
    run(process.execPath, ['scripts/collect-workbook-' + report + '.mjs', '--output', resolve(directory, report + '.json')]);
  writeReport('acceptance.json', { complete: true, fullNativeReferences: true, fullNativeExport: true,
    fixedResourceEditing: true, logicalCpus, matrixCases: 128, cells: 200000, formulas: 20000, chunks: 40 });
}

if (part === 'all' || part === 'native') nativeCapacity();
if (part === 'all') for (const engine of workbookCapacityBrowsers) browserCapacity(engine);
if (part === 'browser') browserCapacity(browser);
if (part === 'all' || part === 'summary') summarize();
