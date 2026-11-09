/** Stage-only capacity acceptance. Daily CI keeps the independent small cases. */
import { spawnSync, execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { userInfo } from 'node:os';

const root = resolve(import.meta.dirname, '..'), directory = resolve(root, '.cache/workbook-scale-acceptance');
mkdirSync(directory, { recursive: true });
// A failed local rerun must not leave a previous success as its latest result.
const acceptance = resolve(directory, 'acceptance.json');
rmSync(acceptance, { force: true });
function run(command, args, env = process.env) {
  const result = spawnSync(command, args, { cwd: root, stdio: 'inherit', env });
  if (result.error) throw result.error;
  if (result.status !== 0) throw Error('Scale acceptance failed: ' + command + ' ' + args.join(' '));
}
const pnpm = process.env.npm_execpath;
if (!pnpm) throw Error('Run capacity acceptance through pnpm.');
if (process.env.DOFFI_CI_MODE === 'daily') throw Error('Stage capacity acceptance cannot run in daily mode.');
run(process.execPath, ['scripts/check-acceptance-fonts.mjs']);
run(process.execPath, [pnpm, 'run', 'check:xlsx-fixtures']);
run(process.execPath, [pnpm, 'run', 'check:xlsx-reference']);
const unitReport = resolve(directory, 'unit-scale.json');
run(process.execPath, [pnpm, 'exec', 'vitest', 'run', '--testNamePattern=@scale', '--reporter=default', '--reporter=json', '--outputFile=' + unitReport]);
const unit = JSON.parse(readFileSync(unitReport, 'utf8'));
if (unit.numFailedTests !== 0 || unit.numPassedTests < 8) throw Error('The six native-capacity and two editing-capacity unit cases must actually execute.');
const files = ['workbook-storage', 'workbook-preview', 'workbook-read-worker', 'workbook-playground'].map(name => name + '.spec.ts');
const selection = '@scale|completes the public reading matrix';
run(process.execPath, ['scripts/playwright.mjs', 'test', ...files, '--grep=' + selection, '--reporter=list', '--output=.cache/workbook-reading-matrix-modern']);
run(process.execPath, ['scripts/playwright.mjs', 'test', ...files, '--grep=' + selection, '--project=chromium', '--reporter=list', '--output=.cache/workbook-reading-matrix-legacy'],
  { ...process.env, DOFFI_BROWSER_EXECUTABLE: resolve(root, '.cache/legacy-browser/chromium-1041/chrome-linux/chrome') });
run(process.execPath, ['scripts/collect-workbook-reading-matrix.mjs', '--output', resolve(directory, 'reading-matrix.json')]);

// Use four available logical CPUs, preferring distinct guest-reported cores.
const allowed = /^Cpus_allowed_list:\s*(.+)$/m.exec(readFileSync('/proc/self/status', 'utf8'))?.[1];
if (!allowed) throw Error('Capacity acceptance requires Linux CPU affinity.');
const cpus = allowed.split(',').flatMap(span => {
  const [start, end = start] = span.split('-').map(Number);
  return Array.from({ length: end - start + 1 }, (_, index) => start + index);
});
if (cpus.length < 4) throw Error('Capacity acceptance requires at least four available logical CPUs.');
const topology = execFileSync('lscpu', ['-p=CPU,CORE,SOCKET'], { encoding: 'utf8' }).split('\n').filter(line => line && !line.startsWith('#')).map(line => line.split(','));
const selected = [], cores = new Set();
for (const cpu of cpus) {
  const row = topology.find(row => Number(row[0]) === cpu), key = row ? row[2] + '/' + row[1] : String(cpu);
  if (!cores.has(key)) { selected.push(cpu); cores.add(key); }
  if (selected.length === 4) break;
}
for (const cpu of cpus) { if (selected.length === 4) break; if (!selected.includes(cpu)) selected.push(cpu); }
const identity = userInfo(), environment = ['PATH', 'FONTCONFIG_FILE', 'FONTCONFIG_PATH', 'XDG_DATA_HOME', 'LD_LIBRARY_PATH', 'PLAYWRIGHT_BROWSERS_PATH']
  .filter(name => process.env[name] !== undefined).map(name => '--setenv=' + name + '=' + process.env[name]);
for (const browser of ['chromium', 'firefox', 'webkit', 'legacy-chromium']) {
  run('sudo', ['-n', 'systemd-run', '--wait', '--pipe', '--collect', '--unit=doffi-scale-' + browser + '-' + process.pid,
    '--uid=' + identity.uid, '--gid=' + identity.gid, '--working-directory=' + root, ...environment,
    '-p', 'MemoryMax=8G', '-p', 'MemorySwapMax=0', '-p', 'CPUQuota=400%', '-p', 'AllowedCPUs=' + selected.join(','),
    process.execPath, 'scripts/workbook-reading-budget.mjs', '--browser', browser]);
}
run(process.execPath, ['scripts/collect-workbook-reading-budget.mjs', '--output', resolve(directory, 'reading-budget.json')]);
run(process.execPath, ['scripts/workbook-performance-acceptance.mjs']);
writeFileSync(acceptance, JSON.stringify({ sourceHead: process.env.DOFFI_SOURCE_HEAD ?? execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim(),
  complete: true, fullNativeReferences: true, fixedResourceEditing: true, logicalCpus: selected, matrixCases: 128, cells: 200000, formulas: 20000, chunks: 40 }, null, 2) + '\n');
