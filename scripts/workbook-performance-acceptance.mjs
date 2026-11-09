import { execFileSync, spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs';
import { resolve } from 'node:path';
import { userInfo } from 'node:os';

if (process.env.DOFFI_CI_MODE === 'daily') throw Error('Editing capacity requires explicit stage acceptance.');
const root = resolve(import.meta.dirname, '..'), directory = resolve(root, '.cache/workbook-scale-acceptance');
mkdirSync(directory, { recursive: true }); rmSync(resolve(directory, 'editing-budget.json'), { force: true });
const allowed = /^Cpus_allowed_list:\s*(.+)$/m.exec(readFileSync('/proc/self/status', 'utf8'))?.[1];
if (!allowed) throw Error('Editing capacity requires Linux CPU affinity.');
const cpus = allowed.split(',').flatMap(span => { const [from, to = from] = span.split('-').map(Number); return Array.from({ length: to - from + 1 }, (_, at) => from + at); });
if (cpus.length < 4) throw Error('Editing capacity requires at least four available CPUs.');
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
function run(command, args) {
  const result = spawnSync(command, args, { cwd: root, stdio: 'inherit', env: process.env });
  if (result.error) throw result.error;
  if (result.status !== 0) throw Error('Editing capacity failed: ' + command + ' ' + args.join(' '));
}
run('python3', ['scripts/generate-workbook-editing-scale.py']);
const native = resolve(root, '.cache/workbook-editing-scale-schema'); mkdirSync(native, { recursive: true });
copyFileSync(resolve(root, 'packages/office/tests/fixtures/workbook-editing-scale.xlsx'), resolve(native, 'workbook-editing-scale.xlsx'));
const schemaPython = resolve(root, '.cache/xlsx-venv/bin/python');
run(existsSync(schemaPython) ? schemaPython : '/usr/bin/python3', ['scripts/check-xlsx-output.py', native]);
for (const browser of ['chromium', 'firefox', 'webkit', 'legacy-chromium']) {
  run('sudo', ['-n', 'systemd-run', '--wait', '--pipe', '--collect', '--unit=doffi-editing-' + browser + '-' + process.pid,
    '--uid=' + identity.uid, '--gid=' + identity.gid, '--working-directory=' + root, ...environment,
    '-p', 'MemoryMax=8G', '-p', 'MemorySwapMax=0', '-p', 'CPUQuota=400%', '-p', 'AllowedCPUs=' + selected.join(','),
    process.execPath, 'scripts/workbook-reading-budget.mjs', '--mode', 'editing', '--browser', browser]);
}
run(process.execPath, ['scripts/collect-workbook-editing-budget.mjs', '--output', resolve(directory, 'editing-budget.json')]);
