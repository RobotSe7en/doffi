import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { userInfo } from 'node:os';

export const workbookCapacityBrowsers = ['chromium', 'firefox', 'webkit', 'legacy-chromium'];

function cpuIds(list) {
  return list.split(',').flatMap(span => {
    const [from, to = from] = span.split('-').map(Number);
    return Array.from({ length: to - from + 1 }, (_, index) => from + index);
  });
}

/** Compare core grouping and available siblings independently of a runner's CPU numbering. */
export function workbookCapacityCpuTopology(resources) {
  const topology = resources.os.reportedCpuTopology, allowed = cpuIds(resources.cpuSet);
  assert.equal(topology.length, 4);
  assert.equal(allowed.length, 4);
  const cpus = topology.map(cpu => cpu.cpu).sort((a, b) => a - b);
  assert.deepEqual(cpus, [...new Set(allowed)].sort((a, b) => a - b));
  const sockets = new Map();
  for (const { cpu, socket, core, siblings } of topology) {
    assert.ok(socket && core);
    const threads = cpuIds(siblings).sort((a, b) => a - b);
    assert.ok(threads.includes(cpu));
    if (!sockets.has(socket)) sockets.set(socket, new Map());
    const cores = sockets.get(socket);
    if (!cores.has(core)) cores.set(core, { selected: [], threads });
    const group = cores.get(core);
    assert.deepEqual(group.threads, threads);
    group.selected.push(cpu);
  }
  return [...sockets.values()].map(cores =>
    [...cores.values()].map(({ selected, threads }) => {
      assert.deepEqual(selected.slice().sort((a, b) => a - b), threads.filter(cpu => allowed.includes(cpu)));
      return selected.length + '/' + threads.length;
    }).sort().join(',')).sort();
}

export function runWorkbookCapacityBudget(browser, mode) {
  if (!workbookCapacityBrowsers.includes(browser) || !['reading', 'editing'].includes(mode))
    throw Error('Select a supported capacity browser and mode.');
  const root = resolve(import.meta.dirname, '..');
  const allowed = /^Cpus_allowed_list:\s*(.+)$/m.exec(readFileSync('/proc/self/status', 'utf8'))?.[1];
  if (!allowed) throw Error('Capacity acceptance requires Linux CPU affinity.');
  const cpus = cpuIds(allowed);
  if (cpus.length < 4) throw Error('Capacity acceptance requires at least four available logical CPUs.');
  const topology = new Map(execFileSync('lscpu', ['-p=CPU,CORE,SOCKET'], { encoding: 'utf8' }).split('\n')
    .filter(row => row && !row.startsWith('#')).map(row => {
      const [cpu, core, socket] = row.split(',');
      return [Number(cpu), socket + '/' + core];
    }));
  const selected = [], cores = new Set();
  for (const cpu of cpus) {
    const core = topology.get(cpu) ?? cpu;
    if (!cores.has(core)) { selected.push(cpu); cores.add(core); }
    if (selected.length === 4) break;
  }
  for (const cpu of cpus) {
    if (selected.length === 4) break;
    if (!selected.includes(cpu)) selected.push(cpu);
  }
  const { uid, gid } = userInfo();
  const environment = ['PATH', 'FONTCONFIG_FILE', 'FONTCONFIG_PATH', 'XDG_DATA_HOME', 'LD_LIBRARY_PATH', 'PLAYWRIGHT_BROWSERS_PATH']
    .filter(name => process.env[name] !== undefined).map(name => '--setenv=' + name + '=' + process.env[name]);
  const result = spawnSync('sudo', ['-n', 'systemd-run', '--wait', '--pipe', '--collect',
    '--unit=doffi-' + mode + '-' + browser + '-' + process.pid, '--uid=' + uid, '--gid=' + gid,
    '--working-directory=' + root, ...environment, '-p', 'MemoryMax=8G', '-p', 'MemorySwapMax=0',
    '-p', 'CPUQuota=400%', '-p', 'AllowedCPUs=' + selected.join(','),
    process.execPath, 'scripts/workbook-reading-budget.mjs', '--browser', browser, '--mode', mode],
  { cwd: root, stdio: 'inherit' });
  if (result.error) throw result.error;
  if (result.status !== 0) throw Error('Fixed-resource ' + mode + ' capacity failed for ' + browser + '.');
  return selected;
}
