import { existsSync, readFileSync, writeFileSync, mkdirSync, readdirSync, readlinkSync, renameSync } from 'node:fs';
import { spawn, execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { parseArgs } from 'node:util';
import { resolve, join, basename } from 'node:path';
import { arch, platform, release, cpus as hostCpus } from 'node:os';
const { values } = parseArgs({ options: { browser: { type: 'string', default: 'chromium' }, mode: { type: 'string', default: 'reading' } } });
const browser = values.browser;
const mode = values.mode;
if (!['reading', 'editing'].includes(mode)) throw Error('Select reading or editing acceptance.');
const startedAt = new Date().toISOString();
if (!['chromium', 'firefox', 'webkit', 'legacy-chromium'].includes(browser)) throw Error('Select chromium, firefox, webkit or legacy-chromium.');
const legacy = browser === 'legacy-chromium';
const relative = readFileSync('/proc/self/cgroup', 'utf8').split('\n').find(line => line.startsWith('0::'))?.slice(3);
if (!relative) throw Error('A real cgroup v2 is required.');
const group = join('/sys/fs/cgroup', relative), read = name => readFileSync(join(group, name), 'utf8').trim();
const status = pid => readFileSync('/proc/' + pid + '/status', 'utf8');
const cpus = text => /Cpus_allowed_list:\s*(\S+)/.exec(text)?.[1];
const countCpus = list => list.split(',').reduce((sum, span) => { const [from, to = from] = span.split('-').map(Number); return sum + to - from + 1; }, 0);
const affinity = cpus(status('self')), [quota, period] = read('cpu.max').split(' ').map(Number), memoryMax = Number(read('memory.max'));
const cpuSet = read('cpuset.cpus.effective');
if (memoryMax !== 8 * 1024 ** 3 || quota / period !== 4 || !affinity || countCpus(affinity) !== 4 || cpuSet !== affinity || read('memory.swap.max') !== '0')
  throw Error('Run under MemoryMax=8G, CPUQuota=400%, AllowedCPUs with exactly four CPUs and no swap.');
const cpuIds = list => list.split(',').flatMap(span => { const [from, to = from] = span.split('-').map(Number); return Array.from({ length: to - from + 1 }, (_, index) => from + index); });
const allowed = new Set(cpuIds(affinity));
const boundedAffinity = mask => Boolean(mask && cpuIds(mask).every(cpu => allowed.has(cpu)));
const output = resolve('.cache/workbook-' + mode + '-budget', browser);
if (mode === 'editing' && existsSync(output)) {
  const archive = resolve('.cache/workbook-editing-budget-history', browser); mkdirSync(archive, { recursive: true });
  renameSync(output, join(archive, startedAt.replace(/[:.]/g, '-')));
}
mkdirSync(output, { recursive: true });
const fonts = execFileSync('fc-list', [':', 'file', 'family', 'style'], { encoding: 'utf8' }).split('\n').filter(Boolean).sort().join('\n');
const fontPaths = [...new Set(execFileSync('fc-list', ['-f', '%{file}\n'], { encoding: 'utf8' }).trim().split('\n'))].sort();
const fontFiles = fontPaths.map(path => { const bytes = readFileSync(path); return { file: basename(path), bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') }; });
const peaks = new Map(), executables = new Set(), badAffinity = new Set(); let rssPeak = 0, groupPeak = 0, samples = 0, threadChecks = 0;
function sample(checkThreads = false) {
  let total = 0;
  for (const pid of read('cgroup.procs').split('\n')) {
    try {
      const executable = readlinkSync('/proc/' + pid + '/exe');
      if (!/(?:chrome|chromium|firefox|MiniBrowser|WebKit|WPE)/i.test(executable ?? '')) continue;
      executables.add(executable);
      const current = status(pid), rss = Number(/VmRSS:\s*(\d+)/.exec(current)?.[1] ?? 0) * 1024, high = Number(/VmHWM:\s*(\d+)/.exec(current)?.[1] ?? 0) * 1024;
      total += rss; peaks.set(pid, { executable: basename(executable), high: Math.max(high, peaks.get(pid)?.high ?? 0) });
      if (!boundedAffinity(cpus(current))) badAffinity.add(pid + ':' + cpus(current));
      if (checkThreads) for (const tid of readdirSync('/proc/' + pid + '/task')) {
        try { const mask = cpus(readFileSync('/proc/' + pid + '/task/' + tid + '/status', 'utf8')); if (!boundedAffinity(mask)) badAffinity.add(pid + '/' + tid + ':' + mask); threadChecks++; } catch { /* An exited thread is no longer resident. */ }
      }
    } catch { /* A process may exit between the cgroup and proc reads. */ }
  }
  rssPeak = Math.max(rssPeak, total); groupPeak = Math.max(groupPeak, Number(read('memory.current'))); samples++;
}
const child = spawn(process.execPath, ['scripts/playwright.mjs', 'test', 'workbook-' + mode + '-performance.spec.ts', '--project=' + (legacy ? 'chromium' : browser), '--output=' + join(output, 'browser')],
  { stdio: 'inherit', env: { ...process.env, DOFFI_WORKBOOK_BUDGET: '1', ...(legacy ? { DOFFI_BROWSER_EXECUTABLE: resolve('.cache/legacy-browser/chromium-1041/chrome-linux/chrome') } : {}) } });
const timer = setInterval(() => sample(samples % 50 === 0), 20);
const code = await new Promise(resolve => { child.once('error', error => { process.stderr.write(String(error) + '\n'); resolve(1); }); child.once('exit', (code, signal) => resolve(code ?? (signal ? 1 : 0))); });
clearInterval(timer); sample(true);
const highSum = [...peaks.values()].reduce((sum, item) => sum + item.high, 0);
const report = { browser, mode, startedAt, completedAt: new Date().toISOString(), cgroup: relative, memoryMax, cpuMax: read('cpu.max'), affinity, cpuSet, swapMax: read('memory.swap.max'), swapCurrent: read('memory.swap.current'),
  memoryPeak: Number(read('memory.peak')), sampledGroupPeak: groupPeak, browserRssPeak: rssPeak, browserSummedProcessHighWater: highSum,
  samples, sampleIntervalMs: 20, threadChecks, invalidAffinity: [...badAffinity], processes: [...peaks.values()],
  fontconfigRecordCount: fonts.split('\n').length, fontListSha256: createHash('sha256').update(fonts).digest('hex'),
  fontFiles, fontCache: 'OS font files were read for hashing before a fresh browser/Worker document open.',
  method: 'All browser subprocesses in the imposed cgroup; workers share renderer RSS. Sum of observed per-process kernel VmHWM is a conservative bound over their nonconcurrent peaks and includes shared pages repeatedly. Processes/threads are sampled; short-lived processes can escape observation. Kernel cgroup memory.peak supplements the sampled record. Absolute browser bounds are stronger than document-incremental bounds.',
  childExitCode: code, node: process.version };
report.os = { platform: platform(), release: release(), arch: arch(), cpuModels: [...new Set(hostCpus().map(cpu => cpu.model))] };
report.os.reportedCpuTopology = cpuIds(affinity).map(cpu => {
  const topology = name => readFileSync('/sys/devices/system/cpu/cpu' + cpu + '/topology/' + name, 'utf8').trim();
  return { cpu, core: topology('core_id'), socket: topology('physical_package_id'), siblings: topology('thread_siblings_list') };
});
report.binaries = [...executables].sort().map(path => ({ file: basename(path), sha256: createHash('sha256').update(readFileSync(path)).digest('hex') }));
writeFileSync(join(output, 'resources.json'), JSON.stringify(report, null, 2) + '\n');
process.stdout.write(JSON.stringify({ browser, childExitCode: code, browserRssPeak: rssPeak, browserSummedProcessHighWater: highSum, memoryPeak: report.memoryPeak, invalidAffinity: report.invalidAffinity }) + '\n');
if (code || badAffinity.size || !peaks.size || highSum > 2 * 1024 ** 3) process.exitCode = 1;
