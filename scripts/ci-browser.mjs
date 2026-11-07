/** Preserve the full browser file matrix while isolating independent suites. */
import { spawnSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseArgs } from 'node:util';

const { values } = parseArgs({ options: {
  browser: { type: 'string' }, suite: { type: 'string' },
  mode: { type: 'string', default: 'daily' }, list: { type: 'boolean', default: false },
} });
const { browser, suite, mode } = values;
if (!['chromium', 'firefox', 'webkit', 'legacy-chromium'].includes(browser)
  || !['word', 'presentation', 'workbook'].includes(suite)
  || !['daily', 'full'].includes(mode)) throw Error('Select a supported browser, suite and daily/full mode.');
const legacy = browser === 'legacy-chromium', project = legacy ? 'chromium' : browser;
const files = {
  word: legacy ? ['workspace', 'pagination'] : ['workspace', 'interaction', 'workbench', 'text-editor', 'pagination', 'rendering', 'remote-access', 'word-closure'],
  presentation: ['presentation', 'presentation-editor', 'presentation-workbench', 'presentation-closure'],
  workbook: ['workbook-storage', 'workbook-grid', 'workbook-drawing', 'workbook-drawing-window', 'workbook-preview', 'workbook-input', 'workbook-read-worker', 'workbook-playground'],
}[suite].map(file => file + '.spec.ts');
const directory = resolve('.cache/ci-browser', browser, suite);
mkdirSync(directory, { recursive: true });
const env = { ...process.env, PLAYWRIGHT_JSON_OUTPUT_NAME: resolve(directory, 'report.json'),
  ...(legacy ? { DOFFI_BROWSER_EXECUTABLE: resolve('.cache/legacy-browser/chromium-1041/chrome-linux/chrome') } : {}) };
// CLI file filters are regexes: an unanchored workbench.spec.ts would also
// select presentation-workbench.spec.ts and duplicate it across suites.
const patterns = files.map(file => '(^|/)' + file.replaceAll('.', '\\.') + '$');
const args = ['scripts/playwright.mjs', 'test', ...patterns, '--project=' + project,
  '--output=' + resolve(directory, 'results'), ...(mode === 'daily' ? ['--grep-invert=@scale'] : []),
  ...(values.list ? ['--list'] : ['--reporter=list,json'])];
const started = Date.now();
console.log(JSON.stringify({ browser, project, suite, mode, files }));
const result = spawnSync(process.execPath, args, { stdio: 'inherit', env });
if (result.error) throw result.error;
let status = result.status ?? 1;
if (!status && !values.list && suite === 'presentation' && !legacy) {
  const reference = spawnSync('/usr/bin/python3', ['scripts/check-presentation-reference.py', '--browser', browser], { stdio: 'inherit', env });
  if (reference.error) throw reference.error;
  status = reference.status ?? 1;
}
if (!values.list) writeFileSync(resolve(directory, 'execution.json'), JSON.stringify({ browser, project, suite, mode, files,
  sourceHead: process.env.DOFFI_SOURCE_HEAD ?? null, elapsedMs: Date.now() - started, exitCode: status }, null, 2) + '\n');
process.exitCode = status;
