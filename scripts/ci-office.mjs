/** Independent Office suites keep their existing checks and pinned producer. */
import { spawnSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
const suite = process.argv[2];
if (process.argv.length !== 3 || !['word', 'presentation', 'workbook'].includes(suite)) throw Error('Select word, presentation or workbook.');
const pnpm = process.env.npm_execpath;
if (!pnpm) throw Error('Run this script through pnpm.');
const script = name => [process.execPath, pnpm, 'run', name];
const python = name => ['/usr/bin/python3', 'scripts/' + name + '.py'];
const tasks = {
  word: [script('check:docx-export:built'), script('check:word-closure:built')],
  presentation: [script('check:pptx-fixtures'), script('check:pptx-reference'), script('check:pptx-export:built'), script('check:presentation-closure:built')],
  workbook: [script('check:xlsx-fixtures'), script('check:xlsx-reference'), script('check:workbook-calculation'), script('check:workbook-filter-calculation'), ...[
    'chart-formulas', 'row-appearance', 'hidden-row-heights', 'accounting-underlines', 'picture-rotation',
    'gradient-rotation', 'comment-mentions', 'comment-uids', 'dual-note-boxes', 'text-overflow',
    'chart-colors', 'chart-markers', 'chart-auto-markers', 'chart-grid', 'chart-axes', 'chart-finite-scales', 'chart-labels',
  ].map(name => python('check-xlsx-' + name))],
}[suite];
const records = [];
let status = 0;
for (const [command, ...args] of tasks) {
  const start = Date.now(); console.log('Acceptance:', [command, ...args].join(' '));
  const result = spawnSync(command, args, { stdio: 'inherit', env: process.env });
  if (result.error) throw result.error;
  status = result.status ?? 1;
  records.push({ command: [command, ...args], elapsedMs: Date.now() - start, exitCode: status });
  if (status) break;
}
const directory = resolve('.cache/ci-office', suite); mkdirSync(directory, { recursive: true });
writeFileSync(resolve(directory, 'execution.json'), JSON.stringify({ suite, sourceHead: process.env.DOFFI_SOURCE_HEAD ?? null, records }, null, 2) + '\n');
process.exitCode = status;
