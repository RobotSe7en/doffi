import { spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, rmSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseArgs } from 'node:util';
import { runWorkbookCapacityBudget, workbookCapacityBrowsers } from './workbook-capacity-resources.mjs';

if (process.env.DOFFI_CI_MODE === 'daily') throw Error('Editing capacity requires explicit stage acceptance.');
const { values } = parseArgs({ options: { 'prepare-only': { type: 'boolean' } } });
const root = resolve(import.meta.dirname, '..'), directory = resolve(root, '.cache/workbook-scale-acceptance');
mkdirSync(directory, { recursive: true }); rmSync(resolve(directory, 'editing-budget.json'), { force: true });
function run(command, args) {
  const result = spawnSync(command, args, { cwd: root, stdio: 'inherit' });
  if (result.error) throw result.error;
  if (result.status !== 0) throw Error('Editing capacity failed: ' + command + ' ' + args.join(' '));
}
run('python3', ['scripts/generate-workbook-editing-scale.py']);
const native = resolve(root, '.cache/workbook-editing-scale-schema'); mkdirSync(native, { recursive: true });
copyFileSync(resolve(root, 'packages/office/tests/fixtures/workbook-editing-scale.xlsx'), resolve(native, 'workbook-editing-scale.xlsx'));
const schemaPython = resolve(root, '.cache/xlsx-venv/bin/python');
run(existsSync(schemaPython) ? schemaPython : '/usr/bin/python3', ['scripts/check-xlsx-output.py', native]);
if (!values['prepare-only']) {
  for (const browser of workbookCapacityBrowsers) runWorkbookCapacityBudget(browser, 'editing');
  run(process.execPath, ['scripts/collect-workbook-editing-budget.mjs', '--output', resolve(directory, 'editing-budget.json')]);
}
