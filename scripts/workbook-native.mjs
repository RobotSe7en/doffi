/** Development acceptance only: use the same archive for the UNO client/server. */
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { basename, dirname, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';

const root = resolve(import.meta.dirname, '..'), script = resolve(root, process.argv[2] ?? '');
if (dirname(script) !== resolve(root, 'scripts') || !/^check-workbook-[a-z-]+\.py$/.test(basename(script))) throw Error('Select a workbook native acceptance script.');
const archive = resolve(root, '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz');
if (createHash('sha256').update(readFileSync(archive)).digest('hex') !== 'f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c') throw Error('Pinned Office archive differs.');
const environment = { ...process.env };
for (const name of ['PYTHONHOME', 'PYTHONPATH', 'URE_BOOTSTRAP', 'UNO_PATH', 'UNO_TYPES', 'UNO_SERVICES']) delete environment[name];
const python = resolve(root, '.cache/acceptance-office/installed/opt/libreoffice25.8/program/python');
const result = spawnSync(python, [script, ...process.argv.slice(3)], { cwd: root, env: environment, stdio: 'inherit' });
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
