/** Acceptance-host preflight, outside every SDK package and document runtime. */
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..'), directory = resolve(root, '.cache/acceptance-fonts');
const manifest = JSON.parse(readFileSync(resolve(directory, 'environment.json'), 'utf8'));
const expected = JSON.parse(readFileSync(resolve(root, 'scripts/fixtures/acceptance-fonts.json'), 'utf8'));
if (!process.env.FONTCONFIG_FILE || resolve(process.env.FONTCONFIG_FILE) !== manifest.config) throw Error('Export the verified acceptance FONTCONFIG_FILE before capacity acceptance.');
for (const font of expected.packages.flatMap(pkg => pkg.fonts)) {
  const resolved = manifest.resolved.find(item => item.pattern === font.pattern);
  if (!resolved || resolved.sha256 !== font.sha256 || createHash('sha256').update(readFileSync(resolved.file)).digest('hex') !== font.sha256)
    throw Error('Acceptance font file differs from the pinned manifest: ' + font.pattern);
  const selected = execFileSync('fc-match', ['-f', '%{file}', font.pattern], { encoding: 'utf8' }).trim();
  if (resolve(selected) !== resolve(resolved.file)) throw Error('Fontconfig changed after the verified font installation: ' + font.pattern);
}
process.env.PLAYWRIGHT_BROWSERS_PATH ??= resolve(root, '.playwright');
const { firefox } = await import('@playwright/test');
const browser = await firefox.launch();
let report;
try {
  const page = await browser.newPage();
  report = await page.evaluate(() => {
    const context = document.createElement('canvas').getContext('2d');
    if (!context) throw Error('Acceptance requires actual browser canvas font measurement.');
    const sample = '汉字中文，。AaWim0123', bases = ['monospace', 'serif', 'sans-serif'];
    const measure = (family, weight) => {
      const widths = bases.map(base => {
        context.font = weight + ' 72px ' + base; const fallback = context.measureText(sample).width;
        context.font = weight + ' 72px ' + JSON.stringify(family) + ',' + base;
        return { base, fallback, selected: context.measureText(sample).width };
      });
      return { family, weight, widths, available: widths.some(value => Math.abs(value.selected - value.fallback) > .01) };
    };
    return { userAgent: navigator.userAgent, faces: ['Noto Sans CJK SC', 'Noto Serif CJK SC'].flatMap(family => ['400', '700'].map(weight => measure(family, weight))),
      negative: measure('DOFFI deliberately unavailable font', '400') };
  });
} finally { await browser.close(); }
const output = resolve(root, '.cache/workbook-scale-acceptance'); mkdirSync(output, { recursive: true });
writeFileSync(resolve(output, 'font-environment.json'), JSON.stringify({ sourceHead: process.env.DOFFI_SOURCE_HEAD ?? null, config: manifest.config,
  installation: manifest.installation, resolved: manifest.resolved, ...report }, null, 2) + '\n');
if (report.faces.some(face => !face.available) || report.negative.available) throw Error('Firefox cannot use the pinned acceptance fonts; fc-match alone does not prove browser availability.');
console.log('Pinned font bytes and four actual Firefox CJK faces verified; unavailable-face negative retained.');
