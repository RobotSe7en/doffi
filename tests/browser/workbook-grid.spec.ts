import { encodeTestPackage } from './package-transport.js';
import { openSdkHost } from './sdk-host.js';
import { expect, test } from "@playwright/test";
import { openDocument } from "@doffi/core";
import type { WorkbookDocumentMetadata } from "@doffi/core";
import { build } from "vite";
import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import { readXlsxRecords } from "../../packages/office/src/xlsx/read.js";
import { projectXlsxRecords } from "../../packages/office/tests/helpers/xlsx-record-model.js";
import type { WorkbookViewport } from "../../packages/render/src/workbook/viewport.js";
import type { WorkbookTextFontRun } from "../../packages/render/src/workbook/text-fonts.js";
import type { WorkbookBorderEdge } from "../../packages/render/src/workbook/border-view.js";
import type { WorkbookBorderKind } from "../../packages/render/src/workbook/border-priority.js";
import borderReference from "../../docs/evidence/xlsx-border-semantics.json" with { type: "json" };
import alignmentReference from "../../docs/evidence/xlsx-alignment-semantics.json" with { type: "json" };
import type { TextLayoutCase, textLayoutMeasurement } from "./fixtures/workbook-text-layout.js";
import type { borderSourceMeasurement } from "./fixtures/workbook-border-source.js";
import type { gridLineMeasurement } from "./fixtures/workbook-grid-lines.js";
import type { assetMeasurement } from "./fixtures/workbook-assets.js";
import { workbookRasterCases } from "../workbook-image-fixture.js";
import type { measureWorkbookFills } from "./fixtures/workbook-fill.js";
import fillReference from "../../docs/evidence/xlsx-fill-pattern-spec.json" with { type: "json" };
import appearanceReference from "../../docs/evidence/xlsx-cell-appearance.json" with { type: "json" };
import borderSpec from "../../docs/evidence/xlsx-border-spec.json" with { type: "json" };
import type { diagonalMeasurement } from "./fixtures/workbook-diagonals.js";
import rotationReference from "../../docs/evidence/xlsx-rotation-semantics.json" with { type: "json" };
import type { numberTextMeasurement, NumberTextCase } from "./fixtures/workbook-number-text.js";
import type { rotatedCellMeasurement, RotatedCellCase, shapedFillMeasurement } from "./fixtures/workbook-rotated-cell.js";
import rotatedCellsReference from "../../docs/evidence/xlsx-rotated-cells.json" with { type: "json" };

import { importXlsx } from '../../packages/office/src/xlsx-import.js';
import type { textCacheMeasurement } from './fixtures/workbook-text-cache.js';
import type { fontProbeMeasurement } from './fixtures/workbook-font-probe.js';
import type { autoRowMeasurement, autoRowSourceSafety, windowLayoutMeasurement } from './fixtures/workbook-window.js';

interface GridHost {
  fontProbeMeasurement: typeof fontProbeMeasurement;
  textCacheMeasurement: typeof textCacheMeasurement;
  windowLayoutMeasurement: typeof windowLayoutMeasurement;
  autoRowMeasurement: typeof autoRowMeasurement;
  autoRowSourceSafety: typeof autoRowSourceSafety;
  numberTextMeasurement(cases: readonly NumberTextCase[], locale?: string): ReturnType<typeof numberTextMeasurement>;
  rotatedCellMeasurement(cases: readonly RotatedCellCase[]): ReturnType<typeof rotatedCellMeasurement>;
  shapedFillMeasurement(): ReturnType<typeof shapedFillMeasurement>;
  diagonalMeasurement(css: string, classes: readonly string[]): ReturnType<typeof diagonalMeasurement>;
  measureWorkbookFills(css: string, classes: readonly string[]): ReturnType<typeof measureWorkbookFills>;
  assetMeasurement(base64: string, mimeType: string): ReturnType<typeof assetMeasurement>;
  borderSourceMeasurement(): ReturnType<typeof borderSourceMeasurement>;
  gridLineMeasurement(): ReturnType<typeof gridLineMeasurement>;
  textLayoutMeasurement(cases: readonly TextLayoutCase[], limits?: { readonly maxNodes?: number; readonly maxTextUnits?: number }): ReturnType<typeof textLayoutMeasurement>;
  render(viewport: WorkbookViewport, rtl?: boolean): Promise<boolean>;
  inspect(): { cellCount: number; borderLines: number; accessibleCells: number; rows: number[]; sourceUnchanged: boolean; merges: { pane?: string; text: string | null; clip: { x: number; y: number; width: number; height: number }; body: { x: number; y: number; width: number; height: number }; hit?: { row: number; column: number } }[] };
  failedPaint(): boolean; dispose(): void;
  numericMeasurement(): Promise<{ rounded: { kind: string; text?: string }; scientific: { kind: string; text?: string }; actual: number; predicted: number; overflow: boolean; fill?: { kind: string; text?: string } }>;
  columnFontMeasurement(): Promise<{ measured: { cssClass: string; maximumDigitWidthPx: number; fonts: { requestedFamily: string; substituted: boolean }[] }; fallback: { cssClass: string; maximumDigitWidthPx: number }; actual: number; unrelatedUnchanged: boolean; retainedProbes: number }>;
  colorMeasurement(): { values: (string | undefined)[]; errors: string[]; fills: string[]; fillErrors: string[]; unchanged: boolean };
  sourceStyleMeasurement(): Promise<{ text: string | null; font: string; size: string; italic: string; background: string; bold: string; foreground: string; plain: string; blockRow?: string; row?: string; bound: boolean; before: string; disposed: boolean; noPartial: boolean; retained: number; rejected: string[] }>;
  textFontMeasurement(): Promise<{ runs: readonly WorkbookTextFontRun[]; rich: readonly WorkbookTextFontRun[]; withBreak: readonly WorkbookTextFontRun[]; limited: boolean; unchanged: boolean; frozen: boolean }>;
  textViewMeasurement(): Promise<{ rich: TextRunView[]; mixed: TextRunView[]; decorated: TextRunView[]; decoration: { strike: string; strikeStyle: string; underline: string; underlineStyle: string; plain: string }; text: string | null; accountingTypes: string[]; failures: string[]; unchanged: boolean; frozen: boolean }>;
  borderViewMeasurement(edges: readonly WorkbookBorderEdge[]): Promise<{ strokes: { from: number[]; to: number[]; width: number; dash: number[]; cap: string | null }[]; difference: number; maskDifference: number; splitDifference: number; splitMaskDifference: number; painted: number; doubleRows: number[]; failures: string[]; rasterImages: string[]; rasterSizes: number[][] }>;
  textDirectionMeasurement(): { values: string[]; explicit: string[]; wholeStringDirection: string; retainedText: string | null; failures: string[]; remaining: number; ownerDirection: string };
}
interface TextRunView { text: string; from: number; to: number; family: string; weight: string; color: string; x: number; y: number; width: number; height: number }
declare global { interface Window { workbookGridHost: GridHost } }
let code = '';
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-grid-host.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookGridProbe' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw new Error('Unexpected workbook grid probe bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output;
  code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
});
test.beforeEach(async ({ page }) => { await openSdkHost(page); await page.addScriptTag({ content: code }); await page.waitForFunction(() => !!window.workbookGridHost); });

for (const rtl of [false, true]) test(`merged frozen fragments preserve full content geometry in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const report = await page.evaluate(async rtl => {
    await window.workbookGridHost.render({ widthPt: 100, heightPt: 80, scrollLeftPt: 40, scrollTopPt: 30 }, rtl);
    return window.workbookGridHost.inspect();
  }, rtl);
  expect(report.merges).toHaveLength(4); expect(report.cellCount).toBe(24); expect(report.accessibleCells).toBe(21);
  expect(report.rows).toEqual([...report.rows].sort((a, b) => a - b)); expect(report.sourceUnchanged).toBe(true);
  for (const cell of report.merges) {
    expect(cell.text).toBe('合并😀禰󠄀'); expect(cell.hit).toEqual({ row: 0, column: 0 });
    expect(cell.body.width).toBeCloseTo(140, 1); expect(cell.body.height).toBeCloseTo(70, 1);
  }
  const body = report.merges.find(cell => cell.pane === 'body')!;
  expect(body.clip.x).toBeCloseTo(rtl ? 0 : 20, 1); expect(body.clip.y).toBeCloseTo(20, 1); expect(body.clip.width).toBeCloseTo(80, 1);
  expect(body.body.x).toBeCloseTo(rtl ? 0 : -40, 1); expect(body.body.y).toBeCloseTo(-30, 1);
  expect(await page.evaluate(() => window.workbookGridHost.failedPaint())).toBe(true);
});

test('viewport DOM stays bounded at row 10,000 and the far worksheet corner', async ({ page }) => {
  const first = await page.evaluate(async () => { await window.workbookGridHost.render({ widthPt: 960, heightPt: 600, scrollLeftPt: 0, scrollTopPt: 0 }); return window.workbookGridHost.inspect(); });
  const middle = await page.evaluate(async () => { await window.workbookGridHost.render({ widthPt: 960, heightPt: 600, scrollLeftPt: 0, scrollTopPt: 99_970 }); return window.workbookGridHost.inspect(); });
  await expect(page.locator('[data-row="9999"][data-column="9"]')).toHaveText('一万行');
  const last = await page.evaluate(async () => { await window.workbookGridHost.render({ widthPt: 960, heightPt: 600, scrollLeftPt: Number.MAX_SAFE_INTEGER, scrollTopPt: Number.MAX_SAFE_INTEGER }); return window.workbookGridHost.inspect(); });
  await expect(page.locator('[data-row="1048575"][data-column="16383"]')).toHaveText('末格');
  for (const report of [first, middle, last]) {
    expect(report.cellCount).toBeLessThanOrEqual(5000); expect(report.sourceUnchanged).toBe(true);
    expect(report.borderLines).toBeGreaterThan(0); expect(report.borderLines).toBeLessThanOrEqual(report.cellCount * 4);
  }
  expect(last.cellCount).toBeLessThanOrEqual(48 * 60); expect(middle.cellCount).toBeLessThanOrEqual(48 * 60);
  await page.evaluate(() => window.workbookGridHost.dispose()); await expect(page.locator('#workbook-grid-probe')).toHaveCount(0);
});

test('General fitting and numeric fill agree with independently laid out fixed-font text', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookGridHost.numericMeasurement());
  expect(result.rounded).toMatchObject({ kind: 'text', text: '26' });
  expect(result.scientific).toMatchObject({ kind: 'text', text: '1.23457E+11' });
  expect(result.overflow).toBe(false); expect(result.fill?.text).toMatch(/^f+$/);
  expect(result.actual).toBeLessThanOrEqual(140.25); expect(Math.abs(result.actual - result.predicted)).toBeLessThan(.25);
});

test('whole grid lines preserve clipped hidden-axis and frozen-pane pixels while removing only segment seams', async ({ page }, info) => {
  const cases = await page.evaluate(() => window.workbookGridHost.gridLineMeasurement());
  const png = await page.locator('#workbook-grid-line-probe').screenshot({ path: info.outputPath('uniform-grid-lines.png') });
  const compared = await page.evaluate(async ({ bytes, cases }) => {
    const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' })), image = new Image();
    try {
      image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
      const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0); const data = context.getImageData(0, 0, image.width, image.height).data;
      return cases.map((item, index) => {
        let max = 0, awayFromJoin = 0, changed = 0;
        for (let y = 0; y < 150; y++) for (let x = 0; x < 240; x++) {
          const first = ((index * 150 + y) * image.width + x) * 4, second = first + 240 * 4;
          const difference = Math.max(...[0, 1, 2].map(channel => Math.abs(data[first + channel]! - data[second + channel]!)));
          max = Math.max(max, difference);
          if (difference > 1) {
            changed++;
            if (!item.endpoints.some(point => Math.abs(point.x / .75 - x) <= 1.5 && Math.abs(point.y / .75 - y) <= 1.5)) awayFromJoin++;
          }
        }
        return { max, awayFromJoin, changed };
      });
    } finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
  }, { bytes: [...png], cases });
  for (const [index, item] of cases.entries()) {
    expect(item.full).toBeLessThan(item.original / 2);
    // Joining solid segments removes their butt-end antialiasing seam. Every
    // other pixel must retain the previous uniform grid's geometry and color.
    expect(compared[index]!.max).toBeLessThanOrEqual(12); expect(compared[index]!.awayFromJoin).toBe(0);
  }
});

test('cached text keeps independent live UTF-16 projections, ownership, font invalidation and complete window budgets', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookGridHost.textCacheMeasurement());
  expect(result.offsets.map(part => part[2]).join('')).toBe('A😀\n禰󠄀B');
  expect(result.offsets[0]![0]).toBe(0); expect(result.offsets.at(-1)![1]).toBe('A😀\n禰󠄀B'.length);
  for (const key of ['projectionsOwned', 'same', 'isolated', 'detached', 'atomic', 'evicted', 'warningsRefreshed', 'bounded', 'invalidated', 'cleared', 'unchanged'] as const) expect(result[key], key).toBe(true);
  expect(result.failures).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded']);
});

test('owned font probes retain identical glyph diagnostics without repeating raster readbacks', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookGridHost.fontProbeMeasurement());
  expect(result.initial).toBeGreaterThan(0); expect(result.repeated).toBe(result.initial); expect(result.cleared).toBeGreaterThan(result.repeated);
  expect(result.first).toEqual(result.second); expect(result.first).toEqual(result.third);
  expect(result.first[0]).toMatchObject({ requestedFamily: 'serif', selectedFamily: 'serif', substituted: false, glyphAuditComplete: true });
});

test('a cached window validates every new source and invalidates text when a provider changes its content', async ({ page }) => {
  const files = await importXlsx(readFileSync('examples/xlsx/01-values-text.xlsx'), { documentId: 'cached-window-source' });
  const result = await page.evaluate(entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 350, heightPt: 120, scrollLeftPt: 0, scrollTopPt: 0 } }], false, true), encodeTestPackage(files));
  expect(result.unchanged).toBe(true); expect(result.cacheSafety).toHaveLength(1);
  expect(result.cacheSafety[0]).toMatchObject({ rejected: 'InvalidDocument', rejectedNumeric: 'InvalidDocument', forkRetainsText: true, retained: true, changedText: '新的😀禰󠄀', leaked: 0 });
  expect(result.cacheSafety[0]!.offsets.map(part => part[2]).join('')).toBe('新的😀禰󠄀');
});

test('virtual grid renders formatted scalar values instead of imported display text in both directions', async ({ page }) => {
  for (const rtl of [false, true]) {
    const result = await page.evaluate(async rtl => {
      await window.workbookGridHost.render({ widthPt: 320, heightPt: 100, scrollLeftPt: 0, scrollTopPt: 4970 }, rtl);
      return window.workbookGridHost.inspect();
    }, rtl);
    await expect(page.locator('[data-office-view="workbook-cell"][data-row="500"][data-column="8"]')).toHaveText('12,345.68');
    expect(result.sourceUnchanged).toBe(true); expect(result.cellCount).toBeLessThanOrEqual(5000);
  }
});

test('numeric cell layouts format current values, fit real fonts and retain cell-based projection', async ({ page }) => {
  const cases: NumberTextCase[] = [
    { value: 1234.5, format: { code: '#,##0.00' }, widthPt: 90 },
    { value: -1.005, format: { code: '0.00;[Red](0.00)' }, widthPt: 80 },
    { value: 60, format: { builtinId: 14 }, widthPt: 100 },
    { value: 12.5, format: { code: '[$-407]0.00' }, widthPt: 80 },
    { value: .125, format: { code: '0.0%' }, widthPt: 80 },
    { value: 25.76, widthPt: 16 },
    { value: 12345.678, format: { code: '#,##0.00' }, widthPt: 20 },
    { value: -1, format: { builtinId: 14 }, widthPt: 80 },
    { value: 5, format: { code: '[>10]0;[<0]0' }, widthPt: 80 },
    { value: 12.5, format: { code: '"¥"* 0.00_)' }, widthPt: 100 },
    { value: 1.2345, widthPt: 15, styles: { '--office-cell-shrink-to-fit': 'true' } },
    { value: 12.5, widthPt: 75, heightPt: 75, styles: { '--office-cell-text-rotation': '45' } },
    { value: 12345, widthPt: 30, heightPt: 120, styles: { '--office-cell-text-rotation': '255' } },
    { value: 1, widthPt: 75, format: { code: '[Green]0.00' } },
    { value: 1, widthPt: 75, format: { code: '[Color3]0.00' }, styles: { '--office-indexed-10': '#123456' } },
    { value: 12.5, widthPt: 160, format: { code: '"预算 "0.0' } },
    { value: 1, widthPt: 75, format: { code: ';;' } },
  ];
  const report = await page.evaluate(cases => window.workbookGridHost.numberTextMeasurement(cases), cases), items = report.measurements;
  expect(items.slice(0, 6).map(item => item.text)).toEqual(['1,234.50', '(1.01)', '1900/2/29', '12,50', '12.5%', '26']);
  expect(items[1]!.color).toBe('rgb(255, 0, 0)'); expect(items[3]!.locale).toBe('de-DE');
  for (const index of [6, 7, 8]) expect(items[index]!.text).toMatch(/^#+$/);
  expect(items.slice(6, 9).map(item => item.overflow)).toEqual(['width', 'invalidDate', 'unmatchedCondition']);
  expect(items[9]!.text).toMatch(/^¥ +12\.50$/); expect(items[9]!.spaces).toHaveLength(1);
  expect(Math.abs(items[9]!.spaces[0]!.actualWidth - items[9]!.spaces[0]!.sourceWidth)).toBeLessThan(.1);
  expect(items[10]!.text).toBe('1.2345'); expect(items[10]!.scale).toBeLessThan(1);
  expect(items[11]!.transform).toContain('rotate(-45deg)'); expect(items[12]!.text).toBe('12345'); expect(items[12]!.breaks).toBe(4);
  expect(items[13]!.color).toBe('rgb(0, 255, 0)'); expect(items[14]!.color).toBe('rgb(18, 52, 86)');
  expect(items[15]!.text).toBe('预算 12.5'); expect(items[15]!.fonts).toContainEqual(expect.objectContaining({ role: 'east-asia', family: expect.stringMatching(/^"?Noto Sans CJK SC"?$/) }));
  expect(items[16]!.text).toBe('');
  for (const [index, item] of items.entries()) {
    expect(item.projections).toBe(0);
    if (![11, 12].includes(index)) { expect(item.widthPt).toBeLessThanOrEqual(cases[index]!.widthPt - 2.25 + .1); expect(Math.abs(item.rightGapPt - 1.125)).toBeLessThan(.1); }
  }
  expect(report.unchanged).toBe(true); expect(report.atomic).toBe(true); expect(report.retained).toBe(0);
  expect(report.failures).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'UnsupportedFeature', 'InvalidOperation', 'InvalidOperation']);
});

test('system date and time patterns share explicit locale semantics across runtime versions', async ({ page }) => {
  const report = await page.evaluate(() => window.workbookGridHost.numberTextMeasurement([
    { value: 45351.5, format: { code: '[$-F800]yyyy-mm-dd' }, widthPt: 240 },
    { value: .5, format: { code: '[$-F400]h:mm' }, widthPt: 100 },
    { value: .5, format: { code: 'h:mm"\u202f"AM/PM' }, widthPt: 100 },
  ], 'en-US'));
  expect(report.measurements.map(item => item.text)).toEqual(['Thursday, February 29, 2024', '12:00:00 PM', '12:00\u202fPM']);
  expect(report.unchanged).toBe(true); expect(report.retained).toBe(0);
});

test('numeric displays retain requested-font substitutions after fitting unchanged text and narrowing to hashes', async ({ page }) => {
  const styles = { 'font-family': '"MissingNumericFont", "Liberation Sans"', '--office-font-ascii': '"MissingNumericFont", "Liberation Sans"' };
  const result = await page.evaluate(styles => window.workbookGridHost.numberTextMeasurement([
    { value: 12.5, widthPt: 80, styles },
    { value: 12345.678, widthPt: 8, format: { code: '0.000' }, styles },
    { value: 12.5, widthPt: 160, format: { code: '"预算 "0.0' }, styles },
  ]), styles);
  expect(result.measurements[0]!.text).toBe('12.5'); expect(result.measurements[1]!.text).toMatch(/^#+$/); expect(result.measurements[2]!.text).toBe('预算 12.5');
  expect(result.fonts).toContainEqual(expect.objectContaining({ requestedFamily: 'MissingNumericFont', selectedFamily: 'Liberation Sans', substituted: true }));
  expect(result).toMatchObject({ unchanged: true, atomic: true, retained: 0 });
});

test('column font metrics use their declared Normal font and keep legacy defaults isolated', async ({ page }) => {
  const report = await page.evaluate(() => window.workbookGridHost.columnFontMeasurement());
  expect(report.measured.cssClass).toBe('normal-metric'); expect(report.fallback.cssClass).toBe('cell-default');
  expect(report.measured.maximumDigitWidthPx).toBeGreaterThan(report.fallback.maximumDigitWidthPx * 1.4);
  expect(report.measured.fonts).toContainEqual(expect.objectContaining({ requestedFamily: 'Liberation Serif', substituted: false }));
  expect(Math.abs(report.actual - report.measured.maximumDigitWidthPx)).toBeLessThan(.1); expect(report.retainedProbes).toBe(0);
  expect(report.unrelatedUnchanged).toBe(true);
});

test('worksheet colors resolve theme variables and luminance without treating native ARGB bytes as opacity', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookGridHost.colorMeasurement());
  expect(result.values).toEqual(['rgba(143,170,220,1)', 'rgba(51,90,161,1)', 'rgba(20,40,60,1)', 'rgba(255,0,0,1)', 'rgba(20,40,60,0.5)', 'rgba(0,255,0,1)', undefined]);
  expect(result.errors).toEqual(['InvalidDocument', 'InvalidDocument', 'InvalidDocument']); expect(result.unchanged).toBe(true);
  expect(result.fills).toEqual(['rgb(1, 2, 3)', 'transparent', 'rgba(143,170,220,1)', 'pattern']); expect(result.fillErrors).toEqual(['InvalidDocument']);
});

test('native cell fills match standard patterns and preserve full-cell gradients across clipped panes', async ({ page }) => {
  const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync('packages/office/tests/fixtures/xlsx-cell-appearance.xlsx')));
  const inputs = appearanceReference.cases.filter(input => input.kind === 'pattern' || input.kind === 'gradient');
  const classes = inputs.map((_, index) => model.getCell(`sheet-${index + 1}`, 1, 1)!.cssClass!);
  const report = await page.evaluate(({ css, classes }) => window.workbookGridHost.measureWorkbookFills(css, classes), { css: model.snapshot().styles.map(style => style.css).join('\n'), classes });
  expect(report.reports).toHaveLength(27);
  expect(report.invalid).toEqual(['InvalidDocument', 'ResourceLimitExceeded', 'InvalidDocument', 'InvalidDocument']);
  for (const [index, actual] of report.reports.entries()) {
    const input = inputs[index]!;
    expect(actual.sourceUnchanged).toBe(true); expect(actual.immutable).toBe(true); expect(actual.nodes).toBe(actual.actual);
    expect(actual.extent).toEqual(['24pt', '18pt']); expect(actual.largeNodes).toBeLessThanOrEqual(actual.nodes);
    expect(actual.errors).toEqual(['ResourceLimitExceeded']); expect(actual.difference, JSON.stringify(input)).toBeLessThanOrEqual(3);
    if (input.kind === 'pattern') {
      if (input.pattern === 'none') { expect(actual.samples.every(sample => sample.color[3] === 0)).toBe(true); continue; }
      if (input.pattern === 'solid') { expect(actual.samples.every(sample => sample.color.join(',') === '32,64,128,255')).toBe(true); continue; }
      const expected = fillReference.cases.find(item => item.pattern === input.pattern)!;
      const rows = expected.rows, width = rows[0]!.length, height = rows.length;
      let matches = false;
      for (let dy = 0; dy < height; dy++) for (let dx = 0; dx < width; dx++) if (actual.rows.every((row, y) => [...row].every((pixel, x) => pixel === rows[(y + dy) % height]![(x + dx) % width]))) matches = true;
      expect(matches, input.pattern).toBe(true);
      expect(actual.rows.join('').split('1').length - 1).toBe(expected.foregroundFraction * 1024);
    } else {
      for (const sample of actual.samples) {
        const x = (sample.x + .5) / 96, y = (sample.y + .5) / 72;
        let fraction: number;
        if (input.type === 'linear') {
          const angle = input.degree! * Math.PI / 180, dx = Math.cos(angle), dy = Math.sin(angle);
          fraction = .5 + ((x - .5) * 96 * dx + (y - .5) * 72 * dy) / (96 * Math.abs(dx) + 72 * Math.abs(dy));
        } else {
          const left = input.left!, right = input.right!, top = input.top!, bottom = input.bottom!;
          fraction = Math.max(0, x < left ? (left - x) / left : x > right ? (x - right) / (1 - right) : 0, y < top ? (top - y) / top : y > bottom ? (y - bottom) / (1 - bottom) : 0);
        }
        const expected = fraction <= .5 ? [255 * (1 - 2 * fraction), 510 * fraction, 0, 255] : [0, 510 * (1 - fraction), 255 * (2 * fraction - 1), 255];
        for (let channel = 0; channel < 4; channel++) expect(Math.abs(sample.color[channel]! - Math.round(expected[channel]!)), `${JSON.stringify(input)} at ${sample.x},${sample.y} channel ${channel}`).toBeLessThanOrEqual(3);
      }
    }
  }
});

test('native diagonal borders preserve direction, double spacing and standard slant-dash-dot strokes', async ({ page }) => {
  const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync('packages/office/tests/fixtures/xlsx-cell-appearance.xlsx')));
  const inputs = appearanceReference.cases.map((input, index) => ({ input, cssClass: model.getCell(`sheet-${index + 1}`, 1, 1)!.cssClass! })).filter(({ input }) => input.kind === 'diagonal');
  const report = await page.evaluate(({ css, classes }) => window.workbookGridHost.diagonalMeasurement(css, classes), { css: model.snapshot().styles.map(style => style.css).join('\n'), classes: inputs.map(input => input.cssClass) });
  expect(report.result).toHaveLength(9); expect(report.failures).toEqual(Array(4).fill('InvalidDocument'));
  for (const [index, actual] of report.result.entries()) {
    const input = inputs[index]!.input, perDirection = input.style === 'thin' ? 1 : 2;
    expect(actual.unchanged).toBe(true); expect(actual.difference, JSON.stringify(input)).toBeLessThanOrEqual(2); expect(actual.maskDifference, JSON.stringify(input)).toBeLessThanOrEqual(1);
    expect(actual.strokes).toHaveLength((Number(input.up) + Number(input.down)) * perDirection);
    expect(actual.strokes.filter(line => line.to[1]! < line.from[1]!)).toHaveLength(input.up ? perDirection : 0);
    expect(actual.strokes.filter(line => line.to[1]! > line.from[1]!)).toHaveLength(input.down ? perDirection : 0);
    for (const line of actual.strokes) { expect(line.color).toBe('rgba(0,68,170,1)'); expect(line.width).toBe(input.style === 'double' ? .5 : .75); }
    if (input.style === 'double') {
      const [a, b] = actual.strokes; const dx = a!.to[0]! - a!.from[0]!, dy = a!.to[1]! - a!.from[1]!;
      expect(Math.abs((b!.from[0]! - a!.from[0]!) * dy - (b!.from[1]! - a!.from[1]!) * dx) / Math.hypot(dx, dy)).toBeCloseTo(1.25);
    }
  }
  const width = borderSpec.rows[0]!.length;
  expect(Array.from({ length: width }, (_, phase) => report.rows.every((row, y) => [...row].every((pixel, x) => pixel === borderSpec.rows[y]![(x + phase) % width]))).some(Boolean)).toBe(true);
});

test('bounded cell style projection retains source chunk selectors, rich runs and lifecycle isolation', async ({ page }) => {
  const requests: string[] = []; page.on('request', request => { if (request.url().includes('workbook-injected')) requests.push(request.url()); });
  const report = await page.evaluate(() => window.workbookGridHost.sourceStyleMeasurement());
  expect(report).toMatchObject({ text: '一万行', italic: 'italic', background: 'rgb(68, 114, 196)', bold: '700', foreground: 'rgb(153, 51, 85)', plain: 'cell-default', blockRow: '9984', row: '9999', bound: true, disposed: true, noPartial: true, retained: 0 });
  expect(report.font).toContain('Liberation Serif'); expect(parseFloat(report.size)).toBeCloseTo(13 / .75, 1);
  expect(report.before).toContain('<span class="rich-style">一万</span>行');
  expect(report.rejected).toEqual(Array(6).fill('InvalidDocument')); expect(requests).toEqual([]);
  expect(await page.evaluate(() => (window as unknown as Record<string, unknown>).workbookInjected)).toBeUndefined();
});

test('cell font roles retain UTF-16 and variation selectors across source spans without changing CSS ancestry', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookGridHost.textFontMeasurement());
  expect(result).toMatchObject({ limited: true, unchanged: true, frozen: true });
  expect(result.runs.map(({ text, from, to, role }) => ({ text, from, to, role }))).toEqual([
    { text: '合并', from: 0, to: 2, role: 'east-asia' }, { text: '😀', from: 2, to: 4, role: 'h-ansi' },
    { text: '禰', from: 4, to: 5, role: 'east-asia' }, { text: '󠄀', from: 5, to: 7, role: 'east-asia' },
  ]);
  for (const run of result.runs.filter(run => run.role === 'east-asia')) expect(run.font.family).toContain('Noto Sans CJK SC');
  expect(result.rich.map(run => run.font.weight)).toEqual(['700', '400']);
  for (const run of result.rich) { expect(run.font.style).toBe('italic'); expect(run.font.family).toContain('Liberation Serif'); expect(run.font.sizePx).toBeCloseTo(13 / .75, 1); }
  expect(result.withBreak.map(run => run.text).join('')).toBe('合并😀禰󠄀\nA');
  expect(result.withBreak.at(-1)).toMatchObject({ text: 'A', from: 8, to: 9, role: 'ascii' });
});

test('rich text display preserves source offsets, local style and independent native decorations', async ({ page }, testInfo) => {
  const result = await page.evaluate(() => window.workbookGridHost.textViewMeasurement());
  expect(result).toMatchObject({ text: 'A\nB', unchanged: true, frozen: true,
    decoration: { strike: 'line-through', strikeStyle: 'solid', underline: 'underline', underlineStyle: 'double', plain: 'none' },
    accountingTypes: ['doubleAccounting'], failures: ['ResourceLimitExceeded', 'InvalidDocument'] });
  expect(result.rich.map(run => [run.text, run.from, run.to, run.weight, run.color])).toEqual([['一万', 0, 2, '700', 'rgb(153, 51, 85)'], ['行', 2, 3, '400', 'rgb(0, 0, 0)']]);
  expect(result.rich[0]!.x + result.rich[0]!.width).toBeCloseTo(result.rich[1]!.x, 0);
  expect(result.mixed.map(run => [run.text, run.from, run.to])).toEqual([['合并', 0, 2], ['😀', 2, 4], ['禰', 4, 5], ['󠄀', 5, 7]]);
  for (const run of result.mixed.filter(run => run.text !== '😀')) expect(run.family).toContain('Noto Sans CJK SC');
  expect(result.decorated.map(run => [run.text, run.from, run.to, run.color])).toEqual([['A', 0, 1, 'rgb(143, 170, 220)'], ['\n', 1, 2, 'rgb(143, 170, 220)'], ['B', 2, 3, 'rgb(143, 170, 220)']]);
  expect(result.decorated[2]!.y).toBeGreaterThan(result.decorated[0]!.y + result.decorated[0]!.height * .8);
  const screenshot = testInfo.outputPath('workbook-rich-text.png'); await page.locator('#workbook-text-view-probe').screenshot({ path: screenshot });
  await testInfo.attach('workbook-rich-text', { path: screenshot, contentType: 'image/png' });
});

test('border strokes match native point geometry and keep dash phase across clipped panes', async ({ page }, testInfo) => {
  const native = Object.entries(borderReference.kinds.pdfStrokes).filter(([kind]) => kind !== 'none' && kind !== 'slantDashDot');
  const edges = native.map(([kind, strokes]) => ({ kind: kind as WorkbookBorderKind, color: 'rgba(255,0,0,1)',
    from: { x: strokes[0]!.from[0]!, y: strokes.reduce((sum, stroke) => sum + stroke.from[1]!, 0) / strokes.length },
    to: { x: strokes[0]!.to[0]!, y: strokes.reduce((sum, stroke) => sum + stroke.to[1]!, 0) / strokes.length } }));
  const report = await page.evaluate(edges => window.workbookGridHost.borderViewMeasurement(edges), edges);
  expect(report.strokes).toHaveLength(native.flatMap(([, strokes]) => strokes).length);
  let index = 0;
  for (const [, strokes] of native) for (const expected of [...strokes].sort((a, b) => a.from[1]! - b.from[1]!)) {
    const actual = report.strokes[index++]!;
    expect(Math.abs(actual.width - expected.widthPt)).toBeLessThan(.02); expect(actual.cap).toBe(expected.cap);
    expect(actual.dash).toHaveLength(expected.dashPt.length);
    for (let n = 0; n < actual.dash.length; n++) expect(Math.abs(actual.dash[n]! - expected.dashPt[n]!)).toBeLessThan(.02);
    for (const point of ['from', 'to'] as const) for (let axis = 0; axis < 2; axis++) expect(Math.abs(actual[point][axis]! - expected[point][axis]!)).toBeLessThan(.01);
  }
  for (let i = 0; i < report.rasterImages.length; i++) await testInfo.attach(`border-raster-${i}`, { body: Buffer.from(report.rasterImages[i]!.split(',')[1]!, 'base64'), contentType: 'image/png' });
  expect(report.rasterSizes).toEqual(Array(4).fill([160, 80]));
  expect(report.splitMaskDifference).toBe(0); expect(report.splitDifference).toBeLessThanOrEqual(1);
  expect(report.maskDifference).toBe(0); expect(report.difference).toBeLessThanOrEqual(1); expect(report.painted).toBeGreaterThan(70); expect(report.painted).toBeLessThan(110);
  expect(report.doubleRows.length).toBeGreaterThanOrEqual(2);
  expect(report.doubleRows.some((row, index) => index && row - report.doubleRows[index - 1]! >= 3)).toBe(true);
  expect(report.failures).toEqual(['UnsupportedFeature', 'InvalidDocument', 'InvalidDocument', 'ResourceLimitExceeded']);
  const screenshot = testInfo.outputPath('workbook-border-strokes.png'); await page.locator('#workbook-border-view-probe').screenshot({ path: screenshot });
  await testInfo.attach('native-calibrated-border-strokes', { path: screenshot, contentType: 'image/png' });
});

test('worksheet context reading order uses its first non-whitespace scalar with isolated Unicode direction', async ({ page }) => {
  const report = await page.evaluate(() => window.workbookGridHost.textDirectionMeasurement());
  expect(report).toEqual({ values: ['ltr', 'rtl', 'ltr', 'ltr', 'rtl', 'ltr', 'ltr', 'rtl', 'rtl', 'ltr', 'rtl'], explicit: ['ltr', 'rtl'], wholeStringDirection: 'rtl',
    retainedText: '', failures: ['InvalidDocument', 'InvalidOperation', 'InvalidOperation'], remaining: 0, ownerDirection: 'rtl' });
});

test('worksheet text layout measures native alignment, Normal indentation and local fonts before clipping', async ({ page }, testInfo) => {
  const cases: TextLayoutCase[] = alignmentReference.cases.map(item => {
    const font = alignmentReference.fonts[item.font]!;
    return { text: item.kind === 'b' ? 'TRUE' : String(item.value), type: item.kind === 'n' ? 'number' : item.kind === 'b' ? 'boolean' : item.kind === 'e' ? 'error' : 'text',
      family: font.family, sizePt: font.sizePt, widthPt: item.cellBoundsPt[2]! - item.cellBoundsPt[0]!, heightPt: item.cellBoundsPt[3]! - item.cellBoundsPt[1]!,
      horizontal: item.horizontal, vertical: item.vertical, indent: item.indent, readingOrder: ['context', 'ltr', 'rtl'][item.readingOrder]!, rightToLeft: item.rightToLeft };
  });
  const result = await page.evaluate(cases => window.workbookGridHost.textLayoutMeasurement(cases), cases);
  await testInfo.attach('alignment-measurements', { body: JSON.stringify(result, null, 2), contentType: 'application/json' });
  expect(result).toMatchObject({ unchanged: true, retained: 0, atomic: true, frozen: true,
    failures: ['InvalidDocument', 'InvalidDocument', 'InvalidDocument', 'InvalidDocument', 'InvalidDocument', 'InvalidOperation', 'InvalidOperation', 'ResourceLimitExceeded', 'InvalidDocument'] });
  for (const [index, item] of cases.entries()) {
    const measured = result.measurements[index]!, native = alignmentReference.cases[index]!;
    expect(measured.text).toBe(item.text); expect(measured.scale).toBe(1);
    const direction = item.type === 'text' && (item.readingOrder === 'rtl' || item.readingOrder === 'context' && item.text === 'אבג') ? 'rtl' : 'ltr';
    expect(measured.direction).toBe(direction);
    if (item.type === 'error' || item.text.startsWith('-')) expect(measured.firstX).toBeLessThan(measured.lastX);
    const horizontal = item.horizontal !== 'general' ? item.horizontal : item.type === 'number' ? 'right' : item.type === 'boolean' || item.type === 'error' ? 'center' : direction === 'rtl' ? 'right' : 'left';
    expect(measured.horizontal).toBe(horizontal);
    const verticalSide = item.vertical === 'bottom' ? 3 : 1;
    expect(Math.abs(measured.gapsPt[verticalSide]! - native.textGapsPt[verticalSide]!), `vertical case ${index}`).toBeLessThan(.8);
    // Explicit placements and Latin General can be compared with this producer.
    // Its Boolean/error alignment and explicit-LTR Hebrew differ from OOXML.
    if (item.type === 'boolean' || item.type === 'error' || item.horizontal === 'general' && item.text === 'אבג' && item.readingOrder === 'ltr') continue;
    const aligned = horizontal === 'right' ? 2 : 0;
    // Calc's RTL PDF places the same content about .44pt farther right relative
    // to its red cell border. Record that producer difference separately; the
    // SDK must keep explicit physical alignment invariant under sheet reversal.
    if (item.rightToLeft) {
      const ltr = result.measurements[index - cases.length / 2]!;
      for (let side = 0; side < 4; side++) expect(measured.gapsPt[side]).toBeCloseTo(ltr.gapsPt[side]!, 1);
    } else if (horizontal !== 'center') expect(Math.abs(measured.gapsPt[aligned]! - native.textGapsPt[aligned]!), `horizontal case ${index}`).toBeLessThan(.65);
    else expect(Math.abs(measured.gapsPt[0]! - measured.gapsPt[2]!)).toBeLessThan(.1);
    expect(Math.abs(measured.widthPt - (native.cellBoundsPt[2]! - native.cellBoundsPt[0]! - native.textGapsPt[0]! - native.textGapsPt[2]!)), `text width case ${index}`).toBeLessThan(.15);
  }
  expect(result.fonts.every(font => !font.substituted)).toBe(true);
  await page.locator('#workbook-text-layout-probe').screenshot({ path: testInfo.outputPath('workbook-text-alignment.png') });
});

test('rotated and stacked text matches native ink bounds while retaining complete graphemes', async ({ page }, testInfo) => {
  const cases: TextLayoutCase[] = rotationReference.cases.map(item => ({ text: item.text, type: 'text', family: rotationReference.font.family, sizePt: rotationReference.font.sizePt,
    widthPt: item.cellSizePt[0]!, heightPt: item.cellSizePt[1]!, horizontal: 'center', vertical: 'center', indent: 0, readingOrder: 'context', rightToLeft: false, rotation: item.rotation }));
  cases.push({ ...cases[12]!, splitOffsets: [1, 3, 4, 5] });
  const report = await page.evaluate(cases => window.workbookGridHost.textLayoutMeasurement(cases), cases);
  expect(report).toMatchObject({ unchanged: true, retained: 0, atomic: true, frozen: true, usedFallbackLineMetrics: false,
    stackFailures: ['InvalidDocument', 'InvalidDocument', 'ResourceLimitExceeded'] });
  const inks = [];
  for (const [index, input] of cases.entries()) {
    const measured = report.measurements[index]!;
    expect(measured.text).toBe(input.text); expect(measured.scale).toBe(1); expect(measured.offsets[0]![0]).toBe(0); expect(measured.offsets.at(-1)![1]).toBe(input.text.length);
    for (let i = 1; i < measured.offsets.length; i++) expect(measured.offsets[i]![0]).toBe(measured.offsets[i - 1]![1]);
    if (input.rotation === 255) expect(measured.breaks).toBe(index >= 12 ? 2 : index === 10 ? 2 : 4);
    else if (input.rotation) expect(measured.transform).toBe(`rotate(${input.rotation <= 90 ? -input.rotation : input.rotation - 90}deg) scale(1)`);
    const bytes = await page.locator('#workbook-text-layout-probe > div').nth(index).screenshot();
    const ink = await page.evaluate(async base64 => {
      const image = new Image(); image.src = 'data:image/png;base64,' + base64; await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height; const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
      const pixels = context.getImageData(0, 0, image.width, image.height).data;
      return [64, 224].map(threshold => {
        let left = Infinity, top = Infinity, right = 0, bottom = 0;
        for (let y = 0; y < image.height; y++) for (let x = 0; x < image.width; x++) {
          const offset = (y * image.width + x) * 4;
          // Include LCD subpixel antialiasing. The gallery's pure-red outline
          // already lies outside the intensity threshold; filtering colored
          // edge pixels would remove real glyph coverage in Chromium.
          const maximum = Math.max(pixels[offset]!, pixels[offset + 1]!, pixels[offset + 2]!);
          if (pixels[offset + 3]! >= 128 && maximum < threshold) { left = Math.min(left, x); top = Math.min(top, y); right = Math.max(right, x + 1); bottom = Math.max(bottom, y + 1); }
        }
        return [left, top, right, bottom].map(value => value * .75);
      });
    }, bytes.toString('base64'));
    inks.push(ink);
    // The producer splits UTF-16 units in the last case, including surrogates
    // and combining marks. Its corrupted glyph boxes are not a reference.
    if (index < 12) for (let side = 0; side < 4; side++) expect.soft(Math.abs(ink[1]![side]! - rotationReference.cases[index]!.coverageBoundsPt[side]!), `rotation ${input.rotation}, case ${index}, side ${side}: ${JSON.stringify(ink)}`).toBeLessThanOrEqual(2);
  }
  await testInfo.attach('rotation-measurements', { body: JSON.stringify({ report, inks }, null, 2), contentType: 'application/json' });
  await testInfo.attach('rotation-gallery', { body: await page.locator('#workbook-text-layout-probe').screenshot(), contentType: 'image/png' });
});

test('rotated cell borders, fill and text share native bottom-anchored geometry', async ({ page }, testInfo) => {
  const report = await page.evaluate(cases => window.workbookGridHost.rotatedCellMeasurement(cases), rotatedCellsReference.cases);
  expect(report).toMatchObject({ unchanged: true, retained: 0 });
  const inks = [];
  for (const [index, reference] of rotatedCellsReference.cases.entries()) {
    const measured = report.measurements[index]!;
    expect(measured.projections[0]![0]).toBe(0); expect(measured.projections.at(-1)![1]).toBe(2);
    for (let part = 1; part < measured.projections.length; part++) expect(measured.projections[part]![0]).toBe(measured.projections[part - 1]![1]);
    for (let point = 0; point < 4; point++) for (let axis = 0; axis < 2; axis++) expect.soft(Math.abs(measured.points[point]![axis]! - reference.fillPointsPt[point]![axis]!), `fill ${index}, ${point}/${axis}`).toBeLessThan(.12);
    const cell = page.locator('#workbook-rotated-cell-probe > div').nth(index);
    // Pin every crop to the same device-pixel origin; scrolling a long gallery
    // otherwise changes glyph antialias phase and screenshot rounding per row.
    await cell.evaluate(element => { Object.assign(element.style, { position: 'fixed', top: '0', left: '0', zIndex: '1000' }); });
    const bytes = await cell.screenshot();
    await cell.evaluate(element => { element.style.position = 'relative'; element.style.removeProperty('top'); element.style.removeProperty('left'); element.style.removeProperty('z-index'); });
    const ink = await page.evaluate(async base64 => {
      const image = new Image(); image.src = 'data:image/png;base64,' + base64; await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height; const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
      const pixels = context.getImageData(0, 0, image.width, image.height).data; let left = Infinity, top = Infinity, right = 0, bottom = 0;
      for (let y = 1; y < image.height - 1; y++) for (let x = 0; x < image.width; x++) {
        const offset = (y * image.width + x) * 4, red = pixels[offset]!, green = pixels[offset + 1]!, blue = pixels[offset + 2]!;
        if (Math.max(red, green, blue) < 224 && !(red > green + 8 && blue > green + 8)) { left = Math.min(left, x); top = Math.min(top, y); right = Math.max(right, x + 1); bottom = Math.max(bottom, y + 1); }
      }
      return Number.isFinite(left) ? [left * .75 - 100, top * .75, right * .75 - 100, bottom * .75] : null;
    }, bytes.toString('base64'));
    inks.push(ink);
    if (!reference.coverageBoundsPt) expect(ink, `offscreen rotated text ${index}`).toBe(null);
    else {
      expect(ink).not.toBe(null);
      for (let side = 0; side < 4; side++) expect.soft(Math.abs(ink![side]! - reference.coverageBoundsPt[side]!), `text ${index}, rotation ${reference.rotation}, ${reference.mode}, side ${side}: ${JSON.stringify(ink)}`).toBeLessThanOrEqual(2);
    }
  }
  await testInfo.attach('rotated-cells', { body: JSON.stringify({ report, inks }, null, 2), contentType: 'application/json' });
});

test('rotated fill fragments retain pattern and gradient phase with bounded visible geometry', async ({ page }) => {
  const reports = await page.evaluate(() => window.workbookGridHost.shapedFillMeasurement());
  expect(reports).toHaveLength(8);
  for (const report of reports) {
    expect(report.difference, JSON.stringify(report)).toBeLessThanOrEqual(3);
    expect(report.hugeNodes).toBeLessThan(20); expect(report.extent).toEqual(['54pt', '30pt']);
    expect(report.failures).toEqual(['ResourceLimitExceeded']);
  }
});

test('worksheet wrapping and shrink preserve text offsets and report missing fonts without changing source', async ({ page }) => {
  const base: TextLayoutCase = { text: 'Long text 汉字𠮷禰󠄀', type: 'text', family: 'Liberation Sans', sizePt: 11, widthPt: 36, heightPt: 100,
    horizontal: 'left', vertical: 'top', indent: 0, readingOrder: 'context', rightToLeft: false };
  const result = await page.evaluate(cases => window.workbookGridHost.textLayoutMeasurement(cases), [
    base, { ...base, shrink: true }, { ...base, wrap: true, shrink: true }, { ...base, text: 'Long line\nSecond line', shrink: true },
    { ...base, family: 'Missing Workbook Layout Face', widthPt: 200 }, { ...base, text: 'A', vertical: 'bottom', widthPt: 100 }, { ...base, text: 'Long line', widthPt: 100 },
  ]);
  expect(result).toMatchObject({ unchanged: true, retained: 0, atomic: true });
  const [plain, shrunk, wrapped, multiline] = result.measurements;
  expect(shrunk!.scale).toBeLessThan(1); expect(shrunk!.widthPt).toBeLessThanOrEqual(36 - 2.25 + .1);
  expect(wrapped!.scale).toBe(1); expect(wrapped!.requiredHeightPt).toBeGreaterThan(plain!.requiredHeightPt * 2);
  expect(wrapped!.widthPt).toBeLessThanOrEqual(36 - 2.25 + .1); expect(multiline!.scale).toBe(1);
  expect(multiline!.requiredHeightPt).toBeCloseTo((result.measurements[6]!.requiredHeightPt - 1.5) * 2 + 1.5, 1);
  for (const item of [plain!, shrunk!, wrapped!]) { expect(item.text).toBe(base.text); expect(item.offsets.at(-1)![1]).toBe(base.text.length); }
  expect(result.fonts.find(font => font.requestedFamily === 'Missing Workbook Layout Face')).toMatchObject({ substituted: true });
});

test('source border resolution keeps native logical sides and physical CSS stroke widths distinct', async ({ page }, testInfo) => {
  const result = await page.evaluate(() => window.workbookGridHost.borderSourceMeasurement());
  expect(result).toMatchObject({ unchanged: true, failures: ['InvalidDocument', 'UnsupportedFeature', 'UnsupportedFeature', 'UnsupportedFeature', 'InvalidDocument'] });
  expect(result.native[0]).toEqual(result.native[1]);
  expect(result.native[0]).toEqual({ before: { kind: 'thin', color: 'rgba(143,170,220,1)' }, after: { kind: 'double', color: 'rgba(10,20,30,1)' } });
  expect(result.css[0]!.before).toEqual({ kind: 'css', style: 'double', widthPt: 3, color: 'rgba(10,20,30,0.5)' });
  expect(result.css[0]!.after).toEqual({ kind: 'css', style: 'dashed', widthPt: 1.5, color: 'rgba(0,0,255,1)' });
  expect(result.css[0]!.top).toEqual({ kind: 'css', style: 'dotted', widthPt: .75, color: 'rgba(0,128,0,1)' });
  expect(result.css[1]!.after).toEqual(result.css[0]!.before); expect(result.css[1]!.before).toEqual(result.css[0]!.after);
  expect(result.aliases[0]!.before!.kind).toBe('medium'); expect(result.aliases[0]!.after!.kind).toBe('thin');
  expect(result.aliases[1]).toEqual(result.aliases[2]);
  expect(result.strokes.map(item => [item.width, item.dash, item.cap])).toEqual([
    ['1', null, 'butt'], ['1', null, 'butt'], ['1.5', '4.5 1.5', 'butt'], ['0.75', '0 1.5', 'round'], ['0.5', null, 'butt'], ['0.5', null, 'butt'],
  ]);
  await page.locator('#workbook-css-border-probe').screenshot({ path: testInfo.outputPath('workbook-css-borders.png') });
});

test('workbook asset leases decode authorized images and isolate URL lifetimes without copying cell chunks', async ({ page }) => {
  for (const { bytes, mimeType, frames } of workbookRasterCases()) {
    const result = await page.evaluate(({ base64, mimeType }) => window.workbookGridHost.assetMeasurement(base64, mimeType), { base64: bytes.toString('base64'), mimeType });
    expect(result).toEqual({ count: 1, byteLength: bytes.length, pixels: 64 * frames, distinct: true,
      dimensions: Array(2).fill([8, 8]), revoked: true, isolated: true, failure: 'OperationCancelled' });
  }
  await expect(page.locator('#workbook-grid-probe img')).toHaveCount(0);
});

for (const rtl of [false, true]) test(`native table styles compose per property with direct formats and visible row banding in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-table-style-precedence.xlsx'), { documentId: 'table-composition' });
  const model = openDocument(files, { kind: 'workbook' }), sheets = model.snapshot().metadata.workbook.sheets;
  const report = await page.evaluate(async ({ entries, sheets, rtl }) => window.workbookGridHost.windowLayoutMeasurement(entries,
    sheets.map(sheetId => ({ sheetId, viewport: { widthPt: 240, heightPt: 180, scrollLeftPt: 0, scrollTopPt: 0 } })), rtl),
  { entries: encodeTestPackage(files), sheets: sheets.map(sheet => sheet.sheetId), rtl });
  expect(report.unchanged).toBe(true);
  const byName = new Map(sheets.map((sheet, index) => [sheet.name, report.results[index]!]));
  const cell = (name: string, row = 0, column = 0) => byName.get(name)!.cells.find(cell => cell.row === row && cell.column === column)!;
  const fill = (name: string, row = 0, column = 0) => byName.get(name)!.fillPaints.find(cell => cell.row === row && cell.column === column)!.color;
  for (const name of ['implicit', 'explicit_zero', 'bold', 'family', 'duplicate_font', 'duplicate_none_fill', 'duplicate_empty_border', 'number_format']) {
    expect(cell(name).weight, name).toBe('700'); expect(cell(name).color, name).toBe('rgb(255, 255, 255)');
    expect(fill(name), name).toBe('rgba(68,114,196,1)');
  }
  expect(cell('red_font').weight).toBe('700'); expect(cell('red_font').color).toBe('rgb(204, 0, 0)');
  expect(cell('red_fill').weight).toBe('700'); expect(fill('red_fill')).toBe('rgba(255,0,0,1)');
  expect(cell('family').family).toContain('DejaVu Sans');
  expect(cell('bold', 1).weight).toBe('700'); expect(cell('implicit', 1).weight).toBe('400');
  expect(cell('number_format', 1).text).toBe('2.00');
  expect(fill('hidden_rows', 1)).toBe(fill('implicit', 1)); expect(fill('hidden_rows', 3)).toBe(fill('implicit', 2));
  expect(fill('hidden_rows', 5)).toBe(fill('implicit', 3)); expect(byName.get('hidden_rows')!.cells.some(cell => cell.row === 2 || cell.row === 4)).toBe(false);
  expect(cell('no_header').weight).toBe('400'); expect(fill('no_header')).toBe(fill('implicit', 1));
  expect(byName.get('bottom_border')!.borderColors).toContain('rgba(0,255,0,1)');
  for (const result of report.results) {
    expect(result.retainedSources).toBe(0); expect(result.retainedMeasurements).toBe(0); expect(result.detached).toBe(true); expect(result.same).toBe(true);
    expect(result.errors).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded']);
  }
});

for (const rtl of [false, true]) test(`custom table differentials retain clearing, priority, names and formats in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const files = new Map(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-table-style-precedence.xlsx'), { documentId: 'custom-table-composition' }));
  const encoder = new TextEncoder(), decoder = new TextDecoder(), metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  const table = metadata.workbook.sheets[0]!.tables![0]!;
  table.style = { kind: 'custom', styleId: 'custom-table' }; table.dataCellStyleName = 'Unresolved native alias'; table.columns[0]!.dataCellStyleName = '区域别名';
  metadata.workbook.cellStyles!.definitions.push({ styleId: 'region-name', cssClass: 'named-region', numberFormatId: 'num-0',
    apply: { font: true, fill: false, border: false, alignment: false, numberFormat: false, protection: false } });
  metadata.workbook.cellStyles!.names.push({ name: '区域别名', styleId: 'region-name', hidden: false, customBuiltin: false });
  metadata.workbook.tableStyles = [{ styleId: 'custom-table', name: 'Custom', pivotCompatible: false, elements: [
    { region: 'firstHeaderCell', size: 1, style: { cssClass: 'corner' } },
    { region: 'headerRow', size: 1, style: { cssClass: 'header' } },
    { region: 'firstRowStripe', size: 1, style: { cssClass: 'body-number', numberFormatId: 'num-2' } },
    { region: 'wholeTable', size: 1, style: { cssClass: 'whole' } },
  ] }];
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  files.set('styles.css', encoder.encode(decoder.decode(files.get('styles.css')!) + `
    .corner{color:#cc0000}
    .header{font-weight:normal;--office-cell-fill-kind:none;--office-cell-border-bottom-kind:medium;--office-cell-border-bottom-color:#00ff00}
    .whole{font-weight:bold;background-color:#ffff00;border-left:1pt solid #ff0000;--office-cell-border-vertical-kind:thin;--office-cell-border-vertical-color:#0000ff}
    .body-number{font-style:italic}
    .named-region{font-weight:bold;font-size:11pt;color:#800080;--office-cell-foreground-color:#800080;--office-cell-foreground-tint:0}
  `));
  const report = await page.evaluate(({ entries, rtl }) => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 240, heightPt: 180, scrollLeftPt: 0, scrollTopPt: 0 } }], rtl), { entries: encodeTestPackage(files), rtl });
  const result = report.results[0]!, cell = (row: number, column: number) => result.cells.find(cell => cell.row === row && cell.column === column)!;
  expect(cell(0, 0).color).toBe('rgb(204, 0, 0)'); expect(cell(0, 0).weight).toBe('400'); expect(cell(0, 1).weight).toBe('400');
  expect(result.fillPaints.find(cell => cell.row === 0 && cell.column === 0)!.color).toBe('transparent');
  expect(cell(1, 0).color).toBe('rgb(128, 0, 128)'); expect(cell(1, 0).text).toBe('2.00');
  expect(result.fillPaints.find(cell => cell.row === 1 && cell.column === 0)!.color).toBe('rgb(255, 255, 0)');
  expect(result.borderColors).toEqual(expect.arrayContaining(['rgba(255,0,0,1)', 'rgba(0,0,255,1)', 'rgba(0,255,0,1)']));
  expect(report.unchanged).toBe(true); expect(result.same).toBe(true); expect(result.retainedSources).toBe(0);
});

for (const rtl of [false, true]) test(`composed native workbook windows retain values, formatting and source in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  for (const name of ['01-values-text', '02-styles-dates', '03-structure']) {
    const files = await importXlsx(readFileSync(`examples/xlsx/${name}.xlsx`), { documentId: 'composed-' + name });
    const report = await page.evaluate(async ({ entries, rtl }) => window.workbookGridHost.windowLayoutMeasurement(entries,
      [{ sheetId: 'sheet-1', viewport: { widthPt: 680, heightPt: 230, scrollLeftPt: 0, scrollTopPt: 0 } }], rtl), { entries: encodeTestPackage(files), rtl });
    expect(report.unchanged).toBe(true);
    const result = report.results[0]!;
    expect(result.detached).toBe(true); expect(result.same).toBe(true); expect(result.retainedSources).toBe(0); expect(result.retainedMeasurements).toBe(0);
    expect(result.errors).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded']);
    const cell = (row: number, column: number) => result.cells.find(cell => cell.row === row && cell.column === column)!;
    if (name === '01-values-text') {
      expect(cell(2, 0).text).toBe('12.5'); expect(cell(2, 2).text).toBe('TRUE'); expect(cell(2, 3).text).toBe('FALSE');
      expect(cell(3, 2).text).toBe('#DIV/0!'); expect(cell(3, 3).text).toBe('#N/A');
      expect(cell(2, 2).offsets).toEqual([]); expect(cell(2, 0).offsets).toEqual([]);
      expect(cell(1, 1).text).toBe('001234567890123456789');
      const rich = cell(5, 0); expect(rich.text).toBe('中文 Latin 😀禰󠄀\n下一行');
      expect(rich.offsets.map(part => part[2]).join('')).toBe(rich.text);
    }
    if (name === '02-styles-dates') {
      expect(cell(1, 0).text).toBe('12,345.68'); expect(cell(1, 1).text).toBe('2025-12-09'); expect(cell(1, 2).text).toBe('12.50%');
      expect(cell(1, 3).text).toBe('12:00:00'); expect(cell(1, 4).text).toBe('1.25E+08'); expect(cell(1, 5).text).toBe('-¥1,234.50');
      expect(cell(6, 1).text).toBe('1900-02-29'); expect(result.borderCount).toBeGreaterThan(0);
    }
    if (name === '03-structure') { expect(cell(1, 1).text).toBe('合并标题'); expect(result.cells.some(cell => cell.row === 3)).toBe(false); }
  }
});

test('composed native rotated cells share text fill and border geometry inside clipped windows', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-rotated-cells.xlsx'), { documentId: 'rotated-window' });
  const report = await page.evaluate(async entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    ['sheet-3', 'sheet-8', 'sheet-13', 'sheet-18'].map(sheetId => ({ sheetId, viewport: { widthPt: 240, heightPt: 120, scrollLeftPt: 50, scrollTopPt: 18 } }))), encodeTestPackage(files));
  expect(report.unchanged).toBe(true);
  for (const result of report.results) {
    const cell = result.cells.find(cell => cell.row === 1 && cell.column === 1)!;
    expect(cell.text).toBe('Ab'); expect(Math.abs(cell.shape.topOffsetPt)).toBeGreaterThan(1);
    expect(result.fills.some(fill => fill.transform?.startsWith('matrix(1 0 '))).toBe(true);
    expect(result.borderCount).toBeGreaterThan(0); expect(result.detached).toBe(true); expect(result.same).toBe(true);
    expect(result.retainedSources).toBe(0); expect(result.retainedMeasurements).toBe(0);
    expect(result.fills.every(fill => fill.width <= 321)).toBe(true);
  }
});

for (const rtl of [false, true]) test(`composed merged fills and text keep full geometry across four panes in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const files = new Map(await importXlsx(readFileSync('examples/xlsx/03-structure.xlsx'), { documentId: 'merge-window' }));
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!));
  metadata.workbook.sheets[0].merges[0] = { from: { row: 1, column: 1 }, to: { row: 5, column: 5 } };
  metadata.workbook.sheets[0].frozen = { rows: 3, columns: 3 };
  metadata.workbook.sheets[0].showGridLines = true;
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  files.set('styles.css', new TextEncoder().encode(new TextDecoder().decode(files.get('styles.css')!) +
    'tr[data-row="1"] td[data-column="1"]{--office-cell-fill-kind:darkTrellis;--office-cell-fill-color:#445566;--office-cell-fill-background-color:#ddeeff;--office-cell-horizontal:center;--office-cell-vertical:center}' +
    'div{width:999pt!important;height:888pt!important}svg{width:777pt!important;height:666pt!important}'));
  const widthPt = 350, heightPt = 120;
  const report = await page.evaluate(async ({ entries, rtl, widthPt, heightPt }) => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt, heightPt, scrollLeftPt: 30, scrollTopPt: 15 } }], rtl), { entries: encodeTestPackage(files), rtl, widthPt, heightPt });
  const result = report.results[0]!, merged = result.cells.filter(cell => cell.row === 1 && cell.column === 1);
  expect(report.unchanged).toBe(true); expect(merged).toHaveLength(4); expect(result.widthPt).toBeCloseTo(widthPt, 1); expect(result.heightPt).toBeCloseTo(heightPt, 1);
  expect(result.gridLines).toBe(4);
  const offsets = merged.map(cell => ({ x: cell.bounds.x - (rtl ? widthPt - cell.rect.x - cell.rect.width : cell.rect.x) / .75, y: cell.bounds.y - cell.rect.y / .75 }));
  for (const [index, cell] of merged.entries()) {
    expect(cell.text).toBe('合并标题'); expect(cell.offsets.map(part => part[2]).join('')).toBe('合并标题');
    expect(Math.abs(offsets[index]!.x - offsets[0]!.x)).toBeLessThan(.1); expect(Math.abs(offsets[index]!.y - offsets[0]!.y)).toBeLessThan(.1);
  }
  expect(result.same).toBe(true); expect(result.retainedMeasurements).toBe(0); expect(result.retainedSources).toBe(0);
});

test('composed rotated merge uses its covered perimeter for text and fill coupling', async ({ page }) => {
  const files = new Map(await importXlsx(readFileSync('examples/xlsx/03-structure.xlsx'), { documentId: 'merge-corner-window' }));
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!));
  metadata.workbook.sheets[0].merges[0] = { from: { row: 1, column: 1 }, to: { row: 5, column: 5 } };
  metadata.workbook.sheets[0].frozen = { rows: 3, columns: 3 };
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  files.set('styles.css', new TextEncoder().encode(new TextDecoder().decode(files.get('styles.css')!) +
    'tr[data-row="1"] td[data-column="1"]{--office-cell-text-rotation:45;--office-cell-border-left-kind:none;--office-cell-border-right-kind:none;--office-cell-border-top-kind:none;--office-cell-border-bottom-kind:none;--office-cell-fill-kind:solid;--office-cell-fill-color:#ddeeff}' +
    'tr[data-row="5"] td[data-column="1"]{--office-cell-border-bottom-kind:thin;--office-cell-border-bottom-color:#778899}'));
  const report = await page.evaluate(async entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 350, heightPt: 120, scrollLeftPt: 30, scrollTopPt: 15 } }]), encodeTestPackage(files));
  const result = report.results[0]!, merged = result.cells.filter(cell => cell.row === 1 && cell.column === 1);
  expect(merged).toHaveLength(4);
  for (const cell of merged) expect(cell.shape.topOffsetPt).toBeCloseTo(cell.shape.heightPt, 6);
  expect(result.fills.filter(fill => fill.transform?.startsWith('matrix(1 0 '))).toHaveLength(4);
  expect(result.retainedSources).toBe(0); expect(result.same).toBe(true); expect(report.unchanged).toBe(true);
});

test('composed native plain and rich text formats preserve source styles and fields through affixes, repetition and stacking', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-text-formats.xlsx'), { documentId: 'text-format-window' });
  const report = await page.evaluate(async entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 300, heightPt: 400, scrollLeftPt: 0, scrollTopPt: 0 } }]), encodeTestPackage(files));
  const result = report.results[0]!, cell = (row: number) => result.cells.find(cell => cell.row === row && cell.column === 0)!;
  expect(result.same).toBe(true); expect(result.detached).toBe(true); expect(report.unchanged).toBe(true);
  expect(cell(0).text).toBe('ID 001😀'); expect(cell(0).offsets.map(part => part[2]).join('')).toBe('001😀');
  expect(cell(1).text).toBe('A/B / A/B'); expect(cell(1).offsets).toEqual([[0, 3, 'A/B'], [0, 3, 'A/B']]);
  expect(cell(2).text).toBe('ID AB'); expect(cell(2).color).toBe('rgb(255, 0, 0)'); expect(cell(2).offsets).toEqual([[0, 2, 'AB']]);
  expect(cell(3).text).toBe('fixed'); expect(cell(3).offsets).toEqual([]); expect(cell(4).text).toBe(''); expect(cell(4).offsets).toEqual([]);
  expect(cell(5).text).toMatch(/^AB\.+$/); expect(cell(5).spaces).toHaveLength(1); expect(cell(5).spaces[0]).toBeGreaterThan(0);
  expect(cell(6).text).toMatch(/^x+AB$/); expect(cell(7).text).toMatch(/^ABx+$/); expect(cell(8).text).toMatch(/^AB\.+$/);
  for (const row of [5, 6, 7, 8]) { expect(cell(row).offsets).toEqual([[0, 2, 'AB']]); expect(cell(row).bounds.width * .75).toBeLessThanOrEqual(cell(row).shape.widthPt - 2.25 + .1); }
  expect(cell(9).text).toBe('ID A😀'); expect(cell(9).breaks).toBe(4); expect(cell(9).offsets).toEqual([[0, 1, 'A'], [1, 3, '😀']]);
  expect(cell(10).text).toBe('ID AB'); expect(cell(10).offsets).toEqual([[0, 1, 'A'], [1, 2, 'B']]);
  expect(cell(10).runs.map(run => run.color)).toEqual(['rgb(255, 0, 0)', 'rgb(255, 0, 0)']);
  expect(cell(11).text).toBe('ID AB'); expect(cell(11).offsets).toEqual([[0, 1, 'A'], [1, 2, 'B']]);
  expect(cell(11).runs.map(run => [run.color, run.weight, run.style])).toEqual([['rgb(0, 0, 255)', '700', 'normal'], ['rgb(0, 170, 0)', '400', 'italic']]);
  expect(cell(12).text).toBe('AB/AB'); expect(cell(12).offsets).toEqual([[0, 1, 'A'], [1, 2, 'B'], [0, 1, 'A'], [1, 2, 'B']]);
  for (const row of [10, 11, 12]) {
    expect(cell(row).runs.map(run => run.weight)).toEqual(row === 12 ? ['700', '400', '700', '400'] : ['700', '400']);
    expect(cell(row).runs.filter(run => run.text === 'A').every(run => run.family.includes('Liberation Serif'))).toBe(true);
  }
  expect(result.retainedMeasurements).toBe(0); expect(result.retainedSources).toBe(0);
});

for (const rtl of [false, true]) test(`composed native conditional rules and validation use current source in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const files = await importXlsx(readFileSync('examples/xlsx/06-rules-validation.xlsx'), { documentId: 'rules-window' });
  const report = await page.evaluate(async ({ entries, rtl }) => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 420, heightPt: 220, scrollLeftPt: 0, scrollTopPt: 0 } }], rtl), { entries: encodeTestPackage(files), rtl });
  const result = report.results[0]!, cell = (row: number, column: number) => result.cells.find(cell => cell.row === row && cell.column === column)!;
  expect(report.unchanged).toBe(true); expect(result.retainedSources).toBe(0); expect(result.same).toBe(true);
  // Independent source DXF 0 is FF9C0006; DXF 1 changes only fill.
  expect(cell(4, 0).color).toBe('rgb(156, 0, 6)'); expect(cell(1, 0).color).not.toBe('rgb(156, 0, 6)');
  expect(result.fillPaints.find(cell => cell.row === 2 && cell.column === 1)!.color).toBe('rgba(198,239,206,1)');
  expect(result.fillPaints.find(cell => cell.row === 1 && cell.column === 1)!.color).not.toBe('rgba(198,239,206,1)');
  expect(result.rules!.cells.filter(cell => cell.column === 4).map(cell => cell.validation?.state)).toEqual(['valid', 'invalid', 'valid', 'valid', 'valid', 'valid', 'valid']);
  const fills = (column: number) => result.fillPaints.filter(cell => cell.column === column && cell.row >= 1 && cell.row <= 5).map(cell => cell.color);
  expect(fills(2)[0]).toBe('rgb(248, 105, 107)'); expect(fills(2).at(-1)).toBe('rgb(99, 190, 123)');
  expect(fills(3)[0]).toBe('rgb(248, 105, 107)'); expect(fills(3).at(-1)).toBe('rgb(99, 190, 123)');
});

test('native color scales preserve constant and repeated threshold paints from independent PDFs', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-color-scales.xlsx'), { documentId: 'scale-paints' });
  const report = await page.evaluate(async entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    Array.from({ length: 8 }, (_, index) => ({ sheetId: `sheet-${index + 1}`, viewport: { widthPt: 150, heightPt: 110, scrollLeftPt: 0, scrollTopPt: 0 } }))), encodeTestPackage(files));
  const red = 'rgb(255, 0, 0)', yellow = 'rgb(255, 255, 0)', green = 'rgb(0, 255, 0)';
  const expected = [[red, red, red], [red, red, red], [red, 'rgb(128, 127, 0)', green], [green, green, green], [green, green, green], [red, yellow, green], [red, red, green], [red, green, green]];
  for (const [index, result] of report.results.entries()) expect(result.fillPaints.filter(cell => cell.column === 0 && cell.row < 3).map(cell => cell.color)).toEqual(expected[index]);
  expect(report.unchanged).toBe(true);
});

for (const rtl of [false, true]) test(`conditional property priority overrides table and direct formatting without speculative fallback in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const files = new Map(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-table-style-precedence.xlsx'), { documentId: 'conditional-precedence' }));
  const encoder = new TextEncoder(), decoder = new TextDecoder(), metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  const origin = { row: 0, column: 0 }, ranges = [{ from: origin, to: { row: 2, column: 1 } }];
  metadata.workbook.sheets[0]!.conditionalRules = [
    { ruleId: 'unknown', priority: 1, kind: 'expression', expression: '=FUTUREFUNC()', origin, ranges: [{ from: { row: 0, column: 1 }, to: { row: 0, column: 1 } }], stopIfTrue: false, cssClass: 'cf-unknown' },
    { ruleId: 'first', priority: 2, kind: 'expression', expression: '=TRUE', origin, ranges, stopIfTrue: false, cssClass: 'cf-first', numberFormatId: 'num-2' },
    { ruleId: 'stop', priority: 3, kind: 'expression', expression: '=ROW()=2', origin, ranges, stopIfTrue: true },
    { ruleId: 'last', priority: 4, kind: 'expression', expression: '=TRUE', origin, ranges, stopIfTrue: false, cssClass: 'cf-last', numberFormatId: 'num-0' },
  ];
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  files.set('styles.css', encoder.encode(decoder.decode(files.get('styles.css')!) + '.cf-first{font-weight:normal;color:#112233}.cf-last{font-style:italic;font-weight:bold;color:#ff0000;background-color:#ffff00}.cf-unknown{color:#ff00ff}'));
  const report = await page.evaluate(({ entries, rtl }) => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 240, heightPt: 160, scrollLeftPt: 0, scrollTopPt: 0 } }], rtl), { entries: encodeTestPackage(files), rtl });
  const result = report.results[0]!, cell = (row: number, column: number) => result.cells.find(cell => cell.row === row && cell.column === column)!;
  expect(cell(0, 0).weight).toBe('400'); expect(cell(0, 0).color).toBe('rgb(17, 34, 51)');
  expect(cell(0, 1).color).toBe('rgb(255, 255, 255)');
  expect(cell(1, 0).text).toBe('2.00'); expect(cell(1, 0).color).toBe('rgb(17, 34, 51)');
  expect(result.fillPaints.find(cell => cell.row === 0 && cell.column === 0)!.color).toBe('rgb(255, 255, 0)');
  expect(result.fillPaints.find(cell => cell.row === 1 && cell.column === 0)!.color).not.toBe('rgb(255, 255, 0)');
  expect(result.rules!.cells.find(cell => cell.row === 1 && cell.column === 0)!.conditional.map(rule => rule.ruleId)).toEqual(['first', 'stop']);
  expect(report.unchanged).toBe(true); expect(result.retainedSources).toBe(0);
});

test('rules display current formula values while preserving deliberately wrong native caches and derived text has no source offsets', async ({ page }) => {
  const files = new Map(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-scalar-semantics.xlsx'), { documentId: 'current-rule-display' }));
  const encoder = new TextEncoder(), metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.conditionalRules = [{ ruleId: 'current', kind: 'expression', expression: '=TRUE', priority: 1, stopIfTrue: false, origin: { row: 0, column: 4 }, ranges: [{ from: { row: 0, column: 4 }, to: { row: 14, column: 4 } }] }];
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  const report = await page.evaluate(entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 420, heightPt: 300, scrollLeftPt: 0, scrollTopPt: 0 } }]), encodeTestPackage(files));
  const cells = report.results[0]!.cells.filter(cell => cell.column === 4 && cell.row <= 14), expected = JSON.parse(readFileSync('docs/evidence/xlsx-scalar-semantics.json', 'utf8')).cases;
  for (const cell of cells) {
    const value = expected[cell.row].expected;
    expect(cell.text).toBe(value.type === 'boolean' ? value.value ? 'TRUE' : 'FALSE' : String(value.value));
    expect(cell.offsets).toEqual([]);
  }
  expect(cells.length).toBeGreaterThan(10); expect(report.unchanged).toBe(true);
  expect(openDocument(files, { kind: 'workbook' }).getCell('sheet-1', 0, 4)!.value).toEqual({ type: 'number', value: 999 });
});

test('conditional font attributes override rich runs while text-format explicit color keeps its declared precedence', async ({ page }) => {
  const files = new Map(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-text-formats.xlsx'), { documentId: 'rich-rule-font' }));
  const encoder = new TextEncoder(), decoder = new TextDecoder(), metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.conditionalRules = [{ ruleId: 'font', kind: 'expression', expression: '=TRUE', priority: 1, stopIfTrue: false, origin: { row: 10, column: 0 },
    ranges: [{ from: { row: 10, column: 0 }, to: { row: 11, column: 0 } }], cssClass: 'conditional-font' }];
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  files.set('styles.css', encoder.encode(decoder.decode(files.get('styles.css')!) + '.conditional-font{font-weight:normal;color:#112233}'));
  const report = await page.evaluate(entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 300, heightPt: 400, scrollLeftPt: 0, scrollTopPt: 0 } }]), encodeTestPackage(files));
  const cells = report.results[0]!.cells;
  expect(cells.find(cell => cell.row === 10 && cell.column === 0)!.runs.map(run => [run.color, run.weight, run.style])).toEqual([['rgb(255, 0, 0)', '400', 'normal'], ['rgb(255, 0, 0)', '400', 'italic']]);
  expect(cells.find(cell => cell.row === 11 && cell.column === 0)!.runs.map(run => [run.color, run.weight, run.style])).toEqual([['rgb(17, 34, 51)', '400', 'normal'], ['rgb(17, 34, 51)', '400', 'italic']]);
  expect(report.unchanged).toBe(true);
});

test('whole-value fill and horizontal line alignment retain exact source offsets and bounded display text', async ({ page }) => {
  const base: TextLayoutCase = { text: 'Ab', type: 'text', family: 'Liberation Sans', sizePt: 11, widthPt: 80, heightPt: 70, horizontal: 'fill', vertical: 'bottom', readingOrder: 'context', indent: 0, rightToLeft: false };
  const cases = [base, { ...base, text: '12', type: 'number' as const, wrap: true }, { ...base, wrap: true },
    { ...base, text: 'Alpha beta gamma delta epsilon', horizontal: 'justify', widthPt: 80 },
    { ...base, text: 'Alpha beta', horizontal: 'distributed', justifyLastLine: true },
    { ...base, text: 'Alpha beta', horizontal: 'distributed', justifyLastLine: false }];
  const report = await page.evaluate(cases => window.workbookGridHost.textLayoutMeasurement(cases), cases);
  expect(report.unchanged).toBe(true); expect(report.retained).toBe(0);
  for (const index of [0, 1]) {
    const value = report.measurements[index]!; expect(value.text).toMatch(index === 0 ? /^(Ab){3,}$/ : /^(12){3,}$/);
    expect(value.widthPt).toBeLessThanOrEqual(80 - 2.25 + .2); expect(value.offsets.every(range => range[0] === 0 && range[1] === 2)).toBe(true);
  }
  expect(report.measurements[2]!.text).toBe('Ab');
  expect(report.measurements[3]!.horizontal).toBe('justify'); expect(report.measurements[3]!.heightPt).toBeGreaterThan(20);
  expect(report.measurements[4]!.lastX).toBeGreaterThan(report.measurements[5]!.lastX + 20);
});

test('vertical line distribution preserves rich UTF-16 fragments and natural automatic-height measurements', async ({ page }) => {
  const base: TextLayoutCase = { text: 'Alpha\nbeta\ngamma', type: 'text', family: 'Liberation Sans', sizePt: 11, widthPt: 100, heightPt: 100, horizontal: 'left', vertical: 'justify', readingOrder: 'context', indent: 0, rightToLeft: false, wrap: true };
  const report = await page.evaluate(cases => window.workbookGridHost.textLayoutMeasurement(cases), [base, { ...base, vertical: 'distributed', splitOffsets: [2, 7, 12] }, { ...base, vertical: 'top' }]);
  expect(report.unchanged).toBe(true); expect(report.retained).toBe(0);
  for (const index of [0, 1]) {
    const value = report.measurements[index]!, original = report.measurements[2]!;
    expect(value.text).toBe(base.text); expect(value.requiredHeightPt).toBeCloseTo(original.requiredHeightPt, 1);
    const ys = [...new Set(value.lineRects.map(rect => Math.round(rect.y)))].sort((a, b) => a - b);
    expect(ys).toHaveLength(3); expect(ys[1]! - ys[0]!).toBeGreaterThan(35); expect(ys[2]! - ys[1]!).toBeGreaterThan(35);
    expect(value.offsets[0]![0]).toBe(0); expect(value.offsets.at(-1)![1]).toBe(base.text.length);
  }
});

test('filled and distributed displays reject excessive text and nodes while retaining the accepted gallery', async ({ page }) => {
  const base: TextLayoutCase = { text: 'Ab', type: 'text', family: 'Liberation Sans', sizePt: 11, widthPt: 80, heightPt: 100, horizontal: 'fill', vertical: 'bottom', readingOrder: 'context', indent: 0, rightToLeft: false };
  const result = await page.evaluate(async base => {
    await window.workbookGridHost.textLayoutMeasurement([base]); const accepted = document.querySelector('#workbook-text-layout-probe')!;
    const errors: string[] = [];
    for (const input of [{ item: base, limits: { maxTextUnits: 2 } }, { item: { ...base, widthPt: 1_000_000 }, limits: { maxNodes: 50 } },
      { item: { ...base, text: 'Alpha\nbeta\ngamma', horizontal: 'left', vertical: 'distributed' }, limits: { maxNodes: 12 } }]) {
      try { await window.workbookGridHost.textLayoutMeasurement([input.item], input.limits); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    }
    return { errors, retained: accepted.isConnected, galleries: document.querySelectorAll('#workbook-text-layout-probe').length,
      probes: document.querySelectorAll('[data-office-view$="measurement"]').length };
  }, base);
  expect(result).toEqual({ errors: ['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded'], retained: true, galleries: 1, probes: 0 });
});

test('native whole-value fills match independent fixed-font PDF words in both worksheet directions', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), { documentId: 'native-cell-fills' });
  const reference = JSON.parse(readFileSync('docs/evidence/xlsx-cell-flow.json', 'utf8')).cases as { group: string; horizontal: string; words: { text: string }[] }[];
  const selected = reference.map((value, index) => ({ value, index })).filter(({ value }) => value.group === 'horizontal' && value.horizontal === 'fill');
  const report = await page.evaluate(({ entries, ids }) => window.workbookGridHost.windowLayoutMeasurement(entries,
    ids.map(sheetId => ({ sheetId, viewport: { widthPt: 280, heightPt: 130, scrollLeftPt: 0, scrollTopPt: 0 } }))),
    { entries: encodeTestPackage(files), ids: selected.map(({ index }) => 'sheet-' + (index + 1)) });
  for (const [index, result] of report.results.entries()) {
    const expected = selected[index]!.value.words.filter(word => !['Origin', 'After'].includes(word.text)).map(word => word.text).join('');
    expect(result.cells.find(cell => cell.row === 1 && cell.column === 1)!.text).toBe(expected);
  }
  expect(report.unchanged).toBe(true);
});

test('native text spills through true blanks and styled blanks while values, formulas and merges block it', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), { documentId: 'native-text-overflow' });
  const report = await page.evaluate(entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [1, 2, 3, 4, 5].map(index => ({ sheetId: 'sheet-' + index, viewport: { widthPt: 280, heightPt: 120, scrollLeftPt: 0, scrollTopPt: 0 } }))), encodeTestPackage(files));
  for (const [index, result] of report.results.entries()) {
    const cell = result.cells.find(cell => cell.row === 1 && cell.column === 1)!;
    expect(cell.text).toBe('Alpha beta gamma delta epsilon'); expect(cell.offsets.map(item => item[2]).join('')).toBe(cell.text);
    if ([0, 2].includes(index)) expect(cell.clipWidthPt).toBeGreaterThan(cell.rect.width * 3);
    else expect(cell.clipWidthPt).toBeCloseTo(cell.rect.width, 1);
  }
  expect(report.unchanged).toBe(true);
});

test('offscreen source anchors remain visible only when their bounded ink reaches an unblocked pane', async ({ page }) => {
  const files = new Map(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), { documentId: 'offscreen-overflow' }));
  const encoded = encodeTestPackage(files);
  const report = await page.evaluate(entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [1, 2, 3, 4, 5, 6].map(index => ({ sheetId: 'sheet-' + index, viewport: { widthPt: 120, heightPt: 110, scrollLeftPt: index === 6 ? 110 : 160, scrollTopPt: 0 } }))), encoded);
  for (const [index, result] of report.results.entries()) {
    const cell = result.cells.find(cell => cell.row === 1 && cell.column === 1);
    if ([0, 2, 5].includes(index)) {
      expect(cell, result.sheetId).toBeDefined(); expect(cell!.rect.x + cell!.rect.width).toBeLessThan(0); expect(result.overflowCount).toBeGreaterThan(0);
      expect(cell!.clipWidthPt).toBeLessThanOrEqual(120); expect(cell!.offsets.map(part => part[2]).join('')).toBe(index === 5 ? 'Ab' : 'Alpha beta gamma delta epsilon');
    } else expect(cell).toBeUndefined();
    expect(result.completeRows).toContain(1); expect(result.retainedSources).toBe(0);
  }
  const encoder = new TextEncoder(); files.set('styles.css', encoder.encode(new TextDecoder().decode(files.get('styles.css')!) + 'tr[data-row="1"] td[data-column="1"]{--office-cell-text-rotation:135}'));
  const rotated = await page.evaluate(entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 120, heightPt: 110, scrollLeftPt: 140, scrollTopPt: 0 } }]), encodeTestPackage(files));
  expect(rotated.results[0]!.overflowCount).toBeGreaterThan(0); expect(rotated.results[0]!.cells.find(cell => cell.row === 1 && cell.column === 1)).toBeDefined();
  expect(rotated.unchanged).toBe(true); expect(report.unchanged).toBe(true);
});

for (const rtl of [false, true]) test(`native center across selection preserves cell geometry and source in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), { documentId: 'continuous-span' });
  const report = await page.evaluate(({ entries, rtl }) => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-6', viewport: { widthPt: 300, heightPt: 120, scrollLeftPt: 0, scrollTopPt: 0 } }], rtl), { entries: encodeTestPackage(files), rtl });
  const cell = report.results[0]!.cells.find(cell => cell.row === 1 && cell.column === 1)!;
  expect(cell.shape.widthPt).toBeCloseTo(cell.rect.width * 3, 1); expect(cell.clipWidthPt).toBeCloseTo(cell.shape.widthPt, 1);
  expect(cell.text).toBe('Ab'); expect(cell.offsets).toEqual([[0, 2, 'Ab']]);
  expect(report.unchanged).toBe(true);
});

test('RTL overflow follows physical text direction even when the offscreen source is in a later logical column', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), { documentId: 'rtl-offscreen' });
  const report = await page.evaluate(entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [25, 26, 27, 28, 29].map(index => ({ sheetId: 'sheet-' + index, viewport: { widthPt: 30, heightPt: 110, scrollLeftPt: 0, scrollTopPt: 0 } })), true), encodeTestPackage(files));
  for (const result of report.results) {
    const cell = result.cells.find(cell => cell.row === 1 && cell.column === 1)!;
    expect(cell, result.sheetId).toBeDefined(); expect(cell.rect.x).toBeGreaterThan(30); expect(cell.clipWidthPt).toBeLessThanOrEqual(30);
    expect(result.overflowCount).toBeGreaterThan(0); expect(cell.offsets.map(part => part[2]).join('')).toBe('Alpha beta gamma delta epsilon');
  }
  expect(report.unchanged).toBe(true);
});

test('offscreen implicit rotated cells keep their skewed fills and shared borders within the pane budget', async ({ page }) => {
  const files = new Map(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), { documentId: 'implicit-skew' }));
  const encoder = new TextEncoder(), decoder = new TextDecoder(), metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.columnDimensions = [{ from: 0, to: 0, unit: 'characters', cssClass: 'blank-skew' }];
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  files.set('styles.css', encoder.encode(decoder.decode(files.get('styles.css')!) + '.blank-skew{--office-cell-text-rotation:15;--office-cell-fill-kind:solid;--office-cell-fill-color:#123456;'
    + ['left', 'right', 'top', 'bottom'].map(side => `--office-cell-border-${side}-kind:thin;--office-cell-border-${side}-color:#ff00ff;`).join('') + '}'));
  const model = openDocument(files, { kind: 'workbook' }); expect(model.getCell('sheet-1', 1, 0)).toBeUndefined();
  const report = await page.evaluate(entries => window.workbookGridHost.windowLayoutMeasurement(entries,
    [{ sheetId: 'sheet-1', viewport: { widthPt: 120, heightPt: 110, scrollLeftPt: 160, scrollTopPt: 0 } }]), encodeTestPackage(files));
  const result = report.results[0]!;
  expect(result.overflowFills).toEqual(expect.arrayContaining([expect.objectContaining({ color: 'rgba(18,52,86,1)' })]));
  expect(result.overflowFills.every(fill => fill.widthPt <= 120)).toBe(true); expect(result.borderColors).toContain('rgba(255,0,255,1)');
  expect(result.accessible).toBeLessThan(100); expect(result.retainedSources).toBe(0); expect(report.unchanged).toBe(true);
});

test('the combined cell-fragment budget counts offscreen origins in addition to the visible grid', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), { documentId: 'fragment-budget' });
  const entries = encodeTestPackage(files);
  const result = await page.evaluate(async entries => {
    const view = { sheetId: 'sheet-25', viewport: { widthPt: 30, heightPt: 60, scrollLeftPt: 0, scrollTopPt: 15 }, maxCellFragments: 1 };
    let error = ''; try { await window.workbookGridHost.windowLayoutMeasurement(entries, [view], true); error = 'accepted'; } catch (failure) { error = (failure as { code: string }).code; }
    const accepted = await window.workbookGridHost.windowLayoutMeasurement(entries, [{ ...view, maxCellFragments: 2 }], true);
    return { error, count: accepted.results[0]!.accessible, overflow: accepted.results[0]!.overflowCount, unchanged: accepted.unchanged };
  }, entries);
  expect(result).toEqual({ error: 'ResourceLimitExceeded', count: 1, overflow: 1, unchanged: true });
});


test('stable automatic row heights use all actual columns, native wrapping and font sizes without changing source', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), { documentId: 'automatic-row-flow' });
  const report = await page.evaluate(entries => window.workbookGridHost.autoRowMeasurement(entries, ['sheet-49', 'sheet-50', 'sheet-51']),
    encodeTestPackage(files));
  expect(report).toMatchObject({ unchanged: true, retained: 0, errors: ['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'OperationCancelled'] });
  expect(report.largest).toBeLessThanOrEqual(8192);
  const [wrapped, lines, large] = report.results;
  expect(wrapped!.heights[1]).toBeGreaterThan(55); expect(wrapped!.heights[1]).toBeLessThan(72);
  expect(lines!.heights[1]).toBeGreaterThan(30); expect(lines!.heights[1]).toBeLessThan(48);
  expect(large!.heights[1]).toBeGreaterThan(21); expect(large!.heights[1]).toBeLessThan(29);
  for (const result of report.results) { expect(result.heights[0]).toBe(15); expect(result.after).toBe(15 + result.heights[1]!); expect(result.segments).toBe(3); }
});


test('automatic row indexing respects native fixed defaults, exceptions, hidden rows, merges and distant content', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-row-heights.xlsx'), { documentId: 'native-auto-height-cases' });
  const report = await page.evaluate(entries => window.workbookGridHost.autoRowMeasurement(entries, Array.from({ length: 21 }, (_, i) => 'sheet-' + (i + 1))),
    encodeTestPackage(files));
  expect(report.unchanged).toBe(true); expect(report.retained).toBe(0);
  const heights = report.results.map(result => result.heights[1]);
  expect(heights[0]).toBeGreaterThan(55); expect(heights[1]).toBe(18); expect(heights[2]).toBe(heights[0]);
  expect(heights[3]).toBe(12); expect(heights[4]).toBe(0); expect(heights[5]).toBe(heights[0]);
  expect(heights[6]).toBe(15); expect(heights[7]).toBe(15); expect(heights[8]).toBeGreaterThan(90);
  expect(heights[9]).toBeGreaterThan(30); expect(heights[10]).toBeGreaterThan(21); expect(heights[11]).toBeGreaterThan(30);
  expect(heights[12]).toBe(heights[0]); expect(heights[13]).toBe(15); expect(report.results[13]!.farHeight).toBe(heights[0]);
  for (const index of [14, 15, 16, 17, 18, 20]) { expect(heights[index]).toBeGreaterThan(21); expect(heights[index]).toBeLessThan(29); }
  expect(heights[19]).toBeGreaterThanOrEqual(heights[18]!);
  expect(report.results[16]!.heights.slice(0, 3)).toEqual([heights[16], heights[16], heights[16]]);
});


test('automatic row fonts retain explicit sheet defaults independently of a larger Normal font and smaller explicit cells', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-row-font-floor.xlsx'), { documentId: 'native-row-font-floor' });
  const report = await page.evaluate(entries => window.workbookGridHost.autoRowMeasurement(entries, Array.from({ length: 12 }, (_, index) => 'sheet-' + (index + 1))),
    encodeTestPackage(files));
  expect(report.unchanged).toBe(true); expect(report.retained).toBe(0);
  for (const [index, result] of report.results.entries()) {
    expect(result.heights[1]).toBe(index < 6 ? 15 : 30);
    expect(result.heights[0]).toBe(index < 6 ? report.results[0]!.heights[0] : 30);
    expect(result.heights[0]).toBeGreaterThanOrEqual(21);
  }
});

test('automatic rows discover implicit table and conditional fonts and validation-only current formula text', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-row-appearance.xlsx'), { documentId: 'native-row-appearance' });
  const report = await page.evaluate(entries => window.workbookGridHost.autoRowMeasurement(entries, Array.from({ length: 9 }, (_, index) => 'sheet-' + (index + 1))),
    encodeTestPackage(files));
  expect(report.unchanged).toBe(true); expect(report.retained).toBe(0);
  const heights = report.results.map(result => result.heights[1]!);
  for (const index of [0, 2, 3, 4, 8]) { expect(heights[index]).toBeGreaterThan(21); expect(heights[index]).toBeLessThan(29); }
  expect(heights[1]).toBe(15); expect(report.results[2]!.heights[2]).toBe(heights[2]);
  expect(report.results[3]!.heights[3]).toBe(heights[3]); expect(heights[5]).toBeGreaterThan(32);
  expect(heights[6]).toBeGreaterThan(30); expect(heights[7]).toBe(15);
  expect(report.reads).toBeLessThan(50); expect(report.errors).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'OperationCancelled']);
});

test('automatic row prototypes validate every source and release all nodes on malformed replies and in-flight cancellation', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-row-heights.xlsx'), { documentId: 'row-prototype-validation' });
  const report = await page.evaluate(entries => window.workbookGridHost.autoRowSourceSafety(entries), encodeTestPackage(files));
  expect(report).toEqual({ errors: ['InvalidDocument', 'InvalidDocument', 'InvalidDocument', 'InvalidDocument', 'OperationCancelled'], retained: [0, 0, 0, 0, 0], executed: false, unchanged: true });
});


test('hidden columns use their stored widths for automatic row height while display columns and drawing anchors remain shared', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-hidden-row-heights.xlsx'), { documentId: 'hidden-auto-height' });
  const entries = encodeTestPackage(files);
  const report = await page.evaluate(entries => window.workbookGridHost.autoRowMeasurement(entries, Array.from({ length: 12 }, (_, i) => 'sheet-' + (i + 1))), entries);
  expect(report.unchanged).toBe(true); expect(report.retained).toBe(0);
  const heights = report.results.map(result => result.heights[1]!);
  for (const index of [0, 2, 4]) expect(heights[index + 1]).toBe(heights[index]);
  expect(heights[2]).toBeGreaterThan(heights[0]!); expect(heights[4]).toBeLessThan(heights[0]!); expect(heights[6]).toBe(heights[0]);
  for (const index of [7, 8]) { expect(heights[index]).toBeGreaterThan(21); expect(heights[index]).toBeLessThan(29); }
  expect(report.results[7]!.heights.slice(0, 3)).toEqual([heights[7], heights[7], heights[7]]);
  expect(heights.slice(9)).toEqual([12, 0, 15]);
  for (const [index, result] of report.results.entries()) {
    expect(result.columnWidths[1] === 0).toBe(index >= 6 || index % 2 === 1);
    expect(result.drawing).toEqual({ x: result.columnWidths[0]! + result.columnWidths[1]! + 2, y: result.after + 3, width: 30, height: 20 });
  }
  expect(report.largest).toBeLessThanOrEqual(4096); expect(report.errors).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'OperationCancelled']);
  const safety = await page.evaluate(entries => window.workbookGridHost.autoRowSourceSafety(entries, true), entries);
  expect(safety).toEqual({ errors: ['InvalidDocument', 'InvalidDocument', 'InvalidDocument', 'InvalidDocument', 'OperationCancelled'], retained: [0, 0, 0, 0, 0], executed: false, unchanged: true });
});


test('accounting underlines fill the cell while preserving font metrics, source text and complete decoration budgets', async ({ page }) => {
  const base: TextLayoutCase = { text: 'Ab', type: 'text', family: 'Liberation Sans', sizePt: 11, widthPt: 120, heightPt: 60, horizontal: 'left', vertical: 'bottom', readingOrder: 'context', indent: 0, rightToLeft: false, inkOnly: true };
  const variants: Partial<TextLayoutCase>[] = [{}, { horizontal: 'right' }, { horizontal: 'center' }, { horizontal: 'general', rightToLeft: true },
    { wrap: true, widthPt: 50, text: 'Alpha beta gamma delta' }, { text: 'First\nSecond', vertical: 'top' },
    { shrink: true, widthPt: 22, text: 'Long accounting text' }, { rotation: 45, heightPt: 100 }, { rotation: 135, heightPt: 100 },
    { rotation: 90, heightPt: 130 }, { rotation: 255, heightPt: 130, text: 'ABC' }, { horizontal: 'fill', text: 'A' },
    { vertical: 'distributed', text: 'First\nSecond', heightPt: 100 }, { text: 'A中禰󠄀', splitOffsets: [1, 2] }, { text: '😀', family: 'Noto Color Emoji' }, { text: '󠄀' }];
  const cases: TextLayoutCase[] = ['singleAccounting', 'doubleAccounting'].flatMap(underline => variants.map(variant => ({ ...base, ...variant, underline: underline as 'singleAccounting' | 'doubleAccounting' })));
  const report = await page.evaluate(cases => window.workbookGridHost.textLayoutMeasurement(cases), cases);
  expect(report.unchanged).toBe(true); expect(report.retained).toBe(0); expect(report.atomic).toBe(true);
  for (const [index, item] of report.measurements.entries()) {
    if (cases[index]!.horizontal === 'fill') expect(item.text).toMatch(/^A+$/); else expect(item.text).toBe(cases[index]!.text); expect(item.visibleText).toBeDefined(); expect(item.materializedNodes).toBe(item.actualNodes); expect(item.codeUnits).toBe(item.actualCodeUnits);
    expect(item.accounting.length).toBeGreaterThan(0); expect(item.accounting.every(line => line.style === (cases[index]!.underline === 'doubleAccounting' ? 'double' : 'solid') && line.text === '' && line.strokes === (cases[index]!.underline === 'doubleAccounting' ? 2 : 1))).toBe(true);
  }
  const gallery = page.locator('#workbook-text-layout-probe'), screenshot = await gallery.screenshot();
  const bounds = await gallery.boundingBox(); expect(bounds).not.toBeNull();
  const pixels = await page.evaluate(async ({ base64, cells, origin }) => {
    const image = new Image(); image.src = 'data:image/png;base64,' + base64; await image.decode();
    const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height; const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
    const data = context.getImageData(0, 0, image.width, image.height).data;
    return cells.map(cell => {
      const rows: number[] = []; let painted = 0, escaped = 0, left = Infinity, right = -Infinity;
      const fromX = Math.round(cell.x - origin.x), fromY = Math.round(cell.y - origin.y), width = Math.round(cell.width), height = Math.round(cell.height);
      for (let y = -2; y < height + 2; y++) { let count = 0; for (let x = -2; x < width + 2; x++) {
        const px = fromX + x, py = fromY + y; if (px < 0 || py < 0 || px >= image.width || py >= image.height) continue;
        const at = (py * image.width + px) * 4;
        if (data[at]! > 180 && data[at + 1]! < 240 && data[at + 2]! > 180) { painted++; if (x < 0 || x >= width || y < 0 || y >= height) escaped++; else { count++; left = Math.min(left, x); right = Math.max(right, x); } }
      } if (count >= width - 2) rows.push(y); }
      let groups = 0; rows.forEach((row, i) => { if (!i || row > rows[i - 1]! + 1) groups++; });
      return { painted, escaped, left, right, rows, groups, width };
    });
  }, { base64: screenshot.toString('base64'), cells: report.measurements.map(item => item.cell), origin: bounds! });
  for (const [index, ink] of pixels.entries()) {
    expect(ink.painted, 'case ' + index).toBeGreaterThan(10); expect(ink.escaped, 'case ' + index).toBe(0);
    if (![7, 8, 9, 10].includes(index % variants.length)) { expect(ink.left, 'case ' + index).toBeLessThanOrEqual(1); expect(ink.right, 'case ' + index).toBeGreaterThanOrEqual(ink.width - 2); }
    // At very small shrink factors two strokes can occupy one raster row.
    if ([0, 1, 2, 3, 6, 11].includes(index % variants.length) && report.measurements[index]!.scale >= .5) expect(ink.groups, 'case ' + index).toBe(index < variants.length ? 1 : 2);
  }
  for (const limits of [{ maxNodes: 5 }, { maxTextUnits: 1 }]) await expect(page.evaluate(({ item, limits }) => window.workbookGridHost.textLayoutMeasurement([item], limits), { item: cases[0]!, limits })).rejects.toThrow();
});


test('independent native accounting styles render from text and numeric cells in LTR and RTL windows including cached candidates', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-cell-flow.xlsx'), { documentId: 'native-accounting-window' });
  const entries = encodeTestPackage(files);
  const ids = [21, 22, 23, 24, 45, 46, 47, 48], viewport = { widthPt: 220, heightPt: 130, scrollLeftPt: 0, scrollTopPt: 0, overscanRows: 0, overscanColumns: 0 };
  const report = await page.evaluate(({ entries, ids, viewport }) => window.workbookGridHost.windowLayoutMeasurement(entries, ids.map(id => ({ sheetId: 'sheet-' + id, viewport }))), { entries, ids, viewport });
  expect(report.unchanged).toBe(true);
  for (const [index, item] of report.results.entries()) {
    const cell = item.cells.find(cell => cell.row === 1 && cell.column === 1)!;
    expect(cell.text).toBe(index % 2 ? '12' : 'Ab'); expect(cell.accounting).toBeDefined();
    expect(cell.accounting!.nodes).toBe(cell.accounting!.actualNodes); expect(cell.accounting!.codeUnits).toBe(cell.accounting!.actualCodeUnits);
    expect(cell.accounting!.width).toBeCloseTo(cell.rect.width, 1);
    expect(cell.accounting!.lines).toHaveLength(1);
    expect(cell.accounting!.lines[0]!.style).toBe([21, 22, 45, 46].includes(ids[index]!) ? 'solid' : 'double');
    expect(cell.accounting!.lines[0]!.width).toBeCloseTo(cell.rect.width, 1);
    expect(item.retainedSources).toBe(0); expect(item.retainedMeasurements).toBe(0);
  }
  const cached = await page.evaluate(({ entries, viewport }) => window.workbookGridHost.windowLayoutMeasurement(entries, [22, 24, 22, 24].map(id => ({ sheetId: 'sheet-' + id, viewport })), true), { entries, viewport });
  expect(cached.unchanged).toBe(true);
  for (const item of cached.results) { const cell = item.cells.find(cell => cell.row === 1 && cell.column === 1)!; expect(cell.accounting!.nodes).toBe(cell.accounting!.actualNodes); expect(cell.text).toBe('12'); }
  expect(cached.cacheSafety.every(result => result.rejected === 'InvalidDocument' && result.rejectedNumeric === 'InvalidDocument')).toBe(true);
});


test('native underline ink participates in automatic row height while Calc format losses remain separate source facts', async ({ page }) => {
  const results = [];
  for (const suffix of ['', '-producer']) {
    const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-accounting-underlines' + suffix + '.xlsx'), { documentId: 'accounting-auto' + suffix });
    const result = await page.evaluate(entries => window.workbookGridHost.autoRowMeasurement(entries, Array.from({ length: 20 }, (_, index) => 'sheet-' + (index + 1))), encodeTestPackage(files));
    expect(result.unchanged).toBe(true); expect(result.retained).toBe(0);
    const heights = result.results.map(item => item.heights[1]!);
    expect(heights[8]).toBe(15); expect(heights[9]).toBe(15);
    for (let index = 0; index < 4; index++) { expect(heights[index]).toBeGreaterThanOrEqual(15); expect(heights[index + 4]).toBeGreaterThan(heights[index]!); }
    expect(heights[14]).toBe(15); expect(heights[19]).toBe(15);
    results.push(heights);
  }
  expect(results[0]!.slice(0, 10)).toEqual(results[1]!.slice(0, 10));
});
