import { openSdkHost } from './sdk-host.js';
import type { measureChartClipping } from './fixtures/workbook-chart-clipping.js';
import type { measureChartNumberFields, measureNativeChartLabels, measureControlledChartLabels } from './fixtures/workbook-chart-labels.js';
import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { importXlsx } from "@doffi/office";
import { openDocument, createWorkbookNumberFormatter } from "@doffi/core";
import type { WorkbookSnapshot } from '@doffi/core';
import { readXlsxRecords } from '../../packages/office/src/xlsx/read.js';
import { readDocumentZip, writeOpcZip } from '../../packages/office/src/document-zip.js';
import { projectXlsxRecords } from '../../packages/office/tests/helpers/xlsx-record-model.js';
import type { measureContentSource, rejectContentSource } from './fixtures/workbook-content-source.js';
import type { measureImageView, measurePictureFills, disposeImageView } from './fixtures/workbook-image-view.js';
import type { measureObjectGeometry } from './fixtures/workbook-object-geometry.js';
import type { measureDrawingGradients } from './fixtures/drawing-gradient.js';
import type { measureDrawingFontProbeReuse, measureDrawingText, measureDrawingTextLayout, measureNativeDrawingText, measureDrawingTextFallback, measureDrawingTextOverflow, measureDrawingOverflowEdges } from './fixtures/workbook-drawing-text.js';
import type { measureWorkbookSurface, mountWorkbookSurfacePixels, disposeWorkbookSurfacePixels } from './fixtures/workbook-surface.js';
import type { measureWorkbookShadows, mountWorkbookShadowPixels, disposeWorkbookShadowPixels } from './fixtures/workbook-shadow.js';
import type { measureWorkbookCharts, measureWorkbookChartAxes, measureWorkbookChartFiniteScales, measureNativeFiniteCharts, measureWorkbookChartTextFonts, measureWorkbookChartMarkers, measureWorkbookChartGrid, measureWorkbookChartColors, measureWorkbookChartTitles, measureWorkbookChartPositions, measureWorkbookChartTitleContents, measureWorkbookComposedChart, measureWorkbookChartFormats } from './fixtures/workbook-chart.js';
import { bindWorkbookChartData } from "../../packages/render/src/workbook/chart-values.js";
import type { BasicChartData } from '../../packages/render/src/chart.js';
import { modelChartSource, DEFAULT_CHART_READ_LIMITS } from '../../packages/render/src/workbook/chart-source.js';

declare global { interface Window { workbookChartHost: { measureChartNumberFields: typeof measureChartNumberFields; measureChartClipping: typeof measureChartClipping; measureControlledChartLabels: typeof measureControlledChartLabels; measureNativeChartLabels: typeof measureNativeChartLabels; measureWorkbookChartAxes: typeof measureWorkbookChartAxes; measureWorkbookChartFiniteScales: typeof measureWorkbookChartFiniteScales; measureNativeFiniteCharts: typeof measureNativeFiniteCharts; measureWorkbookChartTextFonts: typeof measureWorkbookChartTextFonts; measureWorkbookCharts: typeof measureWorkbookCharts; measureWorkbookChartGrid: typeof measureWorkbookChartGrid; measureWorkbookChartMarkers: typeof measureWorkbookChartMarkers; measureWorkbookChartColors: typeof measureWorkbookChartColors; measureWorkbookChartTitles: typeof measureWorkbookChartTitles; measureWorkbookChartTitleContents: typeof measureWorkbookChartTitleContents; measureWorkbookComposedChart: typeof measureWorkbookComposedChart; measureWorkbookChartFormats: typeof measureWorkbookChartFormats; measureWorkbookChartPositions: typeof measureWorkbookChartPositions } } }

declare global { interface Window { workbookShadowHost: { measureWorkbookShadows: typeof measureWorkbookShadows; mountWorkbookShadowPixels: typeof mountWorkbookShadowPixels; disposeWorkbookShadowPixels: typeof disposeWorkbookShadowPixels } } }

declare global { interface Window { workbookContentHost: { measureContentSource: typeof measureContentSource; rejectContentSource: typeof rejectContentSource; measureImageView: typeof measureImageView; measurePictureFills: typeof measurePictureFills; disposeImageView: typeof disposeImageView; measureObjectGeometry: typeof measureObjectGeometry; measureDrawingGradients: typeof measureDrawingGradients; measureDrawingFontProbeReuse: typeof measureDrawingFontProbeReuse; measureDrawingText: typeof measureDrawingText; measureDrawingTextLayout: typeof measureDrawingTextLayout; measureNativeDrawingText: typeof measureNativeDrawingText; measureDrawingTextFallback: typeof measureDrawingTextFallback; measureDrawingTextOverflow: typeof measureDrawingTextOverflow; measureDrawingOverflowEdges: typeof measureDrawingOverflowEdges; measureWorkbookSurface: typeof measureWorkbookSurface; mountWorkbookSurfacePixels: typeof mountWorkbookSurfacePixels; disposeWorkbookSurfacePixels: typeof disposeWorkbookSurfacePixels } } }
const inputs: { name: string; snapshot: WorkbookSnapshot; text: Record<string, string>; resources: Record<string, string> }[] = [];
const chartPlots = new Map<string, BasicChartData[]>();
let referenceTitles: { snapshot: WorkbookSnapshot; bindings: ReturnType<typeof bindWorkbookChartData> };
let formatCharts: typeof referenceTitles;
let code = '';
const axisCharts: { snapshot: WorkbookSnapshot; plots: BasicChartData[] }[] = [];
const gridCharts: { snapshot: WorkbookSnapshot; plots: BasicChartData[] }[] = [];
const markerCharts: { producer: boolean; snapshot: WorkbookSnapshot; plots: BasicChartData[] }[] = [];
const automaticMarkerCharts: { producer: boolean; snapshot: WorkbookSnapshot; plots: BasicChartData[] }[] = [];
const colorCharts: { theme: string; snapshot: WorkbookSnapshot; plots: BasicChartData[] }[] = [];
let nativeText: WorkbookSnapshot;
let overflowText: WorkbookSnapshot;
let anchoredText: WorkbookSnapshot;
const anchoringEvidence = JSON.parse(readFileSync('docs/evidence/xlsx-text-anchoring.json', 'utf8')) as { sourceSha256: string; cases: { paragraphs: string[]; anchorCenter: boolean; observed: { words: { text: string; boundsPt: number[] }[] } }[] };
let nativeShadows: WorkbookSnapshot, nativeBlur: WorkbookSnapshot;
const textEvidence = JSON.parse(readFileSync('docs/evidence/xlsx-drawing-text-semantics.json', 'utf8')) as { sourceSha256: string; cases: { fontPt: number | null; endPt: number | null; paragraphs: string[]; listStyle: string; observed: { lineAdvancePt: number; words: { text: string; boundsPt: number[] }[] } }[] };
const shadowEvidence = JSON.parse(readFileSync('docs/evidence/xlsx-drawing-shadow-semantics.json', 'utf8')) as { sourceSha256: string; cases: { rotation: number; flipH: boolean; flipV: boolean; rotateWithShape: boolean; direction: number; distancePt: number; observed: { offsetPt: number[] } }[] };
const blurEvidence = JSON.parse(readFileSync('docs/evidence/xlsx-shadow-blur-semantics.json', 'utf8')) as { sourceSha256: string; cases: { blurPt: number; observed: { sharpEdgeOffsetPt?: number; samples: { offsetPt: number; rgb: number[] }[] } }[] };
test.beforeAll(async () => {
  for (const suffix of ['', '-producer']) {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-axes' + suffix + '.xlsx'), { documentId: 'axes' + suffix }), { kind: 'workbook' }), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    const plots = bindWorkbookChartData(snapshot, response, formatter).map(binding => { if (binding.kind !== 'available') throw Error('Axis data unavailable'); return binding.data; });
    formatter.dispose(); axisCharts.push({ snapshot, plots });
  }

  for (const producer of [false, true]) {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-auto-markers' + (producer ? '-producer' : '') + '.xlsx'), { documentId: 'auto-markers-' + producer }), { kind: 'workbook' }), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    const plots = bindWorkbookChartData(snapshot, response, formatter).map(binding => { if (binding.kind !== 'available') throw Error('Automatic marker data unavailable'); return binding.data; });
    formatter.dispose(); automaticMarkerCharts.push({ producer, snapshot, plots });
  }
  for (const suffix of ['', '-producer']) {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-grid' + suffix + '.xlsx'), { documentId: 'grid' + suffix }), { kind: 'workbook' }), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    const plots = bindWorkbookChartData(snapshot, response, formatter).map(binding => { if (binding.kind !== 'available') throw Error('Grid data unavailable'); return binding.data; });
    formatter.dispose(); gridCharts.push({ snapshot, plots });
  }
  for (const producer of [false, true]) {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-markers' + (producer ? '-producer' : '') + '.xlsx'), { documentId: 'markers-' + producer }), { kind: 'workbook' }), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    const plots = bindWorkbookChartData(snapshot, response, formatter).map(binding => { if (binding.kind !== 'available') throw Error('Marker data unavailable'); return binding.data; });
    formatter.dispose(); markerCharts.push({ producer, snapshot, plots });
  }
  for (const theme of ['default', 'alternate']) {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-colors-' + theme + '.xlsx'), { documentId: 'chart-colors-' + theme }), { kind: 'workbook' }), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    const plots = bindWorkbookChartData(snapshot, response, formatter).map(binding => { if (binding.kind !== 'available') throw Error('Independent color source is unavailable'); return binding.data; });
    formatter.dispose(); colorCharts.push({ theme, snapshot, plots });
  }
  overflowText = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-text-overflow.xlsx'), { documentId: 'text-overflow' }), { kind: 'workbook' }).snapshot();
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-content-source.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookDrawingProbe' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected drawing fixture bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  const nativeBytes = execFileSync('python3', ['-c', "import importlib.util,sys; s=importlib.util.spec_from_file_location('drawing_text_inputs','scripts/check-xlsx-drawing-text.py'); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); sys.stdout.buffer.write(m.generate(m.inputs()))"], { maxBuffer: 16 * 1024 * 1024 });
  expect(createHash('sha256').update(nativeBytes).digest('hex')).toBe(textEvidence.sourceSha256);
  nativeText = projectXlsxRecords(await readXlsxRecords(nativeBytes)).model.snapshot();
  const anchoringBytes = readFileSync('packages/office/tests/fixtures/xlsx-text-anchoring.xlsx');
  expect(createHash('sha256').update(anchoringBytes).digest('hex')).toBe(anchoringEvidence.sourceSha256);
  anchoredText = openDocument(await importXlsx(anchoringBytes, { documentId: 'native-text-anchoring' }), { kind: 'workbook' }).snapshot();
  for (const [name, evidence] of [['drawing-effects', shadowEvidence], ['shadow-blur', blurEvidence]] as const) {
    const bytes = execFileSync('python3', ['-c', `import importlib.util,sys; s=importlib.util.spec_from_file_location('shadow_inputs','scripts/check-xlsx-${name}.py'); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); sys.stdout.buffer.write(m.generate(m.inputs()))`], { maxBuffer: 16 * 1024 * 1024 });
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
    const snapshot = projectXlsxRecords(await readXlsxRecords(bytes)).model.snapshot(); if (name === 'drawing-effects') nativeShadows = snapshot; else nativeBlur = snapshot;
  }
  for (const prefix of ['', 'producer/']) for (const name of ['07-three-charts', '08-images-shapes', '09-classic-notes', '10-modern-threads']) {
    const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`))), snapshot = model.snapshot();
    const ids = [...(snapshot.metadata.workbook.stories ?? []).map(story => story.rootNodeId), ...(snapshot.metadata.workbook.objects ?? []).map(object => object.nodeId)];
    inputs.push({ name: prefix + name, snapshot, text: Object.fromEntries(ids.map(id => [id, model.getNode(id)!.text])), resources: Object.fromEntries(snapshot.manifest.assets.map(asset => [asset.path, Buffer.from(model.getAsset(asset.path)!).toString('base64')])) });
    if (name === '07-three-charts') {
      const definitions = snapshot.metadata.workbook.charts!, response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: definitions.map(chart => chart.chartId) }, new AbortController().signal);
      const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
      chartPlots.set(prefix + name, bindWorkbookChartData(snapshot, response, formatter).map(chart => {
        if (chart.kind !== 'available') throw Error('Independent chart source must be available'); return chart.data;
      })); formatter.dispose();
    }
  }
  {
    const files = new Map(readDocumentZip(readFileSync('examples/xlsx/07-three-charts.xlsx'))), part = 'xl/charts/chart1.xml', encoder = new TextEncoder();
    const title = (text: string, scale: number) => `<c:title><c:tx><c:rich><a:bodyPr><a:normAutofit fontScale="${scale}"/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
    const original = new TextDecoder().decode(files.get(part)!);
    const xml = original.replace('<c:axPos val="b"/>', '<c:axPos val="b"/>' + title('季度😀', 75000)).replace('<c:axPos val="l"/>', '<c:axPos val="l"/>' + title('收入禰󠄀', 50000));
    expect(xml).not.toBe(original); files.set(part, encoder.encode(xml));
    const snapshot = projectXlsxRecords(await readXlsxRecords(writeOpcZip(files))).model.snapshot();
    const text = Object.fromEntries(snapshot.metadata.workbook.charts!.map((chart, index) => [chart.nodeId, '图表 ' + chart.kind + (index === 0 ? '季度😀收入禰󠄀' : '')]));
    inputs.push({ name: 'chart-title-layout', snapshot, text, resources: {} });
  }
  {
    const files = new Map(readDocumentZip(readFileSync('examples/xlsx/07-three-charts.xlsx'))), part = 'xl/charts/chart1.xml', encoder = new TextEncoder(), decoder = new TextDecoder();
    const title = (cell: string) => `<c:title><c:tx><c:strRef><c:f>Data!$${cell}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>stale cache</c:v></c:pt></c:strCache></c:strRef></c:tx><c:txPr><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:normAutofit fontScale="75000"/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"><a:defRPr sz="1600"/></a:pPr></a:p></c:txPr></c:title>`;
    files.set(part, encoder.encode(decoder.decode(files.get(part)!).replace(/<c:title>.*?<\/c:title>/, title('A')).replace('<c:axPos val="b"/>', '<c:axPos val="b"/>' + title('B')).replace('<c:axPos val="l"/>', '<c:axPos val="l"/>' + title('C'))));
    const sheet = 'xl/worksheets/sheet1.xml'; files.set(sheet, encoder.encode(decoder.decode(files.get(sheet)!).replace('季度</t>', '当前😀禰󠄀\n标题</t>')));
    const { model } = projectXlsxRecords(await readXlsxRecords(writeOpcZip(files))), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    referenceTitles = { snapshot, bindings: bindWorkbookChartData(snapshot, response, formatter) }; formatter.dispose();
  }
  {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-formats.xlsx'), { documentId: 'chart-formats' }), { kind: 'workbook' }), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    formatCharts = { snapshot, bindings: bindWorkbookChartData(snapshot, response, formatter) }; formatter.dispose();
  }
  for (const kind of ['default', 'inherited', 'zero']) {
    const files = new Map(readDocumentZip(readFileSync('examples/xlsx/08-images-shapes.xlsx'))), decoder = new TextDecoder(), encoder = new TextEncoder();
    const path = 'xl/drawings/drawing1.xml', theme = 'xl/theme/theme1.xml';
    let xml = decoder.decode(files.get(path)!).replace('<a:ln w="12700"', kind === 'zero' ? '<a:ln w="0"' : '<a:ln');
    if (kind !== 'default') xml = xml.replace('<xdr:txBody>', '<xdr:style><a:lnRef idx="1"/><a:fillRef idx="0"/><a:effectRef idx="0"/><a:fontRef idx="none"/></xdr:style><xdr:txBody>');
    files.set(path, encoder.encode(xml)); files.set(theme, encoder.encode(decoder.decode(files.get(theme)!).replace('<a:ln w="12700"', '<a:ln w="50800"')));
    const { model } = projectXlsxRecords(await readXlsxRecords(writeOpcZip(files)));
    inputs.push({ name: 'line-' + kind, snapshot: model.snapshot(), text: {}, resources: {} });
  }
});

test('workbook chart graphics use current values and isolated source styles with bounded failure cleanup', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('07-three-charts'))) {
    const results = await page.evaluate(({ snapshot, plots }) => window.workbookChartHost.measureWorkbookCharts(snapshot, plots), { snapshot: input.snapshot, plots: chartPlots.get(input.name)! });
    expect(results.map(result => result.kind)).toEqual(['column', 'line', 'pie']);
    for (const result of results) {
      expect(result.reported).toBe(result.nodes); expect(result.peak).toBeGreaterThan(result.nodes); expect(result.exact).toBe(true); expect(result.unchanged).toBe(true); expect(result.sourceClasses).toBe(0);
      expect(result.svgSize).toEqual([480, 293]); expect(result.fonts.every(value => Math.abs(parseFloat(value) - 13) < .02), JSON.stringify(result.fonts)).toBe(true);
      expect(result.colors.every(value => value.replace(/\s/g, '') === 'rgb(12,34,56)'), JSON.stringify(result.colors)).toBe(true);
      expect(result.errors).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'format rejected', 'InvalidOperation', 'InvalidOperation', 'InvalidOperation', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'accepted']);
      expect(result.negativePies).toHaveLength(result.kind === 'pie' ? 6 : 0);
      for (const pie of result.negativePies) {
        expect(pie.geometry).toEqual(pie.absolute); expect(pie.labels).toEqual(pie.values.filter(value => value !== null && value !== 0));
        expect(pie.fractions.every(value => value > 0 && value <= 1)).toBe(true);
        expect(pie.fractions.reduce((a, b) => a + b, 0)).toBeCloseTo(pie.labels.length ? 1 : 0, 12);
        const negativePoints = pie.values.filter(value => value !== null && value < 0).length;
        expect(pie.diagnostics).toEqual(negativePoints ? [{ code: 'negative-pie-magnitude', chartId: expect.any(String), negativePoints }] : []);
        expect(pie.positiveDiagnostics).toEqual([]); expect(pie.immutable).toBe(true);
      }
      for (const fraction of result.smallScales) expect(fraction).toBeGreaterThan(.7);
      expect(result.zeroScales).toEqual(result.kind === 'column' ? [0, 0] : []);
      expect(result.titleProjection.map(part => part.text).join('')).toBe(result.titleText);
      for (const part of result.titleProjection) { expect(part.owned).toBe(true); expect(part.to - part.from).toBe(part.text.length); expect(result.titleText!.slice(part.from, part.to)).toBe(part.text); }
      expect(result.text).toEqual(expect.arrayContaining(['第一季度', '第二季度', '第三季度']));
    }
    expect(results[0]!.bars).toHaveLength(6); expect(results[0]!.bars[2]! / results[0]!.bars[0]!).toBeCloseTo(4, 4);
  }
});

test('workbook rich chart titles share source-aware text layout, fit budgets and failure isolation', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('07-three-charts') || input.name === 'chart-title-layout')) {
    const result = await page.evaluate(snapshot => window.workbookChartHost.measureWorkbookChartTitles(snapshot), input.snapshot);
    expect(result).toMatchObject({ unchanged: true, cleaned: true, errors: ['ResourceLimitExceeded', 'InvalidOperation', 'InvalidDocument', 'InvalidOperation', 'ResourceLimitExceeded'] });
    expect(result.native.length).toBeGreaterThanOrEqual(3);
    for (const title of result.native) {
      expect(title.unchanged && title.owned).toBe(true); expect(title.nodes).toBe(title.declared); expect(title.fonts).toBeGreaterThan(0);
      expect(title.sourceId).toBe(title.expectedId); expect(title.widthPt).toBe(240); expect(title.units).toEqual(title.lengths);
      expect(title.reflected).toBe('matrix(1, 0, 0, 1, 0, 0)');
    }
    if (input.name === 'chart-title-layout') {
      expect(result.native.filter(title => title.role !== 'title').map(title => ({ role: title.role, text: title.text, scale: title.scale }))).toEqual([
        { role: 'category-axis-title', text: '季度😀', scale: .75 }, { role: 'value-axis-title', text: '收入禰󠄀', scale: .5 },
      ]);
    }
    expect(result.fits.every(item => item.unchanged && item.sourceText === 'Long title with 中文😀禰󠄀 and several words that need multiple lines.')).toBe(true);
    expect(result.fits[0]!.overflow).toBe(true); expect(result.fits[0]!.scale).toBe(1);
    expect(result.fits[1]!.scale).toBeLessThan(1); expect(result.fits[1]!.width).toBe(120); expect(result.fits[1]!.height).toBe(30);
    expect(result.fits[2]!.height).toBeGreaterThan(30); expect(result.fits[2]!.scale).toBe(1);
  }
});

for (const rtl of [false, true]) test(`cropped worksheet pictures retain full rotated and flipped geometry across four ${rtl ? 'RTL' : 'LTR'} panes`, async ({ page }, info) => {
  for (const input of inputs.filter(input => input.name.endsWith('08-images-shapes'))) {
    const result = await page.evaluate(({ snapshot, resources, rtl }) => window.workbookContentHost.measureImageView(snapshot, resources, rtl), { ...input, rtl });
    expect(result).toMatchObject({ invalid: 'InvalidDocument', cancelled: 'OperationCancelled', connected: 'InvalidOperation', kept: true, cleared: true, badGeometry: ['InvalidDocument', 'InvalidDocument', 'ResourceLimitExceeded'], extremeAccepted: true, peak: 8, decoded: 17 });
    const [cropX, cropY] = input.name.startsWith('producer/') ? [.10035, .05039] : [.1, .05];
    expect(result.panes).toHaveLength(4); expect(result.sourceRotation).toBe(15); expect(result.sourceCrops).toEqual([cropX, cropY, cropX, cropY]);
    const cosine = Math.cos(Math.PI / 12), sine = Math.sin(Math.PI / 12);
    for (const pane of result.panes) {
      expect(pane.body.slice(2)).toEqual([200, 150]); expect(pane.natural).toEqual([32, 24]);
      for (const [index, value] of [-cosine, -sine, -sine, cosine].entries()) expect(pane.matrix[index]).toBeCloseTo(value, 5);
      expect(pane.transformed[0]).toBeCloseTo(200 * cosine + 150 * sine, 1); expect(pane.transformed[1]).toBeCloseTo(200 * sine + 150 * cosine, 1);
      for (const [index, value] of [-cropX! / (1 - 2 * cropX!) * 100, -cropY! / (1 - 2 * cropY!) * 100, 100 / (1 - 2 * cropX!), 100 / (1 - 2 * cropY!)].entries()) expect(pane.crop[index]).toBeCloseTo(value, 3);
      const sourceLeft = pane.pane === 'body' || pane.pane === 'frozenRows' ? -10 : 10;
      expect(pane.clip[0]! + pane.body[0]!).toBeCloseTo(rtl ? 240 - sourceLeft - 200 : sourceLeft, 0);
      expect(pane.clip[1]! + pane.body[1]!).toBeCloseTo(pane.pane === 'body' || pane.pane === 'frozenColumns' ? -5 : 5, 0);
    }
    await page.locator('#workbook-image-probe').screenshot({ path: info.outputPath(`${input.name.replace('/', '-')}-${rtl ? 'rtl' : 'ltr'}.png`) });
    await page.evaluate(() => window.workbookContentHost.disposeImageView());
    await expect(page.locator('#workbook-image-probe')).toHaveCount(0);
  }
});
for (const rtl of [false, true]) test(`native picture fills keep page orientation independently of frame rotation and reflections in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-picture-rotation.xlsx'), { documentId: 'picture-fill' }), { kind: 'workbook' }), snapshot = model.snapshot();
  const resources = Object.fromEntries(snapshot.manifest.assets.map(asset => [asset.path, Buffer.from(model.getAsset(asset.path)!).toString('base64')]));
  const result = await page.evaluate(({ snapshot, resources, rtl }) => window.workbookContentHost.measurePictureFills(snapshot, resources, rtl), { snapshot, resources, rtl });
  expect(result).toMatchObject({ unchanged: true, degenerate: [1, 1], errors: ['InvalidDocument', 'InvalidDocument', 'InvalidDocument', 'InvalidOperation', 'InvalidOperation', 'ResourceLimitExceeded'] });
  expect(result.results).toHaveLength(20);
  for (const record of result.results) {
    expect(record.nodes).toBe(record.actualNodes); expect(record.nodes).toBe(record.rotates ? 2 : 3);
    const angle = record.orientation.rotationDegrees*Math.PI/180, sx = record.orientation.flipH ? -1 : 1, sy = record.orientation.flipV ? -1 : 1;
    const expected = record.rotates ? [Math.cos(angle)*sx, Math.sin(angle)*sx, -Math.sin(angle)*sy, Math.cos(angle)*sy] : [sx, 0, 0, sy];
    for (const [index, number] of expected.entries()) expect(record.matrix[index]).toBeCloseTo(number, 5);
  }
  const gallery = page.locator('#workbook-picture-fill-probe'), image = await gallery.screenshot(), box = (await gallery.boundingBox())!;
  const colors = await page.evaluate(async ({ base64, points, origin }) => {
    const image = new Image(); image.src = 'data:image/png;base64,' + base64; await image.decode();
    const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height; const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
    return points.map(samples => samples.map(point => [...context.getImageData(Math.round(point.x - origin.x), Math.round(point.y - origin.y), 1, 1).data]));
  }, { base64: image.toString('base64'), points: result.results.map(record => record.samples), origin: box });
  for (const [index, samples] of colors.entries()) for (const [at, color] of samples.entries()) {
    const expected = [[255, 0, 0], [0, 255, 0], [0, 0, 255], [255, 255, 0]][at]!;
    for (const [channel, value] of expected.entries()) expect(Math.abs(color[channel]! - value), `picture ${index}, sample ${at}`).toBeLessThanOrEqual(5);
    expect(color[3]).toBe(255);
  }
  await page.evaluate(() => window.workbookContentHost.disposeImageView()); await expect(gallery).toHaveCount(0);
});
test.beforeEach(async ({ page }) => { await openSdkHost(page); await page.addScriptTag({ content: code }); });

test('finite out-of-axis segments retain intersections, blank gaps and clipped marker ink without unbounded coordinates',async({page})=>{
  const snapshot=inputs.find(input=>input.name==='07-three-charts')!.snapshot,result=await page.evaluate(snapshot=>window.workbookChartHost.measureChartClipping(snapshot),snapshot);
  expect(result).toMatchObject({unchanged:true,cleaned:true});expect(result.results).toHaveLength(11);
  for(const item of result.results){
    expect(item.nodes).toBe(item.actual);expect(item.invalid).toBe(false);expect(item.escaped).toBe(0);expect(item.preserved).toBe(true);
    if(item.index===0||item.index===8)expect(item.errors).toEqual(['ResourceLimitExceeded','ResourceLimitExceeded']);
    if([2,4].includes(item.index)){expect(item.path).toBe('');expect(item.painted).toBe(0);}else if(item.index<8||item.index===10)expect(item.painted).toBeGreaterThan(50);
    if(item.index<2){expect(item.coordinates.length).toBeGreaterThan(1);for(const point of item.coordinates){expect(point.x).toBeCloseTo(item.plot.x+item.plot.width/2,4);expect(point.y).toBeGreaterThanOrEqual(item.plot.y-3);expect(point.y).toBeLessThanOrEqual(item.plot.y+item.plot.height+3);}}
    if(item.index===10)expect(item.markers).toBe(3);
    if(item.gradientPixels.length)for(const [index,pixel]of item.gradientPixels.entries()){
      const red=(item.index===8?index===0:index===1)?191:64;expect(Math.abs(pixel[0]!-red)).toBeLessThan(6);expect(pixel[1]).toBe(0);expect(Math.abs(pixel[2]!-(255-red))).toBeLessThan(6);expect(pixel[3]).toBe(255);
    }
  }
});

test('native label limits, signed white negative columns and pie leader lines use measured current-source geometry',async({page})=>{
  for(const suffix of ['', '-producer']){
    const model=openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-labels'+suffix+'.xlsx'),{documentId:'native-labels'}),{kind:'workbook'}),snapshot=model.snapshot();
    const response=await modelChartSource(model).read({...DEFAULT_CHART_READ_LIMITS,documentId:snapshot.manifest.documentId,revision:0,chartIds:snapshot.metadata.workbook.charts!.map(chart=>chart.chartId)},new AbortController().signal);
    const formatter=createWorkbookNumberFormatter({dateSystem:snapshot.metadata.workbook.dateSystem});
    const bindings=bindWorkbookChartData(snapshot,response,formatter);formatter.dispose();
    const result=await page.evaluate(({snapshot,bindings})=>window.workbookChartHost.measureNativeChartLabels(snapshot,bindings),{snapshot,bindings});
    expect(result).toMatchObject({unchanged:true,cleaned:true});expect(result.results).toHaveLength(12);
    for(const [index,item] of result.results.entries()){
      expect(item.nodes).toBe(item.actual);expect(item.invalid).toBe(false);
      for(const label of item.labels){expect(label.x).toBeGreaterThanOrEqual(-.01);expect(label.y).toBeGreaterThanOrEqual(-.01);expect(label.x+label.width).toBeLessThanOrEqual(480.01);expect(label.y+label.height).toBeLessThanOrEqual(300.01);}
      if(!suffix&&index<6)expect(item.labels.map(label=>label.text)).toEqual(index%3===2?['-2','1','6']:['1']);
      if(!suffix&&index===5){expect(item.labels[0]!.y+item.labels[0]!.height/2).toBeGreaterThan(item.plot.y+item.plot.height);expect(item.labels[2]!.y+item.labels[2]!.height/2).toBeLessThan(item.plot.y);}
      if(index===8||index===9){
        expect(item.values[0]).toBe(-2);expect(item.bars[0]!.fill).toBe(suffix ? 'rgba(68,114,196,1)' : '#ffffff');
        expect(item.negativePixel).toEqual(suffix?[68,114,196,255]:[255,255,255,255]);
        expect(item.bars[0]!.y).toBeGreaterThan(item.bars[1]!.y);expect(item.labels.some(label=>label.text==='-2')).toBe(true);
      }
      if(!suffix&&index===10)expect(item.leaders).toBe(0);
      if(!suffix&&index===11)expect(item.leaders).toBeGreaterThan(0);
    }
  }
});

test('chart text spacing, explicit axis alignment and wrapped legends use final measured fonts',async({page})=>{
  const model=openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-labels.xlsx'),{documentId:'controlled-labels'}),{kind:'workbook'}),snapshot=model.snapshot();
  const response=await modelChartSource(model).read({...DEFAULT_CHART_READ_LIMITS,documentId:snapshot.manifest.documentId,revision:0,chartIds:snapshot.metadata.workbook.charts!.map(chart=>chart.chartId)},new AbortController().signal),formatter=createWorkbookNumberFormatter({dateSystem:snapshot.metadata.workbook.dateSystem});
  const bindings=bindWorkbookChartData(snapshot,response,formatter);formatter.dispose();
  for(const decoration of [undefined,'solid','double','dotted','dashed','wavy'] as const) {
  const result=await page.evaluate(({snapshot,bindings,decoration})=>window.workbookChartHost.measureControlledChartLabels(snapshot,bindings,decoration),{snapshot,bindings,decoration});
  expect(result).toMatchObject({unchanged:true,cleaned:true,overrideFill:'rgba(68,114,196,1)'});
  for(const item of result.results){expect(item.size).toBe('9');expect(Number(item.spacing)).toBeCloseTo(1,5);expect(item.lineOffsets).toEqual([-7.5,15]);expect(item.anchors).toEqual(Array(3).fill(item.alignment==='left'?'start':'end'));if(item.alignment==='left')expect(item.skipped).toBeGreaterThan(0);else expect(item.skipped).toBe(0);}
  for(const item of result.results){expect(item.nodes).toBe(item.actual);if(decoration)for(const paint of item.decorationPaint){expect(paint).toMatchObject({text:'A long mixed category 中😀Second line',accessible:true,hidden:true,strokes:decoration==='double'?8:4});}}
  expect(result.legend.rect.height).toBeGreaterThan(28);
  expect(result.largeMarker.height).toBe(72);expect(result.largeMarker.rect.height).toBeGreaterThanOrEqual(72);expect(result.largeMarker.nodes).toBe(result.largeMarker.actual);
  for(const box of result.legend.groups){expect(box.x).toBeGreaterThanOrEqual(result.legend.rect.x);expect(box.y).toBeGreaterThanOrEqual(result.legend.rect.y);expect(box.x+box.width).toBeLessThanOrEqual(result.legend.rect.x+result.legend.rect.width+.01);expect(box.y+box.height).toBeLessThanOrEqual(result.legend.rect.y+result.legend.rect.height+.01);}
  }
});

test('finite chart endpoints and subnormal values retain bounded coordinates across direct and composed scales', async ({ page }) => {
  const snapshot = inputs.find(input => input.name === '07-three-charts')!.snapshot;
  const result = await page.evaluate(snapshot => window.workbookChartHost.measureWorkbookChartFiniteScales(snapshot), snapshot);
  expect(result).toMatchObject({ unchanged: true, cleaned: true }); expect(result.results).toHaveLength(24);
  for (const item of result.results) for (const view of [item.basic, item.full]) {
    expect(view.actual).toBe(view.declared); expect(view.invalid).toBe(false); expect(view.points.length).toBeGreaterThan(0);
    for (const y of view.points) { expect(y).toBeGreaterThanOrEqual(view.top - .001); expect(y).toBeLessThanOrEqual(view.top + view.height + .001); }
    if (item.kind === 'line') {
      expect(view.points[0]).toBeCloseTo(item.reverseValues ? view.top : view.top + view.height, 1);
      if (item.index < 2) expect(view.points[2]).toBeCloseTo(item.reverseValues ? view.top + view.height : view.top, 1);
    }
  }
});

test('independent native and Calc finite chart intervals compose current values without overflowing geometry', async ({ page }) => {
  for (const suffix of ['', '-producer']) {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-finite-scales' + suffix + '.xlsx'), { documentId: 'native-finite' }), { kind: 'workbook' }), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    try {
      const bindings = bindWorkbookChartData(snapshot, response, formatter), result = await page.evaluate(({ snapshot, bindings }) => window.workbookChartHost.measureNativeFiniteCharts(snapshot, bindings), { snapshot, bindings });
      expect(result.unchanged).toBe(true); expect(result.results).toHaveLength(14);
      for (const [index, item] of result.results.entries()) {
        expect(item.nodes).toBe(item.actual); expect(item.invalid).toBe(false); expect(item.points).toBe(index===10||index===11?1:3);
        expect(item.values).toEqual(index < 8 ? [-9e307, 0, 9e307] : index<10 ? [0, Number.MIN_VALUE, Number.MIN_VALUE * 2] : [-9e307,.5,9e307]);
      }
    } finally { formatter.dispose(); }
  }
});

test('physical DrawingML gradients preserve diagonal stop coverage and the scaled flag in browser pixels', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookContentHost.measureDrawingGradients());
  expect(result.errors).toEqual(Array(5).fill('InvalidDocument'));
  for (const entry of result.cases) {
    expect(entry.units).toBe('userSpaceOnUse');
    const expected = entry.scaled ? [0, 0, 200, 100] : [25, -25, 175, 125];
    entry.coordinates.forEach((value, index) => expect(value).toBeCloseTo(expected[index]!, 8));
    // Expected physical projection at pixel centers, independent of the SVG implementation.
    const coordinates = [[.5, .5], [199.5, .5], [.5, 99.5], [199.5, 99.5], [99.5, 49.5]];
    entry.pixels.forEach((pixel, index) => {
      const [x, y] = coordinates[index]!, t = entry.scaled ? (2 * x! + y!) / 500 : (x! + y!) / 300;
      const expectedPixel = [255 * (1 - t), 0, 255 * t, 255];
      pixel.forEach((channel, index) => expect(Math.abs(channel - expectedPixel[index]!)).toBeLessThanOrEqual(2));
    });
  }
});

test('workbook geometry shares preset paths, native colors and bounded gradient/arrow primitives', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('08-images-shapes'))) {
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureObjectGeometry(snapshot), input.snapshot);
    expect(result.results.map(object => object.preset)).toEqual(['rect', 'roundRect', 'rect']);
    for (const object of result.results.slice(1)) {
      expect(object.paths).toHaveLength(1);
      expect(object.paths[0]).toMatchObject({ fill: 'rgba(232,240,230,1)', stroke: 'rgba(40,116,84,1)', width: '1' });
      expect(object.textRect!.width).toBeGreaterThan(0); expect(object.textRect!.height).toBeGreaterThan(0);
    }
    expect(result.errors).toEqual(Array(3).fill('ResourceLimitExceeded')); expect(result.nodes).toBe(19);
    expect(new Set(result.ids).size).toBe(result.ids.length); expect(new Set(result.arrows).size).toBe(4);
    expect(result.arrows.every(reference => result.ids.includes(reference!.slice(5, -1)))).toBe(true);
    expect(result.stops).toEqual(['rgba(255,0,0,1)', 'rgba(0,0,255,1)']);
  }
});

test('omitted native line widths inherit the theme before the final default, and explicit zero wins', async ({ page }) => {
  for (const [kind, width] of [['default', '.75'], ['inherited', '4'], ['zero', '0']]) {
    const input = inputs.find(input => input.name === 'line-' + kind)!;
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureObjectGeometry(snapshot), input.snapshot);
    expect(Number(result.results[1]!.paths[0]!.width)).toBe(Number(width));
  }
});

test('drawing text keeps native styles and exact UTF-16 projections across paragraphs, spans and styled breaks', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('08-images-shapes'))) {
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingText(snapshot), input.snapshot);
    for (const object of result.native) {
      expect(object.text).toBe(input.text[object.nodeId]); expect(object.pieces.map(piece => piece.text).join('')).toBe(object.text);
      expect(object).toMatchObject({ unchanged: true, detachedSource: true, inheritedIds: 0 }); expect(object.nodes).toBe(object.actualNodes);
      expect(object.fonts.length).toBeGreaterThan(0);
      // Firefox quantizes the used font size to its subpixel layout grid.
      for (const piece of object.pieces) expect(Math.abs(parseFloat(piece.size) - 14 / .75)).toBeLessThan(.02);
      expect(object.pieces.filter(piece => piece.role === 'east-asia').every(piece => piece.family.includes('Noto Sans CJK SC'))).toBe(true);
    }
    expect(result.pieces.map(piece => piece.text).join('')).toBe('A中󠄀\né😀Tail');
    expect(result.pieces.find(piece => piece.text === '󠄀')).toMatchObject({ role: 'east-asia', from: 2, to: 4, sourceNodeId: 'drawing-r2' });
    expect(result.pieces.find(piece => piece.text === '\n')).toMatchObject({ from: 4, to: 5, sourceNodeId: 'drawing-br', size: '24px' });
    expect(result.pieces.find(piece => piece.text.startsWith('e'))!.weight).toBe('700');
    expect(result.pieces.find(piece => piece.text === 'Tail')).toMatchObject({ from: 11, to: 15, sourceNodeId: 'drawing-p2' });
    expect(result.unchanged).toBe(true); expect(result.exact).toBe(result.nodes);
    expect(result.errors).toEqual(['InvalidOperation', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'InvalidDocument', 'InvalidOperation']);
  }
});

test('drawing text layout uses native insets, bounded auto-fit and readable glyphs inside reflected objects', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('08-images-shapes'))) {
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingTextLayout(snapshot), input.snapshot);
    const inset = input.name.startsWith('producer/') ? 90000 / 12700 / .75 : 9.6;
    result.nativeStyles.insets.forEach((value, index) => expect(value).toBeCloseTo(index < 2 ? inset : inset / 2, 3));
    expect(result.nativeStyles.margin).toBe('0px'); expect(result.nativeStyles.fonts).toBeGreaterThan(0); expect(result.nativeStyles.defaultSize).toBe('24px');
    const [none, normal, shape] = result.cases;
    expect(none).toMatchObject({ width: 100, height: 36, scale: 1, overflow: true });
    expect(normal).toMatchObject({ width: 100, height: 36, overflow: false }); expect(normal!.scale).toBeLessThan(1);
    expect(shape!.height).toBeGreaterThan(36); expect(shape!.geometryHeight).toBeCloseTo(shape!.height, 5); expect(shape!.overflow).toBe(false);
    for (const entry of result.cases) {
      [Math.sqrt(3) / 2, .5, -.5, Math.sqrt(3) / 2].forEach((value, index) => expect(entry.matrix[index]).toBeCloseTo(value, 5));
      expect(entry.text).toBe('中文与 Latin 123，自动调整文字框。'.repeat(5)); expect(entry.nodes).toBe(entry.actualNodes);
    }
    expect(result.leading[1]!.font).toBeCloseTo(result.leading[0]!.font * .5, 3);
    expect(result.leading[1]!.height).toBeCloseTo(result.leading[0]!.height * .375, 0);
    expect(result.listValues.markers).toEqual([{ text: '3.', left: '-10px', position: 'absolute' }, { text: '4.', left: '-10px', position: 'absolute' }]);
    expect(result.listValues.sourceText).toBe('OneTwo'); expect(result.listValues.nodes).toBe(result.listValues.actualNodes);
    expect(result.zeroList).toEqual({ margin: '0px', indent: '0px', markerLeft: '0px', markerWidth: '0px', sourceText: 'OneTwo' });
    expect(result.insetOverflow).toEqual({ width: 1, overflow: true });
    expect(result).toMatchObject({ unchanged: true, cleaned: true, errors: ['ResourceLimitExceeded', 'InvalidOperation', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded'] });
  }
});

test('explicit native drawing text follows independent fixed-font line advances and empty paragraph sizes', async ({ page }) => {
  const actual = await page.evaluate(snapshot => window.workbookContentHost.measureNativeDrawingText(snapshot), nativeText);
  expect(actual).toHaveLength(textEvidence.cases.length);
  for (const [index, entry] of actual.entries()) {
    const reference = textEvidence.cases[index]!;
    expect(entry.unchanged).toBe(true); expect(entry.sourceText).toBe(reference.paragraphs.join('')); expect(entry.placeholders).toBe(reference.paragraphs.flatMap(text => text.split('\n')).filter(text => !text).length);
    expect(entry.fallback).toBe(false);
    expect(entry.nodes).toBe(entry.declared); expect(entry.exact).toBe(entry.textNodes); expect(entry.limited).toBe('ResourceLimitExceeded');
    // Producer-only default font/defPPr differences are recorded separately.
    if (reference.fontPt === null && !['level-22', 'both'].includes(reference.listStyle) || reference.paragraphs.includes('') && reference.endPt === null) continue;
    expect(Math.abs(entry.lineAdvancePt - reference.observed.lineAdvancePt), `case ${index}`).toBeLessThan(.75);
    entry.words.forEach((word, wordIndex) => {
      const expected = reference.observed.words[wordIndex]!; expect(word.text).toBe(expected.text);
      word.boundsPt.forEach((value, axis) => expect(Math.abs(value - expected.boundsPt[axis]!), `case ${index}, ${word.text}, axis ${axis}`).toBeLessThan(1.5));
    });
  }
});

test('native text overflow clips actual ink independently by axis and keeps complete source offsets behind derived ellipses', async ({ page }, info) => {
  const reference = JSON.parse(readFileSync('docs/evidence/xlsx-text-overflow.json', 'utf8')) as { cases: { horizontal: string; vertical: string; anchor: string; paragraphs: string[] }[] };
  const actual = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingTextOverflow(snapshot), overflowText);
  expect(actual).toMatchObject({ invalid: 'InvalidDocument', limited: 'ResourceLimitExceeded', geometryLimited: 'ResourceLimitExceeded', cleaned: true });
  expect(actual.results).toHaveLength(24);
  const image = await page.locator('#workbook-overflow-probe').screenshot({ path: info.outputPath('overflow-ink.png') });
  await info.attach('native-text-overflow', { body: image, contentType: 'image/png' });
  const ink = await page.evaluate(async bytes => {
    const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' })), image = new Image();
    try {
      image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
      const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
      const pixels = context.getImageData(0, 0, image.width, image.height).data;
      return Array.from({ length: 24 }, (_, index) => {
        const x0 = index % 4 * 300, y0 = Math.floor(index / 4) * 180;
        let left = 300, top = 180, right = -1, bottom = -1, count = 0;
        for (let y = 0; y < 180; y++) for (let x = 0; x < 300; x++) {
          const offset = ((y0 + y) * image.width + x0 + x) * 4;
          if (Math.max(pixels[offset]!, pixels[offset + 1]!, pixels[offset + 2]!) < 180) {
            left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y); count++;
          }
        }
        return [left, top, right, bottom, count];
      });
    } finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
  }, [...image]);
  for (const [index, result] of actual.results.entries()) {
    const expected = reference.cases[index]!, bounds = ink[index]!;
    expect(result.unchanged).toBe(true); expect(result.sourceText).toBe(expected.paragraphs.join(''));
    expect(result.nodes).toBe(result.declared); expect(result.markerSourceIds).toBe(false); expect(result.overflow).toBe(true);
    expect(result.markerCount).toBe(expected.vertical === 'ellipsis' ? 1 : 0);
    if (expected.vertical === 'ellipsis' && expected.anchor === 't') {
      expect(result.markerBounds[0]!.top).toBeGreaterThan(12); expect(result.markerBounds[0]!.top).toBeLessThan(34);
      expect(result.markerBounds[0]!.left).toBeGreaterThan(20);
    }
    for (const [from, to, length] of result.offsets) expect(to! - from!).toBe(length);
    if (expected.horizontal === 'clip') { expect(bounds[0]).toBeGreaterThanOrEqual(40); expect(bounds[2]).toBeLessThan(201); expect(result.paintBounds.width).toBe(120); }
    if (expected.vertical !== 'overflow') { expect(bounds[1]).toBeGreaterThanOrEqual(60); expect(bounds[3]).toBeLessThan(114); expect(result.paintBounds).toMatchObject({ y: 0, height: 40 }); }
    if (expected.vertical === 'overflow') expect(expected.anchor === 'b' ? bounds[1]! < 60 : bounds[3]! >= 114).toBe(true);
  }
  // A real derived glyph and its reserved gap must change visible ink.
  for (const start of [0, 12]) for (let offset = 0; offset < 4; offset++) expect(ink[start + 8 + offset]).not.toEqual(ink[start + 4 + offset]);
});

test('native drawing block centering preserves paragraph alignment and vertical anchoring', async ({ page }) => {
  const actual = await page.evaluate(snapshot => window.workbookContentHost.measureNativeDrawingText(snapshot), anchoredText);
  expect(actual).toHaveLength(anchoringEvidence.cases.length);
  for (const [index, item] of actual.entries()) {
    const expected = anchoringEvidence.cases[index]!;
    expect(item).toMatchObject({ unchanged: true, sourceText: expected.paragraphs.join(''), limited: 'ResourceLimitExceeded' });
    expect(item.nodes).toBe(item.declared); expect(item.exact).toBe(item.textNodes);
    for (const [wordIndex, word] of item.words.entries()) {
      const reference = expected.observed.words[wordIndex]!; expect(word.text).toBe(reference.text);
      for (const [axis, value] of word.boundsPt.entries()) expect(Math.abs(value - reference.boundsPt[axis]!), `case ${index + 1}, ${word.text}, axis ${axis}`).toBeLessThan(1.5);
    }
    if (expected.anchorCenter) {
      const left = Math.min(...item.words.map(word => word.boundsPt[0]!)), right = Math.max(...item.words.map(word => word.boundsPt[2]!));
      expect(Math.abs((left + right) / 2 - 60)).toBeLessThan(.5);
    }
  }
});

for (const invalidCanvasShorthand of [false, true]) test(`drawing text fallback accounts for transient probes and empty lines with ${invalidCanvasShorthand ? 'valid CSS outside Canvas shorthand' : 'missing Canvas font bounds'}`, async ({ page }) => {
  const result = await page.evaluate(({ snapshot, invalidCanvasShorthand }) => window.workbookContentHost.measureDrawingTextFallback(snapshot, invalidCanvasShorthand), { snapshot: inputs.find(input => input.name === '08-images-shapes')!.snapshot, invalidCanvasShorthand });
  expect(result).toMatchObject({ error: 'ResourceLimitExceeded', fallback: true, unchanged: true, sourceText: 'Alpha\nBeta', placeholders: 2, cleaned: true });
  expect(result.nodes).toBe(result.declared);
  expect(result.paragraphs[1]).toBeGreaterThan(0);
  expect(result.paragraphs[0]).toBeCloseTo(result.paragraphs[1]! * 2, 1);
  expect(result.paragraphs[2]).toBeCloseTo(result.paragraphs[1]!, 1);
});

test('workbook surfaces isolate source important rules, host styles and concurrent candidate lifecycles', async ({ page }) => {
  const result = await page.evaluate(snapshot => window.workbookContentHost.measureWorkbookSurface(snapshot), inputs.find(input => input.name === '08-images-shapes')!.snapshot);
  expect(result).toMatchObject({ errors: ['ResourceLimitExceeded', 'InvalidOperation', 'OperationCancelled', 'OperationCancelled'], kept: true, independent: true, released: true });
  expect(result.values.map(value => value.sourceColor)).toEqual(['rgb(13, 75, 91)', 'rgb(9, 33, 77)']);
  for (const value of result.values) {
    expect(value).toMatchObject({ frame: [400, 200], view: [400, 200], object: [160, 120], geometry: [160, 120], sourceWidth: '1332px', paragraphFont: '0px', border: '0px', hidden: 'hidden', connectedAfterReadyAbort: true, unchanged: true });
    expect(Math.abs(parseFloat(value.sourceFont) - 14 / .75)).toBeLessThan(.02); expect(value.paragraphWidth).toBeLessThan(160);
    expect(value.selection).toMatchObject({ text: value.selection.expected, range: true, errors: ['InvalidOperation', 'InvalidOperation'], unchanged: true });
  }
});

test('isolated workbook surfaces paint local SVG gradient references in actual page pixels', async ({ page }) => {
  await page.evaluate(snapshot => window.workbookContentHost.mountWorkbookSurfacePixels(snapshot), inputs.find(input => input.name === '08-images-shapes')!.snapshot);
  try {
    const png = await page.locator('#workbook-surface-pixel-probe iframe').screenshot();
    const pixels = await page.evaluate(async bytes => {
      const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' })), image = new Image();
      try {
        image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
        const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
        return { width: image.width, height: image.height, values: [40, 200, 360].map(x => [...context.getImageData(x, 100, 1, 1).data]) };
      } finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
    }, [...png]);
    expect(pixels).toMatchObject({ width: 400, height: 200 });
    for (const [index, x] of [40, 200, 360].entries()) {
      const expected = [255 * (1 - (x + .5) / 400), 0, 255 * (x + .5) / 400, 255];
      pixels.values[index]!.forEach((value, channel) => expect(Math.abs(value - expected[channel]!)).toBeLessThan(2));
    }
  } finally { await page.evaluate(() => window.workbookContentHost.disposeWorkbookSurfacePixels()); }
});

test('worksheet shadows preserve page offsets through native rotations and reflections with bounded raster allocation', async ({ page }) => {
  const result = await page.evaluate(snapshot => window.workbookShadowHost.measureWorkbookShadows(snapshot), nativeShadows);
  expect(result.errors.slice(0, 5)).toEqual(Array(5).fill('InvalidDocument'));
  // Old Chromium saturates enormous CSS numbers before computed-style access;
  // it must still fail admission instead of allocating an enormous filter.
  expect(['InvalidDocument', 'ResourceLimitExceeded']).toContain(result.errors[5]);
  expect(result.errors.slice(6)).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'UnsupportedFeature']); expect(result.restored).toBe(true);
  expect(result.none).toMatchObject({ filter: 'none', outsetPt: 0, rasterPixels: 0, approximateBlur: false });
  expect(result.cases).toHaveLength(shadowEvidence.cases.length);
  result.cases.forEach((actual, index) => {
    const source = shadowEvidence.cases[index]!;
    expect(actual).toMatchObject({ outset: 15, approximate: false, unchanged: true, filterEnabled: true });
    expect(actual.pixels).toBeGreaterThan(0); expect(actual.pixels).toBeLessThan(16_777_216);
    if (!source.rotateWithShape) actual.worldOffsetPt.forEach((value, axis) => expect(Math.abs(value - source.observed.offsetPt[axis]!)).toBeLessThan(.02));
    else {
      // The source flag retains object-space direction. Calc's contrary
      // producer behavior is recorded, not copied into these expected values.
      const x = source.direction === 0 ? (source.flipH ? -15 : 15) : 0, y = source.direction === 90 ? (source.flipV ? -15 : 15) : 0;
      const expected = source.rotation === 0 ? [x, y] : [Math.sqrt(3) / 2 * x - .5 * y, .5 * x + Math.sqrt(3) / 2 * y];
      actual.worldOffsetPt.forEach((value, axis) => expect(value).toBeCloseTo(expected[axis]!, 4));
    }
  });
});

test('worksheet Gaussian shadow preview reports approximation and checks native edge profiles in actual pixels', async ({ page }) => {
  let browserEdgePt = 0;
  const nativeEdgePt = blurEvidence.cases[0]!.observed.sharpEdgeOffsetPt!;
  for (const [index, reference] of blurEvidence.cases.entries()) {
    const result = await page.evaluate(({ snapshot, index }) => window.workbookShadowHost.mountWorkbookShadowPixels(snapshot, index), { snapshot: nativeBlur, index });
    try {
      const png = await page.locator('#workbook-shadow-pixel-probe iframe').screenshot();
      const raster = await page.evaluate(async ({ bytes, offsets }) => {
        const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' })), image = new Image();
        try {
          image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
          const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
          const row = Math.round(70 / .75 - .5), values = offsets.map(offset => [...context.getImageData(Math.round((220 + offset) / .75 - .5), row, 1, 1).data].slice(0, 3));
          let edge = NaN;
          for (let x = Math.floor(220 / .75) - 6; x < Math.floor(220 / .75) + 6; x++) {
            const a = context.getImageData(x, row, 1, 1).data[1]!, b = context.getImageData(x + 1, row, 1, 1).data[1]!;
            if (a < 127.5 && b >= 127.5) { edge = (x + .5 + (127.5 - a) / (b - a)) * .75 - 220; break; }
          }
          return { values, edge };
        } finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
      // Compare the blur kernels after independently measuring each raster's
      // sharp edge. Keep the sub-point placement error separately bounded;
      // do not shift the renderer to reproduce a producer's pixel-grid snap.
      }, { bytes: [...png], offsets: reference.observed.samples.map(sample => sample.offsetPt + (index ? browserEdgePt - nativeEdgePt : 0)) });
      if (!index) { expect(Math.abs(raster.edge)).toBeLessThan(1); browserEdgePt = raster.edge; }
      expect(result.approximate).toBe(reference.blurPt > 0);
      const errors = raster.values.flatMap((pixel, sampleIndex) => {
        const sample = reference.observed.samples[sampleIndex]!;
        // Sharp vector edges have a producer/pixel-grid quantization rim.
        if (!reference.blurPt && Math.abs(sample.offsetPt) < 1) return [];
        return pixel.map((value, channel) => Math.abs(value - sample.rgb[channel]!));
      });
      expect(Math.max(...errors)).toBeLessThanOrEqual(24);
    } finally { await page.evaluate(() => window.workbookShadowHost.disposeWorkbookShadowPixels()); }
  }
});

test('bounded auxiliary projections preserve native drawing and comment text without activating image paths', async ({ page }) => {
  const requests: string[] = []; page.on('request', request => requests.push(request.url()));
  for (const input of inputs.filter(input => !input.name.startsWith('line-'))) {
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureContentSource(snapshot), input.snapshot);
    expect(result.rootTexts, input.name).toEqual(input.text);
    expect(result, input.name).toMatchObject({ isolated: true, atomic: true, secondSurvives: true, released: true, limited: input.snapshot.content ? 'ResourceLimitExceeded' : '', unknown: 'InvalidOperation', duplicated: 'InvalidOperation', disposed: 'InvalidOperation' });
    expect(result.imageAttributes.every(image => image.src === null && input.snapshot.manifest.assets.some(asset => asset.path === image.path))).toBe(true);
  }
  await page.waitForTimeout(50); expect(requests).toEqual([]);
  await expect(page.locator('[data-office-view="workbook-content-measurement"]')).toHaveCount(0);
});

test('auxiliary parsing rejects active content, foreign nodes, repeated identities and budgets while inert', async ({ page }) => {
  const snapshot = inputs.find(input => input.name === '08-images-shapes')!.snapshot;
  const requests: string[] = []; page.on('request', request => requests.push(request.url()));
  for (const edit of [
    (html: string) => html.replace('<img ', '<img onload="window.injected=true" '),
    (html: string) => html.replace('<img ', '<svg><image '),
    (html: string) => html.replace(/src="[^"]+"/, 'src="https://example.test/private.png"'),
    (html: string) => html.replace('</section>', '<script data-node-id="injection">window.injected=true</script></section>'),
    (html: string) => html.replace('</section>', '<p data-node-id="injection">unowned</p></section>'),
    (html: string) => html.replace(/data-node-id="[^"]+"/g, 'data-node-id="duplicate"'),
  ]) {
    const broken = { ...snapshot, content: { ...snapshot.content!, html: edit(snapshot.content!.html) } };
    expect(await page.evaluate(snapshot => window.workbookContentHost.rejectContentSource(snapshot), broken)).toBe('InvalidDocument');
  }
  for (const limits of [[1, 8 * 1024 * 1024], [100_000, 1]]) expect(await page.evaluate(({ snapshot, limits }) => window.workbookContentHost.rejectContentSource(snapshot, limits[0], limits[1]), { snapshot, limits })).toBe('ResourceLimitExceeded');
  expect(await page.evaluate(snapshot => window.workbookContentHost.rejectContentSource(snapshot, 0), snapshot)).toBe('InvalidOperation');
  await page.waitForTimeout(50); expect(requests).toEqual([]);
});

test('workbook chart compaction preserves original point styles and distinguishes errors from skipped points', async ({ page }) => {
  const input = inputs.find(input => input.name === '07-three-charts')!;
  const result = await page.evaluate(snapshot => window.workbookChartHost.measureWorkbookChartPositions(snapshot), input.snapshot);
  expect(result.maskedColors).toEqual(['rgba(237,125,49,1)', 'rgba(237,125,49,1)']);
  expect(result.colors).toEqual(['rgb(0, 0, 255)', 'rgb(255, 0, 0)']);
  expect(result.plain.match(/M/g)).toHaveLength(1); expect(result.plain.match(/L/g)).toHaveLength(4);
  expect(result.missing.match(/M/g)).toHaveLength(2); expect(result.missing.match(/L/g)).toHaveLength(1);
  expect(result.hiddenPaths).toBe(0); expect(result.hiddenDiagnostics).toEqual([]);
  expect(result.errors).toEqual(['InvalidOperation', 'InvalidOperation', 'InvalidOperation']); expect(result.unchanged).toBe(true);
});


test('current reference chart titles share measured typography without cached text offsets', async ({ page }) => {
  const result = await page.evaluate(({ snapshot, bindings }) => window.workbookChartHost.measureWorkbookChartTitleContents(snapshot, bindings), referenceTitles);
  expect(result).toMatchObject({ unchanged: true, cleaned: true, empty: { projections: 0, visible: '' }, errors: ['ResourceLimitExceeded', 'ResourceLimitExceeded', 'InvalidOperation', 'InvalidOperation', 'InvalidOperation', 'ResourceLimitExceeded'] });
  expect(result.fonts).toBeGreaterThan(0);
  expect(result.results.slice(0, 3).map(item => item.text)).toEqual(['当前😀禰󠄀\n标题', '交付', '计划']);
  for (const item of result.results) {
    expect(item.nodes).toBe(item.declared); expect(item.sourceClasses).toBe(0); expect(item.sourceId).toBe(item.expectedId);
    if (item.reference) { expect(item.projections).toHaveLength(0); expect(item.scale).toBe(.75); expect(item.fontSizes.every(size => Math.abs(parseFloat(size) - 16) < .02)).toBe(true); }
    else expect(item.projections.map(projection => projection.text).join('')).toBe(item.text);
  }
});


test('composed workbook charts reserve separate title, axis and legend boxes and publish atomically', async ({ page }) => {
  const result = await page.evaluate(({ snapshot, bindings }) => window.workbookChartHost.measureWorkbookComposedChart(snapshot, bindings), referenceTitles);
  expect(result).toMatchObject({ unchanged: true, cleaned: true, errors: ['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'InvalidOperation', 'format rejected'] });
  const overlap = (a: { x: number; y: number; width: number; height: number }, b: typeof a) => Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x) > .1 && Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y) > .1;
  for (const item of result.results) {
    expect(item.nodes).toBe(item.declared); expect(item.styles).toBe(0); expect(item.unchanged).toBe(true);
    for (const key of ['x', 'y', 'width', 'height'] as const) expect(item.plot[key]).toBeCloseTo(item.layout.plot[key], 1);
    expect(item.titles.map(title => title.text)).toEqual(['当前😀禰󠄀\n标题', '交付', '计划']);
    for (const title of item.titles) {
      expect(title.projections).toBe(0); expect(overlap(title.rect, item.plot)).toBe(false);
      expect(title.rect.x).toBeGreaterThanOrEqual(-.1); expect(title.rect.y).toBeGreaterThanOrEqual(-.1);
      expect(title.rect.x + title.rect.width).toBeLessThanOrEqual(420.1); expect(title.rect.y + title.rect.height).toBeLessThanOrEqual(300.1);
      if (item.layout.legend) expect(overlap(title.rect, item.layout.legend)).toBe(false);
    }
    const category = item.titles.find(title => title.role === 'category-axis-title')!.rect, value = item.titles.find(title => title.role === 'value-axis-title')!.rect;
    expect(category.y < item.plot.y).toBe(item.opposite); expect(value.x > item.plot.x).toBe(item.opposite);
  }
});


test('chart axes and labels resolve source-linked formats from current cells and independent native formats', async ({ page }) => {
  const results = await page.evaluate(({ snapshot, bindings }) => window.workbookChartHost.measureWorkbookChartFormats(snapshot, bindings), formatCharts);
  expect(results).toHaveLength(20); expect(results.every(result => result.unchanged)).toBe(true);
  for (let mode = 0; mode < 4; mode++) {
    expect(results[mode]!.valueAxis).toEqual(mode === 0 ? ['0.00', '0.25', '0.50', '0.75', '1.00'] : ['0.0%', '25.0%', '50.0%', '75.0%', '100.0%']);
    expect(results[4 + mode]!.categoryAxis).toEqual(mode === 0 ? ['0.10', '0.20', '0.30'] : ['10.0%', '20.0%', '30.0%']);
    expect(results[8 + mode]!.labels).toEqual(mode === 0 ? ['0.25', '0.50', '0.75'] : ['25.0%', '50.0%', '75.0%']);
    expect(results[12 + mode]!.labels).toEqual(mode === 0 ? ['0.17', '0.33', '0.50'] : ['17%', '33%', '50%']);
  }
  for (const result of results.slice(16, 19)) expect(result.valueAxis).toEqual(['0', '0.25', '0.5', '0.75', '1']);
  expect(results[18]!.sourceIndexes).toEqual([1, 2]);
  expect(results[19]!.labels).toEqual(['0.25', '50.0%', '0.75']);
});


test('text overflow preserves cross-span graphemes, RTL and reflections even when no full line fits', async ({ page }) => {
  const results = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingOverflowEdges(snapshot), overflowText);
  expect(results).toHaveLength(24);
  for (const result of results) {
    expect(result, JSON.stringify(result)).toMatchObject({ splitGraphemes: 0, unchanged: true, marker: '…', inside: true });
    expect(result.text).toBe('A禰󠄀é😀WXYZ אבג ABCSecond lineHidden tail'); expect(result.nodes).toBe(result.declared);
  }
});


test('native automatic chart colors follow document themes, single-series variation, explicit overrides and source point indexes', async ({ page }) => {
  const report = JSON.parse(readFileSync('docs/evidence/xlsx-chart-colors.json', 'utf8')) as { workbooks: { cases: { kind: string; vary: boolean; series: number; fill?: string; producer: { pdfWords: Record<string, number>; pdfFillColors: string[] } }[] }[] };
  const rgb = (hex: string) => 'rgb(' + [0, 2, 4].map(offset => parseInt(hex.slice(offset, offset + 2), 16)).join(', ') + ')';
  for (const input of colorCharts) {
    const actual = await page.evaluate(({ snapshot, plots }) => window.workbookChartHost.measureWorkbookChartColors(snapshot, plots), input);
    const palette = (input.theme === 'alternate' ? ['aa5500', '008877', '5544aa'] : ['4472c4', 'ed7d31', 'a5a5a5']).map(rgb);
    expect(actual.unchanged).toBe(true);
    for (const [index, item] of actual.normal.entries()) {
      const source = report.workbooks[0]!.cases[index]!, varied = source.vary && source.series === 1;
      const colors = Array.from({ length: source.series }, (_, series) => Array.from({ length: 3 }, (_, point) => source.fill === 'series' ? rgb('112233') : source.fill === 'point' && point === 1 ? rgb('123456') : palette[varied ? point : series]!)).flat();
      expect(item.points, input.theme + ' case ' + index).toEqual(colors);
      expect(item.text).toEqual(source.kind === 'pie' || varied ? ['One', 'Two', 'Three'] : Array.from({ length: source.series }, (_, i) => 'Series' + (i + 1)));
      expect(item.keys).toEqual(source.kind === 'pie' || varied ? colors : Array.from({ length: source.series }, (_, i) => palette[i]!));
      expect(item.nodes).toBe(item.actualNodes);
      expect(item.words).toEqual(source.producer.pdfWords);
      if (source.kind !== 'line' && source.fill !== 'series') {
        const hex = (color: string) => color.match(/\d+/g)!.slice(0, 3).map(value => Number(value).toString(16).padStart(2, '0')).join('');
        const reference = report.workbooks[input.theme === 'alternate' ? 1 : 0]!.cases[index]!.producer.pdfFillColors;
        expect([...new Set([...item.points, ...item.keys].map(hex))].sort()).toEqual(reference);
      }
    }
    expect(actual.hidden).toMatchObject({ points: [palette[0], palette[0], palette[0]], keys: [palette[0]], text: ['Series1'] });
    expect(actual.compacted).toMatchObject({ points: [palette[0], palette[2]], keys: [palette[0], palette[2]], text: ['One', 'Three'] });
    expect(actual.withoutMarkers.lineOrder).toEqual([palette[0], palette[1]]);
  }
});


test('native point marker geometry, independent paint and inherited properties retain source indexes and exact budgets', async ({ page }) => {
  const cases = JSON.parse(readFileSync('docs/evidence/xlsx-chart-markers.json', 'utf8')).cases as { marker?: string; size?: number; appearance?: string }[];
  for (const input of markerCharts) {
    const result = await page.evaluate(({ snapshot, plots }) => window.workbookChartHost.measureWorkbookChartMarkers(snapshot, plots), input);
    expect(result.unchanged).toBe(true); expect(result.budget).toBe('ResourceLimitExceeded');
    for (const [index, item] of result.normal.entries()) {
      expect(item.nodes).toBe(item.declared);
      if (input.producer) continue;
      const expected = cases[index]!, symbol = expected.marker ?? 'circle', size = expected.size ?? 10;
      const points = item.markers.filter(marker => marker.role === 'point'), legend = item.markers.filter(marker => marker.role === 'legend');
      expect(points.map(marker => marker.sourceIndex)).toEqual(symbol === 'none' ? ['0'] : ['0', '1']);
      expect(legend).toHaveLength(1); expect(legend[0]).toMatchObject({ symbol: 'circle', width: 10, height: 10, fill: 'rgb(34, 136, 51)', stroke: 'rgb(17, 17, 17)' });
      expect(points[0]).toMatchObject({ symbol: 'circle', width: 10, height: 10, fill: 'rgb(34, 136, 51)', stroke: 'rgb(17, 17, 17)' });
      if (symbol === 'none') { expect(result.compacted[index]!.markers.filter(marker => marker.role === 'point')).toHaveLength(0); continue; }
      const point = points[1]!;
      expect(point).toMatchObject({ symbol, width: symbol === 'dot' ? size / 2 : size, height: symbol === 'dot' || symbol === 'dash' ? size / 5 : size });
      expect(point.gradient).toBe(expected.appearance === 'gradient');
      if (point.gradient) expect(point.stops).toEqual([{ offset: '0', color: 'rgb(255, 0, 0)' }, { offset: '1', color: 'rgb(0, 0, 255)' }]);
      if (expected.appearance !== 'gradient') expect(point.fill).toBe(expected.appearance === 'none' ? 'none' : expected.appearance ? 'rgb(255, 0, 0)' : 'rgb(34, 136, 51)');
      expect(point.stroke).toBe(expected.appearance ? 'rgb(0, 0, 255)' : 'rgb(17, 17, 17)');
      expect(parseFloat(point.strokeWidth)).toBe(expected.appearance === 'zero-line' ? 0 : 1.5);
      if (['star', 'plus', 'x'].includes(symbol)) expect(point.paths.join('').match(/M/g)).toHaveLength(symbol === 'star' ? 4 : 2);
      expect(result.compacted[index]!.nodes).toBe(result.compacted[index]!.declared);
      const { paths, fill, ...appearance } = point;
      const compacted = result.compacted[index]!.markers.filter(marker => marker.role === 'point');
      expect(compacted).toHaveLength(1); expect(compacted[0]).toMatchObject(appearance);
      if (!point.gradient) expect(compacted[0]!.fill).toBe(fill);
      expect(compacted[0]!.paths.map(path => path?.match(/M/g)?.length)).toEqual(paths.map(path => path?.match(/M/g)?.length));
    }
  }
});

test('explicit automatic markers retain native source and choose stable series shapes across hidden points and series', async ({ page }) => {
  const cases = JSON.parse(readFileSync('docs/evidence/xlsx-chart-auto-markers.json', 'utf8')).cases as { series: number; symbol: string; pointSymbol?: string; pointSize?: number; pointPaint?: string }[];
  const produced = JSON.parse(readFileSync('docs/evidence/xlsx-chart-auto-markers.json', 'utf8')).producerMarkers as { series: { symbol: string } | null; points: { index: string; marker: { symbol: string | null } | null }[] }[][];
  for (const input of automaticMarkerCharts) {
    const result = await page.evaluate(({ snapshot, plots }) => window.workbookChartHost.measureWorkbookChartMarkers(snapshot, plots), input);
    expect(result.unchanged).toBe(true); expect(result.budget).toBe('ResourceLimitExceeded'); expect(result.flagsIgnored).toBe(true);
    if (input.producer) {
      for (const [index, item] of result.normal.entries()) {
        expect(item.nodes).toBe(item.declared); expect(input.snapshot.metadata.workbook.charts![index]!.nativeShowMarkers).toBe(true);
        for (const [series, native] of produced[index]!.entries()) {
          const parent = native.series?.symbol ?? 'none', markers = item.markers.filter(marker => marker.seriesId === 'series-' + series);
          const expected = [0, 1, 2].map(point => ({ symbol: native.points.find(entry => Number(entry.index) === point)?.marker?.symbol ?? parent, sourceIndex: String(point) })).filter(point => point.symbol !== 'none');
          expect(markers.filter(marker => marker.role === 'point').map(({ symbol, sourceIndex }) => ({ symbol, sourceIndex }))).toEqual(expected);
          expect(markers.filter(marker => marker.role === 'legend').map(marker => marker.symbol)).toEqual(parent === 'none' ? [] : [parent]);
        }
      }
      continue;
    }
    for (const [index, item] of result.normal.entries()) {
      const expected = cases[index]!; expect(item.nodes).toBe(item.declared);
      for (let series = 0; series < expected.series; series++) {
        const automatic = series ? 'square' : 'circle', base = expected.symbol === 'auto' ? automatic : expected.symbol;
        const markers = item.markers.filter(marker => marker.seriesId === 'series-' + series);
        expect(markers.filter(marker => marker.role === 'legend')).toMatchObject([{ symbol: base, automatic: expected.symbol === 'auto' }]);
        const points = markers.filter(marker => marker.role === 'point');
        expect(points.map(point => point.sourceIndex)).toEqual(expected.pointSymbol === 'none' ? ['0', '2'] : ['0', '1', '2']);
        for (const point of points) {
          const override = point.sourceIndex === '1', requested = override ? expected.pointSymbol ?? expected.symbol : expected.symbol;
          expect(point.symbol).toBe(requested === 'auto' ? automatic : requested); expect(point.automatic).toBe(requested === 'auto');
          expect(point.width).toBeCloseTo(override ? expected.pointSize ?? 10 : 10, 4);
          expect(point.height).toBeCloseTo(point.width, 4);
          if (override && expected.pointPaint === 'none') expect(point.fill).toBe('none');
        }
        const compacted = result.compacted[index]!; expect(compacted.nodes).toBe(compacted.declared);
        expect(compacted.markers.filter(marker => marker.role === 'point' && marker.seriesId === 'series-' + series).map(({ symbol, automatic, sourceIndex }) => ({ symbol, automatic, sourceIndex })))
          .toEqual(points.filter(point => point.sourceIndex !== '0').map(({ symbol, automatic, sourceIndex }) => ({ symbol, automatic, sourceIndex })));
      }
      if (expected.series === 2) {
        const hidden = result.hidden[index]!; expect(hidden.nodes).toBe(hidden.declared);
        expect(hidden.markers.every(marker => marker.seriesId === 'series-1' && marker.symbol === 'square' && marker.automatic)).toBe(true);
        expect(hidden.markers.filter(marker => marker.role === 'point')).toHaveLength(3);
      }
    }
  }
});

test('derived chart labels keep mixed font roles and variation sequences with local diagnostics and bounded probes', async ({ page }) => {
  const input = automaticMarkerCharts[0]!;
  for (const missing of [false, true]) {
    const result = await page.evaluate(({ snapshot, plots, missing }) => window.workbookChartHost.measureWorkbookChartTextFonts(snapshot, plots[0]!, missing), { ...input, missing });
    expect(result).toMatchObject({ unchanged: true, cleaned: true, budget: 'ResourceLimitExceeded', scalar: { category: result.mixed } });
    expect(result.nodes).toBe(result.declared); expect(result.peak).toBeGreaterThan(result.nodes);
    expect(result.runs.some(run => run.role === 'east-asia' && run.text?.includes('禰\u{e0100}'))).toBe(true);
    expect(result.runs.every(run => !run.canonical)).toBe(true);
    expect([...new Set(result.runs.map(run => run.role))].sort()).toEqual(['ascii', 'complex-script', 'east-asia', 'h-ansi']);
    for (const [index, role] of ['ascii', 'h-ansi', 'east-asia', 'complex-script'].entries()) {
      const resolution = result.fonts.find(font => font.requestedFamily === result.names[index])!;
      expect(resolution).toBeTruthy(); expect(resolution.substituted).toBe(missing);
      expect(result.composedFonts).toContainEqual(resolution);
      expect(result.inheritedFonts).toContainEqual(resolution);
      expect(result.runs.filter(run => run.role === role).every(run => run.family?.includes(resolution.selectedFamily))).toBe(true);
    }
    expect(result.scalar.size * result.scalar.scale).toBeCloseTo(24, 3);
    expect(result.inheritedSize).toBeCloseTo(24, 3);
    expect(result.localSizes.length).toBeGreaterThan(3); for (const size of result.localSizes) expect(size).toBeCloseTo(12, 3);
  }
});


test('category chart grids share current data intervals and remain visible with hidden axes while plot paint stays clipped', async ({ page }) => {
  const cases = JSON.parse(readFileSync('docs/evidence/xlsx-chart-grid.json', 'utf8')).cases as { kind: string; between: string; reverse: boolean; single?: boolean; pdfCategoryGridPositions: number[] }[];
  for (const input of gridCharts) {
    const result = await page.evaluate(({ snapshot, plots }) => window.workbookChartHost.measureWorkbookChartGrid(snapshot, plots), input);
    expect(result.unchanged).toBe(true);
    for (const [index, item] of result.normal.entries()) {
      const source = cases[index]!, count = source.single ? 1 : 3, between = source.between === 'between';
      const ticks = between ? count + 1 : Math.max(2, count);
      expect(item.nodes).toBe(item.declared); expect(item.dataClipped).toBe(true); expect(item.clip).toEqual([item.left, item.top, item.width, item.height]);
      expect(item.grid).toHaveLength(ticks);
      item.grid.forEach((line, index) => { expect(line.position).toBeCloseTo(index / (ticks - 1), 7); expect(line).toMatchObject({ from: 0, to: 1, color: 'rgb(0, 255, 0)' }); });
      expect(item.points).toHaveLength(count);
      item.points.forEach((position, index) => { const value = (index + (between ? .5 : 0)) / (between ? count : Math.max(1, count - 1)); expect(position).toBeCloseTo(source.reverse ? 1 - value : value, 7); });
      if (source.kind !== 'column' || between || source.single) {
        expect(item.grid).toHaveLength(source.pdfCategoryGridPositions.length);
        item.grid.forEach((line, index) => expect(Math.abs(line.position - source.pdfCategoryGridPositions[index]!)).toBeLessThan(.0001));
      }
    }
    expect(result.paints).toHaveLength(2);
    for (const paint of result.paints) { expect(paint.clipped).toBe(true); expect(paint.painted).toBeGreaterThan(20); expect(paint.escaped).toBe(0); }
  }
});


test('Office chart axes ignore axPos and preserve perpendicular crossings with current hidden categories', async ({ page }) => {
  const cases = JSON.parse(readFileSync('docs/evidence/xlsx-chart-axes.json', 'utf8')).cases as { kind: string; between: string; valueAt?: number; pdfAxes: { category: number[][]; value: number[][] } }[];
  for (const input of axisCharts) {
    const result = await page.evaluate(({ snapshot, plots }) => window.workbookChartHost.measureWorkbookChartAxes(snapshot, plots), input);
    expect(result.unchanged).toBe(true); expect(result.ignored).toEqual(result.normal);
    for (const [index, actual] of result.normal.entries()) {
      const reference = cases[index]!;
      expect(actual.nodes).toBe(actual.declared); expect(result.composed[index]!.nodes).toBe(result.composed[index]!.declared);
      expect(actual.y).toBeCloseTo(reference.pdfAxes.category[0]![1]!, 3);
      if (reference.kind === 'column' && reference.between === 'midCat') expect(actual.x).toBe(.5);
      else expect(actual.x).toBeCloseTo(reference.pdfAxes.value[0]![0]!, 3);
      expect(result.composed[index]!.x).toBeCloseTo(actual.x, 6); expect(result.composed[index]!.y).toBeCloseTo(actual.y, 6);
    }
    for (const group of result.labels) for (const { position, value } of group.results) {
      expect(value.categoryLabels.length).toBeGreaterThan(0); expect(value.valueLabels.length).toBeGreaterThan(0);
      const categoryAt = position === 'nextTo' ? value.y : (position === 'high') !== !!group.reverse ? 0 : 1;
      const valueAt = position === 'nextTo' ? value.x : (position === 'high') !== !!group.reverse ? 1 : 0;
      const top = categoryAt === (group.reverse ? 1 : 0) ? !group.reverse : !!group.reverse;
      const right = valueAt === (group.reverse ? 0 : 1) ? !group.reverse : !!group.reverse;
      for (const label of value.categoryLabels) expect(label.y).toBeCloseTo(categoryAt + (top ? -16 : 16) / value.plot[3]!, 5);
      for (const label of value.valueLabels) { expect(label.x).toBeCloseTo(valueAt + (right ? 7 : -7) / value.plot[2]!, 5); expect(label.anchor).toBe(right ? 'start' : 'end'); }
    }
  }
});

test('chart numeric spaces preserve actual glyph advances and bounded fills after final font resolution', async ({page})=>{
  const result=await page.evaluate(input=>window.workbookChartHost.measureChartNumberFields(input.snapshot,input.bindings),formatCharts);
  expect(result).toMatchObject({unchanged:true,cleaned:true});expect(result.results).toHaveLength(7);
  for(const [index,item] of result.results.entries()){
    expect(item.nodes).toBe(item.actual);
    if([0,1,2,5,6].includes(index)){
      expect(item.spaces.length).toBeGreaterThan(0);
      for(const space of item.spaces){expect(space.text).toBe(' ');expect(space.hidden).toBeNull();expect(space.declared).toBeGreaterThan(0);expect(space.actual).toBeCloseTo(space.declared,1);}
    }
    if(index===3)expect(item.fills.every(fill=>fill.text!.length>3&&/^x+$/.test(fill.text!))).toBe(true);
    if(index===4){expect(item.fills.filter(fill=>fill.text==='')).toHaveLength(8);expect(item.fills.filter(fill=>fill.text!==''&&/^y+$/.test(fill.text!))).toHaveLength(8);}
    for(const label of item.labels)expect(label).not.toMatch(/[)中😀]/u);
  }
});


test('drawing font probes reuse one owner and invalidate after actual font-set changes without changing source or projections', async ({ page }) => {
  const snapshot = inputs.find(input => input.name === '08-images-shapes')!.snapshot;
  const result = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingFontProbeReuse(snapshot), snapshot);
  expect(result).toMatchObject({ unchanged: true, cleaned: true });
  expect(result.firstCreated).toBeGreaterThan(0); expect(result.secondCreated).toBe(result.firstCreated);
  expect(result.changedCreated).toBeGreaterThan(result.secondCreated); expect(result.clearedCreated).toBeGreaterThan(result.changedCreated);
  expect(result.second).toEqual(result.first); expect(result.changed).toEqual(result.first); expect(result.cleared).toEqual(result.first);
});
