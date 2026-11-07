import { encodeTestPackage } from './package-transport.js';
import { openSdkHost } from './sdk-host.js';
import { expect, test } from '@playwright/test';
import { openDocument } from '@doffi/core';
import { readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { scalarWorkbook } from '../../packages/core/tests/workbook-scalar-fixture.js';
import { importXlsx } from '../../packages/office/src/xlsx-import.js';
import type * as fixture from './fixtures/workbook-preview.js';
declare global { interface Window { workbookPreview: typeof fixture } }
let code = '';
const inputs: { name: string; files: [string, string][] }[] = [];
let readAheadFiles: [string, string][] = [];
let sparseRowFiles: [string, string][] = [];
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-preview.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookPreview' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected preview bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  const dense = scalarWorkbook(Array.from({ length: 80 * 8 }, (_, index) => ({ row: Math.floor(index / 8), column: index % 8, value: { type: 'number' as const, value: index + 1 } })), metadata => {
    const sheet = metadata.workbook.sheets[0]!; sheet.frozen = { rows: 1, columns: 1 }; sheet.defaultRowHeightCustom = true; sheet.rightToLeft = true;
    sheet.rowDimensions = [{ from: 4, to: 6, hidden: true }];
  });
  const sparseRows = scalarWorkbook([0, 128].map(row => ({ row, column: 0, value: { type: 'number' as const, value: row + 1 } })));
  sparseRowFiles = encodeTestPackage(sparseRows.serialize());
  readAheadFiles = encodeTestPackage(dense.serialize());
  for (const name of ['09-classic-notes', '10-modern-threads', '08-images-shapes']) {
    const files = await importXlsx(readFileSync(`examples/xlsx/${name}.xlsx`), { documentId: 'public-preview' });
    inputs.push({ name, files: encodeTestPackage(files) });
  }
});
test.beforeEach(async ({ page }) => { await openSdkHost(page); await page.addScriptTag({ content: code }); });
test('incremental preview remeasures distant hidden-column results and retains its accepted surface on failure and cancellation', async ({ page }) => {
  const model = scalarWorkbook(Array.from({ length: 2048 * 3 }, (_, index) => {
    const row = Math.floor(index / 3), column = index % 3;
    return { row, column, value: { type: 'number' as const, value: row === 0 && column === 0 ? 1 : index + 1 },
      ...(row === 1000 && column === 1 ? { formula: '=IF($A$1=1,"short","long wrapped content long wrapped content long wrapped content long wrapped content")' } : {}) };
  }), metadata => { metadata.workbook.sheets[0]!.columnDimensions = [{ from: 1, to: 1, hidden: true, width: 8.43, unit: 'characters' }]; });
  const files = new Map(model.serialize()), encoder = new TextEncoder();
  files.set('styles.css', encoder.encode(new TextDecoder().decode(files.get('styles.css')!).replace('font-size:11pt;', 'font-size:11pt;--office-cell-wrap:true;')));
  const result = await page.evaluate(entries => window.workbookPreview.incrementalPreview(entries), encodeTestPackage(files));
  expect(result).toMatchObject({ failure: 'RevisionConflict', retained: true, superseded: true, sameFrame: true, frames: 1,
    measuredDistant: true, skippedUntouched: true, errors: ['RevisionConflict'], unhandled: [] });
  expect(result.updated).toEqual(result.fresh); expect(result.updated.scrollTopPt).toBeGreaterThan(result.oldTop);
  expect(result.updatedReadRows).toBeLessThan(result.initialReadRows / 2);
});
test('completed calculation renders an unruled long chain and preserves view selection and active sheet', async ({ page }) => {
  const model = scalarWorkbook(Array.from({ length: 150 }, (_, row) => ({ row, column: 0, value: { type: 'number' as const, value: 9999 }, formula: row ? `=A${row}+1` : '=1' })), metadata => { metadata.workbook.sheets[0]!.defaultRowHeightCustom = true; });
  const result = await page.evaluate(entries => window.workbookPreview.calculatedPreview(entries), encodeTestPackage(model.serialize()));
  expect(result).toEqual({ cell: '150', retained: true, sheet: 'preview-empty', errors: [], unchanged: true });
});
for (const rtl of [false, true]) test(`imperative source-range selection scrolls and guards worksheet races in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const files = await importXlsx(readFileSync('examples/xlsx/03-structure.xlsx'), { documentId: 'range-selection' });
  const result = await page.evaluate(({ entries, rtl }) => window.workbookPreview.imperativeRangeSelection(entries, rtl), { entries: encodeTestPackage(files), rtl });
  expect(result).toMatchObject({ invalid: 'InvalidOperation', retained: true, sameViewport: true, race: 'RevisionConflict', unchanged: true, frozen: true,
    first: { documentId: 'range-selection', baseRevision: 0, anchor: { kind: 'cells', ranges: [{ from: { row: 0, column: 0 }, to: { row: 3, column: 3 } }] } },
    far: { anchor: { ranges: [{ from: { row: 500, column: 40 }, to: { row: 502, column: 42 } }] } } });
  expect(result.moved.scrollLeftPt).toBeGreaterThan(0); expect(result.moved.scrollTopPt).toBeGreaterThan(0);
});
for (const rtl of [false, true]) test(`public workbook preview selects cells and switches sheets without serializing source in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  for (const input of inputs.filter(input => input.name !== '08-images-shapes')) {
    const start = await page.evaluate(({ files, rtl }) => window.workbookPreview.mountPublicPreview(files, rtl), { files: input.files, rtl });
    expect(start).toMatchObject({ unchanged: true, serializations: 0, frames: 1, errors: [], report: { kind: 'workbook', layout: 'grid' } });
    await page.mouse.click(start.x, start.y);
    expect(await page.evaluate(() => window.workbookPreview.previewInspect())).toMatchObject({ address: 'A1', selection: { documentId: 'public-preview', baseRevision: 0, anchor: { kind: 'cells', ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] } } });
    await page.evaluate(() => window.workbookPreview.previewAction('comment'));
    const open = await page.evaluate(() => window.workbookPreview.previewInspect()) as { report: { fonts: unknown[] } };
    expect(open.report.fonts.length).toBeGreaterThan(0);
    await page.evaluate(() => window.workbookPreview.previewAction('close'));
    const empty = await page.evaluate(() => window.workbookPreview.previewAction('empty'));
    expect(empty).toMatchObject({ report: { sheetId: 'preview-empty', storedCellCount: 0 }, address: '', frames: 1 });
    expect(await page.evaluate(() => window.workbookPreview.previewAction('bad-sheet'))).toMatchObject({ report: { sheetId: 'preview-empty' }, errors: ['InvalidOperation'], frames: 1 });
    await page.evaluate(() => window.workbookPreview.previewAction('first'));
    await page.mouse.move(start.x + (rtl ? -50 : 50), start.y + 80); await page.mouse.wheel(0, 200);
    await expect.poll(async () => (await page.evaluate(() => window.workbookPreview.previewInspect()) as { report: { viewport: { scrollTopPt: number } } }).report.viewport.scrollTopPt).toBeGreaterThan(0);
    expect(await page.evaluate(() => window.workbookPreview.previewAction('dispose'))).toEqual({ removed: true, frames: 0, unchanged: true, serializations: 0, fieldsCleared: true });
  }
});
test('public workbook preview retains accepted data on failure and settles superseded and disposed providers', async ({ page }) => {
  const result = await page.evaluate(files => window.workbookPreview.publicPreviewLifecycle(files), inputs.find(input => input.name === '08-images-shapes')!.files);
  expect(result).toEqual({ scrollRace: true, failed: 'RevisionConflict', retained: true, updateRace: true, disposed: 'InvalidOperation', errors: ['RevisionConflict'], unhandled: [], providerAlive: true, released: true, removed: true });
});
test('control construction is atomic and open comments and selections share the complete window budget', async ({ page }) => {
  const result = await page.evaluate(files => window.workbookPreview.previewDecorationMeasurement(files), inputs.find(input => input.name === '09-classic-notes')!.files);
  expect(result).toMatchObject({ errors: ['UnsupportedFeature', 'ResourceLimitExceeded', 'cancelled'], highlights: 1, releasedCandidates: 2, releasedAll: 3, empty: true });
  expect(result.nodes).toBe(result.actualNodes); expect(result.fonts).toBeGreaterThan(0); expect(result.panel?.length).toBeGreaterThan(0);
});
test('initial failures release their candidate and a host retry can satisfy the same readiness promise', async ({ page }) => {
  const result = await page.evaluate(files => window.workbookPreview.initialPreviewFailure(files), inputs.find(input => input.name === '09-classic-notes')!.files);
  expect(result).toEqual({ rejected: 'InvalidOperation', frames: 0, errors: ['InvalidOperation', 'InvalidOperation'], retried: true, finalFrames: 1 });
});

test('public read-only validation shows native prompts and the actual violating value without rewriting source', async ({ page }) => {
  const files = await importXlsx(readFileSync('examples/xlsx/06-rules-validation.xlsx'), { documentId: 'validation-preview' });
  const entries = encodeTestPackage(files);
  await page.evaluate(entries => window.workbookPreview.mountPublicPreview(entries, false), entries);
  for (let row = 2; row <= 8; row++) {
    const point = await page.evaluate(row => window.workbookPreview.previewAction('point-E-' + row), row) as { x: number; y: number };
    await page.mouse.click(point.x, point.y);
    const result = await page.evaluate(() => window.workbookPreview.previewInspect());
    expect(result).toMatchObject({ address: 'E' + row, unchanged: true, errors: [] });
    expect(result).toMatchObject({ validation: expect.stringContaining('填写提示\n保留来源违规值') });
    expect(result).toMatchObject({ validation: expect.stringContaining(row === 3 ? '当前值不符合验证规则。\n输入不符\n请检查范围' : '当前值符合验证规则。') });
    if (row === 3) expect(result).toMatchObject({ value: '99' });
  }
  expect(await page.evaluate(() => window.workbookPreview.previewAction('empty'))).toMatchObject({ validation: '' });
  await page.evaluate(() => window.workbookPreview.previewAction('dispose'));
});
test('failed or late rule responses preserve the last complete public window', async ({ page }) => {
  const files = await importXlsx(readFileSync('examples/xlsx/06-rules-validation.xlsx'), { documentId: 'rules-lifecycle' });
  const result = await page.evaluate(entries => window.workbookPreview.ruleWindowLifecycle(entries), encodeTestPackage(files));
  expect(result).toEqual({ failures: ['InvalidDocument', 'RevisionConflict'], cancelled: true, aborted: true, retainedLatest: true });
});

test('worksheet-wide prompt suppression retains current validation status and source violations in the report', async ({ page }) => {
  const files = new Map(await importXlsx(readFileSync('examples/xlsx/06-rules-validation.xlsx'), { documentId: 'quiet-validation' }));
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)); metadata.workbook.sheets[0].disableValidationPrompts = true;
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  await page.evaluate(entries => window.workbookPreview.mountPublicPreview(entries, false), encodeTestPackage(files));
  const point = await page.evaluate(() => window.workbookPreview.previewAction('point-E-3')) as { x: number; y: number }; await page.mouse.click(point.x, point.y);
  expect(await page.evaluate(() => window.workbookPreview.previewInspect())).toMatchObject({ validation: '当前值不符合验证规则。', value: '99', unchanged: true,
    report: { diagnostics: [expect.objectContaining({ code: 'validation-invalid', row: 2, column: 4 })] } });
  await page.evaluate(() => window.workbookPreview.previewAction('dispose'));
});


test('idle neighbor reading is opt-in, bounded, cancellable and never publishes a background candidate', async ({ page }) => {
  for (const [allow, rows, mode] of [[false, 8, 'normal'], [true, 0, 'normal'], [true, 8, 'normal'], [true, 8, 'hold'], [true, 8, 'failure'], [true, 8, 'revision'], [true, 8, 'active'], [true, 8, 'preference']] as const) {
    const result = await page.evaluate(({ entries, allow, rows, mode }) => window.workbookPreview.publicReadAhead(entries, allow, rows, mode), { entries: readAheadFiles, allow, rows, mode });
    expect(result, JSON.stringify({ allow, rows, mode, result })).toMatchObject({ invalid: ['InvalidOperation', 'InvalidOperation', 'InvalidOperation'], unhandled: [], firstRetained: true, foreground: true, stopped: true, clean: true, activeContent: false, unchanged: true });
    if (!allow || !rows) expect(result.reads).toBe(0); else { expect(result.reads).toBeGreaterThan(0); expect(result.reads).toBeLessThanOrEqual(mode === 'hold' ? 9 : 8); expect(result.largestSlots).toBeLessThanOrEqual(256); }
    if (mode === 'hold') { expect(result.aborted).toBe(true); expect(result.whileHeld).toBe(1); }
  }
});

test('automatic row reads overlap one authorized batch and abort it on cancellation or invalid source', async ({ page }) => {
  for (const mode of ['sequential', 'ahead', 'cancel', 'invalid', 'preference'] as const) {
    const result = await page.evaluate(({ entries, mode }) => window.workbookPreview.autoRowReadPipeline(entries, mode), { entries: readAheadFiles, mode });
    expect(result, mode).toMatchObject({ reads: 2, unhandled: [], clean: true });
    if (mode === 'cancel' || mode === 'invalid') expect(result).toMatchObject({ aborted: true, error: mode === 'cancel' ? 'OperationCancelled' : 'InvalidDocument' });
    else expect(result).toMatchObject({ error: undefined, cells: 640, overlap: mode === 'ahead' });
  }
});

test('prepared rolling candidates require fresh valid values and share the complete live DOM budget', async ({ page }) => {
  for (const mode of ['same', 'value', 'active', 'revision', 'font', 'budget'] as const) {
    const result = await page.evaluate(({ entries, mode }) => window.workbookPreview.rollingCandidateLifecycle(entries, mode), { entries: readAheadFiles, mode });
    expect(result, mode).toMatchObject({ retained: true, active: false, stopped: true, clean: true });
    expect(result.background).toBeGreaterThan(0); expect(result.fresh).toBeGreaterThan(0);
    expect(result.fragments).toBeLessThanOrEqual(5000); expect(result.mountedPeak).toBeLessThanOrEqual(mode === 'budget' ? 40 : 5000);
    if (mode === 'active' || mode === 'revision') expect(result).toMatchObject({ current: 30, failedRetained: true, error: mode === 'revision' ? 'RevisionConflict' : 'InvalidDocument' });
    else expect(result).toMatchObject({ current: 60, error: undefined, updated: mode === 'value', failedRetained: false });
  }
});

test('grid wheels retain host scrolling boundaries and shift-wheel navigation', async ({ page }) => {
  await page.evaluate(() => {
    document.documentElement.style.overflow = 'auto'; document.body.style.overflow = 'visible'; document.body.style.height = '4000px'; window.scrollTo(0, 200);
  });
  expect(await page.evaluate(() => window.scrollY)).toBe(200);
  const start = await page.evaluate(entries => window.workbookPreview.mountPublicPreview(entries, false), inputs.find(input => input.name === '09-classic-notes')!.files);
  const panel = await page.evaluate(() => window.workbookPreview.previewAction('comment-scroll')) as { x: number; y: number; max: number };
  expect(panel.max).toBeGreaterThan(0); await page.mouse.move(panel.x, panel.y); await page.mouse.wheel(0, 100);
  await expect.poll(async () => (await page.evaluate(() => window.workbookPreview.previewAction('comment-scroll-state')) as { top: number }).top).toBeGreaterThan(0);
  expect((await page.evaluate(() => window.workbookPreview.previewAction('comment-scroll-state')) as { viewport: { scrollTopPt: number } }).viewport.scrollTopPt).toBe(0);
  expect(await page.evaluate(() => window.scrollY)).toBe(200); await page.evaluate(() => window.workbookPreview.previewAction('close'));
  await page.mouse.move(start.x + 50, start.y + 80); await page.mouse.wheel(0, -200);
  await page.evaluate(() => window.workbookPreview.previewAction('close'));
  expect(await page.evaluate(() => window.scrollY)).toBe(200);
  expect((await page.evaluate(() => window.workbookPreview.previewInspect()) as { report: { viewport: { scrollTopPt: number } } }).report.viewport.scrollTopPt).toBe(0);
  await page.mouse.wheel(0, 200);
  await expect.poll(async () => (await page.evaluate(() => window.workbookPreview.previewInspect()) as { report: { viewport: { scrollTopPt: number } } }).report.viewport.scrollTopPt).toBeGreaterThan(0);
  await page.keyboard.down('Shift'); await page.mouse.wheel(0, 80); await page.keyboard.up('Shift');
  await expect.poll(async () => (await page.evaluate(() => window.workbookPreview.previewInspect()) as { report: { viewport: { scrollLeftPt: number } } }).report.viewport.scrollLeftPt).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.scrollY)).toBe(200);
  await page.evaluate(() => window.workbookPreview.previewAction('dispose'));
});

test('cancelled automatic rows release prototypes before an uncooperative reply arrives', async ({ page }) => {
  for (const mode of ['late-cancel', 'sequential-late-cancel'] as const) {
    const result = await page.evaluate(({ entries, mode }) => window.workbookPreview.autoRowReadPipeline(entries, mode), { entries: sparseRowFiles, mode });
    expect(result, mode).toMatchObject({ reads: 2, waitingObserved: true, aborted: true, error: 'OperationCancelled', cleanBeforeRelease: true, clean: true, unhandled: [] });
  }
});

for (const escaped of [false, true]) test(`viewport-relative workbook typography resizes through an atomic independent candidate (escaped ${escaped})`, async ({ page }) => {
  const result = await page.evaluate(({ entries, escaped }) => window.workbookPreview.responsiveViewportLifecycle(entries, escaped), { entries: readAheadFiles, escaped });
  expect(result.resized).toEqual(result.fresh); expect(result.scrolled).toEqual(result.fresh);
  expect(result.initial.left).not.toBe(result.fresh.left);
  expect(result).toMatchObject({ failure: 'RevisionConflict', retained: true, frames: 1, serializations: 0, unchanged: true, clean: true, superseded: true, errors: ['RevisionConflict'] });
});

test('inline HTML viewport fonts measure at the final resized viewport', async ({ page }) => {
  const result = await page.evaluate(entries => window.workbookPreview.responsiveInlineViewport(entries), inputs.find(input => input.name === '08-images-shapes')!.files);
  expect(result.initial).toBeGreaterThan(0); expect(result.resized).toBeCloseTo(result.fresh, 6); expect(result.resized).toBeGreaterThan(result.initial * 1.5);
});

test('cancelled automatic rule reads release typography before a late custom response', async ({ page }) => {
  const result = await page.evaluate(entries => window.workbookPreview.autoRowRulesCancellation(entries), sparseRowFiles);
  expect(result).toEqual({ calls: 2, retainedBeforeAbort: true, error: 'OperationCancelled', cleanBeforeRelease: true, aborted: true });
});

const matrixCases=JSON.parse(readFileSync('examples/xlsx/expected.json','utf8')) as {cases:{file:string;sheets:string[];values?:Record<string,number|boolean|string>}[]};
for(const producer of [false,true])for(const sample of matrixCases.cases)test(`${producer?'producer':'source'} ${sample.file} completes the public reading matrix without changing its package`,{tag:sample.file.startsWith('15-')?'@scale':[]},async({page},testInfo)=>{
  if(sample.file.startsWith('15-'))test.slow();
  const files=await importXlsx(readFileSync('examples/xlsx/'+(producer?'producer/':'')+sample.file),{documentId:'matrix-'+sample.file.slice(0,2)});
  const model=openDocument(files,{kind:'workbook'}),snapshot=model.snapshot(),saved=model.serialize();
  expect(snapshot.metadata.workbook.sheets.map(sheet=>sheet.name)).toEqual(sample.sheets);
  expect(openDocument(saved,{kind:'workbook'}).snapshot()).toEqual(snapshot);
  // Independent source values are checked before presentation; producer
  // cached results may be recomputed by Calc and are checked in its own audit.
  if(!producer)for(const [address,value] of Object.entries(sample.values??{})){
    const [,name,column,row]=/^(.*)!([A-Z]+)([0-9]+)$/.exec(address)!;
    const index=[...column!].reduce((sum,char)=>sum*26+char.charCodeAt(0)-64,0)-1;
    const sheet=snapshot.metadata.workbook.sheets.find(sheet=>sheet.name===name)!;
    expect(model.getCell(sheet.sheetId,Number(row)-1,index)?.value).toMatchObject({value});
  }
  const start=await page.evaluate(entries=>window.workbookPreview.matrixStart(entries),[...saved].map(([path,bytes])=>[path,Buffer.from(bytes).toString('base64')] as [string,string]));
  expect(start.snapshot).toEqual(snapshot);
  const observations=[];
  try{
    for(const sheet of snapshot.metadata.workbook.sheets.filter(sheet=>sheet.visibility==='visible'))for(const last of [false,true]){
      const result=await page.evaluate(({id,last})=>window.workbookPreview.matrixWindow(id,last),{id:sheet.sheetId,last});
      expect(result).toMatchObject({errors:[],serializations:0,report:{sheetId:sheet.sheetId,documentId:snapshot.manifest.documentId,revision:0,kind:'workbook',layout:'grid'}});
      expect(result.report.cellFragmentCount).toBeLessThanOrEqual(5000);expect(result.report.visibleCellCount).toBeGreaterThan(0);
      expect(result.tabs).toEqual(snapshot.metadata.workbook.sheets.filter(sheet=>sheet.visibility==='visible').map(sheet=>sheet.name));
      expect(result.report.fonts.length).toBeGreaterThan(0);
      observations.push({sheet:sheet.name,last,report:result.report});
      if(!last&&testInfo.project.name==='chromium'&&!process.env.DOFFI_BROWSER_EXECUTABLE){const path=testInfo.outputPath(sheet.name+'.png');await page.screenshot({path});await testInfo.attach(sheet.name+'.png',{path,contentType:'image/png'});}
    }
    const path=testInfo.outputPath('public-reading.json');writeFileSync(path,JSON.stringify({file:sample.file,producer,sourceSha256:createHash('sha256').update(readFileSync('examples/xlsx/'+(producer?'producer/':'')+sample.file)).digest('hex'),hostBundleSha256:createHash('sha256').update(code).digest('hex'),observations},null,2));await testInfo.attach('public-reading.json',{path,contentType:'application/json'});
  }finally{expect(await page.evaluate(()=>window.workbookPreview.matrixFinish())).toEqual({unchanged:true,serializations:0,errors:[],frames:0,removed:true});}
});
