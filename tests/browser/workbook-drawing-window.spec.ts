import { encodeTestPackage } from './package-transport.js';
import { openSdkHost } from './sdk-host.js';
import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { importXlsx } from '../../packages/office/src/xlsx-import.js';
import type { currentChartWindowMeasurement, drawingWindowMeasurement, gradientWindowMeasurement } from './fixtures/workbook-drawing-window.js';
import type { annotationWindowMeasurement, commentTextMeasurement, windowControllerMeasurement } from './fixtures/workbook-window-controller.js';
import type { commentPanelMeasurement, mentionPanelMeasurement } from './fixtures/workbook-comment-panel.js';
import type { annotationControlsAction, mountAnnotationControls } from './fixtures/workbook-annotation-controls.js';
import type { nativeNoteMeasurement } from './fixtures/workbook-native-notes.js';

declare global { interface Window { workbookDrawingWindow: { nativeNoteMeasurement: typeof nativeNoteMeasurement; drawingWindowMeasurement: typeof drawingWindowMeasurement; gradientWindowMeasurement: typeof gradientWindowMeasurement; currentChartWindowMeasurement: typeof currentChartWindowMeasurement; windowControllerMeasurement: typeof windowControllerMeasurement; annotationWindowMeasurement: typeof annotationWindowMeasurement; commentTextMeasurement: typeof commentTextMeasurement; commentPanelMeasurement: typeof commentPanelMeasurement; mentionPanelMeasurement: typeof mentionPanelMeasurement; annotationControlsAction: typeof annotationControlsAction; mountAnnotationControls: typeof mountAnnotationControls } } }
let code = '';
const inputs: { name: string; files: [string, string][] }[] = [];
const comments: typeof inputs = [];
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-drawing-window.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookDrawingWindow' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected drawing-window bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  for (const prefix of ['', 'producer/']) for (const name of ['07-three-charts', '08-images-shapes']) {
    const files = await importXlsx(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`), { documentId: 'drawing-window' });
    inputs.push({ name: prefix + name, files: encodeTestPackage(files) });
  }
  for (const prefix of ['', 'producer/']) for (const name of ['09-classic-notes', '10-modern-threads']) {
    const files = await importXlsx(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`), { documentId: 'comment-window' });
    comments.push({ name: prefix + name, files: encodeTestPackage(files) });
  }
});
test.beforeEach(async ({ page }) => { await openSdkHost(page); await page.addScriptTag({ content: code }); });
test('visible native note boxes preserve selected VML appearance and exact source text in all frozen panes', async ({ page }) => {
  const files = comments.find(input => input.name === '09-classic-notes')!.files;
  for (const rtl of [false, true]) for (const vertical of ['top', 'center', 'bottom', 'justify', 'distributed']) {
    const result = await page.evaluate(({files,control}) => window.workbookDrawingWindow.nativeNoteMeasurement(files,control), {files,control:{rtl,vertical}});
    expect(result).toMatchObject({unchanged:true,preserved:true,released:true,canonical:0,errors:['ResourceLimitExceeded','ResourceLimitExceeded','ResourceLimitExceeded','ResourceLimitExceeded','OperationCancelled']});
    expect(result.actual).toBe(result.nodes); expect(result.fonts).toBeGreaterThan(0); expect(result.notes.map(note=>note.pane)).toEqual(['corner','frozenRows','frozenColumns','body']);
    for (const note of result.notes) {
      expect(note.exact).toBe(true); expect(note.text).toBe('Note A😀\nNote B禰\u{e0100}\nNote C');
      const first=note.projections[0]!,last=note.projections.at(-1)!;
      if(vertical==='top') expect(first.y).toBeLessThan(8);
      if(vertical==='bottom') expect(last.y+last.height).toBeGreaterThan(note.areaHeight-8);
      if(vertical==='justify'||vertical==='distributed') expect(last.y-first.y).toBeGreaterThan(100);
    }
    expect(result.markers).toBe(vertical === 'distributed' ? 2 : 0);
    expect(result.paint).toEqual([{fill:'#0000ff',opacity:'0.75',stroke:'#0000ff',dash:null},{fill:'#ff0000',opacity:'0.5',stroke:'#008000',dash:'8 6'}]);
    expect(result.pixels[1]![0]).toBeGreaterThan(100); expect(result.pixels[1]![2]).toBeGreaterThan(80);
  }
});
test('hidden notes stay in the panel and native visible VML sources open without fabricated text',async({page})=>{
  for(const input of comments.filter(input=>input.name.endsWith('09-classic-notes'))) {
    const result=await page.evaluate(files=>window.workbookDrawingWindow.nativeNoteMeasurement(files),input.files);
    expect(result).toMatchObject({notes:[],unchanged:true,preserved:true,released:true});expect(result.actual).toBe(result.nodes);
  }
  for(const suffix of ['', '-producer']) {
    const files=await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-dual-note-boxes-visible'+suffix+'.xlsx'),{documentId:'visible-native'});
    const result=await page.evaluate(entries=>window.workbookDrawingWindow.nativeNoteMeasurement(entries),encodeTestPackage(files));
    expect(result).toMatchObject({unchanged:true,preserved:true,released:true});expect(result.actual).toBe(result.nodes);expect(result.notes.length).toBeGreaterThan(0);expect(result.notes.every(note=>note.exact)).toBe(true);
  }
});
test('native chart labels and points use current formulas while wrong cell and chart caches remain unchanged', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-formulas.xlsx'), { documentId: 'current-chart-window' });
  const result = await page.evaluate(entries => window.workbookDrawingWindow.currentChartWindowMeasurement(entries), encodeTestPackage(files));
  expect(result).toMatchObject({ objects: 1, unavailable: 0, sourceUnchanged: true, released: true });
  expect(result.text).toContain('Current values'); expect(result.text).toContain('Point0'); expect(result.text).not.toContain('Cache');
  expect(result.diagnostics).toContainEqual(expect.objectContaining({ code: 'error-point', detail: '#DIV/0!' }));
});
test('modern mentions retain exact text and people in a non-interactive panel while conflicting locations require recovery', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-comment-mentions.xlsx'), { documentId: 'mention-panel' });
  const result = await page.evaluate(entries => window.workbookDrawingWindow.mentionPanelMeasurement(entries), encodeTestPackage(files));
  expect(result).toMatchObject({ unchanged: true, released: true }); expect(result.results).toHaveLength(5);
  expect(result.results.every(item => item.exact && item.interactive === 0)).toBe(true);
  expect(result.results.flatMap(item => item.labels)).toEqual([...Array(3).fill('提及：复核人'), '提及：只被提及']);
  expect(result.results[0]!.text).toContain('回复与主评论的位置不一致'); expect(result.results[0]!.text).toContain('请复核预算😀禰󠄀');
  expect(result.results[2]!.text).toContain('提及的文字范围需要重新定位'); expect(result.results[3]!.text).toContain('缺少原始单元格位置');
  expect(result.results[4]!.status).toBe('open');
});
test('revision UIDs remain opaque source metadata while compatibility identity conflicts have a visible recovery explanation', async ({ page }) => {
  const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-comment-uids.xlsx'), { documentId: 'uid-panel' });
  const result = await page.evaluate(entries => window.workbookDrawingWindow.mentionPanelMeasurement(entries), encodeTestPackage(files));
  expect(result).toMatchObject({ unchanged: true, released: true });
  expect(result.results.map(item => item.status)).toEqual(['resolved', 'needs_reanchor', 'open']);
  expect(result.results.every(item => item.exact && item.interactive === 0 && !item.labels.length)).toBe(true);
  expect(result.results[1]!.text).toContain('兼容记录的身份不一致'); expect(result.results[2]!.text).toContain('请确认：预算😀禰󠄀');
});
for (const rtl of [false, true]) test(`worksheet objects compose above cells across four ${rtl ? 'RTL' : 'LTR'} panes with isolated candidate resources`, async ({ page }) => {
  for (const input of inputs) {
    const result = await page.evaluate(({ files, rtl }) => window.workbookDrawingWindow.drawingWindowMeasurement(files, rtl), { files: input.files, rtl });
    expect(result, input.name).toMatchObject({ detached: true, unchanged: true, preserved: true, cleaned: true, released: true, clearedImages: true, removed: true, uniqueIds: true, cancelled: 'OperationCancelled', unhandled: [] });
    expect(result.errors).toEqual(Array(6).fill('ResourceLimitExceeded')); expect(result.actualNodes).toBe(result.nodes); expect(result.gridCells).toBeGreaterThan(0); expect(result.shadowPixels).toBeGreaterThan(0);
    const identities = [...new Set(result.objects.map(object => object.nodeId))]; expect(identities).toHaveLength(3); expect(identities).toEqual(result.expectedOrder); expect(result.canonicalIds).toBe(0);
    expect(result.fragments.map(fragment => fragment.nodeId)).toEqual(identities.flatMap(id => [id, id, id, id]));
    for (const id of identities) {
      const parts = result.objects.filter(object => object.nodeId === id); expect(parts.map(part => part.pane)).toEqual(['corner', 'frozenRows', 'frozenColumns', 'body']);
      expect(result.fragments.filter(fragment => fragment.nodeId === id).map(fragment => fragment.hidden)).toEqual([null, 'true', 'true', 'true']);
      for (const part of parts) {
        const fragment = result.fragments.find(fragment => fragment.nodeId === id && fragment.pane === part.pane)!;
        // CSSOM serializes point lengths to about six significant digits.
        expect(Math.abs(fragment.x - (rtl ? 420 - part.clip.x - part.clip.width : part.clip.x))).toBeLessThan(.001); expect(Math.abs(fragment.width - part.clip.width)).toBeLessThan(.001);
      }
      for (const part of parts) for (const projection of part.projections ?? []) { expect(projection.owned).toBe(true); expect(projection.to - projection.from).toBe(projection.text.length); }
      expect(parts.map(part => part.text)).toEqual(Array(4).fill(parts[0]!.text));
    }
    if (input.name.endsWith('08-images-shapes')) {
      expect(result.imageCount).toBe(4); expect(result.imageUrls).toBe(1); expect(result.assets.count).toBe(1); expect(result.live).toBe(1); expect(result.unavailable).toBe(0);
      expect(result.cropError).toBe('ResourceLimitExceeded');
    } else { expect(result.assets.count).toBe(0); expect(result.unavailable).toBe(4); expect(result.staleDiagnostics).toBeGreaterThan(0); }
  }
});


test('page-oriented picture fills retain complete pane budgets and isolated failed candidates', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('08-images-shapes'))) for (const rtl of [false, true]) {
    const result = await page.evaluate(({ files, rtl }) => window.workbookDrawingWindow.drawingWindowMeasurement(files, rtl, false, true), { files: input.files, rtl });
    expect(result).toMatchObject({ detached: true, unchanged: true, preserved: true, cleaned: true, released: true, clearedImages: true, removed: true, uniqueIds: true, cancelled: 'OperationCancelled', unhandled: [], pageOrientedImages: 4, imageCount: 4, imageUrls: 1, live: 1, cropError: 'ResourceLimitExceeded' });
    expect(result.nodes).toBe(result.actualNodes); expect(result.errors).toEqual(Array(6).fill('ResourceLimitExceeded'));
  }
});

test('chart plot series and marker gradients share fixed fill orientation across clipped panes', async ({ page }) => {
  const input = inputs.find(input => input.name === '07-three-charts')!;
  const result = await page.evaluate(files => window.workbookDrawingWindow.drawingWindowMeasurement(files, false, false, false, true), input.files);
  expect(result).toMatchObject({ detached: true, unchanged: true, preserved: true, cleaned: true, released: true, removed: true, uniqueIds: true, cancelled: 'OperationCancelled', unhandled: [] });
  expect(result.nodes).toBe(result.actualNodes); expect(result.fixedPaints.length).toBeGreaterThan(20); expect(result.fixedPaints.some(paint => paint.marker)).toBe(true);
  for (const paint of result.fixedPaints) { expect(Math.abs(paint.dy)).toBeLessThan(.0002); expect(paint.dx).toBeGreaterThan(0); }
});

test('fixed gradient fields retain native directions while source scaling and Calc geometry losses stay separate', async ({ page }) => {
  const cases = JSON.parse(readFileSync('docs/evidence/xlsx-gradient-rotation.json', 'utf8')).cases as { rotation: number; rotates: boolean | null; angle: number; scaled: boolean; sourcePdf: { normalizedVector: number[]; aspect: number }; producerPdf: { normalizedVector: number[]; aspect: number } }[];
  for (const suffix of ['', '-producer']) {
    const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-gradient-rotation' + suffix + '.xlsx'), { documentId: 'gradient-window' });
    const result = await page.evaluate(entries => window.workbookDrawingWindow.gradientWindowMeasurement(entries), encodeTestPackage(files));
    expect(result).toMatchObject({ unchanged: true, released: true }); expect(result.results).toHaveLength(32);
    for (const [index, item] of result.results.entries()) {
      expect(item.nodes).toBe(item.actual); expect(item.objects).toBe(1); expect(item.vector.every(Number.isFinite)).toBe(true);
      for (const [sample, pixel] of item.pixels.entries()) {
        const expected = sample ? [89, 0, 166, 255] : [166, 0, 89, 255];
        for (const [channel, value] of expected.entries()) expect(Math.abs(pixel[channel]!-value), `gradient ${index}, sample ${sample}`).toBeLessThanOrEqual(5);
      }
      const source = cases[index]!, native = source[suffix ? 'producerPdf' : 'sourcePdf'];
      if (!suffix && source.rotates !== false) {
        const [width, height] = item.size, expected = source.angle === 0 ? [width!, 0] : source.scaled ? [width!, height!] : [(width!+height!)/2, (width!+height!)/2];
        for (const axis of [0, 1]) expect(item.localVector[axis]).toBeCloseTo(expected[axis]!, 3);
        // The chosen rotating-field policy keeps the original coverage.
        // Calc repaints the page bounding rectangle for this same source.
        if (source.rotation === 45 && source.angle === 0) for (const axis of [0, 1]) { expect(item.vector[axis]).toBeCloseTo(2/3, 4); expect(native.normalizedVector[axis]).toBeCloseTo(1, 2); }
      }
      // Calc drops scaled and redefines rotating fields. Compare the actual
      // fixed/unscaled sources, and every separately saved producer field.
      if (suffix || source.rotates === false && (!source.scaled || source.angle === 0)) {
        const direction = Math.atan2(item.vector[1]!, item.vector[0]!*item.aspect), expected = Math.atan2(native.normalizedVector[1]!, native.normalizedVector[0]!*native.aspect);
        expect(Math.abs(Math.atan2(Math.sin(direction-expected), Math.cos(direction-expected)))).toBeLessThan(.003);
        if (source.rotation !== 90) for (const axis of [0, 1]) expect(item.vector[axis]).toBeCloseTo(native.normalizedVector[axis]!, 2);
      }
    }
  }
});

test('offscreen object anchors still project visible hanging text outside the anchor', async ({ page }) => {
  const input = inputs.find(input => input.name === '08-images-shapes')!;
  const result = await page.evaluate(files => window.workbookDrawingWindow.drawingWindowMeasurement(files, false, true), input.files);
  expect(result.overflowParts.length).toBeGreaterThan(0);
  expect(result.overflowParts.every(part => part.x > 420 && part.left < -80)).toBe(true);
  expect(result.overflowParts.some(part => part.visible)).toBe(true);
  expect(result.diagnostics).toContain('drawing-text-overflow');
  expect(result).toMatchObject({ unchanged: true, preserved: true, released: true, cleaned: true });
});

test('whole workbook windows commit atomically across scrolling, sheet changes, late providers and disposal', async ({ page }) => {
  for (const input of inputs.filter(input => !input.name.includes('/'))) {
    const result = await page.evaluate(files => window.workbookDrawingWindow.windowControllerMeasurement(files), input.files);
    expect(result, input.name).toMatchObject({ first: true, errors: ['RevisionConflict', 'InvalidDocument'], retainedWhilePending: true, switched: true,
      obsoleteCancelled: true, empty: { sheetId: 'controller-empty', objects: 0, width: 380, height: 260 }, lateCellIgnored: true, releasedOld: true,
      budget: 'ResourceLimitExceeded', failedMountEmpty: true, cancelledOnDispose: true, released: true, imagesCleared: true, emptyAfterDispose: true, formatterStillOwned: true, unhandled: [], unchanged: true });
    expect(result.nodes).toBe(result.actualNodes);
    expect(result.races).toHaveLength(2);
    for (const race of result.races) expect(race).toMatchObject({ cancelled: true, aborted: true, preserved: true });
  }
});

test('native classic notes and modern threads retain source content and pane locations in complete windows', async ({ page }) => {
  for (const input of comments) {
    const result = await page.evaluate(files => window.workbookDrawingWindow.annotationWindowMeasurement(files), input.files);
    expect(result.visible.map(item => item.id), input.name).toEqual(result.expected);
    // The pinned Calc producer drops modern threads; its output contains none.
    expect(result.visible.length).toBe(input.name === 'producer/10-modern-threads' ? 0 : input.name === '10-modern-threads' ? 2 : 1);
    expect(result.recovery).toEqual(result.expectedRecovery); expect(result.unchanged).toBe(true);
    for (const item of result.visible) { expect(item.body?.length).toBeGreaterThan(0); expect(item.ranges.length).toBeGreaterThan(0); }
    // These independent sources have no frozen comments: scrolling to the
    // last row/column must remove their geometry from the published window.
    expect(result.offscreen).toBe(0);
  }
});

test('comment body flow preserves native run styles, paragraph offsets and UTF-16 selection without source identities', async ({ page }) => {
  for (const input of comments.filter(input => input.name !== 'producer/10-modern-threads')) for (const rich of input.name === '09-classic-notes' ? [false, true, 'singleAccounting', 'doubleAccounting'] as const : [false] as const) {
    const result = await page.evaluate(({ files, rich }) => window.workbookDrawingWindow.commentTextMeasurement(files, rich), { files: input.files, rich });
    expect(result.unchanged).toBe(true); expect(result.results.length).toBeGreaterThan(0);
    for (const body of result.results) {
      expect(body).toMatchObject({ detached: true, unchangedSource: true, canonicalIds: 0, selected: true, errors: ['ResourceLimitExceeded', 'ResourceLimitExceeded'] });
      expect(body.nodes).toBe(body.actualNodes); expect(body.fonts).toBeGreaterThan(0);
      for (const piece of body.offsets) { expect(piece.text).toBe(piece.expected); expect(piece.to - piece.from).toBe(piece.text.length); expect(piece.sourceNodeId).toBe(body.id); }
      if (!rich) expect(body.text).toBe(body.expected);
      else {
        const start = body.expected.indexOf('混排😀禰󠄀'), runs = body.formats.filter(run => run.from >= start && run.from < start + '混排😀禰󠄀'.length);
        expect(runs.map(run => run.text).join('')).toBe('混排😀禰󠄀');
        for (const run of runs) expect(run).toMatchObject({ underline: 'underline', style: rich === 'singleAccounting' ? 'solid' : 'double', strike: 'line-through', color: 'rgb(255, 0, 0)' });
        expect(body.formats.find(run => run.text === 'No underline')).toMatchObject({ underline: 'none', strike: '' });
      }
    }
  }
});

test('read-only comment candidates retain recovery state and reply parents while failed replacements preserve the active panel', async ({ page }) => {
  for (const name of ['09-classic-notes', '10-modern-threads']) for (const recovered of name === '10-modern-threads' ? [false, true] : [false]) {
    const input = comments.find(input => input.name === name)!;
    const result = await page.evaluate(({ files, recovered }) => window.workbookDrawingWindow.commentPanelMeasurement(files, recovered), { files: input.files, recovered });
    expect(result).toMatchObject({ detached: true, editable: 0, released: true, empty: true, unchanged: true, plain: { elements: 0, bodies: 0 } });
    expect(result.nodes).toBe(result.actualNodes); expect(result.fonts).toBeGreaterThan(0);
    expect(result.errors).toEqual([...Array(recovered ? 4 : 3).fill('ResourceLimitExceeded'), 'RevisionConflict', 'InvalidOperation']);
    expect(result.plain.text).toContain('<script>author</script>'); expect(result.plain.text).toContain('<img src="https://invalid.example/private">仅作为文字');
    for (const body of result.bodies) { expect(body.text?.length).toBeGreaterThan(0); expect(body.projections).toBeGreaterThan(0); }
    if (recovered) { expect(result.warning).toContain('位置记录不一致'); expect(result.warning).toContain('原状态：已处理'); expect(result.replyParents).toHaveLength(1); expect(result.text).toContain('回复中的纯文本'); }
    else expect(result.warning).toBeUndefined();
  }
});

for (const rtl of [false, true]) test(`annotation entries open read-only comments with pointer and keyboard in ${rtl ? 'RTL' : 'LTR'} panes`, async ({ page }) => {
  const mounted = await page.evaluate(rtl => window.workbookDrawingWindow.mountAnnotationControls(rtl), rtl);
  expect(mounted.groups).toEqual([['first', 'same'], ['separate']]);
  expect(mounted).toMatchObject({ budget: 'ResourceLimitExceeded', selectionPreserved: true, sourceBlocked: true, trustedEvents: 1 });
  expect(mounted.violations.filter(directive => directive.startsWith('script-src')).length).toBeGreaterThanOrEqual(2);
  await page.mouse.click(mounted.x, mounted.y);
  const clicked = await page.evaluate(() => window.workbookDrawingWindow.annotationControlsAction('inspect')) as { calls: string[][]; panel: string; highlights: { x: number; y: number; width: number; height: number; top: string; left: string; right: string; bottom: string }[]; nodes: number; actualNodes: number };
  expect(clicked).toMatchObject({ calls: [['first', 'same']], sourceClean: true, unchanged: true, pressed: ['true', 'false'] });
  expect(clicked.panel).toContain('批注正文 first'); expect(clicked.highlights).toHaveLength(5); expect(clicked.nodes).toBe(clicked.actualNodes);
  // The frozen pane cut is not a border of the original source range.
  expect(clicked.highlights[0]!.bottom).toBe(''); expect(clicked.highlights[0]![rtl ? 'left' : 'right']).toBe('');
  await page.evaluate(() => window.workbookDrawingWindow.annotationControlsAction('fail'));
  await page.evaluate(() => window.workbookDrawingWindow.annotationControlsAction('focus')); await page.keyboard.press('Enter');
  const failed = await page.evaluate(() => window.workbookDrawingWindow.annotationControlsAction('inspect'));
  expect(failed).toMatchObject({ calls: [['first', 'same'], ['first', 'same']], panel: clicked.panel, highlights: clicked.highlights, error: 'ResourceLimitExceeded', unchanged: true });
  await page.evaluate(() => window.workbookDrawingWindow.annotationControlsAction('clear'));
  await page.keyboard.press('Enter');
  expect(await page.evaluate(() => window.workbookDrawingWindow.annotationControlsAction('inspect'))).toMatchObject({ calls: [['first', 'same'], ['first', 'same'], ['first', 'same']], panel: clicked.panel });
  expect(await page.evaluate(() => window.workbookDrawingWindow.annotationControlsAction('dispose'))).toEqual({ eventsReleased: true, removed: true, sourceClean: true });
});


test('classic dual-box notes use VML while preserving alternate properties and reporting differences without changing body offsets', async ({ page }) => {
  for (const name of ['conflict', 'same-anchor', 'visible', 'commentPr-only']) {
    const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-dual-note-boxes-' + name + '.xlsx'), { documentId: 'dual-box-panel-' + name });
    const result = await page.evaluate(entries => window.workbookDrawingWindow.mentionPanelMeasurement(entries, true), encodeTestPackage(files));
    expect(result).toMatchObject({ unchanged: true, released: true }); expect(result.results).toHaveLength(1);
    const item = result.results[0]!; expect(item.exact).toBe(true); expect(item.interactive).toBe(0); expect(item.nodes).toBe(item.actualNodes);
    expect(item.errors).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded']);
    expect(item.text).toContain('请确认：预算😀禰󠄀');
    expect(item.box!.source).toBe(name === 'commentPr-only' ? 'commentPr' : 'vml');
    if (name === 'commentPr-only') expect(item.boxWarning).toBeUndefined();
    else {
      expect(item.box!.selected!.anchor).toMatchObject({ kind: 'two-cell', from: { row: 1, column: 2, offsetXPt: 7.5, offsetYPt: 1.5 } });
      expect(item.box!.alternate).toMatchObject({ printable: false, altText: '备用备注框😀' });
      expect(item.boxWarning).toContain('已采用 VML'); expect(item.boxWarning).toContain('打印'); expect(item.boxWarning).toContain('另一份信息已保留');
      expect(item.box!.differences!.includes('anchor')).toBe(name !== 'same-anchor');
    }
  }
});


test('chart point and label overrides survive later general/series edits with exact current-source reopen and node budgets', async ({ page }) => {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: 'style-layers' }), { kind: 'workbook' }));
  const commit = (operations: DocumentOperation[]) => session.commit(session.prepare({ protocolVersion: 1, documentId: 'style-layers', baseRevision: session.snapshot().manifest.revision, operations }));
  commit([{ op: 'insertWorkbookChart', sheetId: 'sheet-1', kind: 'column', anchor: { kind: 'absolute', xPt: 20, yPt: 20, widthPt: 350, heightPt: 230 }, data: { categories: { kind: 'literal', valueType: 'text', count: 2, points: [{ index: 0, value: '甲' }, { index: 1, value: '乙' }] }, series: [{ values: { kind: 'literal', valueType: 'number', count: 2, points: [{ index: 0, value: 2 }, { index: 1, value: 4 }] } }] } }]);
  const chart = session.snapshot().metadata.workbook.charts![0]!, seriesId = chart.series[0]!.seriesId;
  commit([{ op: 'setWorkbookChart', nodeId: chart.nodeId, changes: { showLabels: true } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'point', seriesId, index: 0, properties: { '--office-fill-kind': 'solid', '--office-fill-color': '#c02030' } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'pointLabel', seriesId, index: 0, properties: { color: '#703090', 'font-size': '16pt' } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'series', seriesId, properties: { '--office-fill-kind': 'solid', '--office-fill-color': '#2050c0' } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'seriesLabel', seriesId, properties: { color: '#20a040', 'font-size': '11pt' } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'label', properties: { color: '#a05020', 'font-size': '9pt' } }]);
  const saved = session.serialize(); session.dispose();
  const result = await page.evaluate(files => window.workbookDrawingWindow.currentChartWindowMeasurement(files), encodeTestPackage(saved));
  expect(result).toMatchObject({ sourceUnchanged: true, released: true, unavailable: 0, objects: 1 }); expect(result.declaredNodes).toBe(result.actualNodes);
  expect(result.paints).toHaveLength(1); expect(result.paints[0]!.columns).toEqual(['rgba(192,32,48,1)', 'rgba(32,80,192,1)']);
  expect(result.paints[0]!.labels).toMatchObject([{ text: '2', fill: 'rgba(112,48,144,1)' }, { text: '4', fill: 'rgba(32,160,64,1)' }]);
  // Browser font sizes can be quantized to 1/60 CSS px; source pt values
  // are checked exactly in the transaction/native tests.
  expect(Math.abs(Number(result.paints[0]!.labels[0]!.size) - 16)).toBeLessThanOrEqual(.75 / 60); expect(Math.abs(Number(result.paints[0]!.labels[1]!.size) - 11)).toBeLessThanOrEqual(.75 / 60);
});


test('chart label font decorations inherit through native roles and point none clears the default', async ({ page }) => {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: 'label-decorations' }), { kind: 'workbook' }));
  const commit = (operations: DocumentOperation[]) => session.commit(session.prepare({ protocolVersion: 1, documentId: 'label-decorations', baseRevision: session.snapshot().manifest.revision, operations }));
  commit([{ op: 'insertWorkbookChart', sheetId: 'sheet-1', kind: 'column', anchor: { kind: 'absolute', xPt: 20, yPt: 20, widthPt: 350, heightPt: 230 }, data: { series: [{ values: { kind: 'literal', valueType: 'number', count: 3, points: [{ index: 0, value: 2 }, { index: 1, value: 4 }, { index: 2, value: 6 }] } }] } }]);
  const chart = session.snapshot().metadata.workbook.charts![0]!, seriesId = chart.series[0]!.seriesId;
  commit([{ op: 'setWorkbookChart', nodeId: chart.nodeId, changes: { showLabels: true } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'label', properties: { 'text-decoration': 'var(--decor)', '--office-font-ascii': '"Liberation Sans"', color: '#203040' } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'seriesLabel', seriesId, properties: { 'font-weight': 'bold' } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'pointLabel', seriesId, index: 0, properties: { 'text-decoration-line': 'none' } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'pointLabel', seriesId, index: 1, properties: { 'text-decoration': 'line-through solid' } },
    { op: 'setWorkbookChartStyle', nodeId: chart.nodeId, target: 'pointLabel', seriesId, index: 2, properties: { color: '#2050c0', '--office-text-lum-mod': '.5' } }]);
  const saved = new Map(session.serialize()), stylePath = session.snapshot().manifest.styles[0]!; session.dispose();
  saved.set(stylePath, new TextEncoder().encode(new TextDecoder().decode(saved.get(stylePath)!) + ':root{--decor:underline double currentColor}'));
  const result = await page.evaluate(files => window.workbookDrawingWindow.currentChartWindowMeasurement(files, true), encodeTestPackage(saved));
  expect(result).toMatchObject({ sourceUnchanged: true, released: true, unavailable: 0 }); expect(result.actualNodes).toBe(result.declaredNodes);
  expect(result.paints[0]!.labels).toMatchObject([{ text: '2', decoration: 'none' }, { text: '4', decoration: 'line-through', decorationStyle: 'solid' }, { text: '6', decoration: 'underline', decorationStyle: 'double', decorationColor: 'rgb(16, 40, 96)', fill: 'rgba(16,40,96,1)' }]);
  expect(result.ink[0]!.pixels).toBe(0); expect(result.ink[1]!.pixels).toBeGreaterThan(0); expect(result.ink[2]!.pixels).toBeGreaterThan(0); expect(result.ink[2]!.bands).toBe(2); expect(result.ink[2]!.colors[0]!.rgb).toBe('16,40,96');
});
