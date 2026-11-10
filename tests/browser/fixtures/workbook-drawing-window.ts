import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createWorkbookNumberFormatter, openDocument } from '@doffi/core';
import type { WorkbookDocumentMetadata } from '@doffi/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { measureWorkbookColumnFont } from '../../../packages/render/src/workbook/column-font.js';
import { worksheetDimensions } from '../../../packages/render/src/workbook/dimensions.js';
import { WorkbookWindowLoader } from '../../../packages/render/src/workbook/window-loader.js';
import { modelWindowSource } from '../../../packages/render/src/workbook/window-source.js';
import { modelAssetSource } from '../../../packages/render/src/workbook/assets.js';
import type { WorkbookAssetResponse } from '../../../packages/render/src/workbook/assets.js';
import { modelChartSource } from '../../../packages/render/src/workbook/chart-source.js';
import type { WorkbookChartReadResponse } from '../../../packages/render/src/workbook/chart-source.js';
import { layoutWorkbookWindow } from '../../../packages/render/src/workbook/window-layout.js';
import { layoutWorkbookDrawings } from '../../../packages/render/src/workbook/drawing-layout.js';
import { WorkbookWindowController } from '../../../packages/render/src/workbook/window-controller.js';
import type { WorkbookDrawingLayout } from '../../../packages/render/src/workbook/drawing-layout.js';
import { annotationWindowMeasurement, commentTextMeasurement, windowControllerMeasurement } from './workbook-window-controller.js';
import { commentPanelMeasurement, mentionPanelMeasurement } from './workbook-comment-panel.js';
import { annotationControlsAction, mountAnnotationControls } from './workbook-annotation-controls.js';
import { nativeNoteMeasurement } from './workbook-native-notes.js';

export async function gradientWindowMeasurement(entries: TestPackage) {
  const model = openDocument(decodeTestPackage(entries), { kind: 'workbook' }), snapshot = model.snapshot(), before = model.serialize();
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 }), formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
  const controller = new WorkbookWindowController(surface, { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model) }, formatter);
  try {
    const results = [];
    for (const sheet of snapshot.metadata.workbook.sheets) {
      await controller.render(sheet.sheetId, { widthPt: 420, heightPt: 300, scrollLeftPt: 0, scrollTopPt: 0 });
      const current = controller.current!, body = current.element.querySelector<HTMLElement>('[data-office-view=workbook-drawing-body]')!, gradient = body.querySelector('linearGradient')!;
      const matrix = new DOMMatrix(surface.document.defaultView!.getComputedStyle(body).transform), width = parseFloat(body.style.width), height = parseFloat(body.style.height);
      const [x1, y1, x2, y2] = ['x1', 'y1', 'x2', 'y2'].map(key => Number(gradient.getAttribute(key)));
      const dx = matrix.a*(x2!-x1!) + matrix.c*(y2!-y1!), dy = matrix.b*(x2!-x1!) + matrix.d*(y2!-y1!);
      const bboxWidth = Math.abs(matrix.a)*width + Math.abs(matrix.c)*height, bboxHeight = Math.abs(matrix.b)*width + Math.abs(matrix.d)*height;
      const walker = surface.document.createTreeWalker(current.element, 0xffff_ffff); let actual = 1; while (walker.nextNode()) actual++;
      const ns = 'http://www.w3.org/2000/svg', outer = document.createElementNS(ns, 'svg'), group = document.createElementNS(ns, 'g'), graphic = body.querySelector('svg')!.cloneNode(true) as SVGSVGElement;
      outer.setAttribute('width', '400'); outer.setAttribute('height', '320'); outer.setAttribute('viewBox', '0 0 400 320');
      group.setAttribute('transform', `translate(200 160) matrix(${matrix.a} ${matrix.b} ${matrix.c} ${matrix.d} 0 0) translate(${-width/2} ${-height/2})`);
      graphic.setAttribute('width', String(width)); graphic.setAttribute('height', String(height)); graphic.style.width = width+'px'; graphic.style.height = height+'px'; group.append(graphic); outer.append(group);
      const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(outer)], { type: 'image/svg+xml' })), image = new Image(); let pixels: number[][];
      try {
        image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 400; canvas.height = 320; const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
        pixels = [-.15, .15].map(fraction => [...context.getImageData(Math.round(200+dx*fraction), Math.round(160+dy*fraction), 1, 1).data]);
      } finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
      results.push({ nodes: current.nodes, actual, vector: [dx/bboxWidth, dy/bboxHeight], localVector: [x2!-x1!, y2!-y1!], size: [width, height], aspect: bboxWidth/bboxHeight, pixels, objects: current.objects.length });
    }
    const after = model.serialize(); controller.dispose();
    return { results, unchanged: [...before].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((value, index) => value === after.get(path)![index])), released: !surface.view.childNodes.length && !surface.source.childNodes.length };
  } finally { controller.dispose(); formatter.dispose(); surface.dispose(); container.remove(); }
}

export async function currentChartWindowMeasurement(entries: TestPackage, inspectInk = false) {
  const model = openDocument(decodeTestPackage(entries), { kind: 'workbook' }), snapshot = model.snapshot(), before = model.serialize();
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 620, heightPt: 420 }), formatter = createWorkbookNumberFormatter({ dateSystem: '1900' });
  const controller = new WorkbookWindowController(surface, { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model) }, formatter);
  try {
    await controller.render(snapshot.metadata.workbook.activeSheetId, { widthPt: 620, heightPt: 420, scrollLeftPt: 0, scrollTopPt: 0 });
    const window = controller.current!, text = window.element.querySelector('[data-office-view="workbook-drawings"]')!.textContent;
    const after = model.serialize();
    const paints = [...window.element.querySelectorAll('[data-sdk-chart]')].map(chart => ({
      columns: [...chart.querySelectorAll('[data-sdk-chart-point=column]')].map(point => point.getAttribute('fill')),
      labels: [...chart.querySelectorAll<SVGTextElement>('[data-sdk-chart-label=data-label]')].map(label => ({ text: label.textContent, fill: label.getAttribute('fill'), size: label.getAttribute('font-size'), decoration: label.dataset.sdkChartDecorationLine, decorationStyle: label.dataset.sdkChartDecorationStyle, decorationColor: (() => { const paint = label.closest("[data-sdk-chart-text-decoration]")?.querySelector<SVGElement>("[data-sdk-chart-decoration-paint] line,[data-sdk-chart-decoration-paint] path"); return paint ? surface.document.defaultView!.getComputedStyle(paint).stroke : label.dataset.sdkChartDecorationColor; })() })),
      frames: [...chart.querySelectorAll<SVGRectElement>('[data-sdk-chart-label-frame]')].map(frame => ({ index: frame.parentElement!.querySelector<SVGTextElement>('[data-point-index]')!.dataset.pointIndex, fill: frame.getAttribute('fill'), color: frame.getAttribute('stroke'), width: frame.width.baseVal.value, height: frame.height.baseVal.value })),
      legends: [...chart.querySelectorAll<SVGRectElement>('[data-sdk-chart-legend-frame]')].map(frame => ({ fill: frame.getAttribute('fill'), color: frame.getAttribute('stroke'), width: frame.width.baseVal.value, height: frame.height.baseVal.value })),
    }));
    const walker = surface.document.createTreeWalker(window.element, 0xffff_ffff); let actualNodes = 1; while (walker.nextNode()) actualNodes++;
    const ink = inspectInk ? await Promise.all([...window.element.querySelectorAll<SVGTextElement>('[data-sdk-chart-label=data-label]')].map(chartLabelInk)) : [];
    const framePixels = inspectInk ? await Promise.all([...window.element.querySelectorAll<SVGRectElement>('[data-sdk-chart-label-frame]')].filter(frame => frame.getAttribute('fill')?.startsWith('url(')).map(chartFramePixels)) : [];
    const result = { text, paints, ink, framePixels, actualNodes, declaredNodes: window.nodes, objects: window.objects.length, diagnostics: window.diagnostics, unavailable: window.element.querySelectorAll('[data-office-view="workbook-chart-unavailable"]').length,
      sourceUnchanged: [...before].every(([path, bytes]) => { const saved = after.get(path)!; return bytes.length === saved.length && bytes.every((byte, index) => byte === saved[index]); }) };
    controller.dispose(); return { ...result, released: !surface.view.childNodes.length && !surface.source.childNodes.length };
  } finally { controller.dispose(); formatter.dispose(); surface.dispose(); container.remove(); }
}

async function chartFramePixels(frame: SVGRectElement) {
  const source = frame.ownerSVGElement!, svg = source.cloneNode(false) as SVGSVGElement, bounds = frame.getBBox();
  svg.setAttribute('viewBox', [bounds.x, bounds.y, bounds.width, bounds.height].join(' ')); svg.setAttribute('width', '240'); svg.setAttribute('height', '120'); svg.setAttribute('preserveAspectRatio', 'none'); svg.style.cssText = '';
  for (const defs of source.querySelectorAll(':scope > defs')) svg.append(defs.cloneNode(true)); svg.append(frame.cloneNode(true));
  const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(svg)], { type: 'image/svg+xml' }));
  try {
    const image = new Image(); image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 240; canvas.height = 120;
    const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
    return [36, 204].map(x => [...context.getImageData(x, 60, 1, 1).data]);
  } finally { URL.revokeObjectURL(url); }
}

/** Raster comparison observes actual SVG decoration paint, independently of
 * the computed style and source declarations under test. */
async function chartLabelInk(label: SVGTextElement) {
  // At 6× even a 12pt double-line stroke covers an opaque raster pixel;
  // a lower scale can split the entire thin stroke across two partial pixels.
  const source = label.ownerSVGElement!, width = Math.ceil(source.viewBox.baseVal.width * 6), height = Math.ceil(source.viewBox.baseVal.height * 6);
  const raster = async (plain: boolean) => {
    const svg = source.cloneNode(false) as SVGSVGElement, text = (label.closest("[data-sdk-chart-text-decoration]") ?? label).cloneNode(true) as SVGGraphicsElement;
    svg.setAttribute('width', String(width)); svg.setAttribute('height', String(height)); svg.style.cssText = '';
    if (plain) { for (const paint of text.querySelectorAll('[data-sdk-chart-decoration-paint]')) paint.remove(); } svg.append(text);
    const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(svg)], { type: 'image/svg+xml' }));
    try {
      const image = new Image(); image.src = url; await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = width; canvas.height = height; const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
      return context.getImageData(0, 0, width, height).data;
    } finally { URL.revokeObjectURL(url); }
  };
  const [decorated, plain] = await Promise.all([raster(false), raster(true)]), colors = new Map<string, number>(), rows = new Set<number>(); let pixels = 0;
  for (let offset = 0; offset < decorated.length; offset += 4) if (decorated[offset + 3]! > 240 && plain[offset + 3]! < 16) {
    const key = [decorated[offset]!, decorated[offset + 1]!, decorated[offset + 2]!].join(','); colors.set(key, (colors.get(key) ?? 0) + 1); rows.add(Math.floor(offset / 4 / width)); pixels++;
  }
  const ordered = [...rows].sort((a, b) => a - b); let bands = 0; for (let i = 0; i < ordered.length; i++) if (!i || ordered[i]! > ordered[i - 1]! + 1) bands++;
  return { pixels, bands, colors: [...colors].sort((a, b) => b[1] - a[1]).slice(0, 3).map(([rgb, count]) => ({ rgb, count })) };
}

export async function drawingWindowMeasurement(entries: TestPackage, rtl: boolean, overflow = false, pageOrientedPictures = false, fixedGradients = false) {
  const files = decodeTestPackage(entries), decoder = new TextDecoder(), encoder = new TextEncoder();
  const metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  const sheet = metadata.workbook.sheets[0]!; sheet.frozen = { rows: 2, columns: 1 }; sheet.rightToLeft = rtl;
  for (const object of metadata.workbook.objects ?? []) object.anchor = { kind: 'absolute', xPt: 12, yPt: 12, widthPt: object.kind === 'chart' ? 360 : 220, heightPt: object.kind === 'chart' ? 220 : 150 };
  metadata.workbook.objects?.reverse(); // Source HTML, rather than this array, owns stacking.
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  const manifest = JSON.parse(decoder.decode(files.get('manifest.json')!)) as { styles: string[]; content: string };
  if (fixedGradients) {
    const template = document.createElement('template'); template.innerHTML = decoder.decode(files.get(manifest.content)!);
    for (const chart of metadata.workbook.charts!) {
      chart.plotCssClass = 'fixed-gradient';
      chart.seriesStyles = chart.series.map(series => ({ ...chart.seriesStyles?.find(style => style.seriesId === series.seriesId), seriesId: series.seriesId, cssClass: 'fixed-gradient', ...(chart.kind === 'line' ? { marker: 'circle' as const, markerSize: 7 } : {}) }));
      for (const point of chart.pointStyles ?? []) point.cssClass = 'fixed-gradient';
      const element = template.content.querySelector<HTMLElement>('[data-node-id="' + chart.nodeId + '"]')!; element.classList.add('fixed-gradient'); element.dataset.officeRotation = '45';
    }
    files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
    files.set(manifest.content, encoder.encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="' + manifest.styles[0]! + '"></head><body>' + template.content.querySelector('[data-office-kind="workbook-content"]')!.outerHTML + '</body></html>'));
    const declarations = ['--office-fill', '--office-marker-fill'].flatMap(prefix => [`${prefix}-kind:linear`, `${prefix}-angle:0`, `${prefix}-rotate-with-shape:false`, `${prefix}-scaled:false`, `${prefix}-stop-count:2`, `${prefix}-stop-0-position:0`, `${prefix}-stop-0-color:#ff0000`, `${prefix}-stop-1-position:1`, `${prefix}-stop-1-color:#0000ff`]).join(';');
    files.set(manifest.styles[0]!, encoder.encode(decoder.decode(files.get(manifest.styles[0]!)!) + '.fixed-gradient{' + declarations + '}'));
  }
  let overflowId: string | undefined;
  if (overflow) {
    const object = metadata.workbook.objects!.find(object => object.kind === 'text')!; overflowId = object.nodeId;
    object.anchor = { kind: 'absolute', xPt: 500, yPt: 45, widthPt: 70, heightPt: 30 }; object.text = { wrap: false, autoFit: 'none' };
    const template = document.createElement('template'); template.innerHTML = decoder.decode(files.get(manifest.content)!);
    const source = template.content.querySelector<HTMLElement>('[data-node-id="' + object.nodeId + '"]')!;
    source.setAttribute('data-office-rotation', '0'); source.setAttribute('data-office-flip-h', 'false'); source.setAttribute('data-office-flip-v', 'false');
    const paragraph = document.createElement('p'); paragraph.dataset.nodeId = 'overflow-paragraph';
    paragraph.style.cssText = 'font-family:Liberation Sans;font-size:18pt;line-height:24pt;text-align:left;text-indent:-180pt;margin:0';
    paragraph.style.setProperty('--office-font-ascii', '"Liberation Sans"'); paragraph.textContent = 'Outside heading '.repeat(12); source.replaceChildren(paragraph);
    files.set(manifest.content, encoder.encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="' + manifest.styles[0]! + '"></head><body>' + template.content.querySelector('[data-office-kind="workbook-content"]')!.outerHTML + '</body></html>'));
    files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  }
  const stylePath = manifest.styles[0]!;
  if (pageOrientedPictures) {
    const template = document.createElement('template'); template.innerHTML = decoder.decode(files.get(manifest.content)!);
    for (const image of template.content.querySelectorAll('[data-office-kind=image]')) image.setAttribute('data-office-image-rotate-with-shape', 'false');
    files.set(manifest.content, encoder.encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="' + manifest.styles[0]! + '"></head><body>' + template.content.querySelector('[data-office-kind="workbook-content"]')!.outerHTML + '</body></html>'));
  }
  files.set(stylePath, encoder.encode(decoder.decode(files.get(stylePath)!) + '\ndiv{--office-shadow-kind:outer;--office-shadow-color:#ff0000;--office-shadow-source:direct;--office-shadow-blur:4pt;--office-shadow-distance:3pt;--office-shadow-direction:30;--office-shadow-rotate-with-shape:false}'));
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), initial = model.serialize();
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 });
  const sources = new WorkbookContentSources(surface.document, snapshot), loader = new WorkbookWindowLoader(modelWindowSource(model));
  const formatter = createWorkbookNumberFormatter({ dateSystem: metadata.workbook.dateSystem }), assets = modelAssetSource(model), charts = modelChartSource(model);
  const normal = measureWorkbookColumnFont(surface.source, snapshot.metadata.workbook, sheet.sheetId), dimensions = worksheetDimensions(sheet, normal.maximumDigitWidthPx);
  const created = new Set<string>(), revoked = new Set<string>(), originalCreate = URL.createObjectURL, originalRevoke = URL.revokeObjectURL;
  URL.createObjectURL = blob => { const url = originalCreate.call(URL, blob); created.add(url); return url; };
  URL.revokeObjectURL = url => { revoked.add(url); originalRevoke.call(URL, url); };
  const unhandled: string[] = [], rejection = (event: PromiseRejectionEvent) => { unhandled.push(String(event.reason)); event.preventDefault(); };
  window.addEventListener('unhandledrejection', rejection);
  let active: WorkbookDrawingLayout | undefined;
  try {
    await loader.load(sheet.sheetId, dimensions, { widthPt: 420, heightPt: 300, scrollLeftPt: 20, scrollTopPt: 10 });
    const expected = sources.project(surface.source, metadata.workbook.objects!.map(object => object.nodeId));
    const expectedOrder = [...expected.get(metadata.workbook.objects![0]!.nodeId)!.parentElement!.children].map(node => (node as HTMLElement).dataset.nodeId!); expected.dispose();
    const frame = loader.frame!, grid = layoutWorkbookWindow(surface, snapshot, frame, dimensions, formatter, { normalSpaceWidthPx: normal.spaceWidthPx });
    active = await layoutWorkbookDrawings(surface, snapshot, frame, dimensions, sources, assets, charts, formatter, new AbortController().signal);
    const detached = !active.element.isConnected && surface.view.childNodes.length === 0;
    grid.element.append(active.element); surface.view.replaceChildren(grid.element); surface.frame.style.setProperty('visibility', 'visible', 'important');
    const stable = surface.view.innerHTML, originals = [...active.element.querySelectorAll('img')];
    const objects = active.objects.map(object => ({ nodeId: object.nodeId, pane: object.projection.pane, rect: object.projection.rect, clip: object.projection.clip,
      text: object.text?.text.element.textContent, paint: object.text?.paintBounds,
      projections: object.text?.text.projections.map(projection => ({ from: projection.from, to: projection.to, text: projection.node.data, owned: object.text!.element.contains(projection.node) })) }));
    const overflowParts = active.objects.filter(object => object.nodeId === overflowId).map(object => {
      const text = object.text!, clip = text.element.closest('[data-office-view="workbook-drawing"]')!.getBoundingClientRect(), range = surface.document.createRange();
      range.selectNodeContents(text.text.projections[0]!.node); const ink = range.getBoundingClientRect();
      return { x: object.projection.rect.x, left: text.paintBounds.x, visible: Math.max(ink.left, clip.left) < Math.min(ink.right, clip.right) && Math.max(ink.top, clip.top) < Math.min(ink.bottom, clip.bottom) };
    });
    const ids = [...active.element.querySelectorAll('[id]')].map(node => node.id), walker = surface.document.createTreeWalker(active.element, 0xffff_ffff);
    let count = 1; while (walker.nextNode()) count++;
    const fragments = [...active.element.children].map(node => ({ nodeId: (node as HTMLElement).dataset.sourceNodeId, pane: (node as HTMLElement).dataset.pane,
      x: parseFloat((node as HTMLElement).style.left), width: parseFloat((node as HTMLElement).style.width), hidden: node.getAttribute('aria-hidden') }));
    const errors: string[] = [];
    for (const limits of [{ maxNodes: 2 }, { maxSourceNodes: 1 }, { maxObjects: 1 }, { maxCodeUnits: 1 }, { maxShadowPixels: 1 }, { maxNodes: active.nodes - 1 }]) {
      try { (await layoutWorkbookDrawings(surface, snapshot, frame, dimensions, sources, assets, charts, formatter, new AbortController().signal, limits)).dispose(); errors.push('accepted'); }
      catch (error) { errors.push((error as { code: string }).code); }
      if (surface.view.innerHTML !== stable) throw Error('Failed drawings changed the active cell/object window');
    }
    let cropError: string | undefined;
    if (originals.length) {
      const extremeFiles = new Map(initial), template = document.createElement('template'); template.innerHTML = decoder.decode(initial.get(manifest.content)!);
      const image = template.content.querySelector('[data-office-kind="image"]')!;
      image.setAttribute('data-office-crop-left', '0.99999'); image.setAttribute('data-office-crop-right', '0');
      extremeFiles.set(manifest.content, encoder.encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="' + manifest.styles[0]! + '"></head><body>' + template.content.querySelector('[data-office-kind="workbook-content"]')!.outerHTML + '</body></html>'));
      const extreme = openDocument(extremeFiles, { kind: 'workbook' }).snapshot(), extremeSources = new WorkbookContentSources(surface.document, extreme);
      try { (await layoutWorkbookDrawings(surface, extreme, frame, dimensions, extremeSources, assets, charts, formatter, new AbortController().signal)).dispose(); cropError = 'accepted'; }
      catch (error) { cropError = (error as { code: string }).code; }
      finally { extremeSources.dispose(); }
    }
    const controller = new AbortController(); let begin!: () => void;
    const started = new Promise<void>(resolve => { begin = resolve; });
    let late!: () => Promise<void>, cancelled = '';
    const stalledAssets = { ...assets, readAsset(request: Parameters<typeof assets.readAsset>[0]) {
      return new Promise<WorkbookAssetResponse>(resolve => { late = async () => resolve(await assets.readAsset(request, new AbortController().signal)); begin(); });
    } };
    const stalledCharts = { ...charts, read(request: Parameters<typeof charts.read>[0]) {
      return new Promise<WorkbookChartReadResponse>(resolve => { late = async () => resolve(await charts.read(request, new AbortController().signal)); begin(); });
    } };
    const work = layoutWorkbookDrawings(surface, snapshot, frame, dimensions, sources, stalledAssets, stalledCharts, formatter, controller.signal);
    await started; controller.abort(); try { await work; } catch (error) { cancelled = (error as { code: string }).code; }
    await late(); await new Promise<void>(resolve => setTimeout(resolve, 0));
    let unavailable = 0, staleDiagnostics = 0;
    if (metadata.workbook.charts?.length) {
      const stale = { ...charts, async read(request: Parameters<typeof charts.read>[0], signal: AbortSignal): Promise<WorkbookChartReadResponse> {
        const response = await charts.read(request, signal);
        return { ...response, charts: response.charts.map((chart, index) => index ? chart : { ...chart, references: chart.references.map(reference => reference.role !== 'series-values' || reference.data.kind !== 'cells' ? reference : { ...reference,
          data: { ...reference.data, areas: reference.data.areas.map(area => ({ ...area, cells: area.cells.map((cell, index) => index ? cell : { ...cell, state: 'stale' as const }) })) } }) }) };
      } };
      const candidate = await layoutWorkbookDrawings(surface, snapshot, frame, dimensions, sources, assets, stale, formatter, new AbortController().signal);
      unavailable = candidate.element.querySelectorAll('[data-office-view="workbook-chart-unavailable"]').length;
      staleDiagnostics = candidate.diagnostics.filter(diagnostic => diagnostic.code === 'stale-cell').length; candidate.dispose();
    }
    const after = model.serialize(), unchanged = [...initial].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index]));
    const normalized = [...files].filter(([path, bytes]) => bytes.length !== initial.get(path)!.length || bytes.some((byte, index) => byte !== initial.get(path)![index])).map(([path]) => path);
    const fixedPaints = [...active.element.querySelectorAll<SVGLinearGradientElement>('linearGradient')].map(gradient => {
      const body = gradient.closest<HTMLElement>('[data-office-view=workbook-drawing-body]')!, matrix = new DOMMatrix(surface.document.defaultView!.getComputedStyle(body).transform);
      const dx = Number(gradient.getAttribute('x2'))-Number(gradient.getAttribute('x1')), dy = Number(gradient.getAttribute('y2'))-Number(gradient.getAttribute('y1'));
      return { marker: gradient.id.includes('-marker-'), dx: matrix.a*dx+matrix.c*dy, dy: matrix.b*dx+matrix.d*dy };
    });
    const result = { normalized, expectedOrder, overflowParts, fixedPaints, canonicalIds: active.element.querySelectorAll('[data-node-id]').length, objects, fragments, nodes: active.nodes, actualNodes: count, uniqueIds: new Set(ids).size === ids.length, detached, unchanged,
      gridCells: grid.element.querySelectorAll('[role="gridcell"]').length, live: [...created].filter(url => !revoked.has(url)).length,
      imageCount: originals.length, pageOrientedImages: active.element.querySelectorAll('[data-office-view=workbook-image-fill]').length, imageUrls: new Set(originals.map(image => image.src)).size, assets: active.assets, shadowPixels: active.shadowPixels,
      fonts: active.fonts.length, diagnostics: active.diagnostics.map(diagnostic => diagnostic.code), errors, cropError, cancelled, unavailable, staleDiagnostics,
      preserved: surface.view.innerHTML === stable && originals.every(image => image.isConnected && image.naturalWidth > 0 && image.hasAttribute('src')),
      cleaned: !surface.source.querySelector('[data-office-view="workbook-content-measurement"]') && !surface.view.querySelector('[data-office-view="workbook-drawing-text-measurement"]'), unhandled };
    active.dispose(); active.dispose();
    return { ...result, released: [...created].every(url => revoked.has(url)), clearedImages: originals.every(image => !image.hasAttribute('src')), removed: !active.element.parentNode && active.element.childNodes.length === 0 };
  } finally {
    active?.dispose(); loader.dispose(); sources.dispose(); formatter.dispose(); surface.dispose(); container.remove();
    URL.createObjectURL = originalCreate; URL.revokeObjectURL = originalRevoke; window.removeEventListener('unhandledrejection', rejection);
  }
}
Object.assign(window, { workbookDrawingWindow: { nativeNoteMeasurement, drawingWindowMeasurement, gradientWindowMeasurement, currentChartWindowMeasurement, windowControllerMeasurement, annotationWindowMeasurement, commentTextMeasurement, commentPanelMeasurement, mentionPanelMeasurement, annotationControlsAction, mountAnnotationControls } });
