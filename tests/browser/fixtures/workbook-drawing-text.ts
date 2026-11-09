import type { WorkbookSnapshot } from '@doffi/core';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookDrawingText } from '../../../packages/render/src/workbook/drawing-text.js';
import { resolvePreviewFontsInScope } from '../../../packages/render/src/fonts.js';
import { layoutWorkbookDrawingTexts } from '../../../packages/render/src/workbook/drawing-text-layout.js';
import { createWorkbookDrawingFragment, workbookDrawingOrientation } from '../../../packages/render/src/workbook/drawing-view.js';
import { applyDrawingTextOverflow } from '../../../packages/render/src/workbook/drawing-text-overflow.js';
import { FontProbeCache } from '../../../packages/render/src/font-probe-cache.js';

function surface(snapshot: WorkbookSnapshot) {
  const host = document.createElement('div'), scope = document.createElement('div'), style = document.createElement('style'), sourceStyles = document.createElement('style');
  const root = host.attachShadow({ mode: 'open' });
  style.textContent = ':host{all:initial;display:block}' + snapshot.styles.map(style => style.css).join('\n');
  // Mirror the production frame's admitted source styles for the decoration
  // resolver; the shadow owns paint, so keep this host copy inactive.
  sourceStyles.dataset.officeSourceStyles = ''; sourceStyles.media = 'not all'; sourceStyles.textContent = snapshot.styles.map(style => style.css).join('\n'); document.head.append(sourceStyles);
  root.append(style, scope); document.body.append(host);
  return { scope, dispose: () => { host.remove(); sourceStyles.remove(); } };
}

export async function measureDrawingFontProbeReuse(snapshot: WorkbookSnapshot) {
  await document.fonts.ready;
  const { scope, dispose } = surface(snapshot), sources = new WorkbookContentSources(document, snapshot);
  const object = snapshot.metadata.workbook.objects!.find(object => object.kind === 'text')!, projection = sources.project(scope, [object.nodeId]), source = projection.get(object.nodeId)!;
  let created = 0;
  const probes = new class extends FontProbeCache {
    override available(key: string, measure: () => boolean) { return super.available(key, () => { created++; return measure(); }); }
    override glyph(key: string, measure: () => number) { return super.glyph(key, () => { created++; return measure(); }); }
  }();
  const face = new FontFace('doffi-cache-epoch-marker', 'local("Liberation Sans")'), before = source.outerHTML;
  const measure = () => {
    const text = createWorkbookDrawingText(source, object), batch = layoutWorkbookDrawingTexts(scope, [{ source, object, text, widthPt: 120, heightPt: 90 }], {}, probes);
    const item = batch.objects[0]!;
    return { fonts: batch.fonts, nodes: item.nodes, width: item.widthPt, height: item.heightPt, bounds: item.paintBounds,
      text: text.projections.map(part => ({ text: part.node.data, sourceNodeId: part.sourceNodeId, from: part.from, to: part.to })) };
  };
  try {
    const first = measure(), firstCreated = created, second = measure(), secondCreated = created;
    // Changing the actual owner FontFaceSet invalidates the probe cache even
    // when this unused local face does not alter the document's selected font.
    document.fonts.add(face); const changed = measure(), changedCreated = created;
    probes.clear(); const cleared = measure();
    return { first, second, changed, cleared, firstCreated, secondCreated, changedCreated, clearedCreated: created,
      unchanged: before === source.outerHTML, cleaned: !scope.querySelector('[data-office-view="workbook-drawing-text-measurement"]') };
  } finally { document.fonts.delete(face); probes.clear(); sources.dispose(); dispose(); }
}

export async function measureDrawingTextOverflow(snapshot: WorkbookSnapshot) {
  await document.fonts.ready;
  const { scope, dispose } = surface(snapshot), sources = new WorkbookContentSources(document, snapshot);
  const objects = snapshot.metadata.workbook.objects!, projection = sources.project(scope, objects.map(object => object.nodeId));
  const gallery = document.createElement('div'); gallery.id = 'workbook-overflow-probe';
  gallery.style.cssText = 'display:grid;grid-template-columns:repeat(4,300px);width:1200px;background:white;color:black';
  const results = [];
  try {
    for (const object of objects) {
      const source = projection.get(object.nodeId)!, before = source.outerHTML, text = createWorkbookDrawingText(source, object);
      const result = layoutWorkbookDrawingTexts(scope, [{ source, object, text, widthPt: 120, heightPt: 40 }]).objects[0]!;
      const cell = document.createElement('div'); cell.style.cssText = 'position:relative;width:300px;height:180px;overflow:hidden';
      result.element.style.position = 'absolute'; result.element.style.left = '40px'; result.element.style.top = '60px';
      const border = document.createElement('div'); border.style.cssText = 'position:absolute;left:40px;top:60px;width:160px;height:53.333333px;outline:1px solid lime';
      scope.append(result.element);
      const walker = document.createTreeWalker(result.element, NodeFilter.SHOW_ALL); let nodes = 1; while (walker.nextNode()) nodes++;
      const markers = [...text.element.querySelectorAll<HTMLElement>('[data-office-view="workbook-drawing-ellipsis"]')];
      const box = text.element.getBoundingClientRect(), markerBounds = markers.map(marker => {
        const rect = marker.getBoundingClientRect(); return { left: rect.left - box.left, top: rect.top - box.top, bottom: rect.bottom - box.top };
      });
      results.push({ unchanged: source.outerHTML === before, sourceText: text.projections.map(part => part.node.data).join(''),
        offsets: text.projections.map(part => [part.from, part.to, part.node.data.length]), nodes, declared: result.nodes,
        markerCount: markers.length, markerSourceIds: markers.some(node => node.hasAttribute('data-source-node-id') || node.hasAttribute('data-node-id')),
        paintBounds: result.paintBounds, overflow: result.overflow, markerBounds });
      cell.append(border, result.element); gallery.append(cell);
    }
    document.querySelector('#workbook-overflow-probe')?.remove(); document.body.append(gallery);
    const source = projection.get(objects[0]!.nodeId)!, text = createWorkbookDrawingText(source, objects[0]!);
    source.style.setProperty('--office-text-vertical-overflow', 'bad');
    let invalid = 'accepted';
    try { layoutWorkbookDrawingTexts(scope, [{ source, object: objects[0]!, text, widthPt: 120, heightPt: 40 }]); } catch (e) { invalid = (e as { code: string }).code; }
    source.style.setProperty('--office-text-vertical-overflow', 'ellipsis');
    const failed = createWorkbookDrawingText(source, objects[0]!);
    let limited = 'accepted';
    try { layoutWorkbookDrawingTexts(scope, [{ source, object: objects[0]!, text: failed, widthPt: 120, heightPt: 40 }], { maxNodes: failed.nodes + 2 }); } catch (e) { limited = (e as { code: string }).code; }
    const inspection = document.createElement('div'); inspection.textContent = 'ABCD'; scope.append(inspection);
    let geometryLimited = 'accepted';
    try { applyDrawingTextOverflow(inspection, { left: 0, top: 0, right: 1, bottom: 1 }, 'clip', 'ellipsis', () => {}, 1); } catch (e) { geometryLimited = (e as { code: string }).code; }
    inspection.remove();
    return { results, invalid, limited, geometryLimited, cleaned: !text.element.parentNode && !failed.element.parentNode && !scope.querySelector('[data-office-view="workbook-drawing-text-measurement"]') };
  } catch (error) { gallery.remove(); throw error; }
  finally { sources.dispose(); dispose(); }
}

export async function measureDrawingOverflowEdges(snapshot: WorkbookSnapshot) {
  await document.fonts.ready;
  const { scope, dispose } = surface(snapshot), sources = new WorkbookContentSources(document, snapshot), object = snapshot.metadata.workbook.objects![0]!;
  const projection = sources.project(scope, [object.nodeId]), source = projection.get(object.nodeId)!, results = [];
  try {
    source.innerHTML = '<p data-node-id="overflow-p1"><span data-node-id="overflow-base">A禰</span><span data-node-id="overflow-vs">󠄀e</span><span data-node-id="overflow-accent">́😀</span><span data-node-id="overflow-tail">WXYZ אבג ABC</span></p><p data-node-id="overflow-p2">Second line</p><p data-node-id="overflow-p3">Hidden tail</p>';
    for (const rtl of [false, true]) for (const flipH of [false, true]) for (const flipV of [false, true]) for (const heightPt of [5, 24, 40]) {
      source.style.setProperty('--office-text-horizontal-overflow', 'clip'); source.style.setProperty('--office-text-vertical-overflow', 'ellipsis');
      source.style.setProperty('--office-text-anchor', 't'); source.dataset.officeFlipH = String(flipH); source.dataset.officeFlipV = String(flipV);
      for (const p of source.querySelectorAll<HTMLElement>('p')) { p.style.direction = rtl ? 'rtl' : 'ltr'; p.style.textAlign = rtl ? 'right' : 'left'; }
      const before = source.outerHTML, text = createWorkbookDrawingText(source, object), result = layoutWorkbookDrawingTexts(scope, [{ source, object, text, widthPt: 90, heightPt }]).objects[0]!;
      scope.append(result.element);
      const marker = text.element.querySelector<HTMLElement>('[data-office-view="workbook-drawing-ellipsis"]')!, bounds = text.element.getBoundingClientRect(), cut = marker.getBoundingClientRect();
      const edge = rtl !== flipH ? cut.right : cut.left;
      const first = text.projections.filter(part => part.sourceNodeId !== 'overflow-p2' && part.sourceNodeId !== 'overflow-p3');
      const textValue = first.map(part => part.node.data).join(''), segmenter = new Intl.Segmenter('und', { granularity: 'grapheme' });
      let splitGraphemes = 0; const splits: { text: string; left: number; right: number; edge: number }[] = [];
      for (const segment of segmenter.segment(textValue)) {
        const start = first.find(part => part.to > segment.index)!, end = first.find(part => part.to >= segment.index + segment.segment.length)!;
        if (getComputedStyle(start.node.parentElement!).visibility === 'hidden') continue;
        const range = document.createRange(); range.setStart(start.node, segment.index - start.from); range.setEnd(end.node, segment.index + segment.segment.length - end.from);
        const rect = range.getBoundingClientRect();
        if (heightPt > 5 && rect.top < cut.bottom - 1 && rect.bottom > cut.top + 1 && rect.left < edge - 1 && rect.right > edge + 1) { splitGraphemes++; splits.push({ text: segment.segment, left: rect.left - bounds.left, right: rect.right - bounds.left, edge: edge - bounds.left }); }
      }
      const walker = document.createTreeWalker(result.element, NodeFilter.SHOW_ALL); let nodes = 1; while (walker.nextNode()) nodes++;
      results.push({ rtl, flipH, flipV, heightPt, splitGraphemes, splits, markerStyle: marker.style.cssText, text: text.projections.map(part => part.node.data).join(''), unchanged: source.outerHTML === before,
        nodes, declared: result.nodes, marker: marker.textContent, inside: cut.left >= bounds.left - .1 && cut.right <= bounds.right + .1 });
      result.element.remove();
    }
    return results;
  } finally { sources.dispose(); dispose(); }
}

export function measureDrawingText(snapshot: WorkbookSnapshot) {
  const { scope, dispose } = surface(snapshot);
  const sources = new WorkbookContentSources(document, snapshot), objects = snapshot.metadata.workbook.objects!.filter(object => object.kind === 'geometry' || object.kind === 'text');
  const projection = sources.project(scope, objects.map(object => object.nodeId));
  try {
    const native = objects.map(object => {
      const source = projection.get(object.nodeId)!, before = source.outerHTML, result = createWorkbookDrawingText(source, object);
      scope.append(result.element); const fonts = resolvePreviewFontsInScope(document, result.element, '[data-office-text-style],[data-office-script]');
      const values = { nodeId: object.nodeId, text: result.element.textContent, unchanged: before === source.outerHTML, detachedSource: source.isConnected,
        nodes: result.nodes, actualNodes: result.element.querySelectorAll('*').length + result.projections.length + 1,
        inheritedIds: result.element.querySelectorAll('[data-node-id],[id],[class]').length,
        pieces: result.projections.map(piece => ({ text: piece.node.data, from: piece.from, to: piece.to, sourceNodeId: piece.sourceNodeId, family: getComputedStyle(piece.node.parentElement!).fontFamily, size: getComputedStyle(piece.node.parentElement!).fontSize, role: piece.node.parentElement!.dataset.officeScript })), fonts };
      result.element.remove(); return values;
    });
    const object = objects[0]!, source = projection.get(object.nodeId)!;
    // A source paragraph boundary has no UTF-16 unit. Structural whitespace
    // between paragraphs does occupy source offsets but is not visible text.
    source.innerHTML = '<p data-node-id="drawing-p1" style="font-family:serif;--office-font-ascii:&#34;Liberation Sans&#34;;--office-font-east-asia:&#34;Noto Sans CJK SC&#34;"><span data-node-id="drawing-r1">A中</span><span data-node-id="drawing-r2">󠄀</span><br data-node-id="drawing-br" style="font-size:18pt"><strong data-node-id="drawing-r3">é😀</strong></p>\n <p data-node-id="drawing-p2">Tail</p>';
    const before = source.outerHTML, text = createWorkbookDrawingText(source, object); scope.append(text.element);
    const pieces = text.projections.map(piece => ({ text: piece.node.data, from: piece.from, to: piece.to, sourceNodeId: piece.sourceNodeId, role: piece.node.parentElement!.dataset.officeScript, size: getComputedStyle(piece.node.parentElement!).fontSize, weight: getComputedStyle(piece.node.parentElement!).fontWeight }));
    text.element.remove(); const exact = createWorkbookDrawingText(source, object, text.nodes).nodes;
    const errors: string[] = [];
    for (const budget of [0, 1, text.nodes - 1]) try { createWorkbookDrawingText(source, object, budget); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    source.append(document.createElement('img'));
    try { createWorkbookDrawingText(source, object); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    source.lastElementChild!.remove(); const unchanged = source.outerHTML === before; projection.dispose();
    try { createWorkbookDrawingText(source, object); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    return { native, pieces, exact, nodes: text.nodes, unchanged, errors };
  } finally { sources.dispose(); dispose(); }
}

export async function measureDrawingTextLayout(snapshot: WorkbookSnapshot) {
  await document.fonts.ready;
  const { scope, dispose } = surface(snapshot);
  const sources = new WorkbookContentSources(document, snapshot), object = snapshot.metadata.workbook.objects![2]!;
  if (object.kind !== 'text' && object.kind !== 'geometry') throw Error('Expected a text-bearing fixture object');
  const projection = sources.project(scope, [object.nodeId]), source = projection.get(object.nodeId)!;
  try {
    const native = createWorkbookDrawingText(source, object), nativeResult = layoutWorkbookDrawingTexts(scope, [{ source, object, text: native, widthPt: 200, heightPt: 100 }]);
    scope.append(nativeResult.objects[0]!.element);
    const computed = getComputedStyle(native.element), paragraph = native.element.querySelector('p')!;
    const nativeStyles = { insets: [computed.paddingLeft, computed.paddingRight, computed.paddingTop, computed.paddingBottom].map(Number.parseFloat), margin: getComputedStyle(paragraph).marginTop, fonts: nativeResult.fonts.length, defaultSize: computed.fontSize };
    nativeResult.objects[0]!.element.remove();
    source.querySelector('p')!.textContent = '中文与 Latin 123，自动调整文字框。'.repeat(5);
    source.dataset.officeRotation = '30'; source.dataset.officeFlipH = 'true'; source.dataset.officeFlipV = 'true';
    const before = source.outerHTML, cases = [];
    for (const fit of ['none', 'normal', 'shape'] as const) {
      const definition = { ...object, text: { autoFit: fit, wrap: true } };
      const text = createWorkbookDrawingText(source, definition), result = layoutWorkbookDrawingTexts(scope, [{ source, object: definition, text, widthPt: 100, heightPt: 36 }]).objects[0]!;
      const fragment = createWorkbookDrawingFragment(document, object.nodeId, { pane: 'body', rect: { x: 0, y: 0, width: result.widthPt, height: result.heightPt }, clip: { x: 0, y: 0, width: 400, height: 500 } }, 400, workbookDrawingOrientation(source), false, result.element);
      scope.append(fragment);
      const body = fragment.querySelector<HTMLElement>('[data-office-view="workbook-drawing-body"]')!, matrix = new DOMMatrix(getComputedStyle(body).transform).multiply(new DOMMatrix(getComputedStyle(text.element).transform));
      cases.push({ fit, width: result.widthPt, height: result.heightPt, scale: result.scale, overflow: result.overflow, geometryHeight: result.geometry.textRect!.height,
        matrix: [matrix.a, matrix.b, matrix.c, matrix.d], text: text.projections.map(piece => piece.node.data).join(''), nodes: result.nodes,
        actualNodes: result.element.querySelectorAll('*').length + text.projections.length + 1 });
      fragment.remove();
    }
    const unchanged = source.outerHTML === before, errors: string[] = [];
    const text = createWorkbookDrawingText(source, object), input = { source, object, text, widthPt: 100, heightPt: 36 };
    for (const limits of [{ maxNodes: 1 }, { maxPasses: 0 }, { maxExtentPt: 10 }]) try { layoutWorkbookDrawingTexts(scope, [input], limits); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    const failing = { ...object, text: { autoFit: 'shape' as const } };
    const privateText = createWorkbookDrawingText(source, failing), firstText = createWorkbookDrawingText(source, object);
    try { layoutWorkbookDrawingTexts(scope, [{ ...input, text: firstText }, { ...input, object: failing, text: privateText }], { maxExtentPt: 110 }); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    const cleaned = !privateText.element.parentNode && !firstText.element.parentNode && !scope.querySelector('[data-office-view="workbook-drawing-text-measurement"]');
    source.querySelector('p')!.textContent = 'One\nTwo';
    // Keep this leading-ratio control on one selected font: mixed fonts may
    // produce different ascender/descender unions after CSS baseline alignment.
    source.querySelector('p')!.style.fontFamily = '"Liberation Sans"';
    source.querySelector('p')!.style.setProperty('--office-font-ascii', '"Liberation Sans"');
    const leading = [];
    for (const reduced of [false, true]) {
      const definition = { ...object, text: { autoFit: 'none' as const, ...(reduced ? { fontScale: .5, lineSpacingReduction: .25 } : {}) } }, text = createWorkbookDrawingText(source, definition);
      const result = layoutWorkbookDrawingTexts(scope, [{ source, object: definition, text, widthPt: 200, heightPt: 100 }]).objects[0]!; scope.append(result.element);
      const paragraph = text.element.querySelector('p')!;
      leading.push({ height: paragraph.getBoundingClientRect().height, font: parseFloat(getComputedStyle(paragraph.querySelector('[data-office-script]')!).fontSize) }); result.element.remove();
    }
    const p = source.querySelector('p')!; p.textContent = 'One'; p.dataset.officeNumbering = 'arabicPeriod'; p.dataset.officeNumberingStart = '3'; p.style.marginLeft = '20px'; p.style.textIndent = '-10px';
    const next = p.cloneNode(true) as HTMLElement; next.dataset.nodeId = 'drawing-list-next'; next.textContent = 'Two'; source.append(next);
    const list = createWorkbookDrawingText(source, object), lists = layoutWorkbookDrawingTexts(scope, [{ source, object, text: list, widthPt: 200, heightPt: 100 }]).objects[0]!;
    scope.append(lists.element);
    const walker = document.createTreeWalker(lists.element, NodeFilter.SHOW_ALL); let listNodes = 1; while (walker.nextNode()) listNodes++;
    const markers = [...list.element.querySelectorAll<HTMLElement>('.sdk-ppt-marker')].map(marker => ({ text: marker.textContent, left: getComputedStyle(marker).left, position: getComputedStyle(marker).position }));
    const listValues = { nodes: lists.nodes, actualNodes: listNodes, markers, sourceText: list.projections.map(piece => piece.node.data).join('') }; lists.element.remove();
    const limitedList = createWorkbookDrawingText(source, object);
    try { layoutWorkbookDrawingTexts(scope, [{ source, object, text: limitedList, widthPt: 200, heightPt: 100 }], { maxNodes: lists.nodes }); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
    p.style.marginLeft = '0'; p.style.textIndent = '0';
    const zeroText = createWorkbookDrawingText(source, object), zeroLayout = layoutWorkbookDrawingTexts(scope, [{ source, object, text: zeroText, widthPt: 200, heightPt: 100 }]).objects[0]!;
    scope.append(zeroLayout.element);
    const zeroParagraph = zeroText.element.querySelector('p')!, zeroMarker = zeroParagraph.querySelector<HTMLElement>('.sdk-ppt-marker')!;
    const zeroList = { margin: getComputedStyle(zeroParagraph).marginLeft, indent: getComputedStyle(zeroParagraph).textIndent,
      markerLeft: getComputedStyle(zeroMarker).left, markerWidth: getComputedStyle(zeroMarker).width,
      sourceText: zeroText.projections.map(piece => piece.node.data).join('') }; zeroLayout.element.remove();
    next.remove(); p.textContent = 'X'; p.removeAttribute('data-office-numbering'); p.removeAttribute('data-office-numbering-start'); p.style.marginLeft = '0'; p.style.textIndent = '0';
    const tinyObject = { ...object, text: { autoFit: 'normal' as const } }, tinyText = createWorkbookDrawingText(source, tinyObject);
    const tiny = layoutWorkbookDrawingTexts(scope, [{ source, object: tinyObject, text: tinyText, widthPt: 1, heightPt: 100 }]).objects[0]!;
    const insetOverflow = { width: tiny.widthPt, overflow: tiny.overflow }; tiny.element.remove();
    return { nativeStyles, cases, unchanged, errors, cleaned: cleaned && !limitedList.element.parentNode, leading, listValues, zeroList, insetOverflow };
  } finally { sources.dispose(); dispose(); }
}

export async function measureNativeDrawingText(snapshot: WorkbookSnapshot) {
  await document.fonts.ready;
  const { scope, dispose } = surface(snapshot);
  const sources = new WorkbookContentSources(document, snapshot), objects = snapshot.metadata.workbook.objects!, projection = sources.project(scope, objects.map(object => object.nodeId));
  try {
    return objects.map(object => {
      const source = projection.get(object.nodeId)!, before = source.outerHTML, text = createWorkbookDrawingText(source, object);
      const batch = layoutWorkbookDrawingTexts(scope, [{ source, object, text, widthPt: 120, heightPt: 90 }]), result = batch.objects[0]!; scope.append(result.element);
      const box = result.element.getBoundingClientRect(), words = text.projections.filter(piece => piece.node.data === 'Alpha' || piece.node.data === 'Beta').map(piece => {
        const range = document.createRange(); range.selectNode(piece.node); const rect = range.getBoundingClientRect();
        return { text: piece.node.data, boundsPt: [(rect.left - box.left) * .75, (rect.top - box.top) * .75, (rect.right - box.left) * .75, (rect.bottom - box.top) * .75] };
      });
      const paragraphs = [...text.element.querySelectorAll('p')].map(p => p.getBoundingClientRect().height * .75);
      const walker = document.createTreeWalker(result.element, NodeFilter.SHOW_ALL); let nodes = 1; while (walker.nextNode()) nodes++;
      const exact = createWorkbookDrawingText(source, object, text.nodes).nodes; let limited = 'accepted';
      try { createWorkbookDrawingText(source, object, text.nodes - 1); } catch (e) { limited = (e as { code: string }).code; }
      const value = { words, paragraphs, unchanged: source.outerHTML === before, sourceText: text.projections.map(piece => piece.node.data).join(''), placeholders: text.element.querySelectorAll('[data-office-view="workbook-drawing-empty-line"]').length,
        nodes, declared: result.nodes, exact, textNodes: text.nodes, limited, fallback: batch.usedFallbackLineMetrics, lineAdvancePt: words[1]!.boundsPt[1]! - words[0]!.boundsPt[1]! };
      result.element.remove(); return value;
    });
  } finally { sources.dispose(); dispose(); }
}

export async function measureDrawingTextFallback(snapshot: WorkbookSnapshot, invalidCanvasShorthand = false) {
  await document.fonts.ready;
  const { scope, dispose } = surface(snapshot), sources = new WorkbookContentSources(document, snapshot);
  const object = snapshot.metadata.workbook.objects!.find(object => object.kind === 'text')!;
  const projection = sources.project(scope, [object.nodeId]), source = projection.get(object.nodeId)!;
  const descriptors = ['fontBoundingBoxAscent', 'fontBoundingBoxDescent'].map(name => [name, Object.getOwnPropertyDescriptor(TextMetrics.prototype, name)] as const);
  try {
    source.innerHTML = '<p data-node-id="fallback-p1" style="font-family:Liberation Sans;font-size:14pt">Alpha<br data-node-id="fallback-br" style="font-size:14pt"></p><p data-node-id="fallback-p2" style="font-family:Liberation Sans;font-size:14pt">Beta</p><p data-node-id="fallback-empty" style="font-family:Liberation Sans;font-size:14pt"></p>';
    if (invalidCanvasShorthand) {
      source.style.fontVariant = 'all-small-caps';
      for (const paragraph of source.querySelectorAll('p')) paragraph.style.fontVariant = 'all-small-caps';
    }
    const before = source.outerHTML, text = createWorkbookDrawingText(source, object);
    if (!invalidCanvasShorthand) for (const [name] of descriptors) Object.defineProperty(TextMetrics.prototype, name, { configurable: true, get: () => undefined });
    const failed = createWorkbookDrawingText(source, object); let error = 'accepted';
    try { layoutWorkbookDrawingTexts(scope, [{ source, object, text: failed, widthPt: 120, heightPt: 90 }], { maxNodes: failed.nodes + 2 }); } catch (e) { error = (e as { code: string }).code; }
    const result = layoutWorkbookDrawingTexts(scope, [{ source, object, text, widthPt: 120, heightPt: 90 }], { maxNodes: text.nodes + 4 }), item = result.objects[0]!;
    scope.append(item.element);
    const walker = document.createTreeWalker(item.element, NodeFilter.SHOW_ALL); let nodes = 1; while (walker.nextNode()) nodes++;
    const paragraphs = [...text.element.querySelectorAll('p')].map(p => p.getBoundingClientRect().height);
    return { error, fallback: result.usedFallbackLineMetrics, nodes, declared: item.nodes, paragraphs, unchanged: before === source.outerHTML,
      sourceText: text.projections.map(piece => piece.node.data).join(''), placeholders: text.element.querySelectorAll('[data-office-view="workbook-drawing-empty-line"]').length,
      cleaned: !failed.element.parentNode && !scope.querySelector('[data-office-view="workbook-drawing-text-measurement"]') };
  } finally {
    if (!invalidCanvasShorthand) for (const [name, descriptor] of descriptors) { if (descriptor) Object.defineProperty(TextMetrics.prototype, name, descriptor); else Reflect.deleteProperty(TextMetrics.prototype, name); }
    sources.dispose(); dispose();
  }
}
