import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createWorkbookNumberFormatter, DEFAULT_WORKBOOK_RULE_LIMITS, openDocument, readWorkbookRules } from '@doffi/core';
import type { WorkbookDocumentMetadata } from '@doffi/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { measureWorkbookColumnFont } from '../../../packages/render/src/workbook/column-font.js';
import { worksheetDimensions } from '../../../packages/render/src/workbook/dimensions.js';
import { workbookDrawingRect } from '../../../packages/render/src/workbook/drawing.js';
import { WorkbookWindowLoader } from '../../../packages/render/src/workbook/window-loader.js';
import { modelWindowSource } from '../../../packages/render/src/workbook/window-source.js';
import { layoutWorkbookWindow } from '../../../packages/render/src/workbook/window-layout.js';
import type { WorkbookViewport } from '../../../packages/render/src/workbook/viewport.js';
import { maximumWorkbookSkew } from '../../../packages/render/src/workbook/skew-sources.js';
import { WorkbookCellTextCache } from '../../../packages/render/src/workbook/cell-text-cache.js';
import { modelRuleSource } from '../../../packages/render/src/workbook/rule-source.js';

export async function windowLayoutMeasurement(entries: TestPackage, views: readonly { sheetId: string; viewport: WorkbookViewport; maxCellFragments?: number }[], rtl = false, cacheEnabled = false) {
  const files = decodeTestPackage(entries);
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json'))) as WorkbookDocumentMetadata;
  for (const sheet of metadata.workbook.sheets) sheet.rightToLeft = rtl;
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot();
  const originalSource = model.serialize();
  const formatter = createWorkbookNumberFormatter({ dateSystem: metadata.workbook.dateSystem, hostLocale: 'zh-CN' });
  const container = document.createElement('div'); container.style.cssText = 'position:relative;width:1000px;height:800px'; document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 750, heightPt: 600 });
  const loader = new WorkbookWindowLoader(modelWindowSource(model), {}, true, maximumWorkbookSkew(surface.document)), results = [];
  let cache = cacheEnabled ? new WorkbookCellTextCache(5000, 50_000, 8 * 1024 * 1024, surface.fontProbes) : undefined;
  const cacheSafety = [];
  try {
    for (const { sheetId, viewport, maxCellFragments } of views) {
      const normal = measureWorkbookColumnFont(surface.source, snapshot.metadata.workbook, sheetId);
      const dimensions = worksheetDimensions(snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!, normal.maximumDigitWidthPx);
      await loader.load(sheetId, dimensions, viewport);
      const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!;
      const rules = sheet.conditionalRules?.length || sheet.dataValidations?.length ? await readWorkbookRules(model, { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision,
        sheetId, coordinates: loader.frame!.sourceCoordinates, ...DEFAULT_WORKBOOK_RULE_LIMITS }) : undefined;
      const rendered = layoutWorkbookWindow(surface, snapshot, loader.frame!, dimensions, formatter, { normalSpaceWidthPx: normal.spaceWidthPx, ...(maxCellFragments === undefined ? {} : { maxCellFragments }) }, rules, cache);
      const before = surface.view.firstChild; const detached = !rendered.element.isConnected && surface.view.firstChild === before;
      surface.view.replaceChildren(rendered.element); surface.frame.style.setProperty('visibility', 'visible', 'important');
      const view = surface.document.defaultView!;
      const cells = rendered.cells.filter(cell => cell.text).map(cell => {
        const text = cell.text!, bounds = text.text.element.getBoundingClientRect();
        const clip = text.element.closest<HTMLElement>('[data-office-view="workbook-cell"],[data-office-view="workbook-text-overflow"]')!;
        const clipBounds = clip.getBoundingClientRect();
        const accountingLayer = text.element.querySelector<HTMLElement>('[data-office-view=workbook-accounting-underlines]');
        let accounting;
        if (accountingLayer) {
          const walker = surface.document.createTreeWalker(text.element, 0xffff_ffff); let actual = 1; while (walker.nextNode()) actual++;
          accounting = { nodes: text.nodes, actualNodes: actual, codeUnits: text.codeUnits, actualCodeUnits: text.element.textContent?.length ?? 0,
            width: accountingLayer.getBoundingClientRect().width * .75,
            lines: [...accountingLayer.querySelectorAll<SVGGElement>('[data-office-view=workbook-accounting-line]')].map(line => ({ style: line.dataset.officeAccountingStyle === 'doubleAccounting' ? 'double' : 'solid', color: view.getComputedStyle(line.querySelector('line')!).stroke, width: line.getBoundingClientRect().width * .75 })) };
        }
        return { accounting, row: cell.projection.row, column: cell.projection.column, pane: cell.projection.pane, text: text.text.element.textContent,
          rect: cell.projection.rect, clipRect: cell.projection.clip, offsets: text.text.projections.map(part => [part.from, part.to, part.node.data]), shape: text.shape, scale: text.scale,
          runs: text.text.projections.map(part => { const style = view.getComputedStyle(part.node.parentElement!); return { from: part.from, to: part.to, text: part.node.data, color: style.color, weight: style.fontWeight, style: style.fontStyle, family: style.fontFamily }; }),
          color: view.getComputedStyle(text.text.element).color, weight: view.getComputedStyle(text.text.element).fontWeight, family: view.getComputedStyle(text.text.element).fontFamily,
          breaks: text.text.element.querySelectorAll("br").length, spaces: [...text.text.element.querySelectorAll<HTMLElement>("[data-office-number-space]")].map(node => node.getBoundingClientRect().width), bounds: { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height },
          clipWidthPt: clipBounds.width * .75, clipHeightPt: clipBounds.height * .75,
          clip: clip.style.cssText };
      });
      const errors: string[] = [];
      for (const limits of [{ maxNodes: 5 }, { maxSourceNodes: 1 }, { maxTextCodeUnits: 1 }]) {
        try { layoutWorkbookWindow(surface, snapshot, loader.frame!, dimensions, formatter, { normalSpaceWidthPx: normal.spaceWidthPx, ...limits }); errors.push('accepted'); }
        catch (error) { errors.push((error as { code: string }).code); }
      }
      const same = rendered.element === surface.view.firstChild;
      const fillNodes = new Map([...rendered.element.querySelector('[data-office-view="workbook-fills"]')!.children].map(node => [node.getAttribute('data-office-fill-key'), node]));
      results.push({ sheetId, cells, detached, same, errors, fonts: rendered.fonts.length, rules, completeRows: loader.frame!.completeRows,
        overflowCount: rendered.element.querySelectorAll('[data-office-view="workbook-text-overflow"]').length,
        overflowFills: [...rendered.element.querySelectorAll<SVGSVGElement>('[data-office-overflow-fill]')].map(fill => ({ color: fill.querySelector('rect')?.getAttribute('fill'), widthPt: fill.getBoundingClientRect().width * .75 })),
        fillPaints: frameFills(),
        borderColors: [...rendered.element.querySelectorAll('line')].filter(line => !line.closest('[data-office-view="workbook-gridlines"]')).map(line => line.getAttribute('stroke')),
        gridLines: rendered.element.querySelectorAll('[data-office-view="workbook-gridlines"]').length,
        widthPt: rendered.element.getBoundingClientRect().width * .75, heightPt: rendered.element.getBoundingClientRect().height * .75,
        accessible: rendered.element.querySelectorAll('[role="gridcell"]').length,
        fills: [...rendered.element.querySelectorAll<SVGSVGElement>('[data-office-view="workbook-cell-fill"]')].map(fill => ({ viewBox: fill.getAttribute('viewBox'), transform: fill.querySelector('g')?.getAttribute('transform'), width: fill.clientWidth })),
        borderCount: rendered.element.querySelectorAll('line').length,
        retainedSources: surface.source.querySelectorAll('[data-office-view="workbook-source-measurement"]').length,
        retainedMeasurements: surface.view.querySelectorAll('[data-office-view$="measurement"]').length,
      });
      if (cache) {
        // Release the previous controller while its independently owned cache
        // fork still renders exact text/projection clones on the same surface.
        const previous = cache; cache = previous.fork(); previous.clear();
        const reused = layoutWorkbookWindow(surface, snapshot, loader.frame!, dimensions, formatter, { normalSpaceWidthPx: normal.spaceWidthPx }, rules, cache);
        const forkRetainsText = reused.cells.every((cell, index) => {
          const before = rendered.cells[index]?.text, next = cell.text;
          return !before && !next || before !== undefined && next !== undefined && before.element !== next.element
            && before.text.element.textContent === next.text.element.textContent
            && JSON.stringify(before.text.projections.map(part => [part.from, part.to, part.node.data])) === JSON.stringify(next.text.projections.map(part => [part.from, part.to, part.node.data]))
            && next.text.projections.every(part => next.element.contains(part.node));
        });
        const frame = loader.frame!, point = frame.sourceCoordinates.find(point => { const cell = frame.getCell(point.row, point.column); return cell?.value.type === 'text' && !cell.formula; })!, source = frame.getCell(point.row, point.column)!;
        const changed = { ...source, text: '新的😀禰󠄀', value: { type: 'text' as const, value: '新的😀禰󠄀' }, html: source.html.replace(/>[\s\S]*<\/td>$/, '>新的😀禰󠄀</td>') };
        const replacement = (cell: typeof source) => ({ ...frame, getCell: (row: number, column: number) => row === point.row && column === point.column ? cell : frame.getCell(row, column) });
        let rejected = '', rejectedNumeric = '';
        try { layoutWorkbookWindow(surface, snapshot, replacement({ ...source, html: source.html.replace('<td', '<td onclick="window.workbookInjected=true"') }), dimensions, formatter, { normalSpaceWidthPx: normal.spaceWidthPx }, rules, cache); }
        catch (error) { rejected = (error as { code: string }).code; }
        const numericPoint = [...frame.sourceCoordinates].reverse().find(point => frame.getCell(point.row, point.column)?.value.type === 'number')!, numeric = frame.getCell(numericPoint.row, numericPoint.column)!;
        try { layoutWorkbookWindow(surface, snapshot, { ...frame, getCell: (row, column) => row === numericPoint.row && column === numericPoint.column
          ? { ...numeric, html: numeric.html.replace('<td', '<td onclick="window.workbookInjected=true"') } : frame.getCell(row, column) }, dimensions, formatter, { normalSpaceWidthPx: normal.spaceWidthPx }, rules, cache); }
        catch (error) { rejectedNumeric = (error as { code: string }).code; }
        const next = layoutWorkbookWindow(surface, snapshot, replacement(changed), dimensions, formatter, { normalSpaceWidthPx: normal.spaceWidthPx }, rules, cache);
        cacheSafety.push({ rejected, rejectedNumeric, forkRetainsText, retained: surface.view.firstChild === rendered.element,
          changedText: next.cells.find(cell => cell.projection.row === point.row && cell.projection.column === point.column)!.text!.text.element.textContent,
          offsets: next.cells.find(cell => cell.projection.row === point.row && cell.projection.column === point.column)!.text!.text.projections.map(part => [part.from, part.to, part.node.data]),
          leaked: surface.source.querySelectorAll('[data-office-view="workbook-source-measurement"]').length });
      }
      function frameFills() {
        return loader.frame!.projections.map(cell => { const node = fillNodes.get(`${cell.pane}:${cell.row}:${cell.column}`); return { row: cell.row, column: cell.column, pane: cell.pane,
          color: node?.getAttribute('data-office-fill-color') ?? node?.querySelector('rect')?.getAttribute('fill') ?? 'transparent' }; });
      }
    }
    const after = model.serialize(), unchanged = [...originalSource].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index]));
    return { results, unchanged, cacheSafety };
  } finally { cache?.clear(); loader.dispose(); formatter.dispose(); surface.dispose(); container.remove(); }
}

export async function autoRowMeasurement(entries: TestPackage, ids: readonly string[]) {
  const { measureWorkbookAutoRows } = await import('../../../packages/render/src/workbook/auto-rows.js');
  const files = decodeTestPackage(entries);
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), before = model.serialize();
  const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem, hostLocale: 'zh-CN' });
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 600, heightPt: 360 });
  const provider = modelWindowSource(model), results = [], errors: string[] = [];
  const rules = { source: modelRuleSource(model), limits: DEFAULT_WORKBOOK_RULE_LIMITS };
  let reads = 0, largest = 0;
  const source = { snapshot, async read(request: Parameters<typeof provider.read>[0], signal: AbortSignal) {
    reads++; largest = Math.max(largest, request.windows.reduce((sum, window) => sum + window.rows * window.columns, 0));
    return provider.read(request, signal);
  } };
  try {
    for (const sheetId of ids) {
      const normal = measureWorkbookColumnFont(surface.source, snapshot.metadata.workbook, sheetId);
      const result = await measureWorkbookAutoRows(surface, source, sheetId, normal, formatter, new AbortController().signal, {}, rules);
      results.push({ sheetId, heights: Array.from({ length: 4 }, (_, row) => result.dimensions.rows.size(row)), after: result.dimensions.rows.position(2),
        columnWidths: [0, 1, 2].map(column => result.dimensions.columns.size(column)),
        drawing: workbookDrawingRect({ kind: 'one-cell', from: { row: 2, column: 2, offsetXPt: 2, offsetYPt: 3 }, widthPt: 30, heightPt: 20 }, result.dimensions),
        farHeight: result.dimensions.rows.size(299), cells: result.measuredCells, fonts: result.fonts.length, segments: result.dimensions.rows.segmentCount });
    }
    const sheetId = ids[0]!, normal = measureWorkbookColumnFont(surface.source, snapshot.metadata.workbook, sheetId);
    for (const limits of [{ maxReadCells: 1 }, { maxCells: 1 }, { maxCodeUnits: 1 }]) {
      try { await measureWorkbookAutoRows(surface, source, sheetId, normal, formatter, new AbortController().signal, limits, rules); errors.push('accepted'); }
      catch (error) { errors.push((error as { code: string }).code); }
    }
    const abort = new AbortController(); abort.abort();
    try { await measureWorkbookAutoRows(surface, source, sheetId, normal, formatter, abort.signal, {}, rules); errors.push('accepted'); }
    catch (error) { errors.push((error as { code: string }).code); }
    const after = model.serialize(), unchanged = [...before].every(([path, bytes]) => bytes.every((byte, i) => byte === after.get(path)![i]));
    return { results, reads, largest, errors, unchanged, retained: surface.source.querySelectorAll('[data-office-view$="measurement"]').length + surface.view.querySelectorAll('[data-office-view$="measurement"]').length };
  } finally { formatter.dispose(); surface.dispose(); container.remove(); }
}

export async function autoRowSourceSafety(entries: TestPackage, hidden = false) {
  const { measureWorkbookAutoRows } = await import('../../../packages/render/src/workbook/auto-rows.js');
  const files = decodeTestPackage(entries), model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), before = model.serialize();
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 600, heightPt: 360 });
  const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem }), provider = modelWindowSource(model), sheetId = hidden ? snapshot.metadata.workbook.sheets[1]!.sheetId : snapshot.metadata.workbook.activeSheetId;
  const normal = measureWorkbookColumnFont(surface.source, snapshot.metadata.workbook, sheetId), errors: string[] = [], retained: number[] = [];
  try {
    for (const mode of ['handler', 'script', 'wrong-text', 'wrong-column', 'cancel']) {
      const abort = new AbortController(), view = surface.document.defaultView!, computed = view.getComputedStyle;
      if (mode === 'cancel') view.getComputedStyle = function(element, pseudo) { if (element.closest('office-workbook-source')) abort.abort(); return computed.call(this, element, pseudo); };
      const source = { snapshot, async read(request: Parameters<typeof provider.read>[0], signal: AbortSignal) {
        const response = await provider.read(request, signal);
        return { ...response, cells: response.cells.map(cell => {
          if (cell.row !== (hidden ? 1 : 2) || cell.column !== (hidden ? 1 : 0) || mode === 'cancel') return cell;
          const html = mode === 'handler' ? cell.html.replace('<td', '<td onclick="window.parent.document.body.dataset.rowSourceExecuted=true"')
            : mode === 'script' ? cell.html.replace(hidden ? 'Alpha' : 'After', '<script>window.parent.document.body.dataset.rowSourceExecuted=true</script>Changed')
            : mode === 'wrong-text' ? cell.html.replace(hidden ? 'Alpha' : 'After', 'Different') : cell.html.replace('data-column="' + (hidden ? 1 : 0) + '"', 'data-column="2"');
          return { ...cell, html };
        }) };
      } };
      try { await measureWorkbookAutoRows(surface, source, sheetId, normal, formatter, abort.signal); errors.push('accepted'); }
      catch (error) { errors.push((error as { code: string }).code); }
      finally { view.getComputedStyle = computed; }
      retained.push(surface.source.querySelectorAll('[data-office-view="workbook-source-measurement"]').length + surface.view.querySelectorAll('[data-office-view$="measurement"]').length);
    }
    const after = model.serialize();
    return { errors, retained, executed: Boolean(document.body.dataset.rowSourceExecuted), unchanged: [...before].every(([path, bytes]) => bytes.every((byte, i) => byte === after.get(path)![i])) };
  } finally { formatter.dispose(); surface.dispose(); container.remove(); }
}
