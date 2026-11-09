import { measureChartClipping } from './workbook-chart-clipping.js';
import { measureChartNumberFields, measureNativeChartLabels, measureControlledChartLabels } from './workbook-chart-labels.js';
import type { WorkbookSnapshot } from '@doffi/core';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { disposeImageView, measureImageView, measurePictureFills } from './workbook-image-view.js';
import { measureObjectGeometry } from './workbook-object-geometry.js';
import { measureDrawingGradients } from './drawing-gradient.js';
import { measureDrawingFontProbeReuse, measureDrawingText, measureDrawingTextLayout, measureNativeDrawingText, measureDrawingTextFallback, measureDrawingTextOverflow, measureDrawingOverflowEdges } from './workbook-drawing-text.js';
import { measureWorkbookSurface, mountWorkbookSurfacePixels, disposeWorkbookSurfacePixels } from './workbook-surface.js';
import { measureWorkbookShadows, mountWorkbookShadowPixels, disposeWorkbookShadowPixels } from './workbook-shadow.js';
import { measureWorkbookCharts, measureWorkbookChartAxes, measureWorkbookChartFiniteScales, measureNativeFiniteCharts, measureWorkbookChartTextFonts, measureWorkbookChartMarkers, measureWorkbookChartGrid, measureWorkbookChartColors, measureWorkbookChartTitles, measureWorkbookChartPositions, measureWorkbookChartTitleContents, measureWorkbookComposedChart, measureWorkbookChartFormats } from './workbook-chart.js';

function sourceText(root: Node): string {
  if (root.nodeType === 3) return root.textContent ?? '';
  if (root.nodeType === 1 && (root as Element).localName === 'br') return '\n';
  if (root.nodeType === 1 && (root as Element).localName === 'img') return '\ufffc';
  return [...root.childNodes].map(sourceText).join('');
}
export function measureContentSource(snapshot: WorkbookSnapshot) {
  const scope = document.createElement('div'), style = document.createElement('style');
  style.textContent = snapshot.styles.map(style => style.css).join('\n'); document.head.append(style); document.body.append(scope);
  const source = new WorkbookContentSources(document, snapshot);
  try {
    const ids = [...(snapshot.metadata.workbook.stories ?? []).map(story => story.rootNodeId), ...(snapshot.metadata.workbook.objects ?? []).map(object => object.nodeId)];
    const all = source.project(scope, ids), before = scope.childElementCount;
    const rootTexts = Object.fromEntries(ids.map(id => [id, sourceText(all.get(id)!)]));
    const imageAttributes = all.assets.map(asset => ({ ...asset, src: all.get(asset.nodeId)!.getAttribute('src') }));
    const selectedIds = ids.filter((_, index) => index % 2 === 0), selected = source.project(scope, [...selectedIds].reverse());
    const isolated = ids.every(id => Boolean(selected.get(id)) === selectedIds.includes(id)) && selectedIds.every(id => selected.get(id) !== all.get(id));
    let limited = '', unknown = '', duplicated = '';
    for (const kind of ['limited', 'unknown', 'duplicated'] as const) {
      try { source.project(scope, kind === 'unknown' ? ['missing'] : kind === 'duplicated' ? [ids[0]!, ids[0]!] : ids, kind === 'limited' ? 1 : 50_000).dispose(); }
      catch (error) { const code = (error as { code: string }).code; if (kind === 'limited') limited = code; else if (kind === 'unknown') unknown = code; else duplicated = code; }
    }
    const atomic = scope.childElementCount === before + 1;
    all.dispose(); all.dispose(); const secondSurvives = selectedIds.every(id => selected.get(id)!.isConnected);
    source.dispose(); const released = scope.childElementCount === 0;
    let disposed = ''; try { selected.get(ids[0]!); } catch (error) { disposed = (error as { code: string }).code; }
    return { rootTexts, imageAttributes, isolated, atomic, secondSurvives, released, limited, unknown, duplicated, disposed };
  } finally { source.dispose(); scope.remove(); style.remove(); }
}
export function rejectContentSource(snapshot: WorkbookSnapshot, maxNodes?: number, maxCodeUnits?: number) {
  try { const source = new WorkbookContentSources(document, snapshot, maxNodes, maxCodeUnits); source.dispose(); return ''; }
  catch (error) { return (error as { code: string }).code; }
}
Object.assign(window, { workbookContentHost: { measureContentSource, rejectContentSource, measureImageView, measurePictureFills, disposeImageView, measureObjectGeometry, measureDrawingGradients, measureDrawingFontProbeReuse, measureDrawingText, measureDrawingTextLayout, measureNativeDrawingText, measureDrawingTextFallback, measureDrawingTextOverflow, measureDrawingOverflowEdges, measureWorkbookSurface, mountWorkbookSurfacePixels, disposeWorkbookSurfacePixels } });
Object.assign(window, { workbookShadowHost: { measureWorkbookShadows, mountWorkbookShadowPixels, disposeWorkbookShadowPixels } });
Object.assign(window, { workbookChartHost: { measureChartNumberFields, measureChartClipping, measureControlledChartLabels, measureNativeChartLabels, measureWorkbookChartAxes, measureWorkbookChartFiniteScales, measureNativeFiniteCharts, measureWorkbookChartTextFonts, measureWorkbookCharts, measureWorkbookChartMarkers, measureWorkbookChartGrid, measureWorkbookChartColors, measureWorkbookChartTitles, measureWorkbookChartPositions, measureWorkbookChartTitleContents, measureWorkbookComposedChart, measureWorkbookChartFormats } });
