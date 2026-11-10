import { createWorkbook, createWorkbookNumberFormatter, openDocument } from '@doffi/core';
import { WorkbookCellTextCache, layoutCachedWorkbookCellTexts } from '../../../packages/render/src/workbook/cell-text-cache.js';
import { viewNodeRole } from '../../../packages/render/src/view-nodes.js';

export async function textCacheMeasurement() {
  await document.fonts.ready;
  const scope = document.createElement('div'), table = document.createElement('table'), row = table.insertRow(); scope.append(table); document.body.append(scope);
  const source = row.insertCell(); source.style.cssText = 'font:11pt "Liberation Sans";line-height:normal';
  source.style.setProperty('--office-font-east-asia', '"Noto Sans CJK SC"');
  source.append(document.createTextNode('A😀'), document.createElement('br'));
  const rich = document.createElement('span'); rich.style.color = 'rgb(128,0,64)'; rich.textContent = '禰󠄀B'; source.append(rich);
  const snapshot = openDocument(createWorkbook({ documentId: 'text-cache-probe' }), { kind: 'workbook' }).snapshot();
  const formatter = createWorkbookNumberFormatter({ dateSystem: '1900', hostLocale: 'zh-CN' });
  const cache = new WorkbookCellTextCache(2, 50_000, 8 * 1024 * 1024), original = source.outerHTML;
  const input = { source: { element: source, format: { cssClass: '' } }, value: { type: 'text' as const, value: source.textContent!.replace('A😀', 'A😀\n') }, formula: false, widthPt: 90, heightPt: 60, bordered: false };
  const layout = (nodes = 50_000, units = 8 * 1024 * 1024) => layoutCachedWorkbookCellTexts(scope, snapshot, [input], ['rich'], formatter, { normalSpaceWidthPx: 4 }, nodes, units, cache);
  try {
    const first = layout(), initial = first.cells[0]!.text!, expected = initial.element.outerHTML;
    scope.append(initial.element);
    const second = layout(), copied = second.cells[0]!.text!; scope.append(copied.element);
    const offsets = copied.text.projections.map(part => [part.from, part.to, part.node.data]);
    const projectionsOwned = copied.text.projections.every((part, index) => copied.element.contains(part.node) && part.node !== initial.text.projections[index]!.node);
    const roles = (root: Node) => { const walk = document.createTreeWalker(root, 0xffff_ffff), result = [viewNodeRole(root)]; let node: Node | null; while ((node = walk.nextNode())) result.push(viewNodeRole(node)); return result; };
    const same = expected === copied.element.outerHTML && JSON.stringify(roles(initial.element)) === JSON.stringify(roles(copied.element));
    initial.text.projections[0]!.node.data = 'changed accepted DOM';
    copied.text.projections[0]!.node.data = 'changed returned copy';
    const retained = cache.get('rich')!.cell.text!, detached = !retained.element.parentNode && !retained.element.isConnected;
    const isolated = retained.element.outerHTML === expected;
    const failures: string[] = [];
    for (const limits of [[1, 1000], [50_000, 1]]) try { layout(limits[0], limits[1]); failures.push('accepted'); } catch (error) { failures.push((error as { code: string }).code); }
    const atomic = initial.element.isConnected && copied.element.isConnected;
    cache.put('two', { text: retained }, second.fonts, false); cache.get('rich'); cache.put('three', { text: retained }, second.fonts, false);
    const evicted = !cache.get('two') && !!cache.get('rich') && !!cache.get('three');
    const warned = new WorkbookCellTextCache(2, 50_000, 8 * 1024 * 1024);
    warned.put('warned', { text: retained }, second.fonts.map(font => ({ ...font, missingCharacters: ['x'] })), false);
    const fork = warned.fork(), warningsRefreshed = !fork.get('warned') && !!warned.get('warned'); fork.clear(); warned.clear();
    const tiny = new WorkbookCellTextCache(1, 50_000, 1); tiny.put('rich', { text: retained }, second.fonts, false); const bounded = !tiny.get('rich'); tiny.clear();
    const face = new FontFace('WorkbookCacheProbe', 'local("Liberation Sans")');
    document.fonts.add(face); cache.settled(document.fonts); const invalidated = !cache.get('rich'); document.fonts.delete(face);
    cache.put('rich', { text: retained }, second.fonts, false); cache.clear(); const cleared = !cache.get('rich');
    return { offsets, projectionsOwned, same, isolated, detached, failures, atomic, evicted, warningsRefreshed, bounded, invalidated, cleared, unchanged: source.outerHTML === original };
  } finally { cache.clear(); formatter.dispose(); scope.remove(); }
}
