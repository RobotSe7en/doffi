import { expandSourceCssProperty } from '../document/source-css.js';
import type { WordMetadata } from '../contracts/word.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
import type { SourceStyleNode } from '../document/source-css.js';
import type { SourceStyleEdit } from './source-styles.js';
import { resolveWordTableStyle, resolveWordStyleRun, wordTableRegions,wordTableRowRegions,wordTableRowCellProperties } from '../word-style-context.js';
const canonical = (css: Readonly<Record<string, string>>) => JSON.stringify(Object.fromEntries(Object.entries(css).flatMap(([key, value]) => key.startsWith('--') || value.includes('var(') ? [[key, value]] : expandSourceCssProperty(key, value)).sort(([a], [b]) => a!.localeCompare(b!))));
const runKey = (key: string) => /^(?:font-|color$|background-color$|text-(?:decoration|transform)|vertical-align$|letter-spacing$|word-spacing$|--word-(?:font|run|kerning|language|bold|italic)|--office-(?:font|text))/.test(key);
const run = (css: Record<string, string>) => Object.fromEntries(Object.entries(css).filter(([key]) => runKey(key)));
/** Replace owned style classes, never a cell's direct formatting. */
export function refreshWordTableStyles(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, styles: SourceStyleEdit, ids: ReadonlySet<string>): void {
    const source = parse(), resolver = styles.resolver(), read = (name: string) => resolver.specifiedClass(name), entries: {
        name: string;
        properties: Record<string, string>;
        weight: number;
    }[] = [];
    const project = (id: string): SourceStyleNode => { const node = source.peekNode(id)!; return { tag: node.tagName, attrs: node.attributes, parent: node.parentNodeId ? project(node.parentNodeId) : undefined }; };
    for (const info of word.tables) {
        if (!ids.has(info.nodeId))
            continue;
        const table = source.peekNode(info.nodeId);
        if (!table)
            continue;
        const styleId = table.attributes['data-word-style'], definition = word.styles.find(style => style.styleId === styleId), setting = (key: 'tableLookMask' | 'rowBandSize' | 'columnBandSize') => { let style = definition; const seen = new Set<string>(); while (style && !seen.has(style.styleId)) {
            seen.add(style.styleId);
            if (style[key] !== undefined)
                return style[key];
            style = word.styles.find(item => item.styleId === style!.basedOn);
        } return undefined; };
        const previousContexts = info.styleContexts ?? [], ownClasses = new Set(previousContexts.map(item => item.cssClass)), contexts: NonNullable<typeof info.styleContexts> = [];
        const assign = (id: string, kind: typeof contexts[number]['kind'], properties: Record<string, string>) => { const node = source.peekNode(id)!, name = 'wt-' + id, classes = (node.attributes.class ?? '').split(/\s+/).filter(item => item && !ownClasses.has(item)); classes.push(name); const next = [...new Set(classes)].join(' '); if (next !== node.attributes.class)
            edit.attributes(id, { class: next }); contexts.push({ nodeId: id, kind, cssClass: name }); if (canonical(read(name)) !== canonical(properties))
            entries.push({ name, properties, weight: 5 }); };
        const base = resolveWordTableStyle(word, read, styleId, []);
        assign(info.nodeId, 'table', { 'margin-left': '-5.4pt', ...base.table });
        const rows = table.childNodeIds.flatMap(id => { const node = source.peekNode(id)!; return node.tagName === 'tr' ? [id] : node.childNodeIds.filter(id => source.peekNode(id)!.tagName === 'tr'); });
        let story = table;
        while (story.parentNodeId && !word.stories.some(item => item.rootNodeId === story.nodeId))
            story = source.peekNode(story.parentNodeId)!;
        const defaults = run(resolver.get(project(story.nodeId)));
        if (word.defaultRunCssClass)
            Object.assign(defaults, read(word.defaultRunCssClass));
        const occupied: number[] = Array(info.grid.length).fill(0), mask = info.lookMask ?? setting('tableLookMask') ?? 0;
        rows.forEach((rowId, rowIndex) => {
            const rowRegions=wordTableRowRegions(rowIndex,rows.length,info.grid.length,mask,info.rowBandSize??setting('rowBandSize')??0,info.columnBandSize??setting('columnBandSize')??0),rowStyle=resolveWordTableStyle(word,read,styleId,rowRegions),rowCell=wordTableRowCellProperties(rowStyle.cell);
            assign(rowId,'row',rowStyle.row);
            let column = 0;
            for (const cellId of source.peekNode(rowId)!.childNodeIds) {
                const cell = source.peekNode(cellId)!;
                if (!['td', 'th'].includes(cell.tagName))
                    continue;
                while (occupied[column]! > rowIndex)
                    column++;
                const span = Number(cell.attributes.colspan ?? 1), height = Number(cell.attributes.rowspan ?? 1) || rows.length - rowIndex;
                for (let c = column; c < column + span; c++)
                    occupied[c] = rowIndex + height;
                if(cell.attributes['data-word-grid-gap']){column+=span;continue;}
                const regions = wordTableRegions(rowIndex, column, rows.length, info.grid.length, span, mask, info.rowBandSize ?? setting('rowBandSize') ?? 0, info.columnBandSize ?? setting('columnBandSize') ?? 0), value = resolveWordTableStyle(word, read, styleId, regions), inner: Record<string, string> = {};
                if (base.table['--word-border-inside-h'] && rowIndex > 0)
                    inner['border-top'] = 'var(--word-border-inside-h)';
                if (base.table['--word-border-inside-v'] && column > 0)
                    inner['border-left'] = 'var(--word-border-inside-v)';
                assign(cellId, 'cell', { 'padding-left': '5.4pt', 'padding-right': '5.4pt', ...value.cell,...rowCell, ...inner });
                const visit = (id: string, pRun: Record<string, string>, paragraphRun=pRun) => {
                    const node = source.peekNode(id)!;
                    if (node.tagName === 'table')
                        return;
                    let current = pRun;
                    if (/^p$|^h[1-6]$/.test(node.tagName)) {
                        const style = word.styles.find(style => style.styleId === node.attributes['data-word-style'] && style.kind === 'paragraph')??(!node.attributes['data-word-style']?word.styles.find(style=>style.kind==='paragraph'&&style.default)??word.styles.find(style=>style.kind==='paragraph'&&style.styleId==='Normal'):undefined), p: Record<string, string> = { ...value.p }, chain: typeof word.styles = [];
                        const seen = new Set<string>();
                        for (let item = style; item && !seen.has(item.styleId); item = word.styles.find(next => next.styleId === item!.basedOn)) {
                            seen.add(item.styleId);
                            chain.unshift(item);
                        }
                        for (const item of chain) {
                            if (item.components?.p)
                                Object.assign(p, read(item.components.p));
                            else {
                                const css = read(item.cssClass);
                                for (const key of item.paragraphProperties ?? Object.keys(css).filter(key => !runKey(key)))
                                    if (css[key] !== undefined)
                                        p[key] = css[key]!;
                            }
                        }
                        current = resolveWordStyleRun(word, read, style?.styleId, { ...defaults, ...value.r });
                        assign(id, 'paragraph', { ...p, ...Object.fromEntries(Object.entries(current).filter(([key]) => key !== 'background-color' && !key.startsWith('text-decoration'))) });
                    }
                    if(/^p$|^h[1-6]$/.test(node.tagName))paragraphRun=current;
                    if(node.attributes['data-word-reset-text']==='true')current=paragraphRun;
                    const character = word.styles.find(style => style.styleId === node.attributes['data-word-style'] && style.kind === 'character');
                    if (node.attributes['data-word-run'] || character || node.attributes['data-word-reset-text']==='true') {
                        current = resolveWordStyleRun(word, read, character?.styleId, current);
                        assign(id, 'run', Object.fromEntries(Object.entries(current).filter(([key, value]) => value !== pRun[key] || key === 'background-color' || key.startsWith('text-decoration'))));
                    }
                    for (const child of node.childNodeIds)
                        visit(child, current,paragraphRun);
                };
                for (const id of cell.childNodeIds)
                    visit(id, { ...defaults, ...value.r });
                column += span;
            }
        });
        info.styleContexts = contexts;
        // Remove stale bindings from surviving nodes, e.g. a span whose style was cleared.
        const retained = new Set(contexts.map(item => item.nodeId));
        for (const old of previousContexts) {
            const node = source.peekNode(old.nodeId);
            if (!node || retained.has(old.nodeId))
                continue;
            edit.attributes(old.nodeId, { class: (node.attributes.class ?? '').split(/\s+/).filter(name => name !== old.cssClass).join(' ') });
        }
    }
    if (entries.length)
        styles.replaceClasses(entries);
}
/** Include previous ancestors when a generic removal already detached its target. */
export function wordTableStyleTargets(word: WordMetadata, source: ParsedHtml, before: ParsedHtml, operations: readonly {
    op: string;
    nodeId?: string;
}[]): Set<string> {
    const ids = new Set<string>();
    for (const operation of operations) {
        if (!operation.nodeId)
            continue;
        if (['setWordStyle', 'deleteWordStyle', 'reviewWordRevision', 'reviewWordRevisions', 'cloneWordStory', 'setWordWatermark'].includes(operation.op))
            for (const table of word.tables)
                ids.add(table.nodeId);
        for (const tree of [source, before])
            for (let node = tree.peekNode(operation.nodeId); node; node = node.parentNodeId ? tree.peekNode(node.parentNodeId) : undefined)
                if (node.tagName === 'table')
                    ids.add(node.nodeId);
    }
    for(const table of word.tables)if(!before.hasNode(table.nodeId)&&source.hasNode(table.nodeId))ids.add(table.nodeId);
    return ids;
}
