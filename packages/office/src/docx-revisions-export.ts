import { NATIVE_MARKER } from './docx-native-export.js';
import { W } from './ooxml/xml.js';
import { DocumentError } from '@doffi/core';
import type { DeepReadonly, WordMetadata, WordRevision } from '@doffi/core';
import type { SourceNode, SourceElement } from './export/source.js';
import { element } from './export/source.js';
import type { ExportCss } from './export/css.js';
import type { ExportStyles } from './export/styles.js';
import type { Css } from './ooxml/styles.js';
import { attributes, empty, value, orderedProperties } from './export/xml.js';
import { paragraphProperties, paragraphKeys, runProperties, runKeys, tableProperties, cellProperties, difference } from './export/properties.js';
import { nativeInteger } from './export/values.js';
export class ExportWordRevisions {
    private next = 0;
    private readonly rowMarks = new Map<string, DeepReadonly<WordRevision>[]>();
    private readonly formats = new Map<string, DeepReadonly<WordRevision>[]>();
    constructor(private readonly word: DeepReadonly<WordMetadata> | undefined, private readonly css: ExportCss, private readonly styles: ExportStyles) {
        this.next = (word?.revisions.length ?? 0) + (word?.moves?.length ?? 0);
        for (const revision of word?.revisions ?? []) {
            if (['rowInsert', 'rowDelete'].includes(revision.kind)) {
                const entries = this.rowMarks.get(revision.nodeId) ?? [];
                entries.push(revision);
                this.rowMarks.set(revision.nodeId, entries);
            }
            if (revision.kind === 'format') {
                const entries = this.formats.get(revision.nodeId) ?? [];
                entries.push(revision);
                this.formats.set(revision.nodeId, entries);
            }
        }
    }
    private revision(node: SourceNode, property: WordRevision['property']): {
        revision: DeepReadonly<WordRevision>;
        owner: SourceElement;
    } | undefined {
        for (let current: SourceNode | undefined = node; current; current = current.parent)
            if (element(current)) {
                const revision = this.formats.get(current.attrs['data-node-id'] ?? '')?.find(item => item.property === property);
                if (revision)
                    return { revision, owner: current };
                if (property !== 'run')
                    break;
            }
        return;
    }
    private paragraphStyle(node: SourceElement): string | undefined {
        for (let current: SourceElement | undefined = node; current; current = current.parent)
            if (/^p$|^h[1-6]$/.test(current.tag))
                return current.attrs['data-word-style'];
        return undefined;
    }
    private old(owner: SourceElement, revision: DeepReadonly<WordRevision>, fallback: Css): {
        node: SourceElement;
        css: Css;
    } {
        const attrs = { ...owner.attrs };
        for (const [key, value] of Object.entries(revision.before?.attributes ?? {})) {
            if (value == null)
                delete attrs[key];
            else
                attrs[key] = value;
        }
        const node = { ...owner, attrs, tag: revision.before?.tagName ?? owner.tag };
        return { node, css: revision.before?.properties ? { ...fallback, ...revision.before.properties } : this.css.get(node) };
    }
    private revisionDate(revision: DeepReadonly<WordRevision>): string | undefined { const date = revision.date; if (date !== undefined && (!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)$/.test(date) || !Number.isFinite(Date.parse(date))))
        throw new DocumentError('UnsupportedFeature', 'A revision date is not an XML dateTime.'); return date; }
    private change(revision: DeepReadonly<WordRevision>, name: string, previous: string, owner?: SourceElement): string {
        if (revision.date !== undefined && (!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)$/.test(revision.date) || !Number.isFinite(Date.parse(revision.date))))
            throw new DocumentError('UnsupportedFeature', 'A revision date is not an XML dateTime.');
        let beforeSource: string | undefined;
        if (revision.before?.inheritNativeProperties && owner) {
            for (let current: SourceElement | undefined = owner; current; current = current.parent) {
                const id = current.attrs['data-node-id'];
                if (this.word?.nativeContent?.fragments.some(fragment => fragment.nodeId === id && fragment.kind === 'property' && !fragment.revisionPrevious && !fragment.propertyPath.some(path => path.endsWith('Change')) && (fragment.ownerName===name || fragment.propertyPath[0] === W + ' ' + name || !fragment.propertyPath.length && fragment.propertyName === W + ' ' + name))) {
                    beforeSource = id;
                    break;
                }
            }
        }
        return '<w:' + name + 'Change' + (beforeSource ? ' xmlns:dn="' + NATIVE_MARKER + '"' + attributes({ 'dn:before-source': beforeSource, 'dn:before-property': name }) : '') + attributes({ 'w:id': this.next++, ...(name==='tblGrid'?{}:{'w:author': revision.author ?? '', 'w:date': revision.date}) }) + '><w:' + name + '>' + previous + '</w:' + name + '></w:' + name + 'Change>';
    }
    run(node: SourceNode, expected: Css): string {
        const entry = this.revision(node, 'run');
        if (!entry)
            return '';
        const old = this.old(entry.owner, entry.revision, expected), style = old.node.attrs['data-word-style'], context = this.styles.runContext(this.paragraphStyle(entry.owner), style), props = (style ? value('rStyle', this.styles.ids.get(style) ?? style) : '') + runProperties(difference(old.css, { ...expected, ...context }, [...runKeys, ...Object.keys(old.css).filter(key => key.startsWith('--'))]), { ...expected, ...context });
        return this.change(entry.revision, 'rPr', props, entry.owner);
    }
    paragraph(node: SourceElement, width: number): string {
        const entry = this.revision(node, 'paragraph');
        if (!entry)
            return '';
        const old = this.old(node, entry.revision, {}), style = old.node.attrs['data-word-style'];
        let props = (style ? value('pStyle', this.styles.ids.get(style) ?? style) : '') + paragraphProperties({ ...difference(old.css, this.styles.paragraphContext(style), [...paragraphKeys, ...Object.keys(old.css).filter(key => key.startsWith('border-'))]), 'font-size': old.css['font-size']! }, width);
        const num = old.node.attrs['data-word-numbering'];
        if (num)
            props += '<w:numPr>' + value('ilvl', old.node.attrs['data-word-level'] ?? 0) + value('numId', this.styles.numberingIds.get(num) ?? num) + '</w:numPr>';
        const tabs = entry.revision.before?.tabs;
        if (tabs?.stops.length)
            props += '<w:tabs>' + tabs.stops.map(tab => empty('w:tab', { 'w:pos': nativeInteger(tab.position), 'w:val': tab.alignment, 'w:leader': tab.leader })).join('') + '</w:tabs>';
        if (old.node.attrs['data-word-outline-level'] !== undefined)
            props += value('outlineLvl', old.node.attrs['data-word-outline-level']);
        else if (/^h[1-6]$/.test(old.node.tag))
            props += value('outlineLvl', Number(old.node.tag[1]) - 1);
        return this.change(entry.revision, 'pPr', orderedProperties('pPr', props), entry.owner);
    }
    table(node: SourceElement): string {
        const entry = this.revision(node, 'table');
        if (!entry)
            return '';
        const old = this.old(node, entry.revision, {}), style = old.node.attrs['data-word-style'];
        return this.change(entry.revision, 'tblPr', (style ? value('tblStyle', this.styles.ids.get(style) ?? style) : '') + tableProperties(old.css), entry.owner);
    }
    cell(node: SourceElement): string {
        const entry = this.revision(node, 'cell');
        if (!entry)
            return '';
        const old = this.old(node, entry.revision, {});
        return this.change(entry.revision, 'tcPr', cellProperties(old.css) + value('textDirection', old.node.attrs['data-word-text-direction'] ?? old.css['--word-cell-text-direction']), entry.owner);
    }
    rowMark(node: SourceElement): string { return (this.rowMarks.get(node.attrs['data-node-id'] ?? '') ?? []).sort((a, b) => a.kind === 'rowInsert' ? -1 : b.kind === 'rowInsert' ? 1 : 0).map(revision => empty('w:' + (revision.kind === 'rowInsert' ? 'ins' : 'del'), { 'w:id': this.next++, 'w:author': revision.author ?? '', 'w:date': this.revisionDate(revision) })).join(''); }
    row(node: SourceElement): string {
        const entry = this.revision(node, 'row');
        if (!entry)
            return '';
        const old = this.old(node, entry.revision, {});
        let props = old.css['break-inside'] === 'avoid' ? '<w:cantSplit/>' : '';
        if (old.css.height)
            props += empty('w:trHeight', { 'w:val': nativeInteger(Number.parseFloat(old.css.height)) });
        if (old.node.attrs['data-word-repeat-header'])
            props += '<w:tblHeader/>';
        return this.change(entry.revision, 'trPr', props, entry.owner);
    }
    grid(node:SourceElement):string{const entry=this.revision(node,'tableGrid');if(!entry)return'';return this.change(entry.revision,'tblGrid',(entry.revision.before?.grid??[]).map(width=>empty('w:gridCol',{'w:w':nativeInteger(width)})).join(''),entry.owner);}
    section(node:SourceElement,compile:(properties:DeepReadonly<NonNullable<WordRevision['before']>['section']>)=>string):string{const entry=this.revision(node,'section');if(!entry)return'';const previous=compile(entry.revision.before?.section).replace(/^<w:sectPr[^>]*>/,'').replace(/<\/w:sectPr>$/,'');return this.change(entry.revision,'sectPr',previous,entry.owner);}
}
