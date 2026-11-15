import {readWordRowGapWidths,readWordRowGapCount} from './docx-row-gaps.js';
import {WORD_ROW_GAP_WIDTH_ATTRIBUTES} from '@doffi/core';
import{readWordConditionalFormat}from'./docx-conditional-format.js';
import type { WordMetadata, WordRevision } from '@doffi/core';
import { attr, child, check, fail, on, number, unsupported } from './ooxml/xml.js';
import type { XmlNode } from './ooxml/xml.js';
import { declarations, mergeTabs, pt,tableLookMask } from './ooxml/styles.js';
import type { DocxStyles, Css } from './ooxml/styles.js';
const propertyNames = { run: 'rPr', paragraph: 'pPr', table: 'tblPr', row: 'trPr', cell: 'tcPr' } as const;
export function readWordFormatRevision(word: WordMetadata, properties: XmlNode | undefined, owner: {
    tag: string;
    attrs: Record<string, string>;
}, property: keyof typeof propertyNames, styles: DocxStyles, options: {
    inheritedRun?: Css;
    runClasses?: readonly string[];
    defaultTab?: number;
} = {}): void {
    const name = propertyNames[property], change = child(properties, name + 'Change');
    if (!change)
        return;
    check(change, [name]);
    const previous = child(change, name), sourceId = attr(change, 'id');
    if (!previous || sourceId === undefined || child(previous, name + 'Change'))
        fail(change, 'A formatting revision requires one original property set and an identity.');
    const attributes: Record<string, string | null> = { style: null }, before: NonNullable<WordRevision['before']> = { attributes };
    let css: Css = {};
    if(['paragraph','row','cell'].includes(property))before.conditionalFormat=readWordConditionalFormat(previous)??null;
    if (property === 'run') {
        const styleId = attr(child(previous, 'rStyle'), 'val'), inherited = styles.runStyle(styleId ?? styles.defaultCharacter, options.inheritedRun ?? {});
        css = styles.run(previous, inherited);
        if (styleId) {
            const standalone = styles.get(styleId).r;
            for (const name of ['font-weight', 'font-style', 'font-variant', 'text-transform'])
                if (standalone[name] !== undefined && inherited[name] !== standalone[name] && css[name] === undefined)
                    css[name] = inherited[name]!;
        }
        attributes['data-word-style'] = styleId ?? null;
        attributes.class = [...options.runClasses ?? [], ...styleId ? [styles.styleClass(styleId)] : []].join(' ') || null;
        attributes.lang = attr(child(previous, 'lang'), 'val') ?? null;
    }
    else if (property === 'paragraph') {
        const styleId = attr(child(previous, 'pStyle'), 'val') ?? styles.defaultParagraph, style = styles.get(styleId);
        css = styles.paragraph(previous);
        attributes['data-word-style'] = styleId ?? null;
        attributes.class = styleId ? styles.styleClass(styleId) : null;
        const outline = child(previous, 'outlineLvl'), level = outline ? number(attr(outline, 'val'), outline) : style.outline;
        attributes['data-word-outline-level']=outline?String(level):null;
        before.tagName = level !== undefined && level < 6 ? 'h' + (level + 1) as 'h1' : 'p';
        attributes['data-word-line-rule'] = css['line-height'] && /^\d+(?:\.\d+)?$/.test(css['line-height']) ? 'auto' : null;
        const num = child(previous, 'numPr'), id = attr(child(num, 'numId'), 'val') ?? style.numbering?.id;
        attributes['data-word-numbering'] = id && id !== '0' ? id : null;
        attributes['data-word-numbering-off'] = id === '0' ? 'true' : null;
        attributes['data-word-level'] = id && id !== '0' ? String(number(attr(child(num, 'ilvl'), 'val'), num ?? previous, style.numbering?.level ?? 0)) : null;
        before.tabs = { defaultInterval: options.defaultTab ?? 36, stops: mergeTabs(style.tabs ?? styles.defaults.tabs, child(previous, 'tabs')) };
    }
    else if (property === 'table') {
        const styleId=attr(child(previous,'tblStyle'),'val')??styles.defaultTable;
        attributes['data-word-style']=styleId??null;attributes.class=styleId?styles.styleClass(styleId):null;
        css = styles.table(previous);
        before.tableStyleOptions={};const mask=tableLookMask(child(previous,'tblLook'));if(mask!==undefined)before.tableStyleOptions.lookMask=mask;for(const[name,key]of [['tblStyleRowBandSize','rowBandSize'],['tblStyleColBandSize','columnBandSize']]as const){const item=child(previous,name);if(item)before.tableStyleOptions[key]=number(attr(item,'val'),item);}
    }
    else if (property === 'cell') {
        if (attr(child(previous, 'vMerge'), 'val') !== attr(child(properties, 'vMerge'), 'val'))
            unsupported(change, 'A cell merge revision is preserved with its native grid.');
        css = styles.cell(previous);
        const span=child(previous,'gridSpan');if(span&&number(attr(span,'val'),span)<1)fail(span,'A previous cell span must be positive.');attributes.colspan=span?attr(span,'val')!:null;
        attributes['data-word-text-direction'] = attr(child(previous, 'textDirection'), 'val') ?? null;
    }
    else {
        check(previous, ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden']);
        if (on(child(previous, 'hidden')))
            unsupported(change, 'This row format revision remains native.');
        css=styles.row(previous);
        before.rowOmissions={before:readWordRowGapCount(previous,'before'),after:readWordRowGapCount(previous,'after')};const widths=readWordRowGapWidths(previous);for(const key of WORD_ROW_GAP_WIDTH_ATTRIBUTES)attributes[key]=widths[key]??null;
        if (child(previous,'cantSplit'))css['break-inside']=on(child(previous,'cantSplit'))?'avoid':'auto';
        const height = child(previous, 'trHeight');
        if (height)
            {const rule=attr(height,'hRule')??'atLeast';if(!['auto','atLeast','exact'].includes(rule))unsupported(change,'This previous row height rule remains native.');css.height = pt(number(attr(height, 'val'), height,0) / 20);css['--word-row-height-rule']=rule;}
        attributes['data-word-repeat-header'] = child(previous,'tblHeader')?(on(child(previous,'tblHeader'))?'true':'false'):null;
    }
    attributes.style = declarations(css) || null;
    word.revisions.push({ nodeId: owner.attrs['data-node-id']!, revisionId: owner.attrs['data-node-id'] + '-format-' + property, sourceId, kind: 'format', property, before, ...(attr(change, 'author') !== undefined ? { author: attr(change, 'author')! } : {}), ...(attr(change, 'date') !== undefined ? { date: attr(change, 'date')! } : {}) });
}
