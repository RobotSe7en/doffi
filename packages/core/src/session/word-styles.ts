import type { WordMetadata } from '../contracts/word.js';
import type { WordOperation } from '../contracts/word-operations.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
import type { SourceStyleEdit } from './source-styles.js';
import type { DocumentReadLimits } from '../runtime.js';
import { DocumentError } from '../errors.js';
type Operation = Extract<WordOperation, {
    op: 'setWordStyle' | 'deleteWordStyle';
}>;
type Style = WordMetadata['styles'][number];
const fail = (message: string): never => { throw new DocumentError('InvalidOperation', message); };
const runKey = (key: string): boolean => /^(?:font-|color$|background-color$|text-(?:decoration|transform)|vertical-align$|letter-spacing$|word-spacing$|--word-(?:font|run|kerning)|--office-(?:font|text))/.test(key);
const special = (key: string): boolean => key === 'background-color' || key.startsWith('text-decoration');
/** Rebuild flattened preview classes from own properties after every definition edit. */
export function editWordStyle(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, styles: SourceStyleEdit, operation: Operation, limits: DocumentReadLimits): void {
    const source = parse();
    if (source.peekNode(operation.nodeId)!.parentNodeId !== null)
        fail('Shared Word definitions require the document root.');
    const resolver = styles.resolver(), old = new Map(word.styles.map(style => [style.styleId, style])), oldRunClasses = new Set(word.styles.flatMap(style => style.runCssClass ? [style.runCssClass] : []));
    const own = new Map<string, {
        p: Record<string, string>;
        r: Record<string, string>;
        other: Record<string, string>;
    }>();
    for (const style of word.styles) {
        const paragraph = resolver.specifiedClass(style.cssClass), css = { ...paragraph, ...(style.runCssClass ? resolver.specifiedClass(style.runCssClass) : {}) }, p = style.paragraphProperties ?? Object.keys(css).filter(key => !runKey(key)), r = style.runProperties ?? Object.keys(css).filter(runKey);
        own.set(style.styleId, { p: Object.fromEntries(p.flatMap(key => paragraph[key] !== undefined ? [[key, paragraph[key]]] : [])), r: Object.fromEntries(r.flatMap(key => css[key] !== undefined ? [[key, css[key]]] : [])), other: style.kind === 'table' ? Object.fromEntries(Object.entries(css).filter(([key]) => !p.includes(key) && !r.includes(key))) : {} });
    }
    if (operation.op === 'setWordStyle') {
        const previous = old.get(operation.style.styleId);
        if (previous && previous.kind !== operation.style.kind)
            fail('A style cannot change its native kind.');
        const definition = JSON.parse(JSON.stringify(operation.style)) as Style, values = own.get(definition.styleId) ?? { p: {}, r: {}, other: {} };
        for (const key of operation.removeParagraphProperties ?? [])
            delete values.p[key];
        for (const key of operation.removeRunProperties ?? [])
            delete values.r[key];
        Object.assign(values.p, operation.paragraphProperties);
        Object.assign(values.r, operation.runProperties);
        definition.paragraphProperties = Object.keys(values.p);
        definition.runProperties = Object.keys(values.r);
        own.set(definition.styleId, values);
        if (definition.default)
            for (const style of word.styles)
                if (style.kind === definition.kind)
                    style.default = false;
        const index = word.styles.findIndex(style => style.styleId === definition.styleId);
        if (index < 0)
            word.styles.push(definition);
        else
            word.styles[index] = definition;
    }
    else {
        const removed = old.get(operation.styleId) ?? fail('The deleted style is missing.'), replacement = old.get(operation.replacementStyleId) ?? fail('The replacement style is missing.');
        if (removed === replacement || removed.kind !== replacement.kind)
            fail('A replacement must be a different style of the same kind.');
        if (word.revisions.some(revision => revision.before?.attributes?.['data-word-style'] === removed.styleId))
            fail('Review previous-format references before deleting their style.');
        for (let current: Style | undefined = replacement; current; current = current.basedOn ? old.get(current.basedOn) : undefined)
            if (current === removed)
                fail('A replacement cannot inherit from the deleted style.');
        for (const style of word.styles) {
            if (style.basedOn === removed.styleId)
                style.basedOn = replacement.styleId;
            if (style.nextStyleId === removed.styleId)
                style.nextStyleId = replacement.styleId;
            if (style.linkedStyleId === removed.styleId)
                delete style.linkedStyleId;
        }
        if (removed.default)
            replacement.default = true;
        for (const id of source.nodeIds()) {
            const node = source.peekNode(id)!;
            if (node.attributes['data-word-style'] !== removed.styleId)
                continue;
            edit.attributes(id, { 'data-word-style': replacement.styleId, class: (node.attributes.class ?? '').split(/\s+/).filter(name => name && name !== removed.cssClass && name !== removed.runCssClass).concat(replacement.cssClass).join(' ') });
        }
        word.styles = word.styles.filter(style => style !== removed);
        if (word.nativeContent)
            word.nativeContent.fragments = word.nativeContent.fragments.filter(fragment => fragment.styleId !== removed.styleId);
    }
    const definitions = new Map(word.styles.map(style => [style.styleId, style])), resolved = new Map<string, {
        p: Record<string, string>;
        r: Record<string, string>;
        other: Record<string, string>;
    }>(), ordered: Style[] = [], active = new Set<string>();
    function resolve(style: Style): {
        p: Record<string, string>;
        r: Record<string, string>;
        other: Record<string, string>;
    } { const known = resolved.get(style.styleId); if (known)
        return known; if (active.has(style.styleId))
        fail('Style inheritance contains a cycle.'); if (active.size >= limits.maxDepth)
        throw new DocumentError('ResourceLimitExceeded', 'Style inheritance exceeds the depth budget.'); active.add(style.styleId); const parent = style.basedOn ? definitions.get(style.basedOn) ?? fail('The base style is missing.') : undefined; if (parent && parent.kind !== style.kind)
        fail('A style must inherit from the same native kind.'); const base = parent ? resolve(parent) : { p: {}, r: {}, other: {} }, values = own.get(style.styleId)!; const result = { p: { ...base.p, ...values.p }, r: { ...base.r, ...values.r }, other: { ...base.other, ...values.other } }; resolved.set(style.styleId, result); active.delete(style.styleId); ordered.push(style); return result; }
    for (const style of word.styles)
        resolve(style);
    const entries: {
        name: string;
        properties: Record<string, string>;
        weight: number;
    }[] = [];
    for (const style of ordered) {
        const values = resolved.get(style.styleId)!, weight = style.kind === 'character' ? 4 : style.kind === 'paragraph' ? 3 : 1;
        entries.push({ name: style.cssClass, weight, properties: style.kind === 'character' ? values.r : { ...values.p, ...values.other, ...Object.fromEntries(Object.entries(values.r).filter(([key]) => !special(key))) } });
        if (style.kind !== 'character') {
            style.runCssClass ??= style.cssClass + '-run';
            entries.push({ name: style.runCssClass, weight, properties: Object.fromEntries(Object.entries(values.r).filter(([key]) => special(key))) });
        }
    }
    styles.replaceClasses(entries);
    const current = parse(), oldClasses = new Set([...old.values()].map(style => style.cssClass));
    for (const id of current.nodeIds()) {
        const node = current.peekNode(id)!, definition = definitions.get(node.attributes['data-word-style'] ?? '');
        let classes = (node.attributes.class ?? '').split(/\s+/).filter(Boolean), changed = false;
        if (definition) {
            classes = classes.filter(name => !oldClasses.has(name));
            classes.push(definition.cssClass);
            changed = true;
        }
        if (/^p$|^h[1-6]$/.test(node.tagName) && definition) {
            let style: Style | undefined = definition, outline: number | undefined = node.attributes['data-word-outline-level'] !== undefined ? Number(node.attributes['data-word-outline-level']) : undefined;
            while (style && outline === undefined) {
                outline = style.outlineLevel;
                style = style.basedOn ? definitions.get(style.basedOn) : undefined;
            }
            edit.renameSource(id, outline !== undefined && outline < 6 ? 'h' + (outline + 1) : 'p');
        }
        if (node.attributes['data-word-run']) {
            let parent = node.parentNodeId ? current.peekNode(node.parentNodeId) : undefined;
            while (parent && !/^p$|^h[1-6]$/.test(parent.tagName))
                parent = parent.parentNodeId ? current.peekNode(parent.parentNodeId) : undefined;
            const paragraph = definitions.get(parent?.attributes['data-word-style'] ?? '');
            classes = classes.filter(name => !oldRunClasses.has(name));
            if (paragraph?.runCssClass)
                classes.push(paragraph.runCssClass);
            changed = true;
        }
        if (changed)
            edit.attributes(id, { class: [...new Set(classes)].join(' ') });
    }
}
