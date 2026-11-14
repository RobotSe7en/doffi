import type { WordMetadata } from '../contracts/word.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
import type { DocumentReadLimits } from '../runtime.js';
import type { SourceStyleEdit } from './source-styles.js';
import { wordTextDefaults } from '../document/word-format.js';

const variable = (name: string): string => '--word-reset-' + name.replace(/^--/, '');
export function wordResetTextProperties(defaults: Readonly<Record<string, string>>): Record<string, string> {
    return { ...Object.fromEntries(Object.entries(defaults).map(([name, value]) => [name, 'var(' + variable(name) + ',' + value + ')'])), '--office-font-decoration-reset': 'all' };
}

/** These source classes cache paragraph/table defaults, like owned table style
 * contexts. Recompute them at the transaction boundary, never from editor DOM. */
export function refreshWordResetStyles(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, styles: SourceStyleEdit, limits: DocumentReadLimits): void {
    const source = parse(), paragraphs = new Set<string>();
    for (const id of source.nodeIds()) if (source.peekNode(id)!.attributes['data-word-reset-text'] === 'true') {
        for (let node = source.peekNode(id); node; node = node.parentNodeId ? source.peekNode(node.parentNodeId) : undefined)
            if (/^p$|^h[1-6]$/.test(node.tagName)) { paragraphs.add(node.nodeId); break; }
    }
    if (!paragraphs.size) return;
    const resolver = styles.resolver(), entries: { name: string; properties: Record<string, string> }[] = [];
    for (const id of paragraphs) {
        const node = source.peekNode(id)!, name = 'wr-' + id, previous = node.attributes['data-word-reset-style'];
        const properties = Object.fromEntries(Object.entries(wordTextDefaults(word, source, id, resolver, limits.maxNodes * 8)).map(([key, value]) => [variable(key), value]));
        const classes = (node.attributes.class ?? '').split(/\s+/).filter(value => value && value !== previous && value !== name); classes.push(name);
        if (classes.join(' ') !== node.attributes.class || previous !== name) edit.attributes(id, { class: classes.join(' '), 'data-word-reset-style': name });
        if (JSON.stringify(resolver.specifiedClass(name)) !== JSON.stringify(properties)) entries.push({ name, properties });
    }
    if (entries.length) styles.replaceClasses(entries);
}

/** The native run has one effective character style. A clear marker starts
 * again from the paragraph even while its old ancestors remain for history. */
export function wordRunCharacterStyle(word: WordMetadata, source: ParsedHtml, id: string, before?: Readonly<Record<string, string | null>>): string | undefined {
    for (let node = source.peekNode(id); node; node = node.parentNodeId ? source.peekNode(node.parentNodeId) : undefined) {
        if (/^p$|^h[1-6]$/.test(node.tagName)) break;
        const attributes = node.nodeId === id && before ? { ...node.attributes, ...before } : node.attributes;
        const style = attributes['data-word-style'];
        if (style && word.styles.some(item => item.styleId === style && item.kind === 'character')) return style;
        if (attributes['data-word-reset-text'] === 'true') break;
    }
    return;
}

export function wordRunCharacterStyleAt(word: WordMetadata, source: ParsedHtml, nodeId: string, from: number): string | undefined {
    const offset = source.textOffset(nodeId)! + from; let parent = nodeId;
    for (;;) {
        const child = source.peekNode(parent)!.childNodeIds.find(id => { const start = source.textOffset(id)!; return start <= offset && start + source.getNode(id)!.text.length > offset; });
        if (!child) return wordRunCharacterStyle(word, source, parent);
        parent = child;
    }
}
