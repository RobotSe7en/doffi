import type { WordMetadata } from '../contracts/word.js';
import type { WordRevision } from '../contracts/word-revision.js';
import type { ParsedHtml } from '../document/html.js';
import { DocumentError } from '../errors.js';
/** One native property history is retained until review; overlapping histories
 * require the separate merge policy currently awaiting a product decision. */
export function recordWordFormat(word: WordMetadata, source: ParsedHtml, nodeId: string, property: 'paragraph' | 'table' | 'row' | 'cell'): void {
    if (!word.trackChanges)
        return;
    if (word.revisions.some(revision => revision.nodeId === nodeId && revision.kind === 'format' && revision.property === property))
        throw new DocumentError('InvalidOperation', 'Review the overlapping formatting revision before applying another format.');
    const node = source.peekNode(nodeId)!, attributes: Record<string, string | null> = {};
    for (const name of ['style', 'class', 'lang', 'data-word-style', 'data-word-line-rule', 'data-word-numbering', 'data-word-level', 'data-word-outline-level', 'data-word-text-direction', 'data-word-repeat-header'])
        attributes[name] = node.attributes[name] ?? null;
    const before: NonNullable<WordRevision['before']> = { attributes, tagName: node.tagName as NonNullable<WordRevision['before']>['tagName'], inheritNativeProperties: true };
    if (property === 'paragraph') {
        const tabs = word.tabs.find(tab => tab.paragraphNodeId === nodeId);
        before.tabs = tabs ? { defaultInterval: tabs.defaultInterval, stops: tabs.stops.map(stop => ({ ...stop })) } : { defaultInterval: word.tabs[0]?.defaultInterval ?? 36, stops: [] };
    }
    const id = nodeId + '-format-' + property;
    word.revisions.push({ nodeId, sourceId: id, revisionId: id, kind: 'format', property, before, author: word.trackingAuthor ?? '', date: new Date().toISOString() });
}
