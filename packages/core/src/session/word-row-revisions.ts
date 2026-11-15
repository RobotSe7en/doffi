import {moveWordContinuationFormats} from './word-conditional-format.js';
import type { WordMetadata } from '../contracts/word.js';
import type { ParsedHtml } from '../document/html.js';
import type { DocumentEdit } from '../document/edit.js';
import type { DocumentReadLimits } from '../runtime.js';
import { DocumentError } from '../errors.js';
import { editSourceTable } from './table-edit.js';
export function wordTableRows(source: ParsedHtml, tableId: string): string[] { return source.peekNode(tableId)!.childNodeIds.flatMap(id => { const node = source.peekNode(id)!; return node.tagName === 'tr' ? [id] : node.childNodeIds.filter(child => source.peekNode(child)!.tagName === 'tr'); }); }
/** Review row marks with the normal grid editor, including vertical merge migration. */
export function removeWordRows(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, nodeIds: ReadonlySet<string>, limits: DocumentReadLimits): void {
    const tables = new Set<string>();
    for (const id of nodeIds) {
        let node = parse().peekNode(id);
        while (node && node.tagName !== 'table')
            node = node.parentNodeId ? parse().peekNode(node.parentNodeId) : undefined;
        if (node)
            tables.add(node.nodeId);
    }
    for (const table of tables) {
        const source = parse(), rows = wordTableRows(source, table), remove = rows.map((id, index) => nodeIds.has(id) ? index : -1).filter(index => index >= 0);
        if (!remove.length)
            continue;
        if (remove.length === rows.length) {
            edit.removeSource(table);
            continue;
        }
        const info = word.tables.find(info => info.nodeId === table);
        if (!info)
            throw new DocumentError('InvalidOperation', 'A row review requires its native table grid.');
        const groups: {
            at: number;
            count: number;
        }[] = [];
        for (const index of remove) {
            const last = groups.at(-1);
            if (last && last.at + last.count === index)
                last.count++;
            else
                groups.push({ at: index, count: 1 });
        }
        for (const group of groups.reverse()) {
            const current = wordTableRows(parse(), table);
            moveWordContinuationFormats(info,{kind:'deleteRows',...group});
            editSourceTable({ edit, parse, limits, wordOmissions:true }, table, info.grid, { width: info.grid.reduce((a, b) => a + b, 0), height: current.length * 20 }, { kind: 'deleteRows', ...group });
        }
    }
}
export function trackWordRowDeletion(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, tableId: string, at: number, count: number, limits: DocumentReadLimits): void {
    const rows = wordTableRows(parse(), tableId);
    if (at + count > rows.length)
        throw new DocumentError('InvalidOperation', 'The tracked row selection exceeds the table.');
    const selected = rows.slice(at, at + count), author = word.trackingAuthor ?? '', cancel = new Set<string>(), groupId = selected[0] + '-row-delete', date = new Date().toISOString();
    for (const id of selected) {
        if (word.revisions.some(revision => revision.nodeId === id && revision.kind === 'rowInsert' && revision.author === author)) {
            cancel.add(id);
            continue;
        }
        if (!word.revisions.some(revision => revision.nodeId === id && revision.kind === 'rowDelete'))
            word.revisions.push({ nodeId: id, sourceId: id + '-row-delete', revisionId: id + '-row-delete', groupId, kind: 'rowDelete', author, date });
    }
    if (cancel.size)
        removeWordRows(word, edit, parse, cancel, limits);
}
