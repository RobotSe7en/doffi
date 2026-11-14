import{expandWordRevisionGroups}from'../word-revision-groups.js';
import{reconcileTrackedWord}from'./word-tracked-structure.js';
import type { WordMetadata } from '../contracts/word.js';
import type { WordRevision, WordRevisionOperation } from '../contracts/word-revision.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
import { DocumentError } from '../errors.js';
import { isTextBoundary } from '../document/text.js';
const fail = (message: string): never => { throw new DocumentError('InvalidOperation', message); };
export const wordRevisionId = (revision: Pick<WordRevision, 'nodeId' | 'revisionId'>): string => revision.revisionId ?? revision.nodeId;
export function isWordRevisionOperation(operation: {
    readonly op: string;
}): operation is WordRevisionOperation { return ['reconcileWordTracked','setWordReview', 'reviewWordRevision', 'reviewWordRevisions', 'trackWordText'].includes(operation.op); }
export function trackWordText(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, operation: {
    nodeId: string;
    from: number;
    to: number;
    expectedText: string;
    text: string;
    properties?: Record<string, string> | undefined;
    author?: string | undefined;
    date?: string | undefined;
}): void {
    const source = parse(), node = source.getNode(operation.nodeId) ?? fail('The tracked text source is missing.');
    const { from, to, expectedText, text } = operation;
    if (from > to || node.text.slice(from, to) !== expectedText || !isTextBoundary(node.text, from) || !isTextBoundary(node.text, to) || /[\u0000\r\ud800-\udfff]/u.test(text))
        fail('The tracked text range or Unicode content is invalid.');
    if (expectedText === text)
        return;
    const author = operation.author ?? word.trackingAuthor ?? '', date = operation.date ?? new Date().toISOString(), start = source.textOffset(operation.nodeId)! + from, end = source.textOffset(operation.nodeId)! + to;
    for (const revision of word.revisions) {
        const owner = source.getNode(revision.nodeId);
        if (!owner)
            continue;
        const begin = source.textOffset(owner.nodeId)!, finish = begin + owner.text.length;
        const inside=(id:string,parent:string):boolean=>{for(let current:string|null=id;current;current=source.peekNode(current)!.parentNodeId)if(current===parent)return true;return false;};if(!inside(owner.nodeId,operation.nodeId)&&!inside(operation.nodeId,owner.nodeId))continue;
        if (revision.kind === 'insert' && revision.author === author && start >= begin && end <= finish) {
            edit.replaceSourceText(revision.nodeId, { ...operation, from: start - begin, to: end - begin });
            if (operation.properties && text)
                edit.formatSourceText(operation.nodeId, { from, to: from + text.length, expectedText: text }, operation.properties);
            return;
        }
        if (['delete', 'moveFrom'].includes(revision.kind) && start < finish && end > begin)
            fail('Deleted revision content must be reviewed before editing it.');
    }
    const deleted = edit.wrapSourceText(operation.nodeId, from, to, 'del');
    const escape = (value: string): string => value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
    const inserted = text ? edit.insertSourceAt(operation.nodeId, to, '<ins>' + escape(text).split('\n').join('<br>') + '</ins>').rootNodeIds : [], groupId = deleted[0] ?? inserted[0]!;
    if (operation.properties)
        for (const id of inserted)
            for (const [key, value] of Object.entries(operation.properties))
                edit.sourceStyle(id, key, value);
    for (const [kind, ids] of [['delete', deleted], ['insert', inserted]] as const)
        for (const id of ids)
            word.revisions.push({ nodeId: id, revisionId: id, sourceId: id, groupId, kind, author, date });
}
function nativeReview(word: WordMetadata, revision: WordRevision, decision: 'accept' | 'reject'): void {
    if (!word.nativeContent)
        return;
    const property = ({ run: 'rPr', paragraph: 'pPr', table: 'tblPr', row: 'trPr', cell: 'tcPr', section: 'sectPr', tableGrid: 'tblGrid' } as const)[revision.property ?? 'run'], change = property + 'Change';
    word.nativeContent.fragments = word.nativeContent.fragments.flatMap(fragment => {
        if (fragment.nodeId !== revision.nodeId)
            return [fragment];
        if (fragment.revisionSourceId === revision.sourceId) {
            if (decision === 'reject' && fragment.revisionPrevious) {
                const path = [...fragment.propertyPath],indices=fragment.propertyIndices?[...fragment.propertyIndices]:undefined, index = path.findIndex(name => name.endsWith(' ' + change));
                if (index >= 0) {
                    {path.splice(index, 1);indices?.splice(index,1);}
                    if (path[index] === path[index - 1])
                        {path.splice(index, 1);indices?.splice(index,1);}
                }
                const next = { ...fragment, propertyPath: path,...(indices?{propertyIndices:indices}:{}) };
                delete next.revisionSourceId;
                delete next.revisionPrevious;
                return [next];
            }
            return [];
        }
        if (decision === 'reject' && !revision.before?.inheritNativeProperties && revision.kind === 'format' && fragment.kind === 'property' && (fragment.propertyPath[0]?.endsWith(' ' + property) || !fragment.propertyPath.length && fragment.propertyName?.endsWith(' ' + property)))
            return [];
        return [fragment];
    });
}
export function applyWordRevision(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, operation: WordRevisionOperation): void {
    if (operation.op === 'setWordReview') {
        if (parse().peekNode(operation.nodeId)!.parentNodeId !== null)
            fail('Review settings require the document root.');
        word.trackChanges = operation.trackChanges;
        if (operation.author !== undefined)
            word.trackingAuthor = operation.author;
        return;
    }
    if(operation.op==='reconcileWordTracked'){reconcileTrackedWord(word,edit,parse,operation);return;}
    if (operation.op === 'trackWordText') {
        trackWordText(word, edit, parse, operation);
        return;
    }
    let selected: WordRevision[];
    if (operation.op === 'reviewWordRevisions') {
        if (parse().peekNode(operation.nodeId)!.parentNodeId !== null)
            fail('Reviewing all revisions requires the document root.');
        selected = word.revisions.filter(revision => operation.author === undefined || revision.author === operation.author);
    }
    else {
        const revision = word.revisions.find(revision => wordRevisionId(revision) === operation.revisionId && revision.nodeId === operation.nodeId) ?? fail('The selected revision is missing.');
        selected = word.revisions.filter(item => item === revision || revision.groupId && item.groupId === revision.groupId || revision.moveId && item.moveId === revision.moveId);
    }
    selected=[...expandWordRevisionGroups(word.revisions,selected.map(wordRevisionId))];
    const removals = new Set<string>();
    for (const revision of selected) {
        nativeReview(word, revision, operation.decision);
        const node = parse().peekNode(revision.nodeId);
        if (!node)
            continue;
        if (revision.kind === 'format') {
            if (operation.decision === 'reject') {
                const before = revision.before ?? fail('The revision has no previous formatting.');
                if (before.attributes)
                    edit.attributes(revision.nodeId, Object.fromEntries(Object.entries(before.attributes).map(([name, value]) => [name, value ?? undefined])));
                if (before.tagName)
                    edit.renameSource(revision.nodeId, before.tagName);
                if (before.tabs) {
                    const old = word.tabs.findIndex(tab => tab.paragraphNodeId === revision.nodeId), tab = { paragraphNodeId: revision.nodeId, ...before.tabs };
                    if (old < 0)
                        word.tabs.push(tab);
                    else
                        word.tabs[old] = tab;
                }
                for (const [key, value] of Object.entries(before.properties ?? {}))
                    edit.sourceStyle(revision.nodeId, key, value);
                if (before.grid) {
                    const table = word.tables.find(table => table.nodeId === revision.nodeId);
                    if (table)
                        table.grid = [...before.grid];
                }
            }
            continue;
        }
        const insertion = ['insert', 'moveTo', 'paragraphInsert','paragraphMoveTo', 'rowInsert', 'cellInsert'].includes(revision.kind), remove = operation.decision === 'accept' ? !insertion : insertion;
        if (['insert', 'delete', 'moveFrom', 'moveTo'].includes(revision.kind)) {
            if (remove) {
                edit.removeSource(revision.nodeId);
                removals.add(revision.nodeId);
            }
            else
                edit.unwrapSource(revision.nodeId);
        }
        else if (revision.kind.startsWith('paragraph')) {
            if (remove)
                edit.mergeSourceParagraphMark(revision.nodeId);
        }
        else if (remove) {
            edit.removeSource(revision.nodeId);
            removals.add(revision.nodeId);
        }
    }
    const moves=new Set(selected.flatMap(revision=>revision.moveId?[revision.moveId]:[]));for(const move of word.moves??[])if(moves.has(move.moveId))for(const id of [move.startNodeId,move.endNodeId]){const source=parse(),node=source.peekNode(id);if(!node)continue;const parent=node.parentNodeId?source.peekNode(node.parentNodeId):undefined,wrapper=parent&&parent.tagName==='div'&&parent.childNodeIds.length===1&&Object.keys(parent.attributes).every(name=>name==='data-node-id')&&source.getNode(parent.nodeId)!.text===''&&!word.stories.some(story=>story.rootNodeId===parent.nodeId);edit.removeSource(wrapper?parent!.nodeId:id);}if(word.moves)word.moves=word.moves.filter(move=>!moves.has(move.moveId));
    const reviewed = new Set(selected.map(wordRevisionId));
    word.revisions = word.revisions.filter(revision => !reviewed.has(wordRevisionId(revision)) && !removals.has(revision.nodeId));
}
