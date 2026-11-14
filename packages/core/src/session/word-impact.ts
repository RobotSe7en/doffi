import type { DeepReadonly } from '../runtime.js';
import type { ContentDocumentMetadata } from '../contracts/document.js';
import type { WordMetadata } from '../contracts/word.js';
/** Semantic changes can leave the atomic HTML source byte-for-byte unchanged. */
export function wordSemanticChanges(before: DeepReadonly<ContentDocumentMetadata>, after: DeepReadonly<ContentDocumentMetadata>, rootNodeId: string): Set<string> {
    const ids = new Set<string>();
    if (before.presentation || after.presentation)
        return ids;
    const left = before.word, right = after.word;
    if (!left && !right)
        return ids;
    const compare = <T>(a: readonly T[], b: readonly T[], key: (value: T) => string): void => { const group=(items:readonly T[]):Map<string,string>=>{const rows=new Map<string,T[]>();for(const item of items){const id=key(item),values=rows.get(id)??[];values.push(item);rows.set(id,values);}return new Map([...rows].map(([id,values])=>[id,JSON.stringify(values)]));};const previous=group(a),next=group(b); for (const id of new Set([...previous.keys(), ...next.keys()]))
        if (previous.get(id) !== next.get(id))
            ids.add(id); };
    for (const name of ['conditionalFormats','drawings', 'equations', 'controls', 'fields', 'hyperlinks', 'tables', 'revisions', 'noteReferences'] as const)
        compare<DeepReadonly<{
            nodeId: string;
        }>>(left?.[name] ?? [], right?.[name] ?? [], item => item.nodeId);
    compare(left?.tabs ?? [], right?.tabs ?? [], item => item.paragraphNodeId);
    compare(left?.bookmarks ?? [], right?.bookmarks ?? [], item => item.startNodeId);
    compare(before.sections, after.sections, item => item.rootNodeId);
    if (JSON.stringify(before.pageSetup) !== JSON.stringify(after.pageSetup))
        ids.add(rootNodeId);
    const globals = (word: DeepReadonly<WordMetadata> | undefined) => word ? Object.fromEntries(['bordersDoNotSurroundHeader','bordersDoNotSurroundFooter','styles', 'numbering', 'mirrorMargins','evenAndOddHeaders', 'automaticHyphenation','noColumnBalance', 'defaultRunCssClass','trackChanges','trackingAuthor', 'footnoteOptions', 'endnoteOptions'].map(key => [key, Reflect.get(word, key)])) : undefined;
    if (JSON.stringify(globals(left)) !== JSON.stringify(globals(right)))
        ids.add(rootNodeId);
    const notes = new Map([...(left?.notes ?? []), ...(right?.notes ?? [])].map(note => [note.noteId, note]));
    const leftNotes = new Map(left?.notes?.map(note => [note.noteId, JSON.stringify(note)])), rightNotes = new Map(right?.notes?.map(note => [note.noteId, JSON.stringify(note)]));
    for (const [id, note] of notes)
        if (leftNotes.get(id) !== rightNotes.get(id)) {
            const story = [...(left?.stories ?? []), ...(right?.stories ?? [])].find(story => story.storyId === note.storyId);
            if (story)
                ids.add(story.rootNodeId);
            for (const ref of [...left?.noteReferences ?? [], ...right?.noteReferences ?? []])
                if (ref.noteId === id)
                    ids.add(ref.nodeId);
        }
    return ids;
}
