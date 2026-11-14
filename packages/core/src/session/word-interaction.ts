import { DocumentError } from '../errors.js';
import { wordSemanticChanges } from './word-impact.js';
import type { Annotation } from '../contracts/annotations.js';
import type { WordDocumentMetadata } from '../contracts/document.js';
import type { DocumentState } from '../document/state.js';
import type { DeepReadonly } from '../runtime.js';
import { resolveReference } from '../document/package.js';
import { checkLimit } from '../document/validation.js';
/** Owned stories and complete fields are part of the object being edited. */
export function expandWordContextRoots(state: DocumentState, roots: Set<string>, maximumNodes: number): Set<string> {
    const word = state.snapshot.metadata.word, nodes = new Set<string>(), queue: string[] = [];
    const edges = new Map<string, string[]>(), link = (from: string, to: string): void => { const list = edges.get(from) ?? []; list.push(to); edges.set(from, list); };
    if (word) {
        const stories = new Map(word.stories.map(story => [story.storyId, story.rootNodeId]));
        for (const drawing of word.drawings ?? []) {
            if ('textStoryId' in drawing && drawing.textStoryId)
                link(drawing.nodeId, stories.get(drawing.textStoryId)!);
            if (drawing.kind === 'connector')
                for (const target of [drawing.connector.startNodeId, drawing.connector.endNodeId])
                    if (target) {
                        link(drawing.nodeId, target);
                        link(target, drawing.nodeId);
                    }
        }
        const notes = new Map(word.notes?.map(note => [note.noteId, note]));
        for (const reference of word.noteReferences ?? []) {
            const note = notes.get(reference.noteId);
            if (note)
                link(reference.nodeId, stories.get(note.storyId)!);
        }
        for (const field of word.fields) {
            const ids = [field.nodeId, ...field.endNodeId ? [field.endNodeId] : [], ...field.resultNodeIds ?? []];
            for (const id of ids.slice(1)) {
                link(ids[0]!, id);
                link(id, ids[0]!);
            }
        }
        const revisionGroups=new Map<string,string[]>();for(const revision of word.revisions){const id=revision.groupId??revision.moveId;if(!id)continue;const nodes=revisionGroups.get(id)??[];nodes.push(revision.nodeId);revisionGroups.set(id,nodes);}for(const nodes of revisionGroups.values())for(const node of nodes.slice(1)){link(nodes[0]!,node);link(node,nodes[0]!);}
        for (const hyperlink of word.hyperlinks)
            if (hyperlink.endNodeId) {
                link(hyperlink.nodeId, hyperlink.endNodeId);
                link(hyperlink.endNodeId, hyperlink.nodeId);
            }
    }
    const include = (root: string): void => { const stack = [root]; while (stack.length) {
        const id = stack.pop()!;
        if (nodes.has(id))
            continue;
        const node = state.html.peekNode(id);
        if (!node)
            continue;
        nodes.add(id);
        checkLimit(nodes.size, maximumNodes);
        queue.push(id);
        stack.push(...node.childNodeIds);
    } };
    for (const root of roots)
        include(root);
    const inline = new Set(['span', 'strong', 'em', 'u', 's', 'sub', 'sup', 'ins', 'del', 'img', 'br']);
    for (let index = 0; index < queue.length; index++)
        for (const target of edges.get(queue[index]!) ?? []) {
            if (nodes.has(target))
                continue;
            let root = state.html.peekNode(target);
            if (!root)
                continue;
            while (inline.has(root.tagName) && root.parentNodeId)
                root = state.html.peekNode(root.parentNodeId)!;
            roots.add(root.nodeId);
            include(root.nodeId);
        }
    return nodes;
}
/** Native residuals and embedded workbooks can contain unrelated content. The
 * selected context exposes current semantic data and authorized visible images. */
export function wordInteractionContext(state: DocumentState, nodes: ReadonlySet<string>, annotations: readonly DeepReadonly<Annotation>[]) {
    if (!state.snapshot.metadata.word)
        return { metadata: state.snapshot.metadata, assets: state.snapshot.manifest.assets };
    const metadata = JSON.parse(JSON.stringify(state.snapshot.metadata)) as WordDocumentMetadata, word = metadata.word!;
    delete word.nativeContent;
    word.drawings = word.drawings?.filter(item => nodes.has(item.nodeId));
    for (const drawing of word.drawings ?? [])
        if (drawing.kind === 'chart')
            delete drawing.chart.native;
    word.equations = word.equations?.filter(item => nodes.has(item.nodeId));
    word.controls = word.controls?.filter(item => nodes.has(item.nodeId));
    word.fields = word.fields.filter(item => nodes.has(item.nodeId));
    word.bookmarks = word.bookmarks.filter(item => nodes.has(item.startNodeId) && nodes.has(item.endNodeId));
    word.hyperlinks = word.hyperlinks.filter(item => nodes.has(item.nodeId));
    word.tabs = word.tabs.filter(item => nodes.has(item.paragraphNodeId));
    word.tables = word.tables.filter(item => nodes.has(item.nodeId));
    word.revisions = word.revisions.filter(item => nodes.has(item.nodeId));
    word.noteReferences = word.noteReferences?.filter(item => nodes.has(item.nodeId));
    const notes = new Set(word.noteReferences?.map(item => item.noteId));
    word.notes = word.notes?.filter(item => notes.has(item.noteId));
    const selectedAnnotations = new Set(annotations.map(item => item.annotationId)), bodies = new Set(annotations.flatMap(annotation => [annotation, ...annotation.replies ?? []]).flatMap(message => message.bodyNodeId ? [message.bodyNodeId] : []));
    word.comments = word.comments?.filter(item => selectedAnnotations.has(item.annotationId));
    const authors = new Set([...annotations.flatMap(annotation => [annotation, ...annotation.replies ?? []]).flatMap(message => message.author ? [message.author.name] : []), ...word.revisions.flatMap(revision => revision.author ? [revision.author] : [])]);
    word.people = word.people?.filter(person => authors.has(person.author));
    word.stories = word.stories.filter(story => story.kind === 'main' || nodes.has(story.rootNodeId) || bodies.has(story.rootNodeId));
    const storyIds = new Set(word.stories.map(story => story.storyId));
    for (const section of metadata.sections)
        for (const key of ['headers', 'footers'] as const)
            if (section[key]) {
                section[key] = Object.fromEntries(Object.entries(section[key]!).filter(([, id]) => id !== undefined && storyIds.has(id)));
                if (!Object.keys(section[key]!).length)
                    delete section[key];
            }
    const used = new Set<string>();
    for (const id of nodes) {
        const node = state.html.peekNode(id)!;
        if (node.tagName === 'img')
            used.add(resolveReference(node.attributes.src!, state.snapshot.manifest.entrypoint));
    }
    return { metadata, assets: state.snapshot.manifest.assets.filter(asset => used.has(asset.path)) };
}
/** Recheck the prepared result as owned-story copying can affect more than the
 * operation's named source node. Candidate preparation remains isolated. */
export function validateWordCandidate(state: DocumentState, request: import('../contracts/interaction.js').DocumentInteractionRequest, model: import('../runtime.js').ContentDocumentModel): void {
    const scope = new Set<string>();
    for (const block of request.context.blocks) {
        const stack = [block.nodeId];
        while (stack.length) {
            const id = stack.pop()!;
            if (scope.has(id))
                continue;
            scope.add(id);
            stack.push(...state.html.peekNode(id)!.childNodeIds);
        }
    }
    const allowedAssets = new Set(request.context.assets.map(asset => asset.path));
    const after = model.snapshot(), afterNodes = new Set<string>();
    const collect = (root: string): void => { const stack = [root]; while (stack.length) {
        const id = stack.pop()!;
        if (afterNodes.has(id))
            continue;
        const node = model.getNode(id);
        if (!node)
            continue;
        afterNodes.add(id);
        stack.push(...node.childNodeIds);
    } };
    // Source traversal follows the same bounded HTML tree already validated by prepare.
    const root = state.html.nodeIds().find(id => state.html.peekNode(id)!.parentNodeId === null)!;
    collect(root);
    for (const id of afterNodes) {
        const node = model.getNode(id)!, old = state.html.peekNode(id);
        if (node.tagName === 'img' && (!old || old.attributes.src !== node.attributes.src) && !allowedAssets.has(resolveReference(node.attributes.src!, after.manifest.entrypoint)))
            throw new DocumentError('InvalidProposal', 'A Word proposal introduced a resource outside its request context.');
    }
    if (!state.snapshot.metadata.word)
        return;
    const changed = wordSemanticChanges(state.snapshot.metadata, after.metadata, root!);
    for (const id of changed)
        if (state.html.hasNode(id) && !scope.has(id))
            throw new DocumentError('InvalidProposal', 'A Word proposal changed semantic content outside its request context.');
}
