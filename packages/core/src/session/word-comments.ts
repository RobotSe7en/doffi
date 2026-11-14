import type { Annotation } from '../contracts/annotations.js';
import type { WordMetadata } from '../contracts/word.js';
import type { WordDocumentMetadata } from '../contracts/document.js';
import type { WordCommentOperation } from '../contracts/word-comment.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
import { DocumentError } from '../errors.js';
import { anchorNodes } from './anchors.js';
const invalid = (message: string): never => { throw new DocumentError('InvalidOperation', message); };
export function isWordCommentOperation(operation: {
    readonly op: string;
}): operation is WordCommentOperation { return ['addWordCommentReply', 'removeWordCommentReply', 'setWordCommentBody', 'setWordCommentIdentity'].includes(operation.op); }
export function applyWordComment(metadata: WordDocumentMetadata, annotations: Annotation[], edit: DocumentEdit, parse: () => ParsedHtml, operation: WordCommentOperation): void {
    const word = metadata.word!;
    const annotation = annotations.find(item => item.annotationId === operation.annotationId) ?? invalid('The Word comment thread is missing.');
    const target = 'replyId' in operation && operation.replyId ? annotation.replies?.find(reply => reply.replyId === operation.replyId) ?? invalid('The Word reply is missing.') : annotation;
    if (target.bodyNodeId ? target.bodyNodeId !== operation.nodeId : !anchorNodes(annotation.anchor).includes(operation.nodeId))
        invalid('A Word comment operation must target its own body or anchor.');
    const body = (html: string): string => {
        const source = parse(), root = source.nodeIds().find(id => source.peekNode(id)!.parentNodeId === null)!, main = word.stories.find(story => story.kind === 'main')!;
        if (main.rootNodeId === root) {
            const children = [...source.peekNode(root)!.childNodeIds], wrapper = edit.insertSource(root, '<div data-office-story="main"></div>').rootNodeIds[0]!;
            edit.moveSource(children, wrapper);
            edit.attributes(root, { 'data-office-story': undefined });
            main.rootNodeId = wrapper;
            for (const section of metadata.sections)
                if (section.rootNodeId === root)
                    section.rootNodeId = wrapper;
        }
        const rootNodeId = edit.insertSource(root, '<div data-office-story="comment">' + html + '</div>').rootNodeIds[0]!;
        word.stories.push({ storyId: rootNodeId, kind: 'comment', rootNodeId });
        return rootNodeId;
    };
    let binding = word.comments?.find(item => item.annotationId === annotation.annotationId);
    if (!binding) {
        binding = { annotationId: annotation.annotationId, replies: (annotation.replies ?? []).map(reply => ({ replyId: reply.replyId })) };
        (word.comments ??= []).push(binding);
    }
    if (operation.op === 'addWordCommentReply') {
        if (operation.parentReplyId && !annotation.replies?.some(reply => reply.replyId === operation.parentReplyId))
            invalid('A reply parent must belong to its thread.');
        const bodyNodeId = body(operation.html), replyId = bodyNodeId;
        (annotation.replies ??= []).push({ replyId, bodyNodeId, ...(operation.author ? { author: operation.author } : {}), ...(operation.date !== undefined ? { date: operation.date } : {}) });
        binding.replies.push({ replyId, ...(operation.parentReplyId ? { parentReplyId: operation.parentReplyId } : {}) });
        return;
    }
    if (operation.op === 'removeWordCommentReply') {
        const removed = binding.replies.find(reply => reply.replyId === operation.replyId);
        binding.replies = binding.replies.filter(reply => reply.replyId !== operation.replyId);
        for (const reply of binding.replies)
            if (reply.parentReplyId === operation.replyId) {
                if (removed?.parentReplyId)
                    reply.parentReplyId = removed.parentReplyId;
                else
                    delete reply.parentReplyId;
            }
        annotation.replies = annotation.replies!.filter(reply => reply.replyId !== operation.replyId);
        if (!annotation.replies.length)
            delete annotation.replies;
        return;
    }
    if (operation.op === 'setWordCommentBody') {
        if (target.bodyNodeId)
            edit.replaceSourceContent(target.bodyNodeId, operation.html);
        else {
            const id = body(operation.html);
            delete target.instruction;
            target.bodyNodeId = id;
        }
        return;
    }
    if (operation.author !== undefined)
        target.author = operation.author;
    if (operation.date !== undefined)
        target.date = operation.date;
}
/** Rich comment stories belong to messages, including newly added messages. */
export function pruneWordComments(word: WordMetadata, annotations: readonly Annotation[], edit: DocumentEdit, parse: () => ParsedHtml): void {
    const kept = new Set(annotations.flatMap(annotation => [annotation, ...annotation.replies ?? []]).flatMap(message => message.bodyNodeId ? [message.bodyNodeId] : [])), source = parse();
    for (const story of word.stories)
        if (story.kind === 'comment' && !kept.has(story.rootNodeId) && source.hasNode(story.rootNodeId))
            edit.removeSource(story.rootNodeId);
    const byId = new Map(annotations.map(annotation => [annotation.annotationId, annotation]));
    if (word.comments)
        word.comments = word.comments.filter(binding => { const annotation = byId.get(binding.annotationId); if (!annotation)
            return false; const replies = new Set(annotation.replies?.map(reply => reply.replyId)); binding.replies = binding.replies.filter(reply => replies.has(reply.replyId)); return true; });
}
