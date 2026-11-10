import type { Annotation, AnnotationReply } from "../contracts/annotations.js";
import type { WorkbookComment, WorkbookCommentMention, WorkbookMetadata, WorkbookNoteBox } from "../contracts/workbook.js";
import { WorkbookCommentOperationSchema } from "../contracts/workbook-comment-operations.js";
import type { WorkbookCommentOperation } from "../contracts/workbook-comment-operations.js";
import type { DeepReadonly, WorkbookModel, WorkbookSnapshot } from "../runtime.js";
import { invalid } from "../document/validation.js";
import { isTextBoundary } from "../document/text.js";
import { resolveWorkbookCommentMention } from "../workbook/comment-mentions.js";
import { WorkbookContentEdit } from "./workbook-content.js";

const copy = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T;
const names = new Set<string>(WorkbookCommentOperationSchema.options.map(schema => schema.shape.op.value));
export function isWorkbookCommentOperation(operation: { readonly op: string }): operation is WorkbookCommentOperation { return names.has(operation.op); }
export function validateWorkbookCommentTarget(base: WorkbookModel, operation: WorkbookCommentOperation): void {
  const anchor = operation.op === "addWorkbookComment" ? operation.anchor : operation.op === "updateWorkbookComment" ? operation.changes.anchor : undefined;
  if (anchor && !base.snapshot().metadata.workbook.sheets.some(sheet => sheet.sheetId === anchor.sheetId)) invalid("Comment anchors must belong to a base worksheet.");
  if ("annotationId" in operation) {
    const annotation = base.snapshot().annotations.find(annotation => annotation.annotationId === operation.annotationId);
    if (!annotation) invalid("Comment targets must belong to the base snapshot.");
    if ("replyId" in operation && operation.replyId !== undefined && !annotation.replies?.some(reply => reply.replyId === operation.replyId)) invalid("Comment reply targets must belong to the base thread.");
    const parent = operation.op === "addWorkbookCommentReply" ? operation.parentReplyId : operation.op === "updateWorkbookCommentReply" ? operation.changes.parentReplyId : undefined;
    if (parent && !annotation.replies?.some(reply => reply.replyId === parent)) invalid("A reply parent must belong to the base thread.");
  }
}

/** Removing any annotation entry point also releases its rich stories and
 * native bindings. Other messages, people and immutable assets remain intact. */
export function removeWorkbookCommentSources(workbook: WorkbookMetadata, content: WorkbookContentEdit, annotation: DeepReadonly<Annotation>): void {
  const roots = new Set([annotation.bodyNodeId, ...(annotation.replies ?? []).map(reply => reply.bodyNodeId)].filter((id): id is string => id !== undefined));
  if (roots.size) for (const story of workbook.stories ?? []) if (roots.has(story.rootNodeId)) content.edit().removeSource(story.rootNodeId);
  if (workbook.stories) workbook.stories = workbook.stories.filter(story => !roots.has(story.rootNodeId));
  if (workbook.comments) workbook.comments = workbook.comments.filter(comment => comment.annotationId !== annotation.annotationId);
}
/** Explicit reanchoring acknowledges the old native placement. Retained reply
 * cell fields follow the new point; SDK multi-region positions remain solely
 * in the annotation and are reported when native output cannot encode them. */
export function reanchorWorkbookCommentSources(workbook: WorkbookMetadata, annotation: Annotation): void {
  if (annotation.anchor.kind !== "cells") return;
  const anchor = annotation.anchor, range = anchor.ranges.length === 1 ? anchor.ranges[0] : undefined;
  const point = range && range.from.row === range.to.row && range.from.column === range.to.column ? range.from : undefined;
  for (const message of [annotation, ...annotation.replies ?? []]) if (message.bodyNodeId) {
    const story = workbook.stories?.find(story => story.rootNodeId === message.bodyNodeId); if (story) story.sheetId = anchor.sheetId;
  }
  const binding = workbook.comments?.find(comment => comment.annotationId === annotation.annotationId);
  if (binding?.format === "modern") for (const reply of binding.replies) if (reply.nativeCell) { if (point) reply.nativeCell = { ...point }; else delete reply.nativeCell; }
  if (binding?.format === "classic" && binding.compatibilityThreadId) {
    delete binding.compatibilityThreadId;
    binding.recoveryReasons = binding.recoveryReasons?.filter(reason => reason !== "orphan-placeholder"); if (!binding.recoveryReasons?.length) delete binding.recoveryReasons;
  }
}
/** Generated people are implementation details until used by the final batch.
 * Imported unused people remain provenance. Reversed display-author changes
 * recover the original native person, including equal-name provider identities. */
export function restoreWorkbookCommentSource(workbook: WorkbookMetadata, annotations: readonly Annotation[], before: WorkbookSnapshot): void {
  const originalBindings = new Map(before.metadata.workbook.comments?.map(comment => [comment.annotationId, comment])), originals = new Map(before.annotations.map(annotation => [annotation.annotationId, annotation])), current = new Map(annotations.map(annotation => [annotation.annotationId, annotation]));
  const used = new Set<string>();
  for (const binding of workbook.comments ?? []) {
    const original = originalBindings.get(binding.annotationId), a = originals.get(binding.annotationId), b = current.get(binding.annotationId);
    if (original?.format === binding.format && a?.author?.name === b?.author?.name) { if (original.authorId !== undefined) binding.authorId = original.authorId; else delete binding.authorId; }
    if (binding.authorId) used.add(binding.authorId);
    if (binding.format === "modern") {
      const oldReplies = new Map(original?.format === "modern" ? original.replies.map(reply => [reply.replyId, reply]) : []), oldMessages = new Map(a?.replies?.map(reply => [reply.replyId, reply])), messages = new Map(b?.replies?.map(reply => [reply.replyId, reply]));
      for (const reply of binding.replies) {
        const old = oldReplies.get(reply.replyId);
        if (old && oldMessages.get(reply.replyId)?.author?.name === messages.get(reply.replyId)?.author?.name) { if (old.authorId !== undefined) reply.authorId = old.authorId; else delete reply.authorId; }
        if (reply.authorId) used.add(reply.authorId); for (const mention of reply.mentions ?? []) used.add(mention.authorId);
      }
      for (const mention of binding.mentions ?? []) used.add(mention.authorId);
    }
  }
  if (workbook.commentAuthors) {
    const original = new Set(before.metadata.workbook.commentAuthors?.map(author => author.authorId));
    workbook.commentAuthors = workbook.commentAuthors.filter(author => original.has(author.authorId) || used.has(author.authorId));
    if (!workbook.commentAuthors.length && before.metadata.workbook.commentAuthors === undefined) delete workbook.commentAuthors;
  }
}
function mapMentions(values: readonly WorkbookCommentMention[], before: string, after: string, edits: Extract<WorkbookCommentOperation, { op: "setWorkbookCommentBody" }>["textEdits"]): WorkbookCommentMention[] {
  if (before === after && !edits) return values.map(copy);
  if (!edits) invalid("Changing a mentioned body requires exact text edits or explicit replacement mentions.");
  let end = 0, output = "";
  for (const edit of edits) {
    if (edit.from < end || edit.to < edit.from || !isTextBoundary(before, edit.from) || !isTextBoundary(before, edit.to) || before.slice(edit.from, edit.to) !== edit.expectedText) invalid("Comment text edits require ordered, exact UTF-16 source intervals.");
    output += before.slice(end, edit.from) + edit.text; end = edit.to;
  }
  output += before.slice(end); if (output !== after) invalid("Comment text edits do not prove the new source body.");
  return values.map(mention => {
    const range = resolveWorkbookCommentMention(before, mention); if (!range) invalid("An invalid native mention requires explicit replacement mentions.");
    let shift = 0;
    for (const edit of edits) {
      if (edit.to <= range.from) shift += edit.text.length - (edit.to - edit.from);
      else if (edit.from < range.to) invalid("An edit intersecting a mention requires explicit replacement mentions.");
    }
    const mapped = { ...mention, startIndex: mention.startIndex + shift };
    if (!resolveWorkbookCommentMention(after, mapped)) invalid("The mapped mention splits a Unicode boundary."); return mapped;
  });
}
function noteDifferences(a: WorkbookNoteBox, b: WorkbookNoteBox): NonNullable<Extract<WorkbookComment, { format: "classic" }>["boxDifferences"]> {
  const names = ["anchor", "cssClass", "visible", "printable", "locked", "lockText", "defaultSize", "disabled", "autoFill", "autoLine", "autoScale", "justifyLastLine", "altText"] as const;
  return names.filter(key => JSON.stringify(a[key]) !== JSON.stringify(b[key])).map(key => key === "cssClass" ? "appearance" : key);
}
export function editWorkbookComment(workbook: WorkbookMetadata, annotations: Annotation[], content: WorkbookContentEdit, base: WorkbookModel, operation: WorkbookCommentOperation, allocate: () => string): void {
  const nativeAuthor = (display: Annotation["author"], format: "classic" | "modern", currentId?: string): string | undefined => {
    if (!display) return undefined;
    const current = workbook.commentAuthors?.find(author => author.authorId === currentId && author.format === format && author.name === display.name);
    if (current) return current.authorId;
    const existing = workbook.commentAuthors?.find(author => author.format === format && author.name === display.name);
    if (existing) return existing.authorId;
    const authorId = allocate(); (workbook.commentAuthors ??= []).push({ authorId, format, name: display.name }); return authorId;
  };
  const body = (html: string, sheetId: string): string => {
    const rootNodeId = content.edit().insertSource(content.root(), '<section data-office-kind="comment" data-office-story="comment">' + html + '</section>').rootNodeIds[0]!;
    (workbook.stories ??= []).push({ storyId: allocate(), sheetId, rootNodeId, kind: "comment" }); return rootNodeId;
  };
  if (operation.op === "addWorkbookComment") {
    const format = operation.format ?? "modern"; if (format === "modern" && operation.box) invalid("Modern threads have no classic note box.");
    const annotationId = allocate(), bodyNodeId = body(operation.html, operation.anchor.sheetId), authorId = nativeAuthor(operation.author, format);
    annotations.push({ annotationId, documentId: base.snapshot().manifest.documentId, baseRevision: base.snapshot().manifest.revision, status: "open", anchor: copy(operation.anchor), bodyNodeId,
      ...(operation.author ? { author: copy(operation.author) } : {}), ...(operation.date !== undefined ? { date: operation.date } : {}) });
    (workbook.comments ??= []).push({ annotationId, ...(authorId ? { authorId } : {}), ...(format === "modern" ? { format, replies: [] } : { format, ...(operation.box ? { box: copy(operation.box) } : {}) }) }); return;
  }
  const annotation = annotations.find(annotation => annotation.annotationId === operation.annotationId) ?? invalid("An earlier operation removed the comment.");
  if (annotation.anchor.kind !== "cells") invalid("Workbook comments require cell anchors.");
  let binding = workbook.comments?.find(comment => comment.annotationId === annotation.annotationId);
  function modern(): Extract<WorkbookComment, { format: "modern" }> {
    if (binding?.format === "classic") invalid("Classic notes cannot contain replies or mentions.");
    if (!binding) { binding = { annotationId: annotation.annotationId, format: "modern", replies: (annotation.replies ?? []).map(reply => ({ replyId: reply.replyId })) }; (workbook.comments ??= []).push(binding); }
    return binding;
  }
  const targetReply = (id: string): AnnotationReply => annotation.replies?.find(reply => reply.replyId === id) ?? invalid("An earlier operation removed this reply.");
  if (operation.op === "addWorkbookCommentReply") {
    const thread = modern(); if (operation.parentReplyId) targetReply(operation.parentReplyId);
    const replyId = allocate(), bodyNodeId = body(operation.html, annotation.anchor.sheetId), authorId = nativeAuthor(operation.author, "modern");
    (annotation.replies ??= []).push({ replyId, bodyNodeId, ...(operation.author ? { author: copy(operation.author) } : {}), ...(operation.date !== undefined ? { date: operation.date } : {}) });
    thread.replies.push({ replyId, ...(authorId ? { authorId } : {}), ...(operation.parentReplyId ? { parentReplyId: operation.parentReplyId } : {}) }); return;
  }
  if (operation.op === "removeWorkbookCommentReply") {
    const thread = modern(), reply = targetReply(operation.replyId), removed = thread.replies.find(item => item.replyId === reply.replyId)!;
    if (reply.bodyNodeId) { content.edit().removeSource(reply.bodyNodeId); workbook.stories = workbook.stories?.filter(story => story.rootNodeId !== reply.bodyNodeId); }
    annotation.replies = annotation.replies!.filter(item => item !== reply); if (!annotation.replies.length) delete annotation.replies;
    thread.replies = thread.replies.filter(item => item !== removed);
    for (const item of thread.replies) if (item.parentReplyId === reply.replyId) { if (removed.parentReplyId) item.parentReplyId = removed.parentReplyId; else delete item.parentReplyId; } return;
  }
  if (operation.op === "setWorkbookCommentBody" || operation.op === "replaceWorkbookCommentText") {
    const target = operation.replyId ? targetReply(operation.replyId) : annotation, before = target.instruction ?? content.parse().getNode(target.bodyNodeId!)!.text;
    const mentions = binding?.format === "modern" ? operation.replyId ? binding.replies.find(reply => reply.replyId === operation.replyId) : binding : undefined;
    if (operation.mentions !== undefined && binding?.format === "classic") invalid("A classic note has no mentions.");
    if (operation.op === "replaceWorkbookCommentText") {
      if (!target.bodyNodeId) {
        // Host instruction annotations become rich stories only through an
        // explicit body operation; an ordinary text edit keeps their source.
        if (operation.from > operation.to || !isTextBoundary(before, operation.from) || !isTextBoundary(before, operation.to) || before.slice(operation.from, operation.to) !== operation.expectedText) invalid("Comment text editing requires an exact UTF-16 interval.");
        const text = before.slice(0, operation.from) + operation.text + before.slice(operation.to);
        if (!text) invalid("An instruction annotation cannot have an empty instruction."); target.instruction = text;
      } else content.edit().replaceSourceText(target.bodyNodeId, operation);
    } else if (target.bodyNodeId) content.edit().replaceSourceContent(target.bodyNodeId, operation.html);
    else { const bodyNodeId = body(operation.html, annotation.anchor.sheetId); delete target.instruction; target.bodyNodeId = bodyNodeId; }
    const after = target.instruction ?? content.parse().getNode(target.bodyNodeId!)!.text;
    const edits = operation.op === "replaceWorkbookCommentText" ? [{ from: operation.from, to: operation.to, expectedText: operation.expectedText, text: operation.text }] : operation.textEdits;
    if (edits) mapMentions([], before, after, edits);
    if (operation.mentions !== undefined) {
      const thread = modern(), target = operation.replyId ? thread.replies.find(reply => reply.replyId === operation.replyId)! : thread;
      if (operation.mentions.length) target.mentions = copy(operation.mentions); else delete target.mentions;
    } else if (mentions?.mentions?.length) mentions.mentions = mapMentions(mentions.mentions, before, after, edits);
    if (binding?.format === "modern" && binding.recoveryReasons?.includes("invalid-mention-range")) {
      const valid = (body: Annotation | AnnotationReply, values: readonly WorkbookCommentMention[] | undefined) => {
        const text = body.instruction ?? content.parse().getNode(body.bodyNodeId!)!.text; return !values?.some(mention => !resolveWorkbookCommentMention(text, mention));
      };
      if (valid(annotation, binding.mentions) && binding.replies.every(reply => valid(targetReply(reply.replyId), reply.mentions))) {
        binding.recoveryReasons = binding.recoveryReasons.filter(reason => reason !== "invalid-mention-range"); if (!binding.recoveryReasons.length) delete binding.recoveryReasons;
      }
    }
    return;
  }
  if (operation.op === "setWorkbookNoteBox") {
    if (binding?.format !== "classic") invalid("Note boxes require an existing classic note.");
    if (operation.box) { binding.box = copy(operation.box); if (binding.commentPrBox) { const changes = noteDifferences(binding.box, binding.commentPrBox); if (changes.length) binding.boxDifferences = changes; else delete binding.boxDifferences; } }
    else { delete binding.box; delete binding.boxSource; delete binding.commentPrBox; delete binding.boxDifferences; } return;
  }
  if (operation.op === "updateWorkbookCommentReply") {
    const thread = modern(), reply = targetReply(operation.replyId), link = thread.replies.find(item => item.replyId === operation.replyId)!;
    const { parentReplyId, ...changes } = operation.changes;
    for (const [key, value] of Object.entries(changes)) { if (value === null) Reflect.deleteProperty(reply, key); else if (value !== undefined) Reflect.set(reply, key, copy(value)); }
    if (operation.changes.author !== undefined) { const id = nativeAuthor(reply.author, "modern", link.authorId); if (id) link.authorId = id; else delete link.authorId; }
    if (parentReplyId !== undefined) { if (parentReplyId === null) delete link.parentReplyId; else { targetReply(parentReplyId); link.parentReplyId = parentReplyId; } } return;
  }
  if (operation.op === "updateWorkbookComment") {
    const { format, ...changes } = operation.changes;
    if (annotation.status === "needs_reanchor" && changes.status !== undefined && changes.status !== "needs_reanchor" && !changes.anchor) invalid("Reactivating a comment requires an explicit cell anchor.");
    if (format !== undefined && format !== binding?.format) {
      if (format === "classic" && (annotation.replies?.length || binding?.format === "modern" && binding.mentions?.length)) invalid("Converting a thread to a classic note requires explicitly removing replies and mentions.");
      const authorId = nativeAuthor(changes.author === undefined ? annotation.author : changes.author ?? undefined, format), replacement: WorkbookComment = { annotationId: annotation.annotationId, ...(authorId ? { authorId } : {}), ...(format === "modern" ? { format, replies: (annotation.replies ?? []).map(reply => {
        const authorId = nativeAuthor(reply.author, "modern"); return { replyId: reply.replyId, ...(authorId ? { authorId } : {}) }; }) } : { format }),
        ...((changes.status ?? annotation.status) === "needs_reanchor" ? { recoveryStatus: binding?.recoveryStatus ?? (annotation.status === "needs_reanchor" ? "open" : annotation.status) } : {}) };
      if (binding) workbook.comments![workbook.comments!.indexOf(binding)] = replacement; else (workbook.comments ??= []).push(replacement); binding = replacement;
    }
    const prior = annotation.status;
    for (const [key, value] of Object.entries(changes)) { if (value === null) Reflect.deleteProperty(annotation, key); else if (value !== undefined) Reflect.set(annotation, key, copy(value)); }
    if (binding) {
      if (changes.author !== undefined) { const id = nativeAuthor(annotation.author, binding.format, binding.authorId); if (id) binding.authorId = id; else delete binding.authorId; }
      if (annotation.status === "needs_reanchor" && prior !== "needs_reanchor") binding.recoveryStatus = prior;
      if (annotation.status !== "needs_reanchor") { delete binding.recoveryStatus; delete binding.recoveryReasons; }
    }
    if (changes.anchor) reanchorWorkbookCommentSources(workbook, annotation);
  }
}
