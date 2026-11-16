import type { Annotation, Anchor } from "../contracts/annotations.js";
import type { WordContentSemantics } from "../contracts/word-content.js";
import type { WordMetadata } from "../contracts/word.js";
import type { ParsedHtml } from "../document/html.js";
import { validTextRanges } from "../document/annotations.js";
import { createOpaqueId } from "../identifiers.js";
import { DocumentError } from "../errors.js";
import { mapAnnotations } from "./anchors.js";
import type { DocumentEdit } from "../document/edit.js";
import type { DeepReadonly } from "../runtime.js";

export interface WordContentCommentContext {
  annotations: Annotation[];
  documentId: string;
  baseRevision: number;
  copies?: Map<string, string>;
}
const invalid = (message: string): never => {
  throw new DocumentError("InvalidOperation", message);
};

/** Transport IDs can resolve only to freshly inserted content, never receiver
 * annotations, bodies or nodes. Native durable IDs are allocated on export. */
export function attachWordContentComments(
  word: WordMetadata,
  semantics: WordContentSemantics,
  ids: ReadonlyMap<string, string>,
  storyIds: ReadonlyMap<string, string>,
  owners: Set<string>,
  source: ParsedHtml,
  context?: WordContentCommentContext,
): void {
  const annotations = semantics.annotations ?? [],
    bindings = new Map((semantics.comments ?? []).map((b) => [b.annotationId, b])),
    seen = new Set<string>(),
    bodies = new Set<string>();
  if (bindings.size !== (semantics.comments ?? []).length)
    invalid("Copied comment bindings must be unique.");
  if (annotations.length && !context) invalid("Copied comments require the receiving session.");
  const map = (id: string): string =>
    ids.get(id) ?? invalid("A copied comment may only refer to its selected source graph.");
  const body = (id: string): string => {
    const story = semantics.stories?.find((s) => s.rootNodeId === id && s.kind === "comment");
    if (!story || !storyIds.has(story.storyId) || bodies.has(id))
      return invalid("Each copied comment body requires its independently owned story.");
    bodies.add(id);
    owners.add(story.storyId);
    return map(id);
  };
  for (const original of annotations) {
    if (seen.has(original.annotationId)) invalid("Copied thread identities must be unique.");
    seen.add(original.annotationId);
    const anchor = original.anchor;
    let mapped: Anchor;
    if (anchor.kind === "text" || anchor.kind === "text-set") {
      const ranges = (anchor.kind === "text" ? [anchor] : anchor.ranges).map((range) => ({
        nodeId: map(range.nodeId),
        from: range.from,
        to: range.to,
        quote: range.quote,
      }));
      if (!validTextRanges(ranges, source))
        invalid("Copied comment text must match exact UTF-16 source ranges.");
      mapped =
        anchor.kind === "text" ? { kind: "text", ...ranges[0]! } : { kind: "text-set", ranges };
    } else if (anchor.kind === "node") mapped = { kind: "node", nodeId: map(anchor.nodeId) };
    else if (anchor.kind === "node-set")
      mapped = { kind: "node-set", nodeIds: anchor.nodeIds.map(map) };
    else
      return invalid(
        "Clipboard comments require source text or object anchors, not layout coordinates.",
      );
    const copy = JSON.parse(JSON.stringify(original)) as Annotation,
      annotationId = "word-content-comment-" + createOpaqueId(),
      binding = bindings.get(original.annotationId),
      replies = new Map<string, string>();
    copy.annotationId = annotationId;
    copy.documentId = context!.documentId;
    copy.baseRevision = context!.baseRevision;
    copy.anchor = mapped;
    // Either range endpoint can be opaque even when its original reference is
    // outside the selection. Keep that contextual binding for the retained
    // endpoint; the exporter supplies the missing clipped controls.
    const endpointIds =
      mapped.kind === "text" || mapped.kind === "node"
        ? [mapped.nodeId]
        : mapped.kind === "text-set"
          ? mapped.ranges.map((r) => r.nodeId)
          : mapped.kind === "node-set"
            ? mapped.nodeIds
            : [];
    if (
      !semantics.nativeContent?.fragments.some((f) => {
        if (!f.nodeId || f.kind === "property") return false;
        if (binding?.referenceNodeId === f.nodeId) return true;
        for (let id: string | null = map(f.nodeId); id; id = source.peekNode(id)!.parentNodeId)
          if (endpointIds.includes(id)) return true;
        return false;
      })
    )
      delete copy.sourceId;
    if (copy.bodyNodeId) copy.bodyNodeId = body(copy.bodyNodeId);
    for (const reply of copy.replies ?? []) {
      if (replies.has(reply.replyId))
        invalid("Copied reply identities must be unique within a thread.");
      const id = "word-content-reply-" + createOpaqueId();
      replies.set(reply.replyId, id);
      reply.replyId = id;
      delete reply.sourceId;
      if (reply.bodyNodeId) reply.bodyNodeId = body(reply.bodyNodeId);
    }
    const mapReply = (id: string): string =>
      replies.get(id) ?? invalid("Copied reply bindings must belong to their own thread.");
    (word.comments ??= []).push({
      annotationId,
      ...(binding?.followUp !== undefined ? { followUp: binding.followUp } : {}),
      ...(binding?.referenceNodeId && ids.has(binding.referenceNodeId)
        ? { referenceNodeId: map(binding.referenceNodeId) }
        : {}),
      replies: (binding?.replies ?? []).map((reply) => ({
        replyId: mapReply(reply.replyId),
        ...(reply.parentReplyId ? { parentReplyId: mapReply(reply.parentReplyId) } : {}),
      })),
    });
    context!.annotations.push(copy);
    if (original.documentId === context!.documentId)
      context!.copies?.set(annotationId, original.annotationId);
  }
  if ([...bindings.keys()].some((id) => !seen.has(id)))
    invalid("A copied comment binding requires its selected thread.");
}

/** A verified retained cut follows exact textTransfers. Reuse its thread only
 * when that proof maps the entire old anchor onto the pasted anchor. IDs from
 * HTML alone never select or modify a receiver thread. */
export function mergeWordMovedClipboardComments(
  word: WordMetadata,
  annotations: Annotation[],
  copies: ReadonlyMap<string, string>,
  before: ParsedHtml,
  after: ParsedHtml,
  edit: DocumentEdit,
  previous: readonly DeepReadonly<Annotation>[],
): void {
  if (!edit.textTransfers.length) return;
  const intervals = (anchor: Anchor): unknown => {
    if (
      anchor.kind !== "text" &&
      anchor.kind !== "text-set" &&
      anchor.kind !== "node" &&
      anchor.kind !== "node-set"
    )
      return;
    const ranges =
      anchor.kind === "text"
        ? [anchor]
        : anchor.kind === "text-set"
          ? anchor.ranges
          : (anchor.kind === "node" ? [anchor.nodeId] : anchor.nodeIds).map((nodeId) => ({
              nodeId,
              from: 0,
              to: after.textLength(nodeId)!,
            }));
    return ranges.map((r) => [
      after.textScope(r.nodeId, r.from, r.to),
      after.textOffset(r.nodeId)! + r.from,
      after.textOffset(r.nodeId)! + r.to,
    ]);
  };
  for (const [copiedId, originalId] of copies) {
    const copy = annotations.find((a) => a.annotationId === copiedId),
      original = annotations.find((a) => a.annotationId === originalId);
    if (
      !copy ||
      !original ||
      original.status === "needs_reanchor" ||
      copy.status === "needs_reanchor"
    )
      continue;
    const old = previous.find((a) => a.annotationId === originalId)?.anchor;
    if (!old || old.kind === "region" || old.kind === "cells") continue;
    const oldIds =
      old.kind === "node" || old.kind === "text"
        ? [old.nodeId]
        : old.kind === "text-set"
          ? old.ranges.map((r) => r.nodeId)
          : old.nodeIds;
    if (!oldIds.some((id) => edit.textTransfers.some((t) => t.nodeId === id))) continue;
    // Word node comments follow complete objects as well as text. Prove their
    // move as a source interval before restoring the object anchor kind.
    if (original.anchor.kind === "node" || original.anchor.kind === "node-set") {
      const ids =
        original.anchor.kind === "node" ? [original.anchor.nodeId] : original.anchor.nodeIds;
      if (
        ids.some(
          (id) =>
            !before.hasNode(id) ||
            !edit.textTransfers.some(
              (t) => t.nodeId === id && t.from === 0 && t.to === before.textLength(id),
            ),
        )
      )
        continue;
      const probe = JSON.parse(JSON.stringify(original)) as Annotation;
      probe.anchor = {
        kind: "text-set",
        ranges: ids.map((nodeId) => ({
          nodeId,
          from: 0,
          to: before.textLength(nodeId)!,
          quote: before.getNode(nodeId)!.text,
        })),
      };
      mapAnnotations([probe], before, after, edit, new Set());
      if (
        probe.status === "needs_reanchor" ||
        JSON.stringify(intervals(probe.anchor)) !== JSON.stringify(intervals(copy.anchor))
      )
        continue;
    } else if (
      JSON.stringify(intervals(original.anchor)) !== JSON.stringify(intervals(copy.anchor))
    )
      continue;
    original.anchor = copy.anchor;
    const copiedBinding = word.comments?.find((b) => b.annotationId === copiedId),
      binding = word.comments?.find((b) => b.annotationId === originalId);
    if (binding) {
      if (copiedBinding?.referenceNodeId) binding.referenceNodeId = copiedBinding.referenceNodeId;
      else delete binding.referenceNodeId;
    }
    annotations.splice(annotations.indexOf(copy), 1);
  }
}
