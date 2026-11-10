import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookCoordinate, WorkbookMetadata, WorkbookNoteBox } from "../contracts/workbook.js";
import type { DeepReadonly } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { WorkbookContentEdit } from "./workbook-content.js";

const copy = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T;
type CellAnchor = Extract<Annotation["anchor"], { kind: "cells" }>;
/** Capture original stories and bindings once, before overlapping destinations
 * can remove them. Both whole-sheet and bounded range copies own fresh IDs. */
export function createWorkbookCommentCloner(workbook: WorkbookMetadata, content: WorkbookContentEdit, allocate: () => string) {
  const stories = new Map(workbook.stories?.map(story => [story.rootNodeId, story]));
  const bindings = new Map(workbook.comments?.map(binding => [binding.annotationId, binding]));
  return (original: DeepReadonly<Annotation>, anchor: DeepReadonly<CellAnchor>, revision: number,
    geometry: { box?(box: WorkbookNoteBox): WorkbookNoteBox; coordinate?(at: WorkbookCoordinate): WorkbookCoordinate } = {}) => {
    const value: Annotation = copy(original) as Annotation, replies = new Map((value.replies ?? []).map(reply => [reply.replyId, allocate()]));
    const identities = new Map<string, string>();
    value.annotationId = allocate(); value.baseRevision = revision; value.anchor = copy(anchor) as CellAnchor; delete value.sourceId;
    for (const reply of value.replies ?? []) { reply.replyId = replies.get(reply.replyId)!; delete reply.sourceId; }
    for (const message of [value, ...value.replies ?? []]) if (message.bodyNodeId) {
      const story = stories.get(message.bodyNodeId);
      if (!story) throw new DocumentError("InvalidOperation", "The copied comment story is missing.");
      const cloned = content.edit().cloneSource(message.bodyNodeId, content.root());
      message.bodyNodeId = cloned.rootNodeIds[0]!;
      workbook.stories!.push({ ...copy(story), storyId: allocate(), rootNodeId: message.bodyNodeId, sheetId: anchor.sheetId });
      for (const [from, to] of cloned.identities) identities.set(from, to);
    }
    const originalBinding = bindings.get(original.annotationId);
    if (originalBinding) {
      const binding = copy(originalBinding); binding.annotationId = value.annotationId;
      if (binding.format === "classic") {
        delete binding.nativeShapeId; delete binding.nativeUid;
        if (binding.box && geometry.box) binding.box = geometry.box(binding.box);
        if (binding.commentPrBox && geometry.box) binding.commentPrBox = geometry.box(binding.commentPrBox);
      } else {
        delete binding.compatibilityUids;
        for (const mention of binding.mentions ?? []) mention.mentionId = allocate();
        for (const reply of binding.replies) {
          reply.replyId = replies.get(reply.replyId)!;
          if (reply.parentReplyId) reply.parentReplyId = replies.get(reply.parentReplyId)!;
          if (reply.nativeCell && geometry.coordinate) reply.nativeCell = geometry.coordinate(reply.nativeCell);
          for (const mention of reply.mentions ?? []) mention.mentionId = allocate();
        }
      }
      (workbook.comments ??= []).push(binding);
    }
    return { annotation: value, identities };
  };
}
