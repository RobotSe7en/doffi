import { createDocumentSession, openDocument, expandWordRevisionGroups } from "@doffi/core";
import type { ContentDocumentModel, DocumentOperation } from "@doffi/core";
/** A disposable candidate restores original formatting through the same review
 * boundary as source edits. The host's session and revision stay unchanged. */
export function wordReviewPreview(model: ContentDocumentModel): ContentDocumentModel {
  const snapshot = model.snapshot(),
    word = snapshot.metadata.word;
  if (!word?.revisions.some((revision) => revision.kind === "format")) return model;
  const operations: DocumentOperation[] = [],
    seen = new Set<string>();
  for (const revision of word.revisions) {
    const id = revision.revisionId ?? revision.nodeId;
    if (revision.kind !== "format" || seen.has(id)) continue;
    const group = expandWordRevisionGroups(word.revisions, [id]);
    // Format projection leaves all text and structural review markers in place.
    if (group.some((item) => item.kind !== "format")) continue;
    for (const item of group) seen.add(item.revisionId ?? item.nodeId);
    operations.push({
      op: "reviewWordRevision",
      nodeId: revision.nodeId,
      revisionId: id,
      decision: "reject",
    });
  }
  if (!operations.length) return model;
  const session = createDocumentSession(model);
  try {
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations,
      }),
    );
    return openDocument(session.serialize(), { kind: "document" });
  } finally {
    session.dispose();
  }
}
