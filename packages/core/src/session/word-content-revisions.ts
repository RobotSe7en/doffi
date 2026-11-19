import type { WordMetadata } from "../contracts/word.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";

export interface WordContentRevisionGroup {
  groupId: string;
  author: string;
  date: string;
}

/** Partition by actual source revision boundaries. Own pending insertions are
 * cancelled; retained deletions keep their author, and other insertions receive
 * a separate deletion history. Work backwards so source offsets stay exact. */
export function deleteWordTrackedContent(
  word: WordMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  nodeId: string,
  from: number,
  to: number,
  group: WordContentRevisionGroup,
): number {
  const source = parse(),
    origin = source.textOffset(nodeId)!;
  const inside = (id: string): boolean => {
    for (let current: string | null = id; current; current = source.peekNode(current)!.parentNodeId)
      if (current === nodeId) return true;
    return false;
  };
  const revisions = word.revisions
    .filter(
      (r) =>
        ["insert", "delete", "moveFrom", "moveTo"].includes(r.kind) &&
        source.hasNode(r.nodeId) &&
        inside(r.nodeId),
    )
    .map((revision) => ({
      revision,
      from: source.textOffset(revision.nodeId)! - origin,
      to: source.textOffset(revision.nodeId)! - origin + source.textLength(revision.nodeId)!,
    }));
  const points = [
    ...new Set([
      from,
      to,
      ...revisions.flatMap((r) => [r.from, r.to]).filter((p) => from < p && p < to),
    ]),
  ].sort((a, b) => a - b);
  let removed = 0;
  for (let index = points.length - 2; index >= 0; index--) {
    const start = points[index]!,
      end = points[index + 1]!;
    if (start === end) continue;
    const owners = revisions.filter((r) => r.from <= start && end <= r.to);
    if (owners.some((r) => ["delete", "moveFrom"].includes(r.revision.kind))) continue;
    const insertion = owners
      .filter((r) => ["insert", "moveTo"].includes(r.revision.kind))
      .sort((a, b) => a.to - a.from - (b.to - b.from))[0];
    const cancel =
      insertion?.revision.kind === "insert" && insertion.revision.author === group.author;
    const ids = edit.wrapSourceText(nodeId, start, end, cancel ? "span" : "del");
    if (cancel) {
      for (const id of ids) edit.removeSource(id);
      removed += end - start;
    } else
      for (const id of ids) {
        edit.attributes(id, { "data-word-run-overlay": "true" });
        word.revisions.push({ nodeId: id, revisionId: id, sourceId: id, kind: "delete", ...group });
      }
  }
  return to - removed;
}
