import { DocumentError } from "../errors.js";
import type { WordDocumentMetadata } from "../contracts/document.js";
import type { WordMetadata } from "../contracts/word.js";
import type { ParsedHtml } from "../document/html.js";
import type { DocumentEdit } from "../document/edit.js";
import { isTextBoundary } from "../document/text.js";
import { wordTableRows, removeWordRows } from "./word-row-revisions.js";
import { deleteWordTrackedContent } from "./word-content-revisions.js";

export interface WordContentRange {
  nodeId: string;
  endNodeId?: string | undefined;
  from: number;
  to: number;
  expectedText: string;
}
export function wordContentScope(
  metadata: WordDocumentMetadata,
  source: ParsedHtml,
  id: string,
): string {
  const roots = new Set([
    ...metadata.word!.stories.map((story) => story.rootNodeId),
    ...metadata.sections.map((section) => section.rootNodeId),
  ]);
  for (let current: string | null = id; current; current = source.peekNode(current)!.parentNodeId) {
    const node = source.peekNode(current)!;
    if (
      roots.has(current) ||
      ["td", "th", "li"].includes(node.tagName) ||
      node.attributes["data-word-content-control"]
    )
      return current;
  }
  throw new DocumentError("InvalidOperation", "The Word content range has no story.");
}
export function resolveWordContentRange(
  word: WordMetadata,
  source: ParsedHtml,
  operation: WordContentRange,
) {
  const first = source.getNode(operation.nodeId),
    last = source.getNode(operation.endNodeId ?? operation.nodeId);
  if (
    !first ||
    !last ||
    !/^p$|^h[1-6]$/.test(first.tagName) ||
    !/^p$|^h[1-6]$/.test(last.tagName) ||
    operation.from > first.text.length ||
    operation.to > last.text.length ||
    !isTextBoundary(first.text, operation.from) ||
    !isTextBoundary(last.text, operation.to)
  )
    throw new DocumentError("InvalidOperation", "A Word range requires exact paragraph endpoints.");
  const story = (id: string): string | undefined => {
    const roots = new Set(word.stories.map((item) => item.rootNodeId));
    for (let current: string | null = id; current; current = source.peekNode(current)!.parentNodeId)
      if (roots.has(current)) return current;
    return;
  };
  const root = story(first.nodeId);
  const ordered = source.nodeIds(),
    begin = ordered.indexOf(first.nodeId),
    finish = ordered.indexOf(last.nodeId);
  const from = source.textOffset(first.nodeId)! + operation.from,
    to = source.textOffset(last.nodeId)! + operation.to;
  if (
    !root ||
    root !== story(last.nodeId) ||
    begin > finish ||
    from > to ||
    source
      .getNode(root)!
      .text.slice(from - source.textOffset(root)!, to - source.textOffset(root)!) !==
      operation.expectedText
  )
    throw new DocumentError(
      "InvalidOperation",
      "The Word range is stale, reversed or crosses stories.",
    );
  const paragraphs = ordered
    .slice(begin, finish + 1)
    .filter((id) => /^p$|^h[1-6]$/.test(source.peekNode(id)!.tagName));
  return { first, last, root, from, to, paragraphs };
}

/** Delete only the selected source. Boundary cells and their native grid stay
 * intact; complete interior tables use row review and the existing grid owner. */
export function deleteWordContentRange(
  metadata: WordDocumentMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  operation: WordContentRange,
  group: { groupId: string; author: string; date: string },
): ReturnType<typeof resolveWordContentRange> {
  const word = metadata.word!,
    source = parse(),
    range = resolveWordContentRange(word, source, operation);
  const contains = (root: string, id: string): boolean => {
    for (let current: string | null = id; current; current = source.peekNode(current)!.parentNodeId)
      if (current === root) return true;
    return false;
  };
  const order = source.nodeIds(),
    begin = order.indexOf(range.first.nodeId),
    end = order.indexOf(range.last.nodeId);
  const coveredTables = order
    .map((nodeId) => source.peekNode(nodeId)!)
    .filter(
      (table) =>
        table.tagName === "table" &&
        order.indexOf(table.nodeId) > begin &&
        order.indexOf(table.nodeId) < end &&
        !contains(table.nodeId, range.last.nodeId),
    );
  const rows = new Set<string>();
  for (const table of coveredTables) {
    if (!parse().hasNode(table.nodeId)) continue;
    if (!word.trackChanges) edit.removeSource(table.nodeId);
    else {
      const cancelled = new Set<string>();
      for (const id of wordTableRows(source, table.nodeId)) {
        rows.add(id);
        if (
          word.revisions.some(
            (r) => r.nodeId === id && r.kind === "rowInsert" && r.author === group.author,
          )
        ) {
          cancelled.add(id);
          continue;
        }
        if (word.revisions.some((r) => r.nodeId === id && r.kind === "rowDelete")) continue;
        word.revisions.push({
          nodeId: id,
          revisionId: id + "-range-delete-" + group.groupId,
          sourceId: id + "-range-delete-" + group.groupId,
          kind: "rowDelete",
          ...group,
        });
      }
      if (cancelled.size) removeWordRows(word, edit, parse, cancelled, edit.readLimits);
    }
  }
  for (const id of range.paragraphs) {
    if (!parse().hasNode(id) || [...rows].some((row) => contains(row, id))) continue;
    const text = source.getNode(id)!.text;
    const from = id === range.first.nodeId ? operation.from : 0,
      to = id === range.last.nodeId ? operation.to : text.length;
    if (word.trackChanges) {
      deleteWordTrackedContent(word, edit, parse, id, from, to, group);
    } else if (id !== range.first.nodeId && id !== range.last.nodeId) edit.removeSource(id);
    else
      for (const removed of edit.wrapSourceText(id, from, to, "span")) edit.removeSource(removed);
  }
  if (word.trackChanges) {
    for (const [index, id] of range.paragraphs.entries()) {
      if (
        id === range.last.nodeId ||
        !parse().hasNode(id) ||
        [...rows].some((row) => contains(row, id))
      )
        continue;
      const next = range.paragraphs
        .slice(index + 1)
        .find(
          (next) =>
            parse().hasNode(next) &&
            wordContentScope(metadata, source, next) === wordContentScope(metadata, source, id),
        );
      if (next)
        if (
          !word.revisions.some(
            (r) => r.nodeId === id && ["paragraphDelete", "paragraphMoveFrom"].includes(r.kind),
          )
        )
          word.revisions.push({
            nodeId: id,
            revisionId: id + "-range-mark-" + group.groupId,
            sourceId: id + "-range-mark-" + group.groupId,
            kind: "paragraphDelete",
            ...group,
          });
    }
  } else {
    const current = parse();
    for (const id of current.nodeIds()) {
      const node = current.peekNode(id)!;
      if (!["td", "th"].includes(node.tagName) || node.attributes["data-word-grid-gap"]) continue;
      const pending = [...node.childNodeIds];
      let block = false;
      while (pending.length) {
        const child = current.peekNode(pending.pop()!)!;
        if (/^p$|^h[1-6]$|^table$/.test(child.tagName)) {
          block = true;
          break;
        }
        pending.push(...child.childNodeIds);
      }
      if (!block) edit.insertSource(id, "<p></p>");
    }
  }
  return range;
}
