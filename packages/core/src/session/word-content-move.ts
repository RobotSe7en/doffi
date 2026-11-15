import { DocumentError } from "../errors.js";
import { createOpaqueId } from "../identifiers.js";
import type { WordDocumentMetadata } from "../contracts/document.js";
import type { WordMetadata } from "../contracts/word.js";
import type { WordOperation } from "../contracts/word-operations.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import { deleteWordContentRange, wordContentScope } from "./word-content-range.js";
export function cutWordDrawings(
  metadata: WordDocumentMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  operation: Extract<WordOperation, { op: "cutWordDrawings" }>,
): void {
  const word = metadata.word!,
    source = parse(),
    ids = new Set(operation.nodeIds),
    drawings = new Set(word.drawings?.map((d) => d.nodeId));
  if (
    ids.size !== operation.nodeIds.length ||
    !ids.has(operation.nodeId) ||
    word.revisions.some((r) => r.groupId === operation.cutId)
  )
    throw new DocumentError(
      "InvalidOperation",
      "A drawing cut requires unique selected objects and a fresh cut identity.",
    );
  const ranges: {
    nodeId: string;
    from: number;
    to: number;
    expectedText: string;
    order: number;
  }[] = [];
  for (const id of ids) {
    if (!drawings.has(id)) throw new DocumentError("InvalidOperation", "A cut drawing is missing.");
    let parent = source.peekNode(id)!.parentNodeId,
      paragraph: string | undefined,
      selectedParent = false,
      grouped = false;
    while (parent) {
      if (ids.has(parent)) selectedParent = true;
      if (drawings.has(parent)) grouped = true;
      if (/^p$|^h[1-6]$/.test(source.peekNode(parent)!.tagName)) {
        paragraph = parent;
        break;
      }
      parent = source.peekNode(parent)!.parentNodeId;
    }
    if (selectedParent) continue;
    if (grouped)
      throw new DocumentError(
        "UnsupportedFeature",
        "Cut a group as one native object before changing its internal structure.",
      );
    if (!paragraph)
      throw new DocumentError("InvalidOperation", "A drawing cut requires its source paragraph.");
    const from = source.textOffset(id)! - source.textOffset(paragraph)!,
      to = from + source.textLength(id)!;
    ranges.push({
      nodeId: paragraph,
      from,
      to,
      expectedText: source.getNode(id)!.text,
      order: source.nodeIds().indexOf(id),
    });
  }
  const group = {
    groupId: operation.cutId,
    author: word.trackingAuthor ?? "",
    date: new Date().toISOString(),
  };
  for (const range of ranges.sort((a, b) => b.order - a.order))
    deleteWordContentRange(metadata, edit, parse, range, group);
}

export function cutWordContent(
  metadata: WordDocumentMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  operation: Extract<WordOperation, { op: "cutWordContent" }>,
): void {
  const word = metadata.word!,
    source = parse();
  if (
    word.revisions.some((r) => r.groupId === operation.cutId) ||
    word.moves?.some((m) => m.moveId === operation.cutId)
  )
    throw new DocumentError("InvalidOperation", "A cut requires a fresh revision identity.");
  const sameScope =
    wordContentScope(metadata, source, operation.nodeId) ===
    wordContentScope(metadata, source, operation.endNodeId ?? operation.nodeId);
  deleteWordContentRange(metadata, edit, parse, operation, {
    groupId: operation.cutId,
    author: word.trackingAuthor ?? "",
    date: new Date().toISOString(),
  });
  if (
    !word.trackChanges &&
    operation.endNodeId &&
    operation.endNodeId !== operation.nodeId &&
    sameScope
  )
    edit.mergeSourceParagraphMark(operation.nodeId);
}

/** Link a live cut only when its exact retained text equals the new insertion.
 * A cancelled own insertion or a changed clipboard remains an ordinary paste. */
export function linkWordContentMove(
  word: WordMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  cutId: string | undefined,
  before: ReadonlySet<string>,
): void {
  if (!cutId || !word.trackChanges) return;
  const source = parse(),
    order = new Map(source.nodeIds().map((id, index) => [id, index]));
  const from = word.revisions.filter((r) => r.groupId === cutId),
    added = word.revisions.filter((r) => !before.has(r.revisionId ?? r.nodeId)),
    sort = (a: { nodeId: string }, b: { nodeId: string }) =>
      order.get(a.nodeId)! - order.get(b.nodeId)!;
  const selectedText = (entries: typeof from, kinds: readonly string[]) =>
    entries.filter((r) => kinds.includes(r.kind)).sort(sort);
  const deleted = selectedText(from, ["delete", "rowDelete"]),
    inserted = selectedText(added, ["insert"]);
  if (!deleted.length) deleted.push(...selectedText(from, ["paragraphDelete"]));
  if (!inserted.length) inserted.push(...selectedText(added, ["rowInsert"]));
  if (!inserted.length) inserted.push(...selectedText(added, ["paragraphInsert"]));
  const text = (entry: (typeof from)[number]): string =>
    entry.kind.startsWith("paragraph") ? "" : source.getNode(entry.nodeId)!.text;
  if (
    !deleted.length ||
    !inserted.length ||
    from.some((r) => r.moveId || r.author !== (word.trackingAuthor ?? "")) ||
    deleted.map(text).join("") !== inserted.map(text).join("")
  )
    return;
  const moveId = "word-move-" + createOpaqueId(),
    author = word.trackingAuthor ?? "",
    date = new Date().toISOString();
  for (const [kind, revisions] of [
    ["from", from],
    ["to", added.filter((r) => ["insert", "paragraphInsert", "rowInsert"].includes(r.kind))],
  ] as const) {
    revisions.sort(sort);
    const owner = (id: string): string => {
      const node = source.peekNode(id)!;
      if (!["tr", "td", "th"].includes(node.tagName)) return id;
      let root = id;
      for (
        let current: string | null = node.parentNodeId;
        current;
        current = source.peekNode(current)!.parentNodeId
      )
        if (source.peekNode(current)!.tagName === "table") root = current;
      return root;
    };
    const first = source.peekNode(owner(revisions[0]!.nodeId))!,
      last = source.peekNode(owner(revisions.at(-1)!.nodeId))!;
    const start = edit.insertSource(
      first.parentNodeId!,
      `<span data-word-move-range="${kind}-start"></span>`,
      first.nodeId,
    ).rootNodeIds[0]!;
    const siblings = parse().peekNode(last.parentNodeId!)!.childNodeIds,
      next = siblings[siblings.indexOf(last.nodeId) + 1];
    const end = edit.insertSource(
      last.parentNodeId!,
      `<span data-word-move-range="${kind}-end"></span>`,
      next,
    ).rootNodeIds[0]!;
    (word.moves ??= []).push({
      moveId,
      kind,
      sourceId: start,
      startNodeId: start,
      endNodeId: end,
      author,
      date,
    });
    for (const revision of revisions) {
      if (revision.kind === "delete") revision.kind = "moveFrom";
      else if (revision.kind === "insert") revision.kind = "moveTo";
      else if (revision.kind === "paragraphDelete") revision.kind = "paragraphMoveFrom";
      else if (revision.kind === "paragraphInsert") revision.kind = "paragraphMoveTo";
    }
  }
  for (const revision of [...from, ...added]) revision.moveId = moveId;
  // Each retained source interval maps to the exact new insertion interval.
  // Include its actual ancestors/descendants so paragraph and run anchors agree.
  let targetIndex = 0,
    targetOffset = 0;
  for (const revision of deleted) {
    const begin = source.textOffset(revision.nodeId)!,
      length = revision.kind.startsWith("paragraph") ? 0 : source.textLength(revision.nodeId)!;
    let offset = 0;
    while (offset < length) {
      const target = inserted[targetIndex]!,
        remaining = source.textLength(target.nodeId)! - targetOffset;
      if (!remaining) {
        targetIndex++;
        targetOffset = 0;
        continue;
      }
      const amount = Math.min(length - offset, remaining),
        start = begin + offset,
        end = start + amount;
      const ids = new Set<string>(),
        pending = [revision.nodeId];
      while (pending.length) {
        const id = pending.pop()!;
        ids.add(id);
        pending.push(...source.peekNode(id)!.childNodeIds);
      }
      for (
        let id = source.peekNode(revision.nodeId)!.parentNodeId;
        id;
        id = source.peekNode(id)!.parentNodeId
      )
        ids.add(id);
      for (const id of ids) {
        const base = source.textOffset(id)!,
          finish = base + source.textLength(id)!;
        const left = Math.max(base, start),
          right = Math.min(finish, end);
        if (left < right)
          edit.textTransfers.push({
            nodeId: id,
            from: left - base,
            to: right - base,
            targetNodeId: target.nodeId,
            targetOffset: targetOffset + left - start,
            at: edit.textEdits.length,
          });
      }
      offset += amount;
      targetOffset += amount;
    }
  }
}
