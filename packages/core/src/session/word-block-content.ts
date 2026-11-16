import type { WordContentCommentContext } from "./word-content-comments.js";
import { deleteWordTrackedContent } from "./word-content-revisions.js";
import { createOpaqueId } from "../identifiers.js";
import { DocumentError } from "../errors.js";
import type { WordOperation } from "../contracts/word-operations.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import type { SourceStyleEdit } from "./source-styles.js";
import { isTextBoundary } from "../document/text.js";
import { attachWordContent } from "./word-content.js";
import { copyWordSplitRun } from "./word-format-revisions.js";
import { inheritWordSplitParagraphMarks } from "./word-paragraph-mark.js";
import { normalizeWordListContent } from "./word-list-content.js";
import type { WordDocumentMetadata } from "../contracts/document.js";
import { deleteWordContentRange, wordContentScope } from "./word-content-range.js";
/** Split the receiving paragraph without cloning its existing objects. The
 * imported source graph is allocated independently inside the same candidate. */
export function pasteWordBlockContent(
  metadata: WordDocumentMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  styles: SourceStyleEdit,
  operation: Extract<WordOperation, { op: "pasteWordBlockContent" }>,
  context?: { groupId: string; author: string; date: string },
  comments?: WordContentCommentContext,
): { tailNodeId: string } {
  const word = metadata.word!;
  if (
    word.trackChanges &&
    operation.semantics?.nativeContent?.fragments.some((f) => f.kind === "block")
  )
    throw new DocumentError("UnsupportedFeature", "原生保留区块暂不能记录追踪粘贴，候选未应用。");
  if (operation.endNodeId && operation.endNodeId !== operation.nodeId) {
    const group = context ?? {
      groupId: "word-paste-" + createOpaqueId(),
      author: word.trackingAuthor ?? "",
      date: new Date().toISOString(),
    };
    const source = parse(),
      sameScope =
        wordContentScope(metadata, source, operation.nodeId) ===
        wordContentScope(metadata, source, operation.endNodeId);
    deleteWordContentRange(metadata, edit, parse, operation, group);
    const offset = word.trackChanges
      ? parse().getNode(operation.nodeId)!.text.length
      : operation.from;
    const result = pasteWordBlockContent(
      metadata,
      edit,
      parse,
      styles,
      { ...operation, endNodeId: undefined, from: offset, to: offset, expectedText: "" },
      group,
      comments,
    );
    if (!word.trackChanges && sameScope) {
      edit.textTransfers.push({
        nodeId: result.tailNodeId,
        from: 0,
        to: 0,
        targetNodeId: operation.endNodeId,
        targetOffset: 0,
        at: edit.textEdits.length,
      });
      edit.removeSource(result.tailNodeId);
    }
    return result;
  }
  const before = parse(),
    node = before.getNode(operation.nodeId),
    { from, to, expectedText } = operation;
  if (
    !node ||
    !/^p$|^h[1-6]$/.test(node.tagName) ||
    from > to ||
    to > node.text.length ||
    node.text.slice(from, to) !== expectedText ||
    !isTextBoundary(node.text, from) ||
    !isTextBoundary(node.text, to)
  )
    throw new DocumentError(
      "InvalidOperation",
      "Block Word content requires an exact paragraph range.",
    );
  const author = context?.author ?? word.trackingAuthor ?? "",
    date = context?.date ?? new Date().toISOString(),
    groupId = context?.groupId ?? "word-paste-" + createOpaqueId();
  const splitAt = word.trackChanges
    ? deleteWordTrackedContent(word, edit, parse, operation.nodeId, from, to, {
        groupId,
        author,
        date,
      })
    : from;
  if (!word.trackChanges)
    for (const id of edit.wrapSourceText(operation.nodeId, from, to, "span")) edit.removeSource(id);
  const splitBefore = parse(),
    identities = edit.splitSourceParagraph(operation.nodeId, splitAt),
    tail = identities.get(operation.nodeId)!;
  const inline = new Map([...identities].filter(([id]) => id !== operation.nodeId));
  copyWordSplitRun(word, inline, styles);
  styles.aliasNodeIds(new Map([[operation.nodeId, tail]]));
  for (const revision of word.revisions)
    if (revision.nodeId === operation.nodeId && revision.kind.startsWith("paragraph"))
      revision.nodeId = tail;
  const tabs = word.tabs.find((item) => item.paragraphNodeId === operation.nodeId);
  if (tabs)
    word.tabs.push({
      ...tabs,
      paragraphNodeId: tail,
      stops: tabs.stops.map((stop) => ({ ...stop })),
    });
  inheritWordSplitParagraphMarks(word, splitBefore, edit, parse, styles);
  const histories = new Map<string, string>();
  for (const revision of word.revisions.filter(
    (item) =>
      item.nodeId === operation.nodeId && item.kind === "format" && item.property === "paragraph",
  )) {
    const id = tail + "-paragraph-format";
    histories.set(revision.sourceId, id);
    word.revisions.push({
      ...(JSON.parse(JSON.stringify(revision)) as typeof revision),
      nodeId: tail,
      revisionId: id,
      sourceId: id,
    });
  }
  if (word.nativeContent)
    word.nativeContent.fragments.push(
      ...word.nativeContent.fragments
        .filter((fragment) => fragment.nodeId === operation.nodeId && fragment.kind === "property")
        .map((fragment, index) => ({
          ...(JSON.parse(JSON.stringify(fragment)) as typeof fragment),
          nodeId: tail,
          nativeId: tail + "-split-" + index,
          ...(fragment.revisionSourceId && histories.has(fragment.revisionSourceId)
            ? { revisionSourceId: histories.get(fragment.revisionSourceId)! }
            : {}),
        })),
    );
  const content = edit.insertSource(node.parentNodeId!, operation.html, tail, true);
  const inserted = new Set(content.rootNodeIds);
  attachWordContent(
    metadata,
    edit,
    parse,
    content.identities,
    operation.semantics ?? {},
    styles,
    content.rootNodeIds,
    comments,
  );
  normalizeWordListContent(word, edit, parse(), inserted, styles);
  const after = parse(),
    paragraphs: string[] = [];
  const rows: string[] = [],
    outerParagraphs: string[] = [],
    tables: string[] = [];
  for (const id of after.nodeIds()) {
    let root: string | null = id;
    while (root && !inserted.has(root)) root = after.peekNode(root)!.parentNodeId;
    if (!root) continue;
    const current = after.peekNode(id)!;
    if (["section", "article", "ul", "ol", "li"].includes(current.tagName))
      throw new DocumentError(
        "UnsupportedFeature",
        "This block structure requires its complete Word content graph.",
      );
    if (current.tagName === "tr") rows.push(id);
    if (current.tagName === "table") tables.push(id);
    if (/^p$|^h[1-6]$/.test(current.tagName)) {
      paragraphs.push(id);
      let parent = current.parentNodeId,
        table = false;
      while (parent) {
        if (after.peekNode(parent)!.tagName === "table") {
          table = true;
          break;
        }
        parent = after.peekNode(parent)!.parentNodeId;
      }
      if (!table) outerParagraphs.push(id);
    }
  }
  if (
    !paragraphs.length &&
    !operation.semantics?.nativeContent?.fragments.some((f) => f.kind === "block")
  )
    throw new DocumentError("InvalidOperation", "Word block content requires paragraph structure.");
  for (const id of tables)
    if (!word.tables.some((table) => table.nodeId === id))
      throw new DocumentError(
        "InvalidOperation",
        "A pasted table requires its complete column grid.",
      );
  if (word.trackChanges) {
    for (const id of paragraphs) {
      const wrapper = edit.wrapSourceParagraphContent(id, "ins");
      if (wrapper)
        word.revisions.push({
          nodeId: wrapper,
          sourceId: wrapper,
          revisionId: wrapper,
          kind: "insert",
          author,
          date,
          groupId,
        });
    }
    for (const id of rows) {
      const revisionId = id + "-paste-row";
      word.revisions.push({
        nodeId: id,
        sourceId: revisionId,
        revisionId,
        kind: "rowInsert",
        author,
        date,
        groupId,
      });
    }
    for (const id of [operation.nodeId, ...outerParagraphs]) {
      const revisionId = id + "-paste-mark-" + groupId;
      word.revisions.push({
        nodeId: id,
        sourceId: revisionId,
        revisionId,
        kind: "paragraphInsert",
        author,
        date,
        groupId,
      });
    }
  }
  return { tailNodeId: tail };
}
