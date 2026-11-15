import { createOpaqueId } from "../identifiers.js";
import { DocumentError } from "../errors.js";
import type { WordMetadata } from "../contracts/word.js";
import type { WordOperation } from "../contracts/word-operations.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import type { SourceStyleEdit } from "./source-styles.js";
import { isTextBoundary } from "../document/text.js";
import { attachWordContent } from "./word-content.js";
import { copyWordSplitRun } from "./word-format-revisions.js";
import { inheritWordSplitParagraphMarks } from "./word-paragraph-mark.js";
import { normalizeWordListContent } from "./word-list-content.js";
/** Split the receiving paragraph without cloning its existing objects. The
 * imported source graph is allocated independently inside the same candidate. */
export function pasteWordBlockContent(
  word: WordMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  styles: SourceStyleEdit,
  operation: Extract<WordOperation, { op: "pasteWordBlockContent" }>,
): void {
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
  const under = (source: ParsedHtml, id: string, root: string): boolean => {
    for (
      let current: string | null = id;
      current;
      current = source.peekNode(current)?.parentNodeId ?? null
    )
      if (current === root) return true;
    return false;
  };
  const origin = before.textOffset(operation.nodeId)!,
    start = origin + from,
    end = origin + to;
  if (
    word.revisions.some((revision) => {
      if (
        !["insert", "delete", "moveFrom", "moveTo"].includes(revision.kind) ||
        !under(before, revision.nodeId, operation.nodeId)
      )
        return false;
      const begin = before.textOffset(revision.nodeId)!,
        finish = begin + before.textLength(revision.nodeId)!;
      return from === to ? begin < start && start < finish : start < finish && end > begin;
    })
  )
    throw new DocumentError(
      "UnsupportedFeature",
      "Block content over an existing text revision requires structural review.",
    );
  const author = word.trackingAuthor ?? "",
    date = new Date().toISOString(),
    groupId = "word-paste-" + createOpaqueId(),
    removed = edit.wrapSourceText(operation.nodeId, from, to, word.trackChanges ? "del" : "span");
  if (word.trackChanges)
    for (const id of removed) {
      edit.attributes(id, { "data-word-run-overlay": "true" });
      word.revisions.push({
        nodeId: id,
        revisionId: id,
        sourceId: id,
        kind: "delete",
        author,
        date,
        groupId,
      });
    }
  else for (const id of removed) edit.removeSource(id);
  const splitBefore = parse(),
    splitAt = word.trackChanges ? to : from,
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
    word,
    edit,
    parse,
    content.identities,
    operation.semantics ?? {},
    styles,
    content.rootNodeIds,
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
  if (!paragraphs.length)
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
      const revisionId = id + "-paste-mark";
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
}
