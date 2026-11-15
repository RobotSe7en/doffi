import { copyWordSplitRun } from "./word-format-revisions.js";
import { deleteWordTrackedContent } from "./word-content-revisions.js";
import type { SourceStyleEdit } from "./source-styles.js";
import { DocumentError } from "../errors.js";
import type { WordOperation } from "../contracts/word-operations.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import { isTextBoundary } from "../document/text.js";
import { trackWordInlineContent } from "./word-revisions.js";
import { attachWordContent } from "./word-content.js";
import type { WordDocumentMetadata } from "../contracts/document.js";
import { createOpaqueId } from "../identifiers.js";
import { deleteWordContentRange, wordContentScope } from "./word-content-range.js";
const invalid = (message: string): never => {
  throw new DocumentError("InvalidOperation", message);
};
/** Attach a clipboard graph only to newly allocated source nodes. Transport
 * identifiers never resolve against existing document nodes or resource paths. */
export function pasteWordInlineContent(
  metadata: WordDocumentMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  styles: SourceStyleEdit,
  operation: Extract<WordOperation, { op: "pasteWordInlineContent" }>,
): void {
  const word = metadata.word!;
  if (operation.endNodeId && operation.endNodeId !== operation.nodeId) {
    const group = {
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
    const old = new Set(word.revisions.map((item) => item.revisionId ?? item.nodeId));
    pasteWordInlineContent(metadata, edit, parse, styles, {
      ...operation,
      endNodeId: undefined,
      from: offset,
      to: offset,
      expectedText: "",
    });
    for (const revision of word.revisions)
      if (!old.has(revision.revisionId ?? revision.nodeId)) revision.groupId = group.groupId;
    if (!word.trackChanges && sameScope) edit.mergeSourceParagraphMark(operation.nodeId);
    return;
  }
  const source = parse(),
    target = source.getNode(operation.nodeId),
    { from, to, expectedText } = operation;
  if (
    !target ||
    !/^p$|^h[1-6]$/.test(target.tagName) ||
    from > to ||
    to > target.text.length ||
    target.text.slice(from, to) !== expectedText ||
    !isTextBoundary(target.text, from) ||
    !isTextBoundary(target.text, to)
  )
    invalid("Inline Word content requires an exact paragraph range.");
  edit.assertInlineSource(operation.nodeId, operation.html);
  const semantics = operation.semantics ?? {},
    attach = (ids: ReadonlyMap<string, string>): void =>
      attachWordContent(word, edit, parse, ids, semantics, styles);
  const keepEmpty = Object.values(semantics).some((items) => (items?.length ?? 0) > 0);
  if (word.trackChanges) {
    const origin = source.textOffset(operation.nodeId)!;
    const fold = word.revisions.find(
      (revision) =>
        revision.kind === "insert" &&
        revision.author === (word.trackingAuthor ?? "") &&
        source.hasNode(revision.nodeId) &&
        origin + from >= source.textOffset(revision.nodeId)! &&
        origin + to <= source.textOffset(revision.nodeId)! + source.textLength(revision.nodeId)!,
    );
    if (fold && !operation.cutId) {
      trackWordInlineContent(
        word,
        edit,
        parse,
        {
          op: "trackWordInlineContent",
          nodeId: operation.nodeId,
          from,
          to,
          expectedText,
          html: operation.html,
        },
        { inserted: attach, keepEmpty },
      );
      return;
    }
    const group = {
      groupId: "word-paste-" + createOpaqueId(),
      author: word.trackingAuthor ?? "",
      date: new Date().toISOString(),
    };
    const offset = deleteWordTrackedContent(word, edit, parse, operation.nodeId, from, to, group);
    const content = edit.insertSourceInContainer(
      operation.nodeId,
      offset,
      '<ins data-word-run-overlay="true">' + operation.html + "</ins>",
    );
    copyWordSplitRun(word, content.splitIdentities, styles);
    attach(content.identities);
    const id = content.rootNodeIds[0]!;
    if (parse().getNode(id)!.text.length || keepEmpty)
      word.revisions.push({ nodeId: id, revisionId: id, sourceId: id, kind: "insert", ...group });
    else edit.removeSource(id);
    return;
  }
  for (const id of edit.wrapSourceText(operation.nodeId, from, to, "span")) edit.removeSource(id);
  attach(edit.insertSourceAt(operation.nodeId, from, operation.html, undefined, true).identities);
}
