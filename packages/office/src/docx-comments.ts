import type { Annotation, TextRange, WordMetadata } from "@doffi/core";
import type { OpcPackage } from "./ooxml/package.js";
import { W, attr, child, elements, check, fail, unsupported } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import type { DocxNativeImport } from "./docx-native-import.js";
export const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
export const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
export const CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
export const CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
export interface WordCommentState {
  source: XmlNode;
  bodyNodeId: string;
  ranges: TextRange[];
  point?: string;
  resolved?: boolean;
  parentId?: string;
  durableId?: string;
  dateUtc?: string;
  followUp?: boolean;
}
function attributes(node: XmlNode, names: string[], ns: string): void {
  for (const key of node.attrs.keys())
    if (!names.includes(key.slice(ns.length + 1)) || !key.startsWith(ns + " "))
      unsupported(node, "An unmodeled comment extension attribute is preserved.", key);
}
function hex(value: string | undefined, node: XmlNode, durable = false): string {
  if (
    !value ||
    !/^[0-9a-f]{8}$/i.test(value) ||
    (durable && (parseInt(value, 16) <= 0 || parseInt(value, 16) >= 0x7fffffff))
  )
    fail(node, "A native comment identity is invalid.");
  return value.toUpperCase();
}
function bool(value: string | undefined, node: XmlNode): boolean {
  if (value === undefined) return false;
  if (!["0", "1", "true", "false", "on", "off"].includes(value))
    fail(node, "A native comment Boolean is invalid.");
  return ["1", "true", "on"].includes(value);
}
/** Extension parts identify the last paragraph of each classic message body. */
export function readWordCommentExtensions(
  opc: OpcPackage,
  main: string,
  comments: Map<string, WordCommentState>,
  word: WordMetadata,
  rootNodeId: string,
  native?: DocxNativeImport,
): void {
  const byParagraph = new Map<
      string,
      {
        id: string;
        comment: WordCommentState;
      }
    >(),
    byDurable = new Map<string, WordCommentState>();
  for (const [id, comment] of comments) {
    const stack = [comment.source];
    let paragraph: XmlNode | undefined;
    while (stack.length) {
      const node = stack.pop()!;
      if (node.uri === W && node.name === "p") paragraph = node;
      for (const item of [...node.content].reverse())
        if (typeof item !== "string") stack.push(item);
    }
    const raw = attr(paragraph, "paraId", W14);
    if (raw !== undefined) {
      const key = hex(raw, paragraph!);
      if (byParagraph.has(key)) fail(comment.source, "A comment paragraph identity is repeated.");
      byParagraph.set(key, { id, comment });
    }
  }
  const read = (type: string, ns: string, name: string): XmlNode | undefined => {
    const part = opc.related(main, type);
    if (!part) return;
    const root = opc.read(part);
    if (root.uri !== ns || root.name !== name) fail(root, "A comment extension root is invalid.");
    return root;
  };
  const extended = read("commentsExtended", W15, "commentsEx");
  if (extended) {
    check(extended, ["commentEx"], W15);
    const seen = new Set<string>();
    for (const status of elements(extended, "commentEx", W15)) {
      check(status, [], W15);
      attributes(status, ["paraId", "paraIdParent", "done"], W15);
      const id = hex(attr(status, "paraId", W15), status),
        entry = byParagraph.get(id);
      if (!entry || seen.has(id))
        fail(status, "A comment status has an unknown or repeated paragraph.");
      seen.add(id);
      native?.bind(status, { "data-node-id": entry.comment.bodyNodeId });
      entry.comment.resolved = bool(attr(status, "done", W15), status);
      const parent = attr(status, "paraIdParent", W15);
      if (parent !== undefined) {
        const target = byParagraph.get(hex(parent, status));
        if (!target || target === entry)
          fail(status, "A reply parent is missing or self-referential.");
        entry.comment.parentId = target.id;
      }
    }
  }
  const ids = read("commentsIds", CID, "commentsIds");
  if (ids) {
    check(ids, ["commentId"], CID);
    const seen = new Set<string>();
    for (const item of elements(ids, "commentId", CID)) {
      check(item, [], CID);
      attributes(item, ["paraId", "durableId"], CID);
      const id = hex(attr(item, "paraId", CID), item),
        entry = byParagraph.get(id),
        durable = hex(attr(item, "durableId", CID), item, true);
      if (!entry || seen.has(id) || byDurable.has(durable))
        fail(item, "A durable comment identity is unknown or repeated.");
      seen.add(id);
      entry.comment.durableId = durable;
      byDurable.set(durable, entry.comment);
      native?.bind(item, { "data-node-id": entry.comment.bodyNodeId });
    }
  }
  const extensible = read("commentsExtensible", CEX, "commentsExtensible");
  if (extensible) {
    check(extensible, ["commentExtensible"], CEX);
    const seen = new Set<string>();
    for (const item of elements(extensible, "commentExtensible", CEX)) {
      check(item, [], CEX);
      attributes(item, ["durableId", "dateUtc", "intelligentPlaceholder"], CEX);
      const id = hex(attr(item, "durableId", CEX), item, true),
        comment = byDurable.get(id);
      if (!comment || seen.has(id))
        fail(item, "An extensible comment has an unknown or repeated durable identity.");
      seen.add(id);
      native?.bind(item, { "data-node-id": comment.bodyNodeId });
      const date = attr(item, "dateUtc", CEX);
      if (date !== undefined) {
        if (!date.endsWith("Z") || !Number.isFinite(Date.parse(date)))
          fail(item, "A comment UTC timestamp is invalid.");
        comment.dateUtc = date;
      }
      if (attr(item, "intelligentPlaceholder", CEX) !== undefined) {
        if (comment.parentId) fail(item, "A reply cannot be a follow-up placeholder.");
        comment.followUp = bool(attr(item, "intelligentPlaceholder", CEX), item);
      }
    }
  }
  const people = read("people", W15, "people");
  if (people) {
    check(people, ["person"], W15);
    const names = new Set<string>();
    word.people = [];
    for (const person of elements(people, "person", W15)) {
      check(person, ["presenceInfo"], W15);
      attributes(person, ["author"], W15);
      const author = attr(person, "author", W15);
      if (author === undefined || names.has(author))
        fail(person, "A comment person has a missing or repeated author.");
      names.add(author);
      const presence = child(person, "presenceInfo", W15);
      let identity:
        | {
            providerId: string;
            userId: string;
          }
        | undefined;
      if (presence) {
        check(presence, [], W15);
        attributes(presence, ["providerId", "userId"], W15);
        const providerId = attr(presence, "providerId", W15),
          userId = attr(presence, "userId", W15);
        if (providerId === undefined || userId === undefined)
          fail(presence, "A comment presence identity is incomplete.");
        identity = { providerId, userId };
      }
      word.people.push({ author, ...(identity ? { presence: identity } : {}) });
      const comment = [...comments.values()].find(
        (comment) => attr(comment.source, "author") === author,
      );
      native?.bind(person, { "data-node-id": comment?.bodyNodeId ?? rootNodeId });
    }
  }
}
export function wordCommentAnnotations(
  comments: Map<string, WordCommentState>,
  word: WordMetadata,
  documentId: string,
): Annotation[] {
  const roots = new Map<string, string>();
  for (const [id, comment] of comments) {
    if (roots.has(id)) continue;
    const path: string[] = [],
      seen = new Set<string>();
    let current = id;
    while (!roots.has(current)) {
      if (seen.has(current)) fail(comment.source, "Native reply parents form a cycle.");
      seen.add(current);
      path.push(current);
      const entry = comments.get(current);
      if (!entry) fail(comment.source, "A native reply parent is missing.");
      if (!entry.parentId) {
        roots.set(current, current);
        break;
      }
      current = entry.parentId;
    }
    const root = roots.get(current)!;
    for (const key of path) roots.set(key, root);
  }
  const identity = (comment: WordCommentState) => {
    const name = attr(comment.source, "author"),
      initials = attr(comment.source, "initials"),
      date = comment.dateUtc ?? attr(comment.source, "date");
    return {
      ...(name !== undefined
        ? { author: { name, ...(initials !== undefined ? { initials } : {}) } }
        : {}),
      ...(date !== undefined ? { date } : {}),
    };
  };
  const annotations: Annotation[] = [],
    byId = new Map<string, Annotation>();
  word.comments = [];
  for (const [id, comment] of comments) {
    if (roots.get(id) !== id) continue;
    if (!comment.ranges.length && !comment.point)
      fail(comment.source, "An imported comment has no anchor.");
    const annotation: Annotation = {
      annotationId: "comment-" + id,
      documentId,
      baseRevision: 0,
      status: comment.resolved ? "resolved" : "open",
      bodyNodeId: comment.bodyNodeId,
      sourceId: id,
      ...identity(comment),
      anchor:
        comment.ranges.length === 1
          ? { kind: "text", ...comment.ranges[0]! }
          : comment.ranges.length
            ? { kind: "text-set", ranges: comment.ranges }
            : { kind: "node", nodeId: comment.point! },
    };
    annotations.push(annotation);
    byId.set(id, annotation);
    word.comments.push({
      annotationId: annotation.annotationId,
      ...(comment.point ? { referenceNodeId: comment.point } : {}),
      ...(comment.durableId ? { durableId: comment.durableId } : {}),
      ...(comment.followUp !== undefined ? { followUp: comment.followUp } : {}),
      replies: [],
    });
  }
  for (const [id, comment] of comments) {
    const root = roots.get(id)!;
    if (root === id) continue;
    const annotation = byId.get(root)!,
      replyId = "reply-" + id;
    (annotation.replies ??= []).push({
      replyId,
      sourceId: id,
      bodyNodeId: comment.bodyNodeId,
      status: comment.resolved ? "resolved" : "open",
      ...identity(comment),
    });
    word.comments
      .find((binding) => binding.annotationId === annotation.annotationId)!
      .replies.push({
        replyId,
        ...(comment.durableId ? { durableId: comment.durableId } : {}),
        ...(comment.parentId !== root ? { parentReplyId: "reply-" + comment.parentId } : {}),
      });
  }
  return annotations;
}
