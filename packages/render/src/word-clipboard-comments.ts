import { DocumentError, isDocumentTextBoundary } from "@doffi/core";
import type { Anchor, Annotation, ContentDocumentModel, WordContentSemantics } from "@doffi/core";
import { clipboardNodeText } from "./clipboard-node-text.js";

export type WordClipboardSourceRanges = ReadonlyMap<
  string,
  { from: number; to: number; inexact?: boolean }
>;

/** Intersect source intervals with selected transport nodes. Ancestors absent
 * from inline HTML are resolved through their actual source positions. */
export function copyWordClipboardComments(
  roots: readonly HTMLElement[],
  model: ContentDocumentModel,
  complete: ReadonlySet<string> | undefined,
  selected?: WordClipboardSourceRanges,
): Pick<WordContentSemantics, "annotations" | "comments"> {
  const snapshot = model.snapshot(),
    word = snapshot.metadata.word!,
    positions = new Map<string, { from: number; story: string }>();
  if (!snapshot.annotations.length) return {};
  // childNodeIds contains elements only. Count real Text nodes too, including
  // literals before/after a field, rather than summing element siblings.
  const sourceDocument = new DOMParser().parseFromString(snapshot.html, "text/html");
  const index = (node: Node, from: number, story: string): number => {
    if (node.nodeType === 3) return node.textContent?.length ?? 0;
    const element = node.nodeType === 1 ? (node as HTMLElement) : undefined,
      id = element?.dataset.nodeId;
    if (element?.hasAttribute("data-office-story")) {
      from = 0;
      story = id!;
    }
    if (id) positions.set(id, { from, story });
    if (element?.matches("img,br")) return 1;
    let length = 0;
    for (const child of node.childNodes) length += index(child, from + length, story);
    return length;
  };
  index(sourceDocument.body, 0, "document");
  const position = (id: string): { from: number; story: string } =>
    positions.get(id) ??
    (() => {
      throw new DocumentError("InvalidDocument", "复制批注缺少其源位置。");
    })();
  const fields = new Set(
    word.fields.filter((field) => !field.endNodeId).map((field) => field.nodeId),
  );
  const slices: {
    id: string;
    from: number;
    to: number;
    text: string;
    story: string;
    element: HTMLElement;
    offsets: (at: number) => number | undefined;
  }[] = [];
  for (const [index, root] of roots.entries())
    for (const element of root.querySelectorAll<HTMLElement>("[data-node-id]")) {
      const id = element.dataset.nodeId!,
        source = model.getNode(id);
      if (!source) continue;
      const range = index === 0 ? selected?.get(id) : undefined,
        text = clipboardNodeText(element),
        from = range?.from ?? 0,
        to = range?.to ?? source.text.length;
      if (!range && index === 0 && complete && !complete.has(id)) continue;
      const base = position(id),
        replacements: { from: number; to: number; text: string }[] = [];
      if (source.text.slice(from, to) !== text) {
        for (const child of [element, ...element.querySelectorAll<HTMLElement>("[data-node-id]")]) {
          const fieldId = child.dataset.nodeId!;
          if (
            !fields.has(fieldId) ||
            (complete && !complete.has(fieldId) && !selected?.get(fieldId)?.inexact)
          )
            continue;
          const original = model.getNode(fieldId)!,
            start = position(fieldId).from - base.from;
          const displayed = clipboardNodeText(child);
          if (start >= from && start + original.text.length <= to && displayed !== original.text)
            replacements.push({
              from: start - from,
              to: start - from + original.text.length,
              text: displayed,
            });
        }
        replacements.sort((a, b) => a.from - b.from);
        let expected = source.text.slice(from, to);
        for (const replacement of [...replacements].reverse())
          expected =
            expected.slice(0, replacement.from) + replacement.text + expected.slice(replacement.to);
        if (expected !== text) continue;
      }
      if (!isDocumentTextBoundary(source.text, from) || !isDocumentTextBoundary(source.text, to))
        throw new DocumentError("InvalidOperation", "批注复制不能拆开 UTF-16 字符或变体序列。");
      if (slices.some((slice) => slice.element.contains(element))) continue;
      const offsets = (at: number): number | undefined => {
        let delta = 0;
        for (const replacement of replacements) {
          if (at <= replacement.from) break;
          if (at < replacement.to) return;
          delta += replacement.text.length - replacement.to + replacement.from;
        }
        return at + delta;
      };
      slices.push({
        id,
        from: base.from + from,
        to: base.from + to,
        text,
        story: base.story,
        element,
        offsets,
      });
    }
  const objects = (ids: readonly string[]): string[] => {
    const result = new Set<string>();
    for (const id of ids) {
      const node = model.getNode(id);
      if (!node) continue;
      const base = position(id),
        end = base.from + node.text.length;
      for (const slice of slices) {
        if (slice.story !== base.story) continue;
        // Object anchors retain selected concrete descendants, not an absent
        // or partially copied group/container from the source document.
        for (const element of [
          slice.element,
          ...slice.element.querySelectorAll<HTMLElement>("[data-node-id]"),
        ]) {
          const targetId = element.dataset.nodeId!,
            target = model.getNode(targetId)!;
          const at = position(targetId);
          if (at.from < base.from || at.from + target.text.length > end) continue;
          let ancestor: string | null = targetId;
          while (ancestor && ancestor !== id) ancestor = model.getNode(ancestor)!.parentNodeId;
          if (!ancestor) continue;
          if (
            (complete ? complete.has(targetId) : clipboardNodeText(element) === target.text) ||
            (selected?.has(targetId) && slices.some((s) => s.id === targetId)) ||
            roots.slice(1).some((root) => root.contains(element))
          ) {
            if (
              ![...result].some((existing) => {
                let current: string | null = targetId;
                while (current) {
                  if (current === existing) return true;
                  current = model.getNode(current)!.parentNodeId;
                }
                return false;
              })
            )
              result.add(targetId);
          }
        }
      }
    }
    return [...result];
  };
  const annotations: Annotation[] = [],
    comments: NonNullable<WordContentSemantics["comments"]> = [];
  for (const original of snapshot.annotations) {
    const anchor = original.anchor;
    let mapped: Anchor | undefined;
    const recovery = new Set<string>();
    if (anchor.kind === "text" || anchor.kind === "text-set") {
      const ranges = (anchor.kind === "text" ? [anchor] : anchor.ranges).flatMap((range) => {
        const node = model.getNode(range.nodeId);
        if (!node || node.text.slice(range.from, range.to) !== range.quote) return [];
        const base = position(range.nodeId),
          from = base.from + range.from,
          to = base.from + range.to;
        return slices
          .filter((slice) => slice.story === base.story)
          .flatMap((slice) => {
            const begin = Math.max(from, slice.from),
              end = Math.min(to, slice.to),
              point =
                from === to &&
                from >= slice.from &&
                (from < slice.to || (complete?.has(slice.id) && from === slice.to));
            if (begin >= end && !point) return [];
            const left = slice.offsets(begin - slice.from),
              right = slice.offsets(end - slice.from);
            if (left === undefined || right === undefined) {
              recovery.add(slice.id);
              return [];
            }
            return [
              { nodeId: slice.id, from: left, to: right, quote: slice.text.slice(left, right) },
            ];
          });
      });
      if (ranges.length)
        mapped =
          ranges.length === 1 ? { kind: "text", ...ranges[0]! } : { kind: "text-set", ranges };
      if (recovery.size) {
        for (const range of ranges) recovery.add(range.nodeId);
        mapped =
          recovery.size === 1
            ? { kind: "node", nodeId: [...recovery][0]! }
            : { kind: "node-set", nodeIds: [...recovery] };
      }
    } else if (anchor.kind !== "cells") {
      const ids = objects(anchor.kind === "node" ? [anchor.nodeId] : anchor.nodeIds);
      if (ids.length)
        mapped =
          ids.length === 1 ? { kind: "node", nodeId: ids[0]! } : { kind: "node-set", nodeIds: ids };
    }
    if (!mapped) continue;
    const copy = JSON.parse(JSON.stringify(original)) as Annotation;
    copy.anchor = mapped;
    if (recovery.size) copy.status = "needs_reanchor";
    annotations.push(copy);
    const binding = word.comments?.find((b) => b.annotationId === original.annotationId);
    if (binding) {
      const { durableId: _durable, referenceNodeId, ...rest } = binding;
      const exists =
        referenceNodeId &&
        roots.some((root) =>
          [...root.querySelectorAll<HTMLElement>("[data-node-id]")].some(
            (node) => node.dataset.nodeId === referenceNodeId,
          ),
        );
      comments.push({
        ...rest,
        ...(exists ? { referenceNodeId } : {}),
        replies: binding.replies.map(({ replyId, parentReplyId }) => ({
          replyId,
          ...(parentReplyId ? { parentReplyId } : {}),
        })),
      });
    }
  }
  return {
    ...(annotations.length ? { annotations } : {}),
    ...(comments.length ? { comments } : {}),
  };
}
