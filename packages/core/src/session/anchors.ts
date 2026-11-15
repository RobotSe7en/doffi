import type { Anchor, Annotation, TextRange } from "../contracts/annotations.js";
import { validTextRanges } from "../document/annotations.js";
import type { ParsedHtml } from "../document/html.js";
import { isTextBoundary } from "../document/text.js";
import type { DeepReadonly, TextOffsetMapping, TextRangeTransfer } from "../runtime.js";
import { unsupported } from "../document/validation.js";

export function annotationNodes(annotation: DeepReadonly<Annotation>): readonly string[] {
  return anchorNodes(annotation.anchor);
}
export function anchorNodes(anchor: DeepReadonly<Anchor>): readonly string[] {
  if (anchor.kind === "cells") unsupported("Cell anchors require a workbook editing session.");
  return anchor.kind === "text" || anchor.kind === "node"
    ? [anchor.nodeId]
    : anchor.kind === "text-set"
      ? anchor.ranges.map((range) => range.nodeId)
      : anchor.kind === "region"
        ? [anchor.surfaceId, ...anchor.nodeIds]
        : anchor.nodeIds;
}

function mapRange(
  range: TextRange,
  edits: readonly TextOffsetMapping[],
  after?: ParsedHtml,
): TextRange | undefined {
  let { from, to } = range;
  for (const edit of edits) {
    if (edit.nodeId !== range.nodeId) continue;
    const delta = edit.insertedLength - (edit.to - edit.from);
    if (from === to) {
      // Point anchors attach to the following source; deleted interior points are uncertain.
      if (edit.from < from && from < edit.to) return undefined;
      if (edit.to <= from) from = to = from + delta;
      else if (edit.from === from) from = to = edit.from + edit.insertedLength;
    } else if (edit.to <= from) {
      from += delta;
      to += delta;
    } else if (edit.from >= to) {
      continue;
    } else if (edit.from <= from && edit.to >= to) {
      return undefined;
    } else {
      // Partial replacements remain selected; insertions at either edge stay outside.
      if (edit.from < from) from = edit.from;
      to = edit.to < to ? to + delta : edit.from + edit.insertedLength;
    }
  }
  if (!after) return { ...range, from, to };
  const text = after.getNode(range.nodeId)?.text;
  if (
    text === undefined ||
    from > to ||
    to > text.length ||
    !isTextBoundary(text, from) ||
    !isTextBoundary(text, to)
  )
    return undefined;
  return { nodeId: range.nodeId, from, to, quote: text.slice(from, to) };
}

function transferRange(range: TextRange, move: TextRangeTransfer): TextRange[] {
  if (range.nodeId !== move.nodeId) return [range];
  const { from, to } = range;
  const moved = (from: number, to: number): TextRange => ({
    ...range,
    nodeId: move.targetNodeId,
    from: move.targetOffset + from - move.from,
    to: move.targetOffset + to - move.from,
  });
  if (from === to) return [from >= move.from && from <= move.to ? moved(from, to) : range];
  const begin = Math.max(from, move.from),
    end = Math.min(to, move.to);
  if (begin >= end) return [range];
  return [
    ...(from < begin ? [{ ...range, to: begin }] : []),
    moved(begin, end),
    ...(end < to ? [{ ...range, from: end }] : []),
  ];
}

function mapTransferredRanges(
  ranges: readonly TextRange[],
  before: ParsedHtml,
  after: ParsedHtml,
  edits: {
    readonly textEdits: readonly TextOffsetMapping[];
    readonly textTransfers?: readonly TextRangeTransfer[];
    readonly replacedTextNodes: ReadonlySet<string>;
    readonly structuralTextNodes: ReadonlySet<string>;
  },
): TextRange[] | undefined {
  let current = [...ranges];
  for (const range of ranges) if (edits.replacedTextNodes.has(range.nodeId)) return;
  for (let index = 0; index <= edits.textEdits.length; index++) {
    for (const move of edits.textTransfers ?? [])
      if (move.at === index) current = current.flatMap((range) => transferRange(range, move));
    const edit = edits.textEdits[index];
    if (!edit) continue;
    if (
      !edits.structuralTextNodes.has(edit.nodeId) &&
      before.getNode(edit.nodeId)?.text === after.getNode(edit.nodeId)?.text &&
      !edits.textTransfers?.some(
        (move) => move.nodeId === edit.nodeId || move.targetNodeId === edit.nodeId,
      )
    )
      continue;
    const mapped = current.map((range) => mapRange(range, [edit]));
    if (mapped.some((range) => !range)) return;
    current = mapped as TextRange[];
  }
  const final = current.map((range) => mapRange(range, [], after));
  if (final.some((range) => !range)) return;
  return (final as TextRange[]).sort(
    (a, b) => after.textOffset(a.nodeId)! + a.from - after.textOffset(b.nodeId)! - b.from,
  );
}

/** Map a complete annotation or retain its complete original anchor for manual recovery. */
export function mapAnnotations(
  annotations: Annotation[],
  before: ParsedHtml,
  after: ParsedHtml,
  edits: {
    readonly textEdits: readonly TextOffsetMapping[];
    readonly textTransfers?: readonly TextRangeTransfer[];
    readonly replacedTextNodes: ReadonlySet<string>;
    readonly structuralTextNodes: ReadonlySet<string>;
  },
  explicit: ReadonlySet<string>,
  regionChanged?: (surfaceId: string, anchor: Extract<Anchor, { kind: "region" }>) => boolean,
): void {
  const layoutChanged = before.html !== after.html;
  const byNode = new Map<string, TextOffsetMapping[]>();
  for (const edit of edits.textEdits) {
    const items = byNode.get(edit.nodeId) ?? [];
    items.push(edit);
    byNode.set(edit.nodeId, items);
  }
  for (const annotation of annotations) {
    if (annotation.status === "needs_reanchor" || explicit.has(annotation.annotationId)) continue;
    const anchor = annotation.anchor;
    const transferredText =
      (anchor.kind === "text" || anchor.kind === "text-set") && !!edits.textTransfers?.length;
    let invalid =
      (!transferredText && annotationNodes(annotation).some((id) => !after.hasNode(id))) ||
      (anchor.kind === "region" &&
        (regionChanged ? regionChanged(anchor.surfaceId, anchor) : layoutChanged));
    if (!invalid && (anchor.kind === "text" || anchor.kind === "text-set") && layoutChanged) {
      const ranges = anchor.kind === "text" ? [anchor] : anchor.ranges;
      // Activating recovery data without a new anchor must not manufacture a matching quote.
      if (!validTextRanges(ranges, before)) continue;
      const mapped = transferredText
        ? mapTransferredRanges(ranges, before, after, edits)
        : ranges.map((range) =>
            edits.replacedTextNodes.has(range.nodeId)
              ? undefined
              : mapRange(
                  range,
                  !edits.structuralTextNodes.has(range.nodeId) &&
                    before.getNode(range.nodeId)?.text === after.getNode(range.nodeId)?.text
                    ? []
                    : (byNode.get(range.nodeId) ?? []),
                  after,
                ),
          );
      invalid =
        !mapped ||
        mapped.some((range) => range === undefined) ||
        !validTextRanges(mapped as TextRange[], after);
      if (!invalid)
        annotation.anchor =
          anchor.kind === "text" && mapped!.length === 1
            ? { kind: "text", ...mapped![0]! }
            : { kind: "text-set", ranges: mapped as TextRange[] };
    }
    if (invalid) annotation.status = "needs_reanchor";
  }
}
