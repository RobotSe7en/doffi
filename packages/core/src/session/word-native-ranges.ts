import type { WordMetadata } from "../contracts/word.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";

/** Drop a range when either owner was deleted; opaque bytes stay immutable and
 * the Office adapter omits inactive controls on output. Undo restores both. */
export function pruneWordNativeRanges(
  word: WordMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
): void {
  const native = word.nativeContent;
  if (!native) return;
  const source = parse(),
    counts = new Map<string, number>();
  for (const fragment of native.fragments)
    if (fragment.nodeId && source.hasNode(fragment.nodeId))
      for (const boundary of fragment.rangeBoundaries ?? [])
        if (boundary.active !== false)
          counts.set(boundary.rangeId, (counts.get(boundary.rangeId) ?? 0) + 1);
  for (const fragment of native.fragments) {
    for (const boundary of fragment.rangeBoundaries ?? [])
      if (counts.get(boundary.rangeId) !== 2) boundary.active = false;
    if (
      fragment.rangeBoundary &&
      fragment.nodeId &&
      source.hasNode(fragment.nodeId) &&
      fragment.rangeBoundaries?.every((b) => b.active === false)
    )
      edit.removeSource(fragment.nodeId);
  }
}
