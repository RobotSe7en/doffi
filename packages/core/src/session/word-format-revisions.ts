import type { WordMetadata } from "../contracts/word.js";
import type { WordRevision } from "../contracts/word-revision.js";
import { WORD_ROW_GAP_WIDTH_ATTRIBUTES } from "../contracts/word-revision.js";
import { wordRowOmissions } from "../document/word-row-omissions.js";
import type { ParsedHtml } from "../document/html.js";
import type { DocumentEdit } from "../document/edit.js";
import type { SourceStyleEdit } from "./source-styles.js";
/** Native property histories retain the first value and the latest author. */
export function recordWordFormat(
  word: WordMetadata,
  source: ParsedHtml,
  nodeId: string,
  property: "paragraph" | "table" | "row" | "cell",
  index?: Map<string, WordRevision>,
): void {
  if (!word.trackChanges) return;
  const previous = index
    ? index.get(nodeId + "\n" + property)
    : word.revisions.find(
        (revision) =>
          revision.nodeId === nodeId &&
          revision.kind === "format" &&
          revision.property === property,
      );
  if (previous) {
    previous.author = word.trackingAuthor ?? "";
    previous.date = new Date().toISOString();
    return;
  }
  const node = source.peekNode(nodeId)!,
    attributes: Record<string, string | null> = {};
  for (const name of [
    "style",
    "class",
    "lang",
    "data-word-style",
    "data-word-line-rule",
    "data-word-numbering",
    "data-word-numbering-off",
    "data-word-level",
    "data-word-outline-level",
    "data-word-text-direction",
    "data-word-repeat-header",
  ])
    attributes[name] = node.attributes[name] ?? null;
  const before: NonNullable<WordRevision["before"]> = {
    attributes,
    tagName: node.tagName as NonNullable<WordRevision["before"]>["tagName"],
    inheritNativeProperties: true,
  };
  if (property === "cell") attributes.colspan = node.attributes.colspan ?? null;
  if (property === "row") {
    before.rowOmissions = wordRowOmissions(source, nodeId);
    for (const key of WORD_ROW_GAP_WIDTH_ATTRIBUTES) attributes[key] = node.attributes[key] ?? null;
  }
  if (["paragraph", "row", "cell"].includes(property))
    before.conditionalFormat =
      word.conditionalFormats?.find((item) => item.nodeId === nodeId)?.format ?? null;
  if (property === "table") {
    const table = word.tables.find((table) => table.nodeId === nodeId);
    before.tableStyleOptions = Object.fromEntries(
      ["lookMask", "rowBandSize", "columnBandSize"].flatMap((key) =>
        table && Reflect.get(table, key) !== undefined ? [[key, Reflect.get(table, key)]] : [],
      ),
    );
  }
  if (property === "paragraph") {
    const tabs = word.tabs.find((tab) => tab.paragraphNodeId === nodeId);
    before.tabs = tabs
      ? { defaultInterval: tabs.defaultInterval, stops: tabs.stops.map((stop) => ({ ...stop })) }
      : { defaultInterval: word.tabs[0]?.defaultInterval ?? 36, stops: [] };
  }
  const id = nodeId + "-format-" + property;
  const revision: WordRevision = {
    nodeId,
    sourceId: id,
    revisionId: id,
    kind: "format",
    property,
    before,
    author: word.trackingAuthor ?? "",
    date: new Date().toISOString(),
  };
  word.revisions.push(revision);
  index?.set(nodeId + "\n" + property, revision);
}
/** A formatting-owner split partitions existing histories, rather than
 * treating the surviving right-hand text as newly inserted content. */
export function copyWordSplitRun(
  word: WordMetadata,
  identities: ReadonlyMap<string, string>,
  styles: SourceStyleEdit,
): void {
  styles.aliasNodeIds(identities);
  const copies = word.revisions
    .filter((item) => identities.has(item.nodeId))
    .map((item) => {
      const nodeId = identities.get(item.nodeId)!,
        id = nodeId + "-format";
      return {
        ...(JSON.parse(JSON.stringify(item)) as WordRevision),
        nodeId,
        revisionId: id,
        sourceId: id,
      };
    });
  if (word.nativeContent) {
    const native = word.nativeContent,
      oldRevisions = new Map(
        word.revisions
          .filter((item) => identities.has(item.nodeId))
          .map((item) => [item.nodeId + "\n" + item.sourceId, item]),
      );
    native.fragments.push(
      ...native.fragments
        .filter((item) => item.nodeId && identities.has(item.nodeId) && item.kind === "property")
        .map((item, index) => {
          const nodeId = identities.get(item.nodeId!)!,
            previous = item.revisionSourceId
              ? oldRevisions.get(item.nodeId + "\n" + item.revisionSourceId)
              : undefined;
          return {
            ...(JSON.parse(JSON.stringify(item)) as typeof item),
            nodeId,
            nativeId: nodeId + "-native-" + index,
            ...(previous
              ? {
                  revisionSourceId: copies.find(
                    (copy) => copy.nodeId === nodeId && copy.property === previous.property,
                  )!.sourceId,
                }
              : {}),
          };
        }),
    );
  }
  word.hyperlinks.push(
    ...word.hyperlinks
      .filter((item) => identities.has(item.nodeId))
      .map((item) => ({ ...item, nodeId: identities.get(item.nodeId)! })),
  );
  word.revisions.push(...copies);
}
/** Partition pending run histories at the new selection, so a partial edit
 * changes neither the author nor the review group of adjacent text. */
export function prepareWordRunFormat(
  word: WordMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  styles: SourceStyleEdit,
  range: {
    nodeId: string;
    from: number;
    to: number;
    expectedText: string;
  },
): {
  nodeId: string;
  from: number;
  to: number;
  expectedText: string;
  preferred: Set<string>;
} {
  const original = parse(),
    base = original.textOffset(range.nodeId)!,
    start = base + range.from,
    end = base + range.to;
  let target = original.peekNode(range.nodeId)!;
  for (let parent = target; parent; ) {
    if (/^p$|^h[1-6]$|^li$/.test(parent.tagName)) {
      target = parent;
      break;
    }
    if (!parent.parentNodeId) break;
    parent = original.peekNode(parent.parentNodeId)!;
  }
  for (const boundary of [end, start])
    for (const revision of [...word.revisions]) {
      if (revision.kind !== "format" || revision.property !== "run") continue;
      const source = parse(),
        node = source.getNode(revision.nodeId);
      if (!node) continue;
      const begin = source.textOffset(node.nodeId)!;
      if (boundary <= begin || boundary >= begin + node.text.length) continue;
      const identities = edit.splitSourceInline(node.nodeId, boundary - begin);
      copyWordSplitRun(word, identities, styles);
    }
  const source = parse(),
    offset = source.textOffset(target.nodeId)!,
    preferred = new Set(
      word.revisions
        .filter((item) => {
          if (item.kind !== "format" || item.property !== "run") return false;
          const begin = source.textOffset(item.nodeId),
            node = source.getNode(item.nodeId);
          return begin !== undefined && node && begin >= start && begin + node.text.length <= end;
        })
        .map((item) => item.nodeId),
    );
  return {
    nodeId: target.nodeId,
    from: start - offset,
    to: end - offset,
    expectedText: range.expectedText,
    preferred,
  };
}
