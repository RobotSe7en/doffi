import { defaultTreeAdapter } from "parse5";
import type { DocumentNode, TextOffsetMapping } from "../runtime.js";
import type { DocumentContentOperation } from "../contracts/operations.js";
import { wordScriptRole } from "../word-script.js";
import type { WordScriptRole } from "../word-script.js";
import { parseHtml } from "./html.js";
import type { ParsedHtml } from "./html.js";
import { serializeHtmlOuter } from "./html-parser.js";
import { encode } from "./package.js";
import { snapshotWithText } from "./state.js";
import type { DocumentState } from "./state.js";
import { isTextBoundary } from "./text.js";
import { checkLimit, freeze } from "./validation.js";

type Replacement = Extract<DocumentContentOperation, { op: "replaceText" }>;
interface Run {
  readonly node: DocumentNode;
  readonly original: DocumentNode;
  readonly htmlStart: number;
  readonly bytes: Uint8Array;
  readonly lineage: readonly string[];
  readonly start: number;
  readonly delta: number;
}
interface TextRevision {
  readonly structure: ParsedHtml;
  readonly runs: ReadonlyMap<string, Run>;
  readonly ancestors: ReadonlySet<string>;
}
const revisions = new WeakMap<ParsedHtml, TextRevision>();

/** These revisions keep the entire validated structure, with only named text
 * leaves changed. Unrelated nodes and their signatures are shared. */
export function changedTextRuns(before: ParsedHtml, after: ParsedHtml): readonly string[] | undefined {
  const left = revisions.get(before), right = revisions.get(after);
  if (!(left || right) || (left?.structure ?? before) !== (right?.structure ?? after)) return;
  return [...new Set([...(left?.runs.keys() ?? []), ...(right?.runs.keys() ?? [])])]
    .filter(id => before.directText(id) !== after.directText(id));
}

export function replaceTextLeaf(
  state: DocumentState,
  operation: Replacement,
): { state: DocumentState; textEdits: readonly TextOffsetMapping[] } | undefined {
  const { metadata, manifest, annotations } = state.snapshot, word = metadata.word;
  if (annotations.length || (!word && !metadata.presentation) ||
    word && (word.trackChanges || word.fields.length || word.revisions.length || word.bookmarks.length || word.controls?.length || (word.protection?.enforcement && word.protection.edit !== "none"))) return;
  const { nodeId, from, to, expectedText, text } = operation;
  if (expectedText === text || /[\u0000-\u001f\u007f-\u009f\ud800-\udfff\ufdd0-\ufdef]/u.test(text) ||
    [...text].some(character => (character.codePointAt(0)! & 0xffff) >= 0xfffe)) return;
  const previous = revisions.get(state.html), structure = previous?.structure ?? state.html;
  const original = structure.getNode(nodeId), active = previous?.runs.get(nodeId)?.node ?? original;
  if (!original || !active || !["span", "p", "strong", "em", "u", "s", "sub", "sup"].includes(active.tagName) ||
    active.childNodeIds.length || !active.text || structure.directText(nodeId) !== original.text ||
    !isTextBoundary(active.text, from) || !isTextBoundary(active.text, to) ||
    to < from || active.text.slice(from, to) !== expectedText) return;

  const ancestors = new Set(previous?.ancestors), lineage: string[] = [];
  let main = false, slide = false, block = false, scripted = false;
  let script: WordScriptRole | undefined;
  for (let id: string | null = nodeId; id !== null;) {
    const node: Omit<DocumentNode, "text" | "html"> = structure.peekNode(id)!;
    const attrs = node.attributes;
    if (["ins", "del"].includes(node.tagName) ||
      Object.keys(attrs).some(name => /^data-word-(?:native|drawing|control|field|link|bookmark|comment|tab|break|paragraph-mark|reset|note|equation)/.test(name))) return;
    if (attrs["data-office-field"] !== undefined || attrs["data-office-link"] !== undefined) return;
    if (attrs["data-office-kind"] === "slide") slide = true;
    if (attrs["data-office-story"] !== undefined) {
      if (attrs["data-office-story"] !== "main") return;
      main = true;
    }
    if (/^(?:p|h[1-6])$/.test(node.tagName)) block = true;
    const role = attrs["data-word-script"];
    if (!script && (role === "ascii" || role === "hAnsi" || role === "eastAsia" || role === "cs")) script = role;
    scripted ||= role !== undefined || attrs["data-word-run"] !== undefined;
    lineage.push(id);
    if (id !== nodeId) ancestors.add(id);
    id = node.parentNodeId;
  }
  const role = script ?? "ascii";
  if (!(main || slide) || !block || scripted && [...text].some(character => wordScriptRole(character, role) !== role)) return;
  if (metadata.presentation && (active.attributes["data-office-kind"] !== undefined ||
    metadata.presentation.hyperlinks.some(link => lineage.includes(link.nodeId)) ||
    metadata.presentation.objects.some(object => object.kind === "connector" &&
      (lineage.includes(object.startNodeId ?? "") || lineage.includes(object.endNodeId ?? ""))))) return;
  const replacement = active.text.slice(0, from) + text + active.text.slice(to);
  if (!replacement) return;

  const oldText = serializeHtmlOuter(defaultTreeAdapter.createTextNode(active.text));
  const closing = "</" + active.tagName + ">";
  if (!active.html.endsWith(oldText + closing)) return;
  const escaped = serializeHtmlOuter(defaultTreeAdapter.createTextNode(replacement));
  const nodeHtml = active.html.slice(0, -(oldText.length + closing.length)) + escaped + closing;
  const previousRun = previous?.runs.get(nodeId);
  const baseHtml = structure.html, htmlStart = previousRun?.htmlStart ?? baseHtml.indexOf(original.html);
  // An authored HTML comment can quote the complete markup of this leaf.
  // Ambiguous serialized positions require the ordinary structural editor.
  if (htmlStart < 0 || !previousRun && baseHtml.indexOf(original.html, htmlStart + original.html.length) >= 0) return;
  const files = state.files, bytes = encode(nodeHtml), sharedBytes = files.get(manifest.entrypoint)!.byteLength;
  const byteLength = (state.textEntry?.byteLength ?? sharedBytes) + bytes.length -
    (previousRun?.bytes.byteLength ?? encode(original.html).byteLength);
  checkLimit(byteLength, state.limits.maxFileBytes, "maxFileBytes", "edited HTML");
  let size = 0;
  for (const file of files.values()) size += file.length;
  checkLimit(size - sharedBytes + byteLength, state.limits.maxPackageBytes, "maxPackageBytes", "the edited package");

  const runs = new Map(previous?.runs), startOffset = structure.textOffset(nodeId)!;
  const offsetAt = (offset: number, values: ReadonlyMap<string, Run>): number => {
    let result = offset;
    for (const run of values.values()) if (run.start < offset) result += run.delta;
    return result;
  };
  const documentOffset = offsetAt(startOffset, runs);
  const textEdits = lineage.map(id => {
    const baseOffset = offsetAt(structure.textOffset(id)!, runs);
    return { nodeId: id, from: documentOffset - baseOffset + from, to: documentOffset - baseOffset + to, insertedLength: text.length };
  });
  const node = freeze({ ...active, text: replacement, html: nodeHtml });
  runs.set(nodeId, { node, original, htmlStart, bytes, lineage, start: startOffset, delta: replacement.length - original.text.length });
  const ordered = [...runs.values()].sort((a, b) => a.htmlStart - b.htmlStart);
  const source = (): string => {
    const parts: string[] = [];
    let position = 0;
    for (const run of ordered) {
      parts.push(baseHtml.slice(position, run.htmlStart), run.node.html);
      position = run.htmlStart + run.original.html.length;
    }
    parts.push(baseHtml.slice(position));
    return parts.join("");
  };
  // Structural commands materialize this revision through the normal parser.
  // Ordinary typing, node inspection and comparisons use the retained leaves.
  let parsed: ParsedHtml | undefined;
  const parse = () => parsed ??= parseHtml(source(), manifest, files, state.limits, { nodes: 0 });
  function getNode(id: string): DocumentNode | undefined {
    const patched = runs.get(id);
    if (patched) return patched.node;
    const entry = structure.getNode(id);
    if (!entry || !ancestors.has(id)) return entry;
    const start = structure.textOffset(id)!;
    let nodeHtml = entry.html, nodeText = entry.text;
    for (let index = ordered.length - 1; index >= 0; index--) {
      const run = ordered[index]!;
      if (!run.lineage.includes(id)) continue;
      const relativeHtml = entry.html.indexOf(run.original.html), relativeText = run.start - start;
      nodeHtml = nodeHtml.slice(0, relativeHtml) + run.node.html + nodeHtml.slice(relativeHtml + run.original.html.length);
      nodeText = nodeText.slice(0, relativeText) + run.node.text + nodeText.slice(relativeText + run.original.text.length);
    }
    return freeze({ ...entry, html: nodeHtml, text: nodeText });
  }
  const next: ParsedHtml = {
    ...structure,
    get html() { return source(); },
    getNode,
    directText: id => runs.get(id)?.node.text ?? structure.directText(id),
    signature: id => runs.has(id) ? parse().signature(id) : structure.signature(id),
    textLength(id) {
      const length = structure.textLength(id);
      return length === undefined ? undefined : length + ordered.reduce((delta, run) =>
        delta + (run.lineage.includes(id) ? run.delta : 0), 0);
    },
    textOffset(id) {
      const originalOffset = structure.textOffset(id);
      return originalOffset === undefined ? undefined : offsetAt(originalOffset, runs);
    },
    textScope: (id, begin, end) => parse().textScope(id, begin, end),
  };
  revisions.set(next, { structure, runs, ancestors });
  return {
    state: {
      ...state,
      textEntry: { byteLength, patches: ordered.map(run => run.bytes) },
      html: next,
      snapshot: snapshotWithText(state.snapshot, next),
    },
    textEdits,
  };
}
