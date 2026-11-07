import { defaultTreeAdapter as adapter, html as namespaces } from "parse5";
import type { DefaultTreeAdapterTypes as Tree } from "parse5";
import { DocumentError } from "../errors.js";
import type { WorkbookModel } from "../runtime.js";
import type { WorkbookCellSource } from "../contracts/workbook-operations.js";
import { serializeHtmlTree } from "../document/html-parser.js";
import { isTextBoundary } from "../document/text.js";
import { parseWorkbookCellElement } from "./chunk-index.js";
import { readWorkbookCellSource } from "./cell-source.js";
import type { WorkbookReferenceLocation } from "./reference.js";
import { captureWorkbookModel } from "./source.js";

export interface WorkbookCellTextEdit { readonly from: number; readonly to: number; readonly expectedText: string; readonly text: string }
interface Leaf { readonly node: Tree.TextNode | Tree.Element; readonly from: number; readonly to: number; readonly value: string }
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
function unicode(text: string): void {
  for (const character of text) { const value = character.codePointAt(0)!; if (value === 0 || value >= 0xd800 && value <= 0xdfff) invalid("Replacement text must contain well-formed Unicode without null characters."); }
}
function replaceLeaf(leaf: Leaf, value: string): void {
  if (leaf.value === value) return;
  if (leaf.node.nodeName === "#text" && "value" in leaf.node && !value.includes("\n")) { leaf.node.value = value; return; }
  const parent = leaf.node.parentNode!;
  for (const [index, text] of value.split("\n").entries()) {
    if (index) adapter.insertBefore(parent, adapter.createElement("br", namespaces.NS.HTML, []), leaf.node);
    if (text) adapter.insertBefore(parent, adapter.createTextNode(text), leaf.node);
  }
  adapter.detachNode(leaf.node);
}

/** Apply disjoint original UTF-16 splices to a text constant. Untouched spans
 * retain their source styles; inserted text inherits the first affected run.
 * The returned source still goes through the complete core transaction. */
export function replaceWorkbookCellText(input: WorkbookModel, location: WorkbookReferenceLocation, edits: readonly WorkbookCellTextEdit[]): WorkbookCellSource {
  const model = captureWorkbookModel(input), source = readWorkbookCellSource(model, location), cell = model.getCell(location.sheetId, location.row, location.column);
  if (!cell || cell.formula || cell.value.type !== "text") invalid("Rich cell text replacement requires a source text constant.");
  if (!Array.isArray(edits)) invalid("Cell text edits must be an ordered array.");
  if (edits.length > 32768) throw new DocumentError("ResourceLimitExceeded", "Cell text replacement exceeds its edit budget.");
  const original = cell.text;
  if (cell.value.value !== original) throw new DocumentError("InvalidDocument", "The workbook provider's cell text does not match its value.");
  const changes: WorkbookCellTextEdit[] = []; let previousFrom = -1, previousTo = -1, length = original.length, inserted = 0;
  for (const value of edits) {
    if (!value || typeof value !== "object" || Object.keys(value).some(key => !["from", "to", "expectedText", "text"].includes(key))) invalid("Invalid cell text edit.");
    const { from, to, expectedText, text } = value;
    if (!Number.isSafeInteger(from) || !Number.isSafeInteger(to) || from < 0 || from > to || to > original.length || from < previousTo || from <= previousFrom
      || typeof text !== "string" || typeof expectedText !== "string" || original.slice(from, to) !== expectedText || !isTextBoundary(original, from) || !isTextBoundary(original, to)) invalid("Cell text edits require exact disjoint UTF-16 source ranges, preserving surrogate pairs and variation sequences.");
    unicode(text); length += text.length - (to - from); inserted += text.length;
    if (inserted > 65534) throw new DocumentError("ResourceLimitExceeded", "Cell replacement text exceeds its input budget.");
    previousFrom = from; previousTo = to;
    if (text !== expectedText) changes.push({ from, to, expectedText, text });
  }
  if (length > 32767) throw new DocumentError("ResourceLimitExceeded", "Cell text replacement exceeds its text budget.");
  if (!changes.length) return source;
  if (source.kind !== "richText") {
    const pieces: string[] = []; let cursor = 0;
    for (const edit of changes) { pieces.push(original.slice(cursor, edit.from), edit.text); cursor = edit.to; }
    pieces.push(original.slice(cursor)); return Object.freeze({ kind: "value", value: Object.freeze({ type: "text", value: pieces.join("") }) });
  }
  const root = parseWorkbookCellElement(cell.html), leaves: Leaf[] = [], stack: Tree.Node[] = [root]; let offset = 0, emptyParent = root;
  while (stack.length) {
    const node = stack.pop()!;
    const text = node.nodeName === "#text" && "value" in node ? node.value : "tagName" in node && node.tagName === "br" ? "\n" : undefined;
    if (text !== undefined) { leaves.push({ node: node as Leaf["node"], from: offset, to: offset + text.length, value: text }); offset += text.length; }
    else if ("childNodes" in node) {
      if ("tagName" in node && node.tagName === "span") emptyParent = node;
      for (let index = node.childNodes.length - 1; index >= 0; index--) stack.push(node.childNodes[index]!);
    }
  }
  if (offset !== original.length) throw new DocumentError("InvalidDocument", "Rich cell source does not match its text offsets.");
  if (!leaves.length) { const node = adapter.createTextNode(""); adapter.appendChild(emptyParent, node); leaves.push({ node, from: 0, to: 0, value: "" }); }
  const planned = changes.map(edit => {
    let low = 0, high = leaves.length;
    while (low < high) { const middle = (low + high) >>> 1; if (leaves[middle]!.to <= edit.from) low = middle + 1; else high = middle; }
    return { ...edit, home: Math.min(low, leaves.length - 1) };
  });
  let editIndex = 0;
  for (const [index, leaf] of leaves.entries()) {
    let cursor = 0, value = "";
    while (editIndex < planned.length) {
      const edit = planned[editIndex]!;
      if (edit.from >= leaf.to && edit.home !== index) break;
      if (edit.to <= leaf.from && edit.home < index) { editIndex++; continue; }
      const from = Math.max(0, Math.min(leaf.value.length, edit.from - leaf.from)), to = Math.max(0, Math.min(leaf.value.length, edit.to - leaf.from));
      value += leaf.value.slice(cursor, Math.max(cursor, from));
      if (edit.home === index) value += edit.text;
      cursor = Math.max(cursor, to);
      if (edit.to <= leaf.to) editIndex++; else break;
    }
    value += leaf.value.slice(cursor); replaceLeaf(leaf, value);
  }
  const html = serializeHtmlTree(root);
  if (html.length > 2 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "Rich cell replacement exceeds its HTML budget.");
  return Object.freeze({ kind: "richText", html });
}
