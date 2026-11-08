import { defaultTreeAdapter as adapter, html as namespaces } from "parse5";
import type { DefaultTreeAdapterTypes as Tree } from "parse5";
import type { WorkbookCell } from "../runtime.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import { DocumentError } from "../errors.js";
import { serializeHtmlTree } from "../document/html-parser.js";
import { isTextBoundary } from "../document/text.js";
import { parseWorkbookCellElement } from "../workbook/chunk-index.js";

type Format = Extract<WorkbookOperation, { op: "formatWorkbookText" }>;
export function formatWorkbookCellText(cell: WorkbookCell | undefined, operation: Format, cssClass?: string): string {
  if (!cell || cell.formula || cell.value.type !== "text") throw new DocumentError("InvalidOperation", "Partial font formatting requires a text constant.");
  const { from, to, expectedText } = operation;
  if (from >= to || !isTextBoundary(cell.text, from) || !isTextBoundary(cell.text, to) || cell.text.slice(from, to) !== expectedText)
    throw new DocumentError("InvalidOperation", "Text formatting requires an exact nonempty UTF-16 range preserving surrogate pairs and variation sequences.");
  const root = parseWorkbookCellElement(cell.html); if (!cssClass) return serializeHtmlTree(root);
  const leaves: { node: Tree.TextNode | Tree.Element; from: number; to: number; text?: string }[] = [], pending: Tree.Node[] = [root]; let offset = 0;
  while (pending.length) {
    const node = pending.pop()!;
    if (node.nodeName === "#text" && "value" in node) { leaves.push({ node, from: offset, to: offset + node.value.length, text: node.value }); offset += node.value.length; }
    else if ("tagName" in node && node.tagName === "br") { leaves.push({ node, from: offset, to: ++offset }); }
    else if ("childNodes" in node) for (let index = node.childNodes.length - 1; index >= 0; index--) pending.push(node.childNodes[index]!);
  }
  if (offset !== cell.text.length) throw new DocumentError("InvalidDocument", "Cell source does not match its text offsets.");
  for (const leaf of leaves) {
    if (leaf.to <= from || leaf.from >= to) continue;
    const parent = leaf.node.parentNode!, span = adapter.createElement("span", namespaces.NS.HTML, [{ name: "class", value: cssClass }]);
    if (leaf.text === undefined) { adapter.insertBefore(parent, span, leaf.node); adapter.detachNode(leaf.node); adapter.appendChild(span, leaf.node); continue; }
    const first = Math.max(0, from - leaf.from), last = Math.min(leaf.text.length, to - leaf.from);
    if (first) adapter.insertBefore(parent, adapter.createTextNode(leaf.text.slice(0, first)), leaf.node);
    adapter.appendChild(span, adapter.createTextNode(leaf.text.slice(first, last))); adapter.insertBefore(parent, span, leaf.node);
    if (last < leaf.text.length) adapter.insertBefore(parent, adapter.createTextNode(leaf.text.slice(last)), leaf.node);
    adapter.detachNode(leaf.node);
  }
  const html = serializeHtmlTree(root);
  if (html.length > 2 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "Rich cell formatting exceeds its source budget.");
  return html;
}
