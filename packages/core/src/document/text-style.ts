import { defaultTreeAdapter as adapter, html as namespaces } from "parse5";
import type { DefaultTreeAdapterTypes as Tree } from "parse5";
import { generate, ident, parse } from "css-tree";
import type { DocumentReadLimits } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { checkLimit } from "./validation.js";
import { isTextBoundary } from "./text.js";
import { setDrawingInlineStyle } from "./drawing-css.js";
import { textFormattingProperties } from "./text-decoration.js";

export interface SourceTextRange { readonly from: number; readonly to: number; readonly expectedText: string }
/** Format existing text leaves, retaining all paragraph/run/link identities.
 * New wrappers are core-owned source nodes; no rendered DOM enters the tree. */
export function styleSourceText(target: Tree.Element, range: SourceTextRange, properties: Readonly<Record<string, string>>, limits: DocumentReadLimits, allocate: () => string): void {
  properties = textFormattingProperties(properties);
  const fail: (message: string) => never = message => { throw new DocumentError("InvalidOperation", message); };
  for (const property of Object.keys(properties)) if (!/^(?:color|font-(?:family|size|weight|style)|text-decoration(?:-(?:line|style|color))?|vertical-align|letter-spacing|--office-(?:font|text)-[a-z0-9-]+)$/.test(property)) fail("Text range appearance accepts only font and text properties.");
  const patched = new Map<string, string>();
  const identities = new Map<string, string>();
  const identity = (source: string): string => {
    const known = identities.get(source); if (known !== undefined) return known;
    const ast = parse(source, { context: "declarationList", parseCustomProperty: true }), groups = new Map<string, string[]>();
    if (ast.type !== "DeclarationList") fail("The source font style is invalid.");
    for (const declaration of ast.children) if (declaration.type === "Declaration") {
      const property = ident.decode(declaration.property), key = property.startsWith("text-decoration") ? "text-decoration" : property;
      const values = groups.get(key) ?? []; values.push(generate(declaration)); groups.set(key, values);
    }
    // Independent declarations commute. Decorations retain their shorthand /
    // longhand order, so an explicit longhand after a reset is not a no-op.
    const result = JSON.stringify([...groups].sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)); identities.set(source, result); return result;
  };
  const patch = (source: string): string => {
    const cached = patched.get(source); if (cached !== undefined) return cached;
    const original = source; for (const [property, value] of Object.entries(properties)) source = setDrawingInlineStyle(source, property, value, limits); patched.set(original, source); return source;
  };
  const incoming = patch("");
  const leaves: { node: Tree.TextNode | Tree.Element; from: number; to: number; value: string }[] = [], pending: Tree.Node[] = [target]; let offset = 0;
  while (pending.length) {
    const node = pending.pop()!;
    const value = node.nodeName === "#text" && "value" in node ? node.value : "tagName" in node && node.tagName === "br" ? "\n" : "tagName" in node && node.tagName === "img" ? "\ufffc" : undefined;
    if (value !== undefined) { leaves.push({ node: node as Tree.TextNode | Tree.Element, from: offset, to: offset + value.length, value }); offset += value.length; }
    else if ("childNodes" in node) for (let index = node.childNodes.length - 1; index >= 0; index--) pending.push(node.childNodes[index]!);
  }
  const text = leaves.map(leaf => leaf.value).join("");
  if (range.from >= range.to || !isTextBoundary(text, range.from) || !isTextBoundary(text, range.to) || text.slice(range.from, range.to) !== range.expectedText) fail("Text formatting requires an exact nonempty UTF-16 source range.");
  if (!incoming) return;
  let created = 0;
  const count = (): void => { checkLimit(++created, limits.maxNodes); };
  const makeText = (text: string): Tree.TextNode => { count(); return adapter.createTextNode(text); };
  for (const leaf of leaves) {
    if (leaf.to <= range.from || leaf.from >= range.to) continue;
    if ("tagName" in leaf.node && leaf.node.tagName === "img") fail("Text font formatting cannot change an inline object.");
    const parent = leaf.node.parentNode;
    if (!parent || !("tagName" in parent)) fail("The source text parent is unavailable.");
    const style = parent.attrs.find(attribute => attribute.name === "style")?.value ?? "";
    if (identity(patch(style)) === identity(style)) continue;
    count(); const span = adapter.createElement("span", namespaces.NS.HTML, [{ name: "data-node-id", value: allocate() }, { name: "style", value: incoming }]);
    if (!("value" in leaf.node)) { adapter.insertBefore(parent, span, leaf.node); adapter.detachNode(leaf.node); adapter.appendChild(span, leaf.node); continue; }
    const from = Math.max(0, range.from - leaf.from), to = Math.min(leaf.value.length, range.to - leaf.from);
    if (from) adapter.insertBefore(parent, makeText(leaf.value.slice(0, from)), leaf.node);
    adapter.appendChild(span, makeText(leaf.value.slice(from, to))); adapter.insertBefore(parent, span, leaf.node);
    if (to < leaf.value.length) adapter.insertBefore(parent, makeText(leaf.value.slice(to)), leaf.node);
    adapter.detachNode(leaf.node);
  }
}
