import { defaultTreeAdapter as adapter, html as namespaces } from "parse5";
import type { DefaultTreeAdapterTypes as Tree } from "parse5";
import { generate, ident, parse, walk } from "css-tree";
import type { DocumentReadLimits } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { checkLimit } from "./validation.js";
import { isTextBoundary } from "./text.js";
import { drawingStyleVariable, setDrawingInlineStyle } from "./drawing-css.js";
import { setInlineStyle } from "./css.js";
import { SourceStyleResolver } from "./source-css.js";
import type { SourceStyleNode } from "./source-css.js";
import { TEXT_DECORATION_RESET, textFormattingProperties } from "./text-decoration.js";

export interface SourceTextRange { readonly from: number; readonly to: number; readonly expectedText: string }
/** Format existing text leaves, retaining all paragraph/run/link identities.
 * New wrappers are core-owned source nodes; no rendered DOM enters the tree. */
export function styleSourceText(target: Tree.Element, range: SourceTextRange, properties: Readonly<Record<string, string>>, limits: DocumentReadLimits, allocate: () => string, styles: readonly string[] = []): void {
  properties = textFormattingProperties(properties);
  const fail: (message: string) => never = message => { throw new DocumentError("InvalidOperation", message); };
  for (const property of Object.keys(properties)) if (!/^(?:color|font-(?:family|size|weight|style)|text-decoration(?:-(?:line|style|color))?|vertical-align|letter-spacing|--office-(?:font|text)-[a-z0-9-]+)$/.test(property)) fail("Text range appearance accepts only font and text properties.");
  const patch = (source: string): string => {
    for (const [property, value] of Object.entries(properties)) source = setDrawingInlineStyle(source, property, value, limits); return source;
  };
  const incoming = patch("");
  const declarations = parse(incoming, { context: "declarationList", parseCustomProperty: true });
  if (declarations.type !== "DeclarationList") fail("The source font style is invalid.");
  const requestedDeclarations = declarations.children.toArray().filter(node => node.type === "Declaration");
  const family = requestedDeclarations.find(node => ident.decode(node.property) === "font-family"), familyVariable = family ? drawingStyleVariable(family.value) : undefined;
  const requested = Object.fromEntries(requestedDeclarations.map(node => [ident.decode(node.property), generate(node.value) + (node.important ? "!important" : "")]));
  const variables = new Set<string>();
  for (const declaration of requestedDeclarations) walk(declaration.value, node => {
    if (node.type === "Function" && ident.decode(node.name).toLowerCase() === "var" && node.children.first?.type === "Identifier") variables.add(ident.decode(node.children.first.name));
  });
  const resolver = new SourceStyleResolver(styles, limits), projections = new WeakMap<Tree.Element, SourceStyleNode>();
  function projection(node: Tree.Element): SourceStyleNode {
    let value = projections.get(node); if (value) return value;
    const parent = node.parentNode && "tagName" in node.parentNode ? projection(node.parentNode) : undefined;
    value = { tag: node.tagName, attrs: Object.fromEntries(node.attrs.map(attribute => [attribute.name, attribute.value])), parent }; projections.set(node, value); return value;
  }
  const inheritedFonts = ["font-family", "font-size", "font-weight", "font-style", "font-variant", "text-transform", "color", "line-height", "letter-spacing", "word-spacing", "white-space", "overflow-wrap", "tab-size"];
  const roleFonts = ["ascii", "h-ansi", "east-asia", "complex-script"].flatMap(role => ["font", "major", "minor"].map(prefix => "--office-" + prefix + "-" + role));
  const neutralStyles = new Map<string, string>();
  function neutral(parent: SourceStyleNode, nodeId: string): string {
    const source = { tag: "span", attrs: { style: incoming, "data-node-id": nodeId }, parent }, computed = resolver.get(parent), own = resolver.specified(source);
    const names = new Set([...inheritedFonts, ...variables]);
    for (const key of [...Object.keys(computed), ...Object.keys(own)]) if (key !== TEXT_DECORATION_RESET && /^--office-(?:font-|text-|major-|minor-|underline$|strike$)/.test(key)) names.add(key);
    for (const name of roleFonts) names.add(name);
    const key = JSON.stringify([...names].sort()); let style = neutralStyles.get(key);
    if (style !== undefined) return style;
    style = ""; for (const name of names) style = setInlineStyle(style, name, "inherit!important", limits);
    if (familyVariable?.fallback) for (const role of ["ascii", "h-ansi", "east-asia", "complex-script"]) {
      if (familyVariable.name === "--office-font-" + role) for (const prefix of ["major", "minor"]) style = setInlineStyle(style, "--office-" + prefix + "-" + role, "initial!important", limits);
    }
    for (const [name, value] of Object.entries({ "vertical-align": "baseline", "background-color": "transparent", margin: "0", padding: "0", border: "none" })) style = setInlineStyle(style, name, value + "!important", limits);
    // A wrapper contributes no additional decoration. Authored ancestor paint
    // still propagates unless the requested native edit explicitly resets it.
    style = setInlineStyle(style, "text-decoration-line", "none!important", limits);
    style = setInlineStyle(style, TEXT_DECORATION_RESET, "none!important", limits);
    for (const node of requestedDeclarations) style = setInlineStyle(style, ident.decode(node.property), generate(node.value) + "!important", limits);
    neutralStyles.set(key, style); return style;
  }
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
  const bounds = new Map<Tree.Element, { from: number; to: number }>(), updated = new Set<Tree.Element>();
  for (const leaf of leaves) {
    const parent = leaf.node.parentNode; if (!parent || !("tagName" in parent)) continue;
    const previous = bounds.get(parent); bounds.set(parent, { from: Math.min(previous?.from ?? leaf.from, leaf.from), to: Math.max(previous?.to ?? leaf.to, leaf.to) });
  }
  const protectedFields = [...inheritedFonts, ...roleFonts, TEXT_DECORATION_RESET, "text-decoration-line", "vertical-align", "background-color", "margin", "padding", "border"];
  function updateCompleteSpan(parent: Tree.Element): boolean {
    if (parent.tagName !== "span" || parent.childNodes.some(node => node.nodeName !== "#text")) return false;
    const area = bounds.get(parent); if (!area || range.from > area.from || range.to < area.to) return false;
    const attribute = parent.attrs.find(attribute => attribute.name === "style"); if (!attribute) return false;
    const ast = parse(attribute.value, { context: "declarationList", parseCustomProperty: true }); if (ast.type !== "DeclarationList") return false;
    const own = new Map(ast.children.toArray().flatMap(node => node.type === "Declaration" ? [[ident.decode(node.property), node] as const] : []));
    // A complete self-protected font span can retain its identity and all
    // unrequested declarations. Partial selections still get their own span.
    if (!protectedFields.every(name => own.get(name)?.important)) return false;
    if (!updated.has(parent)) {
      for (const node of requestedDeclarations) attribute.value = setInlineStyle(attribute.value, ident.decode(node.property), generate(node.value) + "!important", limits);
      updated.add(parent);
    }
    return true;
  }
  let created = 0;
  const count = (): void => { checkLimit(++created, limits.maxNodes); };
  const makeText = (text: string): Tree.TextNode => { count(); return adapter.createTextNode(text); };
  for (const leaf of leaves) {
    if (leaf.to <= range.from || leaf.from >= range.to) continue;
    if ("tagName" in leaf.node && leaf.node.tagName === "img") fail("Text font formatting cannot change an inline object.");
    const parent = leaf.node.parentNode;
    if (!parent || !("tagName" in parent)) fail("The source text parent is unavailable.");
    const owner = projection(parent);
    if (resolver.hasSpecified(owner, requested)) continue;
    if (updateCompleteSpan(parent)) continue;
    count(); const nodeId = allocate(), span = adapter.createElement("span", namespaces.NS.HTML, [{ name: "data-node-id", value: nodeId }, { name: "style", value: neutral(owner, nodeId) }]);
    if (!("value" in leaf.node)) { adapter.insertBefore(parent, span, leaf.node); adapter.detachNode(leaf.node); adapter.appendChild(span, leaf.node); continue; }
    const from = Math.max(0, range.from - leaf.from), to = Math.min(leaf.value.length, range.to - leaf.from);
    if (from) adapter.insertBefore(parent, makeText(leaf.value.slice(0, from)), leaf.node);
    adapter.appendChild(span, makeText(leaf.value.slice(from, to))); adapter.insertBefore(parent, span, leaf.node);
    if (to < leaf.value.length) adapter.insertBefore(parent, makeText(leaf.value.slice(to)), leaf.node);
    adapter.detachNode(leaf.node);
  }
}
