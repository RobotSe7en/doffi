import { parse } from "parse5";
import type { DefaultTreeAdapterTypes as Tree } from "parse5";
import { DocumentError } from "@doffi/core";
import type { DocumentReadLimits } from "@doffi/core";

export interface SourceElement {
  readonly tag: string;
  readonly attrs: Readonly<Record<string, string>>;
  readonly children: SourceNode[];
  readonly parent: SourceElement | undefined;
  readonly from: number;
  to: number;
}
export interface SourceText { readonly text: string; readonly parent: SourceElement; readonly from: number; readonly to: number }
export type SourceNode = SourceElement | SourceText;
export interface SourceTree {
  readonly document: SourceElement;
  readonly root: SourceElement;
  readonly byId: ReadonlyMap<string, SourceElement>;
  readonly elements: readonly SourceElement[];
}
export const element = (node: SourceNode): node is SourceElement => "tag" in node;
export const children = (node: SourceElement, tag?: string): SourceElement[] => node.children.filter(element).filter((child) => !tag || child.tag === tag);
export function contains(parent: SourceElement, child: SourceNode): boolean {
  for (let current: SourceNode | undefined = child; current; current = current.parent) if (current === parent) return true;
  return false;
}
export function sourceText(node: SourceNode): string {
  if (!element(node)) return node.text;
  if (node.tag === "br") return "\n"; if (node.tag === "img") return "\ufffc";
  return node.children.map(sourceText).join("");
}
/** Only validated, normalized source reaches this adapter. The second parse never
 * introduces DOM, browser layout, runtime font substitutions or an authority. */
export function parseSource(html: string, limits: DocumentReadLimits, kind: "document" | "presentation" | "workbook-content" = "document"): SourceTree {
  const byId = new Map<string, SourceElement>(), elements: SourceElement[] = [];
  let offset = 0, count = 0;
  function visit(node: Tree.Node, parent: SourceElement | undefined, depth: number): SourceNode | undefined {
    if (++count > limits.maxNodes || depth > limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "Office source tree exceeds its budget.");
    if (node.nodeName === "#text" && "value" in node) {
      if (!parent) return;
      const from = offset; offset += node.value.length; return { text: node.value, parent, from, to: offset };
    }
    if (!("childNodes" in node)) return;
    const tag = "tagName" in node ? node.tagName : "#document";
    const attrs = "attrs" in node ? Object.fromEntries(node.attrs.map((attr) => [attr.name, attr.value])) : {};
    const result: SourceElement = { tag, attrs, children: [], parent, from: offset, to: offset }; elements.push(result);
    if (attrs["data-node-id"]) byId.set(attrs["data-node-id"], result);
    if (tag === "br" || tag === "img") offset++;
    else for (const child of node.childNodes) { const parsed = visit(child, result, depth + 1); if (parsed) result.children.push(parsed); }
    result.to = offset; return result;
  }
  const document = visit(parse(html), undefined, 0) as SourceElement;
  const root = elements.find((node) => node.attrs["data-office-kind"] === kind && node.parent?.tag === "body");
  if (!root) throw new DocumentError("UnsupportedFormat", "The native exporter requires the matching Office source kind.");
  return { document, root, byId, elements };
}
