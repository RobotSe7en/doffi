import { W, MC } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";

export type CommentControl = "commentRangeStart" | "commentRangeEnd" | "commentReference";
export interface NativeCommentBinding {
  sourceId: string;
  id: string;
  startNodeId?: string;
  endNodeId?: string;
  referenceNodeId?: string;
}
export function commentControl(node: XmlNode): CommentControl | undefined {
  return node.uri === W &&
    ["commentRangeStart", "commentRangeEnd", "commentReference"].includes(node.name)
    ? (node.name as CommentControl)
    : undefined;
}
/** Foreign cached XML stays data; real Word textbox content reenters Word scope. */
export function wordScope(node: XmlNode, parent = true): boolean {
  return (node.uri === W && (parent || node.name === "txbxContent")) || (node.uri === MC && parent);
}
/** Retain native endpoints only while the current source still anchors there.
 * A deleted/reanchored thread cannot leave dangling controls in opaque XML. */
export class NativeWordComments {
  private readonly bindings: ReadonlyMap<string, readonly NativeCommentBinding[]>;
  private readonly suppressed = new Set<string>();
  private readonly originals = new WeakSet<XmlNode>();
  constructor(bindings: readonly NativeCommentBinding[]) {
    const indexed = new Map<string, NativeCommentBinding[]>();
    for (const binding of bindings) {
      const items = indexed.get(binding.sourceId) ?? [];
      items.push(binding);
      indexed.set(binding.sourceId, items);
    }
    this.bindings = indexed;
  }
  restore(root: XmlNode, nodeId: string | undefined): void {
    const ranges = new Set<string>();
    const scan = (parent: XmlNode, scope = true): void => {
      scope = wordScope(parent, scope);
      for (const child of parent.content) {
        if (typeof child === "string") continue;
        const kind = wordScope(child, scope) ? commentControl(child) : undefined;
        if (kind === "commentRangeStart" || kind === "commentRangeEnd") {
          const id = child.attrs.get(W + " id") ?? "",
            endpoint = kind === "commentRangeStart" ? "startNodeId" : "endNodeId";
          if (nodeId && this.bindings.get(id)?.some((b) => b[endpoint] === nodeId)) ranges.add(id);
        } else scan(child, scope);
      }
    };
    scan(root);
    const visit = (parent: XmlNode, scope = true): void => {
      scope = wordScope(parent, scope);
      for (let index = parent.content.length - 1; index >= 0; index--) {
        const node = parent.content[index]!;
        if (typeof node === "string") continue;
        const kind = wordScope(node, scope) ? commentControl(node) : undefined;
        if (!kind) {
          visit(node, scope);
          continue;
        }
        const endpoint =
          kind === "commentRangeStart"
            ? "startNodeId"
            : kind === "commentRangeEnd"
              ? "endNodeId"
              : "referenceNodeId";
        const binding = this.bindings
            .get(node.attrs.get(W + " id") ?? "")
            ?.find((binding) => nodeId && binding[endpoint] === nodeId),
          owner =
            kind === "commentRangeStart"
              ? binding?.startNodeId
              : kind === "commentRangeEnd"
                ? binding?.endNodeId
                : binding?.referenceNodeId;
        if (!binding || !nodeId || owner !== nodeId) {
          parent.content.splice(index, 1);
          continue;
        }
        (node.attrs as Map<string, string>).set(W + " id", binding.id);
        this.originals.add(node);
        this.suppressed.add(binding.id + "\n" + kind);
        if (
          kind === "commentReference" &&
          binding.startNodeId === nodeId &&
          binding.endNodeId === nodeId &&
          !ranges.has(binding.sourceId)
        ) {
          this.suppressed.add(binding.id + "\ncommentRangeStart");
          this.suppressed.add(binding.id + "\ncommentRangeEnd");
        }
      }
    };
    visit(root);
  }
  removeDuplicates(roots: Iterable<XmlNode>): void {
    const visit = (parent: XmlNode, scope = true): void => {
      scope = wordScope(parent, scope);
      for (let index = parent.content.length - 1; index >= 0; index--) {
        const node = parent.content[index]!;
        if (typeof node === "string") continue;
        const kind = wordScope(node, scope) ? commentControl(node) : undefined;
        if (
          kind &&
          !this.originals.has(node) &&
          this.suppressed.has((node.attrs.get(W + " id") ?? "") + "\n" + kind)
        )
          parent.content.splice(index, 1);
        else {
          visit(node, scope);
          if (node.attrs.has("urn:doffi:transient:native comment-reference")) {
            (node.attrs as Map<string, string>).delete(
              "urn:doffi:transient:native comment-reference",
            );
            if (!node.content.length) parent.content.splice(index, 1);
          }
        }
      }
    };
    for (const root of roots) visit(root);
  }
}
