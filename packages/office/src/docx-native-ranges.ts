import type { WordNativeContent } from "@doffi/core";
import { W, MC, attr, fail, elements, child } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";

type Boundary = NonNullable<WordNativeContent["fragments"][number]["rangeBoundaries"]>[number];
export function nativeRangeControl(
  node: XmlNode,
): Omit<Boundary, "rangeId" | "active"> | undefined {
  if (node.uri !== W) return;
  const match =
    /^(bookmark|perm)(Start|End)$|^(moveFrom|moveTo|customXmlIns|customXmlDel|customXmlMoveFrom|customXmlMoveTo)Range(Start|End)$/.exec(
      node.name,
    );
  if (!match) return;
  const family = match[1] === "perm" ? "permission" : (match[1] ?? match[3]!),
    raw = attr(node, "id");
  if (raw === undefined || (family !== "permission" && !/^[+-]?\d+$/.test(raw.trim())))
    fail(node, "A native range identity is invalid.");
  const sourceId =
    family === "permission"
      ? raw
      : (raw.trim().startsWith("-") && !/^[-+]?0+$/.test(raw.trim()) ? "-" : "") +
        raw
          .trim()
          .replace(/^[+-]/, "")
          .replace(/^0+(?=\d)/, "");
  if (!sourceId) fail(node, "A native range identity is empty.");
  return { family, sourceId, side: (match[2] ?? match[4]) === "Start" ? "start" : "end" };
}
/** Select the same MCE branch as the importer, and isolate nested Word stories. */
export function visitNativeRanges(
  root: XmlNode,
  visit: (node: XmlNode) => void,
  alternatives = false,
): void {
  const scan = (node: XmlNode, scope = true): void => {
    scope =
      (node.uri === W && (scope || node.name === "txbxContent")) || (node.uri === MC && scope);
    if (scope) visit(node);
    const children =
      node.uri === MC && node.name === "AlternateContent" && !alternatives
        ? [elements(node, "Choice", MC)[0] ?? child(node, "Fallback", MC)].filter(
            (n): n is XmlNode => !!n,
          )
        : node.content.filter((n): n is XmlNode => typeof n !== "string");
    for (const node of children) scan(node, scope);
  };
  scan(root);
}
/** Indexed controls let core manage ownership without interpreting OOXML. */
export class NativeRangeIndex {
  private readonly scopes = new WeakMap<XmlNode, number>();
  private readonly ids = new Map<string, string>();
  private readonly preserved = new Set<string>();
  private serial = 0;
  private readonly endpoints = new Map<string, { start?: XmlNode; end?: XmlNode }>();
  constructor(
    private readonly parent: (node: XmlNode) => XmlNode | undefined,
    private readonly position: (node: XmlNode) => number,
  ) {}
  private key(node: XmlNode): string | undefined {
    const control = nativeRangeControl(node);
    if (!control) return;
    let scope = node;
    for (let current: XmlNode | undefined = node; current; current = this.parent(current)) {
      scope = current;
      if (
        current.uri === W &&
        ["body", "hdr", "ftr", "comment", "footnote", "endnote", "txbxContent"].includes(
          current.name,
        )
      )
        break;
    }
    let id = this.scopes.get(scope);
    if (id === undefined) {
      id = this.serial++;
      this.scopes.set(scope, id);
    }
    return node.part + "\n" + id + "\n" + control.family + "\n" + control.sourceId;
  }
  begin(roots: Iterable<XmlNode>): void {
    this.preserved.clear();
    for (const root of roots)
      visitNativeRanges(root, (node) => {
        const key = this.key(node);
        if (key) this.preserved.add(key);
      });
  }
  preserves(node: XmlNode): boolean {
    const control = nativeRangeControl(node),
      key = control && this.key(node);
    return (
      !!control &&
      (control.family === "permission" ||
        control.family.startsWith("customXml") ||
        this.preserved.has(key!))
    );
  }
  validate(): void {
    for (const pair of this.endpoints.values()) {
      const node = pair.start ?? pair.end!;
      if (!pair.start || !pair.end || this.position(pair.start) >= this.position(pair.end))
        fail(node, "A preserved native range is unpaired or reversed.");
    }
  }
  index(root: XmlNode): Boundary[] {
    const result: Boundary[] = [];
    visitNativeRanges(root, (node) => {
      const control = nativeRangeControl(node),
        key = control && this.key(node);
      if (!control || !key) return;
      let rangeId = this.ids.get(key);
      if (!rangeId) {
        rangeId = "native-range-" + this.ids.size;
        this.ids.set(key, rangeId);
      }
      const pair = this.endpoints.get(key) ?? {};
      if (pair[control.side]) fail(node, "A native range boundary is repeated.");
      pair[control.side] = node;
      this.endpoints.set(key, pair);
      result.push({ ...control, rangeId });
    });
    return result;
  }
}
