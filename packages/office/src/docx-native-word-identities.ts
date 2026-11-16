import { DocumentError } from "@doffi/core";
import { W, MC } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
type Origin = { part: string };
function identity(kind: string, raw: string): string {
  if (kind === "permission") return raw;
  const value = /^([+-]?)(\d+)$/.exec(raw.trim());
  if (!value)
    throw new DocumentError("InvalidDocument", "A preserved Word numeric identity is malformed.");
  const digits = value[2]!.replace(/^0+(?=\d)/, "");
  return (value[1] === "-" && digits !== "0" ? "-" : "") + digits;
}
/** Each OOXML range family owns a separate ID namespace. Paired boundaries
 * can live in different preserved fragments of the same original part. */
function family(node: XmlNode, parent?: XmlNode): string | undefined {
  if (node.uri !== W) return;
  if (node.name === "id" && parent?.uri === W && parent.name === "sdtPr") return "control";
  if (["bookmarkStart", "bookmarkEnd"].includes(node.name)) return "bookmark";
  if (["permStart", "permEnd"].includes(node.name)) return "permission";
  const range =
    /^(moveFrom|moveTo|customXmlIns|customXmlDel|customXmlMoveFrom|customXmlMoveTo)Range(?:Start|End)$/.exec(
      node.name,
    );
  if (range) return range[1];
  if (
    [
      "ins",
      "del",
      "moveFrom",
      "moveTo",
      "rPrChange",
      "pPrChange",
      "tblPrChange",
      "tblPrExChange",
      "trPrChange",
      "tcPrChange",
      "tblGridChange",
      "sectPrChange",
      "cellIns",
      "cellDel",
      "cellMerge",
    ].includes(node.name)
  )
    return "revision";
  return;
}
export class NativeWordIdentities {
  private readonly used = new Map<string, Set<string>>();
  private readonly mapped = new Map<string, string>();
  private readonly next = new Map<string, number>();
  constructor(
    roots: readonly XmlNode[],
    private readonly origin: (node: XmlNode) => Origin | undefined,
    reservedRevisions: Iterable<string> = [],
  ) {
    const visit = (node: XmlNode, word = true, parent?: XmlNode): void => {
      word = (node.uri === W && (word || node.name === "txbxContent")) || (node.uri === MC && word);
      const kind = word ? family(node, parent) : undefined,
        id = node.attrs.get(W + (kind === "control" ? " val" : " id"));
      if (kind && id !== undefined) {
        let used = this.used.get(kind);
        if (!used) {
          used = new Set();
          this.used.set(kind, used);
        }
        used.add(identity(kind, id));
      }
      for (const child of node.content) if (typeof child !== "string") visit(child, word, node);
    };
    for (const root of roots) visit(root);
    const revisions = this.used.get("revision") ?? new Set<string>();
    for (const id of reservedRevisions) revisions.add(id);
    this.used.set("revision", revisions);
  }
  rewrite(roots: readonly XmlNode[]): void {
    const visit = (node: XmlNode, word = true, parent?: XmlNode): void => {
      word = (node.uri === W && (word || node.name === "txbxContent")) || (node.uri === MC && word);
      const origin = this.origin(node),
        kind = word ? family(node, parent) : undefined,
        attribute = W + (kind === "control" ? " val" : " id"),
        old = node.attrs.get(attribute);
      if (origin && kind && old !== undefined) {
        const key = kind + "\n" + origin.part + "\n" + identity(kind, old);
        let id = this.mapped.get(key);
        if (id === undefined) {
          let used = this.used.get(kind);
          if (!used) {
            used = new Set();
            this.used.set(kind, used);
          }
          let next = this.next.get(kind) ?? 0;
          while (used.has(String(next))) next++;
          if (next > 2147483647)
            throw new DocumentError(
              "ResourceLimitExceeded",
              "Native Word range identities exceed their bounded integer namespace.",
            );
          id = String(next);
          used.add(id);
          this.next.set(kind, next + 1);
          this.mapped.set(key, id);
        }
        (node.attrs as Map<string, string>).set(attribute, id);
      }
      for (const child of node.content) if (typeof child !== "string") visit(child, word, node);
    };
    for (const root of roots) visit(root);
  }
}
