import { W } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";

type Insertion = { id: string; author: string; date?: string };
const child = (node: XmlNode, name: string): XmlNode | undefined =>
  node.content.find((n): n is XmlNode => typeof n !== "string" && n.uri === W && n.name === name);
const create = (owner: XmlNode, name: string): XmlNode => ({
  uri: W,
  name,
  attrs: new Map(),
  namespaces: {},
  content: [],
  part: owner.part,
});
const ensure = (owner: XmlNode, name: string): XmlNode => {
  const existing = child(owner, name);
  if (existing) return existing;
  const result = create(owner, name);
  owner.content.unshift(result);
  return result;
};
/** Only known block containers admit a provable whole-object insertion. An
 * unknown sibling might own content outside the inserted paragraphs. */
function insertionUnits(root: XmlNode): XmlNode[] | undefined {
  const units: XmlNode[] = [];
  let complete = true;
  const properties: Record<string, readonly string[]> = {
    customXml: ["customXmlPr"],
    sdt: ["sdtPr", "sdtEndPr"],
    sdtContent: [],
    tbl: ["tblPr", "tblGrid"],
    tr: ["trPr", "tblPrEx"],
    tc: ["tcPr"],
  };
  const visit = (node: XmlNode): void => {
    if (node.uri !== W) {
      complete = false;
      return;
    }
    if (node.name === "p") {
      units.push(node);
      return;
    }
    if (!Object.hasOwn(properties, node.name)) {
      complete = false;
      return;
    }
    if (node.name === "tr") units.push(node);
    for (const child of node.content) {
      if (typeof child === "string") {
        if (child.trim()) complete = false;
      } else if (child.uri !== W || !properties[node.name]!.includes(child.name)) visit(child);
    }
  };
  visit(root);
  return complete && units.length ? units : undefined;
}
/** OOXML has no block-level w:ins wrapper. Mark the native paragraph marks,
 * run content and table rows instead, keeping every insertion in one group. */
export function insertNativeBlock(root: XmlNode, insertion: Insertion): boolean {
  const units = insertionUnits(root);
  if (!units) return false;
  const mark = (owner: XmlNode): XmlNode => ({
    ...create(owner, "ins"),
    attrs: new Map([
      [W + " id", insertion.id],
      [W + " author", insertion.author],
      ...(insertion.date ? [[W + " date", insertion.date] as [string, string]] : []),
    ]),
  });
  for (const node of units) {
    if (node.name === "p") {
      ensure(ensure(node, "pPr"), "rPr").content.push(mark(node));
      const properties = child(node, "pPr")!,
        wrapper = mark(node);
      wrapper.content.push(...node.content.filter((n) => n !== properties));
      node.content.splice(0, node.content.length, properties, wrapper);
    } else {
      ensure(node, "trPr").content.push(mark(node));
    }
  }
  return true;
}
/** Recognize complete native insertions after reopening. A mixed or partly
 * inserted atom stays opaque; it must never acquire a whole-object inverse. */
export function nativeBlockInsertion(root: XmlNode): Insertion | undefined {
  const marks: XmlNode[] = [];
  const units = insertionUnits(root);
  if (!units) return;
  let complete = true;
  for (const node of units) {
    if (node.name === "p") {
      const properties = child(node, "pPr"),
        mark = properties && child(properties, "rPr"),
        insertion = mark && child(mark, "ins");
      if (
        !insertion ||
        node.content.some((n) =>
          typeof n === "string"
            ? !!n.trim()
            : n !== properties &&
              !(
                n.uri === W &&
                n.name === "ins" &&
                ["id", "author", "date"].every(
                  (key) => n.attrs.get(W + " " + key) === insertion.attrs.get(W + " " + key),
                )
              ),
        )
      )
        complete = false;
      else marks.push(insertion);
    } else {
      const properties = child(node, "trPr"),
        insertion = properties && child(properties, "ins");
      if (!insertion) complete = false;
      else marks.push(insertion);
    }
  }
  const first = marks[0],
    id = first?.attrs.get(W + " id");
  if (
    !complete ||
    !first ||
    id === undefined ||
    marks.some((n) =>
      ["id", "author", "date"].some(
        (key) => n.attrs.get(W + " " + key) !== first.attrs.get(W + " " + key),
      ),
    )
  )
    return;
  return {
    id,
    author: first.attrs.get(W + " author") ?? "",
    ...(first.attrs.get(W + " date") ? { date: first.attrs.get(W + " date")! } : {}),
  };
}

/** Accepting source review removes only the insertion created for this opaque
 * block; other native revisions remain preserved. */
export function acceptNativeBlockInsertion(root: XmlNode, id: string): void {
  const visit = (node: XmlNode): void => {
    for (let i = node.content.length - 1; i >= 0; i--) {
      const n = node.content[i]!;
      if (typeof n === "string") continue;
      if (n.uri === W && n.name === "ins" && n.attrs.get(W + " id") === id) {
        node.content.splice(i, 1, ...n.content);
        for (const item of n.content) if (typeof item !== "string") visit(item);
      } else visit(n);
    }
  };
  visit(root);
}
