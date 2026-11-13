import { DocumentError } from "@doffi/core";
import type { DocumentReadLimits } from "@doffi/core";
import { SaxesParser } from "saxes";

export const W: string =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const R: string =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const A: string =
  "http://schemas.openxmlformats.org/drawingml/2006/main";
export const P: string = "http://schemas.openxmlformats.org/presentationml/2006/main";
export const C: string = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const S: string = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const WP: string =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
export const PIC: string =
  "http://schemas.openxmlformats.org/drawingml/2006/picture";
export const MC: string =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const PKG: string =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const aliases: Readonly<Record<string, string>> = {
  "http://purl.oclc.org/ooxml/wordprocessingml/main": W,
  "http://purl.oclc.org/ooxml/officeDocument/relationships": R,
  "http://purl.oclc.org/ooxml/drawingml/main": A,
  "http://purl.oclc.org/ooxml/presentationml/main": P,
  "http://purl.oclc.org/ooxml/drawingml/chart": C,
  "http://purl.oclc.org/ooxml/spreadsheetml/main": S,
  "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing": WP,
  "http://purl.oclc.org/ooxml/drawingml/picture": PIC,
};
export const normalizeXmlNamespace = (uri: string): string => aliases[uri] ?? uri;
export interface XmlNode {
  readonly name: string;
  readonly uri: string;
  readonly attrs: ReadonlyMap<string, string>;
  readonly namespaces: Readonly<Record<string, string>>;
  readonly content: (XmlNode | string)[];
  readonly part: string;
}
export interface XmlBudget {
  nodes: number;
}
const unsupportedSources = new WeakMap<object, { node: XmlNode; attribute?: string }>();
export function unsupportedSource(error: unknown): { node: XmlNode; attribute?: string } | undefined {
  return typeof error === "object" && error !== null ? unsupportedSources.get(error) : undefined;
}
export function fail(
  node: Pick<XmlNode, "part" | "name">,
  message: string,
  code: "InvalidDocument" | "UnsupportedFeature" = "InvalidDocument",
): never {
  throw new DocumentError(code, message, [
    { code, path: [node.part.slice(0, 128), node.name.slice(0, 128)], message },
  ]);
}
export function unsupported(
  node: XmlNode,
  message = "The document contains an unsupported OOXML feature.",
  attribute?: string,
): never {
  const error = new DocumentError("UnsupportedFeature", message, [
    { code: "UnsupportedFeature", path: [node.part.slice(0, 128), node.name.slice(0, 128)], message },
  ]);
  unsupportedSources.set(error, { node, ...(attribute ? { attribute } : {}) });
  throw error;
}
export function attr(
  node: XmlNode | undefined,
  name: string,
  uri: string = W,
): string | undefined {
  return node?.attrs.get(uri + " " + name);
}
/** Namespace bindings are syntax, not unsupported feature attributes. */
export function checkAttributes(node: XmlNode, names: readonly string[], uri = ""): void {
  for (const key of node.attrs.keys()) {
    if (key.startsWith("http://www.w3.org/2000/xmlns/ ")) continue;
    if (!names.some(name => key === uri + " " + name)) {
      try { unsupported(node, "This OOXML attribute is not supported."); }
      catch (error) { if (typeof error === "object" && error) unsupportedSources.set(error, { node, attribute: key }); throw error; }
    }
  }
}
export function elements(
  node: XmlNode | undefined,
  name?: string,
  uri: string = W,
): XmlNode[] {
  return (
    node?.content.filter(
      (item): item is XmlNode =>
        typeof item !== "string" &&
        item.uri === uri &&
        (name === undefined || item.name === name),
    ) ?? []
  );
}
export function child(
  node: XmlNode | undefined,
  name: string,
  uri: string = W,
): XmlNode | undefined {
  const found = elements(node, name, uri);
  if (found.length > 1) fail(node!, "An OOXML singleton is repeated.");
  return found[0];
}
export function text(node: XmlNode | undefined): string {
  return (
    node?.content
      .map((item) => (typeof item === "string" ? item : text(item)))
      .join("") ?? ""
  );
}
export function number(
  value: string | undefined,
  node: XmlNode,
  fallback?: number,
): number {
  if (value === undefined && fallback !== undefined) return fallback;
  if (
    value === undefined ||
    !/^-?\d+$/.test(value) ||
    !Number.isSafeInteger(Number(value))
  )
    fail(node, "An OOXML numeric value is invalid.");
  return Number(value);
}
export function on(node: XmlNode | undefined): boolean {
  if (!node) return false;
  const value = attr(node, "val");
  if (value === undefined || ["1", "true", "on"].includes(value)) return true;
  if (["0", "false", "off"].includes(value)) return false;
  return fail(node, "An OOXML on/off value is invalid.");
}
export function check(
  node: XmlNode | undefined,
  names: readonly string[],
  uri: string = W,
): void {
  if (!node) return;
  for (const item of node.content) {
    if (typeof item === "string") {
      if (item.trim()) fail(node, "Unexpected OOXML text content.");
    } else if (item.uri !== uri || !names.includes(item.name))
      unsupported(item);
  }
}
export function xml(
  bytes: Uint8Array,
  part: string,
  limits: DocumentReadLimits,
  budget: XmlBudget,
  maxPartNodes: number = limits.maxNodes,
): XmlNode {
  const codecs = globalThis as unknown as {
    TextDecoder: new (
      label: string,
      options: { fatal: boolean },
    ) => { decode(value: Uint8Array): string };
  };
  let source: string;
  try {
    const encoding =
      bytes[0] === 0xff && bytes[1] === 0xfe
        ? "utf-16le"
        : bytes[0] === 0xfe && bytes[1] === 0xff
          ? "utf-16be"
          : "utf-8";
    source = new codecs.TextDecoder(encoding, { fatal: true }).decode(bytes);
  } catch {
    return fail(
      { part, name: "XML" },
      "An XML part has invalid text encoding.",
    );
  }
  const parser = new SaxesParser({ xmlns: true });
  const stack: XmlNode[] = [];
  let root: XmlNode | undefined, partNodes = 0;
  parser.on("doctype", () =>
    fail(
      { part, name: "DOCTYPE" },
      "DTD and external entities are not allowed.",
    ),
  );
  parser.on("processinginstruction", () =>
    fail(
      { part, name: "XML" },
      "XML processing instructions are not supported.",
      "UnsupportedFeature",
    ),
  );
  parser.on("error", () =>
    fail({ part, name: "XML" }, "An XML part is malformed."),
  );
  parser.on("opentag", (tag) => {
    if (++budget.nodes > limits.maxNodes || ++partNodes > maxPartNodes || stack.length + 1 > limits.maxDepth)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "XML exceeds the configured node or depth budget.",
      );
    const node: XmlNode = {
      name: tag.local,
      uri: aliases[tag.uri] ?? tag.uri,
      attrs: new Map(
        Object.values(tag.attributes).map((item) => [
          (aliases[item.uri] ?? item.uri) + " " + item.local,
          item.value,
        ]),
      ),
      namespaces: {
        ...(stack[stack.length - 1]?.namespaces ?? {}),
        ...Object.fromEntries(
          Object.entries(tag.ns).map(([key, value]) => [
            key,
            aliases[value] ?? value,
          ]),
        ),
      },
      content: [],
      part,
    };
    if (stack.length) stack[stack.length - 1]!.content.push(node);
    else {
      if (root) fail(node, "An XML part has multiple roots.");
      root = node;
    }
    stack.push(node);
  });
  parser.on("closetag", () => {
    stack.pop();
  });
  const addText = (value: string): void => {
    if (stack.length) {
      if (++budget.nodes > limits.maxNodes || ++partNodes > maxPartNodes)
        throw new DocumentError(
          "ResourceLimitExceeded",
          "XML exceeds the configured node budget.",
        );
      const content = stack[stack.length - 1]!.content;
      if (typeof content[content.length - 1] === "string")
        content[content.length - 1] =
          String(content[content.length - 1]) + value;
      else content.push(value);
    }
  };
  parser.on("text", addText);
  parser.on("cdata", addText);
  try {
    parser.write(source).close();
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    return fail({ part, name: "XML" }, "An XML part could not be parsed.");
  }
  if (!root) return fail({ part, name: "XML" }, "An XML part has no root.");
  return root;
}
