import { DocumentError, createAssetDescriptor } from "@doffi/core";
import type { WordNativeContent } from "@doffi/core";
import type { OpcPackage } from "./ooxml/package.js";
import { W } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";

export const CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
export const bindingKey = (id: string): string => id.replace(/[{}]/g, "").toUpperCase();
type Part = WordNativeContent["parts"][number];

/** Data bindings name a datastore item rather than an r:id. Index those edges
 * in the Office adapter so core and clipboard graph traversal remain opaque. */
export function indexNativeCustomXml(
  opc: OpcPackage,
  main: string,
): Map<string, Part["customXml"] & { sourcePart: string }> {
  const result = new Map<string, Part["customXml"] & { sourcePart: string }>();
  for (const relation of opc.relationships(main).values()) {
    if (relation.external || !relation.type.endsWith("/customXml")) continue;
    const properties = [...opc.relationships(relation.target).values()].filter(
      (r) => !r.external && r.type.endsWith("/customXmlProps"),
    );
    if (properties.length > 1)
      throw new DocumentError("InvalidDocument", "A custom XML item has multiple property parts.");
    if (!properties.length) continue;
    const root = opc.read(properties[0]!.target),
      id = root.attrs.get(CUSTOM_XML + " itemID");
    if (root.uri !== CUSTOM_XML || root.name !== "datastoreItem" || !id)
      throw new DocumentError("InvalidDocument", "A custom XML item has invalid properties.");
    const key = bindingKey(id);
    if (result.has(key))
      throw new DocumentError("InvalidDocument", "Custom XML datastore identities must be unique.");
    result.set(key, {
      storeItemId: id,
      sourcePart: relation.target,
      propertiesPart: properties[0]!.target,
    });
  }
  return result;
}

export function nativeDataBindings(root: XmlNode): string[] {
  const ids = new Set<string>();
  const visit = (node: XmlNode): void => {
    if (node.uri === W && node.name === "dataBinding") {
      const id = node.attrs.get(W + " storeItemID");
      if (id) ids.add(bindingKey(id));
    }
    for (const child of node.content) if (typeof child !== "string") visit(child);
  };
  visit(root);
  return [...ids];
}

export function nativeSemanticReferences(root: XmlNode): {
  styleReferences: Record<string, string>;
  numberingReferences: Record<string, string>;
} {
  const styleReferences: Record<string, string> = Object.create(null) as Record<string, string>,
    numberingReferences: Record<string, string> = Object.create(null) as Record<string, string>;
  const visit = (node: XmlNode): void => {
    const value = node.attrs.get(W + " val");
    if (node.uri === W && value) {
      if (["pStyle", "rStyle", "tblStyle"].includes(node.name)) styleReferences[value] = value;
      if (node.name === "numId" && value !== "0") numberingReferences[value] = value;
    }
    for (const child of node.content) if (typeof child !== "string") visit(child);
  };
  visit(root);
  return { styleReferences, numberingReferences };
}

/** An unresolved copied binding stays unresolved instead of accidentally
 * acquiring a same-named receiver item. The result is stable for this source. */
export function isolatedStoreItemId(scope: string, id: string): string {
  const Encoder = (
      globalThis as unknown as { TextEncoder: new () => { encode(s: string): Uint8Array } }
    ).TextEncoder,
    bytes = new Encoder().encode(scope + "\n" + bindingKey(id)),
    hash = createAssetDescriptor("assets/binding.bin", "application/octet-stream", bytes).sha256;
  return (
    "{" +
    [
      hash.slice(0, 8),
      hash.slice(8, 12),
      "4" + hash.slice(13, 16),
      "8" + hash.slice(17, 20),
      hash.slice(20, 32),
    ]
      .join("-")
      .toUpperCase() +
    "}"
  );
}
