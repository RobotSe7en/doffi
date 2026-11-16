import { DEFAULT_DOCUMENT_READ_LIMITS, DocumentError } from "@doffi/core";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships",
  W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  DS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
export const nativeBindingKey = (id: string): string => id.replace(/[{}]/g, "").toUpperCase();
/** The legacy package resource is generic XML. Select exactly its keyed element;
 * never put the shared collection or its unselected siblings on the clipboard. */
export class ClipboardNativeXml {
  private count = 0;
  parse(source: string): Element {
    if (/<!DOCTYPE|<\?(?!xml\s)/i.test(source))
      throw new DocumentError(
        "InvalidDocument",
        "Native clipboard XML cannot contain DTDs or processing instructions.",
      );
    const document = new DOMParser().parseFromString(source, "application/xml");
    if (document.getElementsByTagName("parsererror").length || !document.documentElement)
      throw new DocumentError("InvalidDocument", "The native clipboard XML is malformed.");
    const visit = (node: Element, depth: number): void => {
      if (
        ++this.count > DEFAULT_DOCUMENT_READ_LIMITS.maxNodes ||
        depth > DEFAULT_DOCUMENT_READ_LIMITS.maxDepth
      )
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Native clipboard XML exceeds its node or depth budget.",
        );
      for (const child of node.children) visit(child, depth + 1);
    };
    visit(document.documentElement, 1);
    return document.documentElement;
  }
  fragment(root: Element, key?: string): string {
    if (!key) return new XMLSerializer().serializeToString(root);
    const entries = new Map<string, Element>();
    for (const item of root.children) {
      if (item.namespaceURI !== "urn:doffi:word:native-fragments" || item.localName !== "item")
        continue;
      const id = item.getAttribute("id");
      if (!id || item.children.length !== 1 || entries.has(id))
        throw new DocumentError("InvalidDocument", "The legacy native collection is invalid.");
      entries.set(id, item.firstElementChild!);
    }
    const selected = entries.get(key);
    if (!selected)
      throw new DocumentError("InvalidDocument", "The legacy native collection entry is missing.");
    return new XMLSerializer().serializeToString(selected);
  }
  index(source: string): {
    relationshipIds: string[];
    bindings: string[];
    styleReferences: Record<string, string>;
    numberingReferences: Record<string, string>;
  } {
    const root = this.parse(source),
      relations = new Set<string>(),
      bindings = new Set<string>(),
      styleReferences: Record<string, string> = Object.create(null) as Record<string, string>,
      numberingReferences: Record<string, string> = Object.create(null) as Record<string, string>;
    for (const node of [root, ...root.querySelectorAll("*")]) {
      for (const attribute of node.attributes)
        if ([R, STRICT_R].includes(attribute.namespaceURI ?? "")) relations.add(attribute.value);
      if (
        [W, "http://purl.oclc.org/ooxml/wordprocessingml/main"].includes(node.namespaceURI ?? "") &&
        node.localName === "dataBinding"
      ) {
        const id = node.getAttributeNS(node.namespaceURI, "storeItemID");
        if (id) bindings.add(nativeBindingKey(id));
      }
      if (
        [W, "http://purl.oclc.org/ooxml/wordprocessingml/main"].includes(node.namespaceURI ?? "")
      ) {
        const value = node.getAttributeNS(node.namespaceURI, "val");
        if (value && ["pStyle", "rStyle", "tblStyle"].includes(node.localName))
          styleReferences[value] = value;
        if (value && node.localName === "numId" && value !== "0")
          numberingReferences[value] = value;
      }
    }
    return {
      relationshipIds: [...relations],
      bindings: [...bindings],
      styleReferences,
      numberingReferences,
    };
  }
  storeItemId(source: string): string | undefined {
    const root = this.parse(source);
    return root.namespaceURI === DS && root.localName === "datastoreItem"
      ? (root.getAttributeNS(DS, "itemID") ?? undefined)
      : undefined;
  }
}
