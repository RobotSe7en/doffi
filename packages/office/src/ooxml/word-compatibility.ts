import { attr } from "./xml.js";
import type { XmlNode } from "./xml.js";

/** Editable settings share the retained compat container without replacing
 * producer flags that the current document model does not interpret. */
export const wordCompatibilityNames = [
  "noColumnBalance",
  "adjustLineHeightInTable",
  "doNotWrapTextWithPunct",
] as const;

export const wordCompatibilityUri = "http://schemas.microsoft.com/office/word";

export function wordCompatibilityKey(node: XmlNode): string {
  return node.name === "compatSetting" &&
    attr(node, "name") === "compatibilityMode" &&
    attr(node, "uri") === wordCompatibilityUri
    ? "compatibilityMode"
    : node.name;
}
