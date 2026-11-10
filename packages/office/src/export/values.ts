import { generate, ident, parse } from "css-tree";
import { DocumentError } from "@doffi/core";
export { resolveSourceCssLength as length, resolveSourceCssColor as color } from "@doffi/core";
export type { SourceCssLengthContext as LengthContext } from "@doffi/core";
function unsupported(value: string): never { throw new DocumentError("UnsupportedFeature", "A CSS value cannot be represented in native DOCX: " + value.slice(0, 80)); }
export function fontName(value: string): string {
  const ast = parse(value, { context: "value" }); if (ast.type !== "Value") return unsupported(value);
  const parts: string[] = [];
  for (const item of ast.children) {
    if (item.type === "Operator" && item.value === ",") break;
    parts.push(item.type === "String" ? item.value : item.type === "Identifier" ? ident.decode(item.name) : generate(item));
  }
  const name = parts.join(" ").trim(); if (!name || name.length > 255) return unsupported(value);
  const generic: Record<string, string> = { serif: "Times New Roman", "sans-serif": "Arial", monospace: "Courier New", "system-ui": "Arial", cursive: "Comic Sans MS", fantasy: "Impact" };
  return generic[name.toLowerCase()] ?? name;
}
export function nativeInteger(value: number, factor = 20, minimum = -2_147_483_648): number {
  const result = Math.round(value * factor); if (!Number.isSafeInteger(result) || result < minimum || result > 2_147_483_647) return unsupported(String(value)); return result;
}
