import { DocumentError } from "../errors.js";

/** A source-level Office run override. Ordinary CSS `none` continues to obey
 * CSS ancestor-paint propagation; only an explicit native edit resets it. */
export const TEXT_DECORATION_RESET = "--office-font-decoration-reset";
export function textFormattingProperties(properties: Readonly<Record<string, string>>): Readonly<Record<string, string>> {
  if (properties[TEXT_DECORATION_RESET] !== undefined) { resetParentDecorations("none", properties[TEXT_DECORATION_RESET]); return properties; }
  if ([properties["text-decoration"], properties["text-decoration-line"]].some(value => value?.replace(/\s*!\s*important\s*$/i, "").trim().toLowerCase() === "none")) return { ...properties, [TEXT_DECORATION_RESET]: "all" };
  return properties;
}
export function resetParentDecorations(lines: string, value: string): string {
  const tokens = value.trim().toLowerCase().split(/\s+/);
  if (!tokens.length || tokens.some(token => !["all", "none", "underline", "overline", "line-through"].includes(token))
    || new Set(tokens).size !== tokens.length || tokens.length > 1 && tokens.some(token => token === "all" || token === "none"))
    throw new DocumentError("UnsupportedFeature", "The native text decoration reset is invalid.");
  if (tokens[0] === "none") return lines;
  if (tokens[0] === "all") return "none";
  return lines.split(/\s+/).filter(line => line !== "none" && !tokens.includes(line)).join(" ") || "none";
}
