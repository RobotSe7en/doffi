import { WordTypographySchema } from "@doffi/core";
import type { DeepReadonly, WordTypography } from "@doffi/core";
import { empty, value } from "../export/xml.js";
import { attr, child, check, fail, on } from "./xml.js";
import type { XmlNode } from "./xml.js";

const rules = ["noLineBreaksAfter", "noLineBreaksBefore"] as const;
export const wordTypographySettings: readonly [
  "characterSpacingControl",
  "strictFirstAndLastChars",
  ...typeof rules,
] = ["characterSpacingControl", "strictFirstAndLastChars", ...rules];

export function readWordTypography(settings: XmlNode | undefined): WordTypography | undefined {
  const result: Record<string, unknown> = {};
  for (const name of wordTypographySettings) {
    const item = child(settings, name);
    if (!item) continue;
    check(item, []);
    result[name] =
      name === "strictFirstAndLastChars"
        ? on(item)
        : name === "characterSpacingControl"
          ? (attr(item, "val") ?? null)
          : { language: attr(item, "lang"), characters: attr(item, "val") };
  }
  if (!Object.keys(result).length) return;
  const parsed = WordTypographySchema.safeParse(result);
  if (!parsed.success) fail(settings!, "An East Asian typography setting is invalid.");
  return parsed.data;
}

export function writeWordTypography(options: DeepReadonly<WordTypography> | undefined): string {
  if (!options) return "";
  return (
    value("characterSpacingControl", options.characterSpacingControl) +
    value(
      "strictFirstAndLastChars",
      options.strictFirstAndLastChars === undefined
        ? undefined
        : Number(options.strictFirstAndLastChars),
    ) +
    rules
      .map((name) => {
        const rule = options[name];
        return rule
          ? empty("w:" + name, { "w:lang": rule.language, "w:val": rule.characters })
          : "";
      })
      .join("")
  );
}
