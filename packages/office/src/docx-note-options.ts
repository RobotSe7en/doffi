import { WordNoteOptionsSchema } from "@doffi/core";
import type { DeepReadonly, WordMetadata } from "@doffi/core";
import { attr, child, check, unsupported, number } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import { empty } from "./export/xml.js";
type Options = NonNullable<WordMetadata["footnoteOptions"]>;
export function readNoteOptions(source: XmlNode | undefined): Options | undefined {
  if (!source) return;
  check(source, ["pos", "numFmt", "numStart", "numRestart", "footnote", "endnote"]);
  const values: Record<string, string | number> = {};
  for (const [key, name] of [
    ["position", "pos"],
    ["format", "numFmt"],
    ["start", "numStart"],
    ["restart", "numRestart"],
  ] as const) {
    const property = child(source, name),
      raw = attr(property, "val");
    if (raw === undefined) continue;
    const value = key === "start" ? number(raw, property!) : raw,
      parsed = WordNoteOptionsSchema.shape[key].safeParse(value);
    if (!parsed.success)
      unsupported(property!, "This note numbering or placement is preserved natively.");
    values[key] = value;
  }
  return WordNoteOptionsSchema.parse(values);
}
export function writeNoteOptions(
  kind: "footnote" | "endnote",
  options: DeepReadonly<Options> | undefined,
  specialIds: readonly number[] = [],
): string {
  if (!options && !specialIds.length) return "";
  const val = (name: string, value: string | number | undefined): string =>
    value === undefined ? "" : empty("w:" + name, { "w:val": value });
  return (
    "<w:" +
    kind +
    "Pr>" +
    val("pos", options?.position) +
    val("numFmt", options?.format) +
    val("numStart", options?.start) +
    val("numRestart", options?.restart) +
    specialIds.map((id) => empty("w:" + kind, { "w:id": id })).join("") +
    "</w:" +
    kind +
    "Pr>"
  );
}
