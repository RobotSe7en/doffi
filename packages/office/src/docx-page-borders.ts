import { WORD_ART_PAGE_BORDER_STYLES, WordPageBordersSchema } from "@doffi/core";
import type { DeepReadonly, WordPageBorders } from "@doffi/core";
import { W, R, attr, check, elements, number, unsupported } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import { attributes, empty } from "./export/xml.js";

const ART_BORDER_STYLES = new Set<string>(WORD_ART_PAGE_BORDER_STYLES);
export function readWordPageBorders(source: XmlNode | undefined): WordPageBorders | undefined {
  if (!source) return;
  check(source, ["top", "left", "bottom", "right"]);
  const result: Record<string, unknown> = {
    offsetFrom: attr(source, "offsetFrom") ?? "text",
    display: attr(source, "display") ?? "allPages",
    zOrder: attr(source, "zOrder") ?? "front",
  };
  for (const border of elements(source)) {
    check(border, []);
    if (
      [...border.attrs.keys()].some(
        (name) =>
          name.startsWith(R + " ") ||
          ["topLeft", "topRight", "bottomLeft", "bottomRight"].some(
            (key) => name === W + " " + key,
          ),
      )
    )
      unsupported(source, "Art page borders retain their native resources.");
    const borderStyle = attr(border, "val") ?? "single";
    const value: Record<string, unknown> = {
      style: borderStyle,
      width: ART_BORDER_STYLES.has(borderStyle)
        ? number(attr(border, "sz"), border, 12)
        : number(attr(border, "sz"), border, 4) / 8,
      space: number(attr(border, "space"), border, 0),
      color: attr(border, "color") ?? "auto",
    };
    for (const key of ["themeColor", "themeTint", "themeShade"] as const)
      if (attr(border, key) !== undefined) value[key] = attr(border, key);
    for (const key of ["shadow", "frame"] as const)
      if (attr(border, key) !== undefined) {
        const raw = attr(border, key)!;
        if (!["true", "1", "on", "false", "0", "off"].includes(raw))
          unsupported(source, "This page-border effect is preserved natively.");
        value[key] = ["true", "1", "on"].includes(raw);
      }
    result[border.name] = value;
  }
  const parsed = WordPageBordersSchema.safeParse(result);
  if (!parsed.success) unsupported(source, "This page-border pattern is preserved natively.");
  return parsed.data;
}
export function writeWordPageBorders(borders: DeepReadonly<WordPageBorders> | undefined): string {
  if (!borders) return "";
  return (
    "<w:pgBorders" +
    attributes({
      "w:offsetFrom": borders.offsetFrom,
      "w:display": borders.display,
      "w:zOrder": borders.zOrder,
    }) +
    ">" +
    (["top", "left", "bottom", "right"] as const)
      .map((side) => {
        const border = borders[side];
        return border
          ? empty("w:" + side, {
              "w:val": border.style,
              "w:sz": ART_BORDER_STYLES.has(border.style)
                ? Math.max(1, Math.min(31, Math.round(border.width)))
                : Math.round(border.width * 8),
              "w:space": border.space,
              "w:color": border.color,
              "w:themeColor": border.themeColor,
              "w:themeTint": border.themeTint,
              "w:themeShade": border.themeShade,
              "w:shadow": border.shadow === undefined ? undefined : border.shadow ? 1 : 0,
              "w:frame": border.frame === undefined ? undefined : border.frame ? 1 : 0,
            })
          : "";
      })
      .join("") +
    "</w:pgBorders>"
  );
}
