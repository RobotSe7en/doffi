import type { DeepReadonly, WordChart } from "@doffi/core";
import { A, C, attr, child, elements, number } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import { attributes, empty } from "./export/xml.js";
type Style = NonNullable<WordChart["textStyle"]>;
export function chartThemeFont(
  value: string | undefined,
  role: string,
  theme: XmlNode | undefined,
  language: string,
): string | undefined {
  if (!value) return undefined;
  const match = /^\+(mj|mn)-(lt|ea|cs)$/.exec(value);
  if (!match) return value;
  const scheme = child(child(theme, "themeElements", A), "fontScheme", A),
    group = child(scheme, match[1] === "mj" ? "majorFont" : "minorFont", A),
    base = attr(
      child(group, match[2] === "lt" ? "latin" : match[2] === "ea" ? "ea" : "cs", A),
      "typeface",
      "",
    );
  if (base) return base;
  const script = /^ja/i.test(language)
    ? "Jpan"
    : /^ko/i.test(language)
      ? "Hang"
      : /^zh-(?:TW|HK|MO|Hant)/i.test(language)
        ? "Hant"
        : /^ar/i.test(language)
          ? "Arab"
          : /^he/i.test(language)
            ? "Hebr"
            : role === "cs"
              ? "Arab"
              : "Hans";
  return attr(
    elements(group, "font", A).find((node) => attr(node, "script", "") === script),
    "typeface",
    "",
  );
}
export function readChartTextStyle(
  source: XmlNode | undefined,
  theme: XmlNode | undefined,
  language = "zh-CN",
): Style | undefined {
  const rich = source?.name === "title" ? child(child(source, "tx", C), "rich", C) : source,
    p = child(rich, "p", A),
    properties = child(child(p, "pPr", A), "defRPr", A) ?? child(child(p, "r", A), "rPr", A);
  if (!properties) return undefined;
  const font = (value: string | undefined, role: string): string | undefined =>
    chartThemeFont(
      value,
      role,
      theme,
      (role === "eastAsia" ? attr(properties, "altLang", "") : undefined) ??
        attr(properties, "lang", "") ??
        language,
    );
  const latin = font(attr(child(properties, "latin", A), "typeface", ""), "ascii"),
    ea = font(attr(child(properties, "ea", A), "typeface", ""), "eastAsia"),
    cs = font(attr(child(properties, "cs", A), "typeface", ""), "cs"),
    size = attr(properties, "sz", ""),
    color = attr(child(child(properties, "solidFill", A), "srgbClr", A), "val", ""),
    boolean = (name: string): boolean | undefined => {
      const value = attr(properties, name, "");
      return value === undefined ? undefined : ["1", "true", "on"].includes(value);
    };
  const result: Style = {
    ...(latin || ea || cs
      ? {
          fonts: {
            ...(latin ? { ascii: latin, hAnsi: latin } : {}),
            ...(ea ? { eastAsia: ea } : {}),
            ...(cs ? { cs } : {}),
          },
        }
      : {}),
    ...(size ? { fontSize: number(size, properties) / 100 } : {}),
    ...(color ? { color: "#" + color } : {}),
    ...(boolean("b") !== undefined ? { bold: boolean("b") } : {}),
    ...(boolean("i") !== undefined ? { italic: boolean("i") } : {}),
    ...(attr(properties, "lang", "") ? { language: attr(properties, "lang", "") } : {}),
  };
  return Object.keys(result).length ? result : undefined;
}
export function chartRunStyle(style: DeepReadonly<Style>): string {
  return (
    attributes({
      sz: style.fontSize === undefined ? undefined : Math.round(style.fontSize * 100),
      b: style.bold === undefined ? undefined : style.bold ? 1 : 0,
      i: style.italic === undefined ? undefined : style.italic ? 1 : 0,
      lang: style.language,
    }) +
    ">" +
    (style.color && style.color !== "auto"
      ? "<a:solidFill>" + empty("a:srgbClr", { val: style.color.slice(1) }) + "</a:solidFill>"
      : "") +
    (style.fonts?.ascii || style.fonts?.hAnsi
      ? empty("a:latin", { typeface: style.fonts.ascii ?? style.fonts.hAnsi })
      : "") +
    (style.fonts?.eastAsia ? empty("a:ea", { typeface: style.fonts.eastAsia }) : "") +
    (style.fonts?.cs ? empty("a:cs", { typeface: style.fonts.cs }) : "")
  );
}
export function writeChartTextStyle(style: DeepReadonly<Style> | undefined): string {
  return style
    ? "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr" +
        chartRunStyle(style) +
        "</a:defRPr></a:pPr><a:endParaRPr/></a:p></c:txPr>"
    : "";
}
