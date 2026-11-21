import { DocumentError } from "@doffi/core";
import type { SourceStyleMap } from "@doffi/core";
import { fontName, color, length } from "../export/values.js";
import { empty } from "../export/xml.js";
import { themeSlots } from "../xlsx/colors.js";
import { bool, element, xstring } from "./xml.js";

type Css = SourceStyleMap;
type Attrs = Parameters<typeof empty>[1];
export const flag = (value: string | undefined): number | undefined =>
  value === "true" ? 1 : value === "false" ? 0 : undefined;
const integer = (value: string | undefined): number | undefined =>
  value !== undefined && /^-?\d+$/.test(value) ? Number(value) : undefined;

export function colorAttrs(raw: Css, css: Css, prefix: string, fallback?: string): Attrs {
  const key = prefix + "-color", source = raw[key];
  const tint = Number(css[prefix + "-tint"] ?? 0) || undefined;
  const theme = /^var\(\s*--office-theme-([\w]+)\s*\)$/.exec(source ?? "")?.[1];
  const index = themeSlots.indexOf(theme ?? "");
  if (index !== -1) return { theme: index, tint };
  const indexed = /^var\(\s*--office-indexed-(\d+)\s*\)$/.exec(source ?? "")?.[1];
  if (indexed !== undefined) return { indexed: Number(indexed), tint };
  if (source === "var(--office-system-foreground)") return { indexed: 64, tint };
  if (source === "var(--office-system-background)") return { indexed: 65, tint };
  const rgb = css[key] || fallback;
  const hex = rgb && color(rgb);
  if (!hex) return undefined;
  const alpha = integer(css[prefix + "-native-alpha"]) ?? 255;
  return { rgb: alpha.toString(16).padStart(2, "0").toUpperCase() + hex.toUpperCase(), tint };
}
export function colorTag(tag: string, raw: Css, css: Css, prefix: string, fallback?: string): string {
  const attrs = colorAttrs(raw, css, prefix, fallback);
  return attrs ? empty(tag, attrs) : "";
}

export function font(raw: Css, css: Css, mode: "cell" | "run" | "differential"): string {
  let children = "";
  const add = (tag: string, val: string | number | undefined): void => {
    if (val !== undefined) children += empty(tag, { val });
  };
  const face = css["--office-font-source-name"] || css["--office-font-ascii"] || css["font-family"];
  if (face) add(mode === "run" ? "rFont" : "name", xstring(fontName(face)));
  add("charset", integer(css["--office-font-charset"]));
  add("family", integer(css["--office-font-family-class"]));
  if (css["font-weight"]) add("b", bool(css["font-weight"] === "bold" || Number(css["font-weight"]) >= 600));
  if (css["font-style"]) add("i", bool(/^italic|^oblique/.test(css["font-style"])));
  const decoration = css["text-decoration-line"];
  add("strike", flag(css["--office-font-strike"]) ?? (decoration ? bool(decoration.includes("line-through")) : undefined));
  const directColor = raw.color && raw.color !== "var(--office-cell-foreground-color)";
  children += directColor
    ? colorTag("color", {}, {}, "--office-cell-foreground", css.color)
    : colorTag("color", raw, css, "--office-cell-foreground", css.color);
  if (css["font-size"]) {
    const size = length(css["font-size"], { font: 11, percent: 11 });
    if (size <= 0 || size > 409) throw new DocumentError("UnsupportedFeature", "A cell font size lies outside the XLSX format range.");
    add("sz", size);
  }
  add("u", css["--office-font-underline"] ?? (decoration ? decoration.includes("underline") ? "single" : "none" : undefined));
  const baseline = css["vertical-align"];
  if (baseline === "super" || baseline === "sub" || baseline === "baseline")
    add("vertAlign", baseline === "super" ? "superscript" : baseline === "sub" ? "subscript" : "baseline");
  add("scheme", css["--office-font-scheme"]);
  return children || mode === "cell" ? element(mode === "run" ? "rPr" : "font", children) : "";
}

export const pattern = (kind: string): string => element("fill", empty("patternFill", { patternType: kind }));
export function fill(raw: Css, css: Css): string {
  const prefix = "--office-cell-fill";
  const kind = css[prefix + "-kind"] ?? (css["background-color"] && color(css["background-color"]) ? "solid" : undefined);
  if (!kind) return "";
  if (kind !== "gradient") {
    const foreground = colorTag("fgColor", raw, css, prefix, kind === "solid" ? css["background-color"] : undefined);
    const background = colorTag("bgColor", raw, css, prefix + "-background");
    if (kind === "solid" && !foreground) throw new DocumentError("UnsupportedFeature", "A solid spreadsheet fill needs an opaque foreground.");
    return element("fill", element("patternFill", foreground + background, { patternType: kind }));
  }
  const gradient = "--office-cell-gradient", count = integer(css[gradient + "-stops"]) ?? 0;
  if (count < 0 || count > 4096) throw new DocumentError("ResourceLimitExceeded", "A cell gradient exceeds its stop budget.");
  const type = css[gradient + "-kind"] ?? "linear";
  const attributes = { type, degree: css[gradient + "-degree"], left: css[gradient + "-left"], right: css[gradient + "-right"],
    top: css[gradient + "-top"], bottom: css[gradient + "-bottom"] };
  const stops = Array.from({ length: count }, (_, index) => {
    const stop = gradient + "-stop-" + index;
    return element("stop", colorTag("color", raw, css, stop), { position: css[stop + "-position"] });
  });
  return element("fill", element("gradientFill", stops.join(""), attributes));
}

const borderSides = ["start", "end", "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"] as const;
export function border(raw: Css, css: Css): string {
  const prefix = "--office-cell-border";
  const attrs = {
    diagonalUp: flag(css[prefix + "-diagonal-up"]), diagonalDown: flag(css[prefix + "-diagonal-down"]),
    outline: flag(css[prefix + "-outline"]),
  };
  const children = borderSides.map(side => {
    const key = prefix + "-" + side;
    const kind = css[key + "-kind"], cssStyle = css["border-" + side + "-style"];
    const style = kind ?? (cssStyle === "solid" ? "thin" : cssStyle);
    const paint = colorTag("color", raw, css, key, style && style !== "none" ? css["border-" + side + "-color"] : undefined);
    return style || paint ? element(side, paint, { style: style === "none" ? undefined : style }) : "";
  }).join("");
  return children || Object.values(attrs).some(value => value !== undefined) ? element("border", children, attrs) : "";
}

export function alignment(css: Css): string {
  const get = (key: string): string | undefined => css["--office-cell-" + key];
  const order = get("reading-order");
  const vertical = css["vertical-align"];
  const attrs = {
    horizontal: get("horizontal") ?? css["text-align"],
    vertical: get("vertical") ?? (vertical === "middle" ? "center" : vertical === "top" || vertical === "bottom" ? vertical : undefined),
    textRotation: integer(get("text-rotation")), wrapText: flag(get("wrap")) ?? (css["white-space"] === "nowrap" ? 0 : undefined),
    shrinkToFit: flag(get("shrink-to-fit")), indent: integer(get("indent")), relativeIndent: integer(get("relative-indent")),
    justifyLastLine: flag(get("justify-last-line")), readingOrder: order === undefined ? undefined : ["context", "ltr", "rtl"].indexOf(order),
  };
  return Object.values(attrs).some(value => value !== undefined) ? empty("alignment", attrs) : "";
}
