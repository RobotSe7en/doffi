import { DocumentError } from "@doffi/core";
import type { SourceStyleMap } from "@doffi/core";
import { fontName, color, length } from "../export/values.js";
import { empty } from "../export/xml.js";
import { cellAlignmentChoices, cellBorderKinds, cellFillPatterns } from "../xlsx/appearance.js";
import { themeSlots } from "../xlsx/colors.js";
import { bool, element, xstring } from "./xml.js";

type Css = SourceStyleMap;
type Attrs = Parameters<typeof empty>[1];
type FormatLoss = (message: string) => void;
export const flag = (value: string | undefined): number | undefined =>
  value === "true" ? 1 : value === "false" ? 0 : undefined;
const integer = (value: string | undefined): number | undefined =>
  value !== undefined && /^[+-]?\d+$/.test(value) ? Number(value) : undefined;
function boundedInteger(value: string | undefined, minimum: number, maximum: number, label: string): number | undefined {
  if (value === undefined) return undefined;
  const number = integer(value);
  if (number === undefined || number < minimum || number > maximum)
    throw new DocumentError("UnsupportedFeature", "The resolved " + label + " is outside the XLSX range.");
  return number;
}
function fraction(value: string | undefined, label: string): number | undefined {
  if (value === undefined) return undefined;
  const number = Number(value);
  if (!value.trim() || !Number.isFinite(number) || number < 0 || number > 1)
    throw new DocumentError("UnsupportedFeature", "The resolved " + label + " must be a fraction between zero and one.");
  return number;
}

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

export function font(raw: Css, css: Css, mode: "cell" | "run" | "differential", loss: FormatLoss): string {
  let children = "";
  const add = (tag: string, val: string | number | undefined): void => {
    if (val !== undefined) children += empty(tag, { val });
  };
  const roles = ["ascii", "h-ansi", "east-asia", "complex-script"];
  const theme = css["--office-font-scheme"];
  const themed = (theme === "major" || theme === "minor") && roles.every(role => {
    const authored = raw["--office-font-" + role];
    return !authored || authored.startsWith(`var(--office-${theme}-${role},`) || authored === `var(--office-${theme}-${role})`;
  });
  const faces = [...new Set(roles.map(role => css["--office-font-" + role]).filter((face): face is string => !!face).map(fontName))];
  const original = css["--office-font-source-name"] && fontName(css["--office-font-source-name"]);
  const fallback = css["font-family"] && fontName(css["font-family"]);
  const face = themed ? original : faces[0] || fallback || original;
  if (face) add(mode === "run" ? "rFont" : "name", xstring(face));
  if (!themed && faces.length > 1) loss("An XLSX font can specify one explicit typeface. Separate script-specific typefaces remain in the document package.");
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
  add("scheme", themed ? theme : theme === undefined ? undefined : "none");
  return children || mode === "cell" ? element(mode === "run" ? "rPr" : "font", children) : "";
}

export const pattern = (kind: string): string => element("fill", empty("patternFill", { patternType: kind }));
export function fill(raw: Css, css: Css): string {
  const prefix = "--office-cell-fill";
  const kind = css[prefix + "-kind"] ?? (css["background-color"] && color(css["background-color"]) ? "solid" : undefined);
  if (!kind) return "";
  if (kind !== "gradient") {
    if (!cellFillPatterns.has(kind)) throw new DocumentError("UnsupportedFeature", "The resolved cell fill pattern is outside the XLSX range.");
    const foreground = colorTag("fgColor", raw, css, prefix, kind === "solid" ? css["background-color"] : undefined);
    const background = colorTag("bgColor", raw, css, prefix + "-background");
    if (kind === "solid" && !foreground) throw new DocumentError("UnsupportedFeature", "A solid spreadsheet fill needs an opaque foreground.");
    return element("fill", element("patternFill", foreground + background, { patternType: kind }));
  }
  const gradient = "--office-cell-gradient", count = boundedInteger(css[gradient + "-stops"], 0, Number.MAX_SAFE_INTEGER, "gradient stop count") ?? 0;
  if (count > 4096) throw new DocumentError("ResourceLimitExceeded", "A cell gradient exceeds its stop budget.");
  const type = css[gradient + "-kind"] ?? "linear";
  if (type !== "linear" && type !== "path") throw new DocumentError("UnsupportedFeature", "The resolved cell gradient type is outside the XLSX range.");
  const degrees = css[gradient + "-degree"], degree = degrees === undefined ? undefined : Number(degrees);
  if (degree !== undefined && (!degrees?.trim() || !Number.isFinite(degree)))
    throw new DocumentError("UnsupportedFeature", "The resolved cell gradient angle must be finite.");
  const attributes = { type, degree, ...Object.fromEntries(["left", "right", "top", "bottom"].map(side =>
    [side, fraction(css[gradient + "-" + side], "cell gradient edge")])) };
  const stops = Array.from({ length: count }, (_, index) => {
    const stop = gradient + "-stop-" + index;
    const paint = colorTag("color", raw, css, stop), position = fraction(css[stop + "-position"], "cell gradient stop position");
    if (!paint || position === undefined)
      throw new DocumentError("UnsupportedFeature", "Each XLSX cell gradient stop requires a color and a position.");
    return element("stop", paint, { position });
  });
  return element("fill", element("gradientFill", stops.join(""), attributes));
}

const borderSides = ["start", "end", "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"] as const;
export function border(raw: Css, css: Css, loss: FormatLoss): string {
  const prefix = "--office-cell-border";
  const attrs = {
    diagonalUp: flag(css[prefix + "-diagonal-up"]), diagonalDown: flag(css[prefix + "-diagonal-down"]),
    outline: flag(css[prefix + "-outline"]),
  };
  const children = borderSides.map(side => {
    const key = prefix + "-" + side;
    const kind = css[key + "-kind"], cssStyle = css["border-" + side + "-style"];
    let style = kind;
    if (kind && !cellBorderKinds.has(kind)) throw new DocumentError("UnsupportedFeature", "The resolved cell border is outside the XLSX range.");
    if (!style && cssStyle) {
      if (cssStyle === "hidden" || cssStyle === "none") style = "none";
      else if (["dashed", "dotted", "double"].includes(cssStyle)) style = cssStyle;
      else {
        style = "thin";
        if (cssStyle !== "solid") loss("An XLSX cell border cannot represent a CSS " + cssStyle + " border. It is saved as a thin solid line.");
      }
    }
    const paint = colorTag("color", raw, css, key, style && style !== "none" ? css["border-" + side + "-color"] : undefined);
    return style || paint ? element(side, paint, { style: style === "none" ? undefined : style }) : "";
  }).join("");
  return children || Object.values(attrs).some(value => value !== undefined) ? element("border", children, attrs) : "";
}

export function alignment(css: Css, loss: FormatLoss): string {
  const get = (key: string): string | undefined => css["--office-cell-" + key];
  const order = get("reading-order");
  const vertical = css["vertical-align"];
  const horizontal = get("horizontal"), nativeVertical = get("vertical"), cssAlign = css["text-align"];
  if (horizontal !== undefined && !cellAlignmentChoices.horizontal.some(choice => choice === horizontal)
    || nativeVertical !== undefined && !cellAlignmentChoices.vertical.some(choice => choice === nativeVertical)
    || order !== undefined && !cellAlignmentChoices.readingOrder.some(choice => choice === order))
    throw new DocumentError("UnsupportedFeature", "A resolved cell alignment is outside the XLSX range.");
  const rotation = boundedInteger(get("text-rotation"), 0, 255, "cell rotation");
  if (rotation !== undefined && rotation > 180 && rotation !== 255)
    throw new DocumentError("UnsupportedFeature", "The resolved cell rotation is outside the XLSX range.");
  if (!horizontal && cssAlign === "end" && order !== "rtl" && order !== "ltr")
    loss("Context-dependent CSS end alignment is saved as right alignment in XLSX.");
  const attrs = {
    horizontal: horizontal ?? (cssAlign === "end" ? order === "rtl" ? "left" : "right" :
      cssAlign === "left" || cssAlign === "center" || cssAlign === "right" || cssAlign === "justify" ? cssAlign : undefined),
    vertical: nativeVertical ?? (vertical === "middle" ? "center" : vertical === "top" || vertical === "bottom" ? vertical : undefined),
    textRotation: rotation, wrapText: flag(get("wrap")) ?? (css["white-space"] === "nowrap" ? 0 : undefined),
    shrinkToFit: flag(get("shrink-to-fit")), indent: boundedInteger(get("indent"), 0, 255, "cell indent"),
    relativeIndent: boundedInteger(get("relative-indent"), -15, 15, "relative cell indent"),
    justifyLastLine: flag(get("justify-last-line")), readingOrder: order === undefined ? undefined : cellAlignmentChoices.readingOrder.findIndex(choice => choice === order),
  };
  return Object.values(attrs).some(value => value !== undefined) ? empty("alignment", attrs) : "";
}
