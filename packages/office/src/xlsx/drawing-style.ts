import { drawingColorTransformNames } from "@doffi/core";
import { A, check, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, integer, one, value } from "./values.js";
import type { Css } from "../ooxml/styles.js";
import { color, effects, fill, onlyColor, outline, paint } from "../drawingml/styles.js";
import type { XlsxStyles } from "./styles.js";

function empty(node: XmlNode, names: readonly string[] = []): void { attributes(node, names); check(node, [], A); }
function enumeration(node: XmlNode, name: string, values: readonly string[]): void {
  const raw = value(node, name); if (raw !== undefined && !values.includes(raw)) fail(node, "A DrawingML appearance value is invalid.");
}
function choice(node: XmlNode, names: readonly string[]): void {
  if (elements(node, undefined, A).filter(item => names.includes(item.name)).length > 1) fail(node, "A DrawingML appearance choice is repeated.");
}
/** Keep the native order for colors represented by the shared transform sequence.
 * Legacy color declarations support a fixed order. */
export function validateDrawingColor(node: XmlNode): void {
  if (node.uri !== A || !["srgbClr", "sysClr", "schemeClr"].includes(node.name)) unsupported(node, "This drawing color is not supported.");
  attributes(node, node.name === "sysClr" ? ["val", "lastClr"] : ["val"]);
  check(node, [...drawingColorTransformNames, "alpha"], A);
  const transforms = elements(node, undefined, A);
  for (const transform of transforms) attributes(transform, ["val"]);
  if (color(node).transforms) return;
  const order = ["lumMod", "lumOff", "shade", "tint"];
  let previous = -1;
  for (const transform of transforms) {
    if (transform.name === "alpha") continue;
    const index = order.indexOf(transform.name);
    if (index < previous) unsupported(node, "This ordered color-transform sequence cannot be represented by the drawing appearance source.");
    previous = index;
  }
}
export function validateDrawingColorContainer(node: XmlNode, optional = false): void {
  check(node, ["srgbClr", "sysClr", "schemeClr"], A);
  const colors = elements(node, undefined, A);
  if (colors.length > 1 || !optional && colors.length !== 1) fail(node, "A drawing appearance requires exactly one color.");
  colors.forEach(validateDrawingColor);
}
export function validateDrawingFill(node: XmlNode): void {
  if (node.uri !== A) unsupported(node);
  if (node.name === "noFill") { empty(node); return; }
  if (node.name === "solidFill") { attributes(node, []); validateDrawingColorContainer(node); return; }
  if (node.name !== "gradFill") unsupported(node, "This drawing fill is not supported.");
  attributes(node, ["flip", "rotWithShape"]); check(node, ["gsLst", "lin", "tileRect"], A);
  if (value(node, "flip") !== undefined && value(node, "flip") !== "none") unsupported(node, "Flipped gradient tiles are not supported.");
  boolean(node, "rotWithShape", true);
  const stops = one(node, "gsLst", A), linear = one(node, "lin", A), tile = one(node, "tileRect", A);
  if (!stops || !linear) unsupported(node, "Only explicit linear drawing gradients are supported.");
  attributes(stops, []); check(stops, ["gs"], A);
  const values = elements(stops, "gs", A);
  if (values.length < 2 || values.length > 1024) fail(stops, "A drawing gradient requires between 2 and 1024 stops.");
  for (const stop of values) { attributes(stop, ["pos"]); integer(stop, "pos", 0, 100000); validateDrawingColorContainer(stop); }
  empty(linear, ["ang", "scaled"]); integer(linear, "ang", 0, 21600000 - 1, 0); boolean(linear, "scaled", false);
  if (tile) { empty(tile, ["l", "r", "t", "b"]); for (const side of ["l", "r", "t", "b"]) if (integer(tile, side, -2147483648, 2147483647, 0)) unsupported(tile, "Offset gradient tiles are not supported."); }
}
export function validateDrawingLine(node: XmlNode): void {
  attributes(node, ["w", "cap", "cmpd", "algn"]);
  check(node, ["noFill", "solidFill", "prstDash", "round", "bevel", "miter", "headEnd", "tailEnd"], A);
  integer(node, "w", 0, 20116800, 9525);
  enumeration(node, "cap", ["rnd", "sq", "flat"]); enumeration(node, "cmpd", ["sng", "dbl", "thickThin", "thinThick", "tri"]); enumeration(node, "algn", ["ctr", "in"]);
  choice(node, ["noFill", "solidFill"]); choice(node, ["round", "bevel", "miter"]);
  for (const item of elements(node, undefined, A)) {
    one(node, item.name, A);
    if (["noFill", "solidFill"].includes(item.name)) validateDrawingFill(item);
    else if (item.name === "prstDash") empty(item, ["val"]);
    else if (item.name === "miter") { empty(item, ["lim"]); integer(item, "lim", 0, 2147483647, 800000); }
    else if (["headEnd", "tailEnd"].includes(item.name)) empty(item, ["type", "w", "len"]);
    else empty(item);
  }
}
export function validateDrawingEffects(node: XmlNode): void {
  attributes(node, []); check(node, ["outerShdw"], A);
  const shadow = one(node, "outerShdw", A); if (!shadow) return;
  attributes(shadow, ["blurRad", "dist", "dir", "sx", "sy", "kx", "ky", "algn", "rotWithShape"]);
  validateDrawingColorContainer(shadow);
  integer(shadow, "dir", 0, 21600000 - 1, 0);
  for (const name of ["blurRad", "dist"]) integer(shadow, name, 0, 27273042316900, 0);
  enumeration(shadow, "algn", ["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"]);
  boolean(shadow, "rotWithShape", true);
}

export function drawingAppearance(source: XmlNode | undefined, properties: XmlNode, styles: XlsxStyles): string {
  const classes: string[] = [], line = one(properties, "ln", A), effect = one(properties, "effectLst", A);
  const fills = elements(properties, undefined, A).filter(node => ["noFill", "solidFill", "gradFill"].includes(node.name));
  if (fills.length > 1) fail(properties, "An object fill is repeated.");
  fills.forEach(validateDrawingFill); if (line) validateDrawingLine(line); if (effect) validateDrawingEffects(effect);
  const css: Css = { ...fill(properties), ...outline(line), ...effects(effect) };
  if (source) {
    attributes(source, []); check(source, ["lnRef", "fillRef", "effectRef", "fontRef"], A);
    for (const [tag, kind] of [["lnRef", "line"], ["fillRef", "fill"], ["effectRef", "effect"]] as const) {
      const reference = one(source, tag, A); if (!reference) fail(source, "A native shape style reference is incomplete.");
      attributes(reference, ["idx"]); validateDrawingColorContainer(reference, true);
      const index = integer(reference, "idx"), background = kind === "fill" && index >= 1000, list = background ? "background" : kind;
      if (index === 0 || background && index === 1000) { if (kind !== "effect") css[`--office-${kind}-kind`] ??= "none"; }
      else {
        const name = styles.drawingTheme[list][index - (background ? 1001 : 1)];
        if (!name) fail(reference, "A drawing references an undefined theme style."); classes.push(name);
      }
      if (elements(reference, undefined, A).length) {
        const color = onlyColor(reference); if (color.css.includes("--office-placeholder-color")) fail(reference, "A shape style color cannot refer to itself.");
        paint(css, `--office-${list}-reference`, color);
      }
    }
    const font = one(source, "fontRef", A); if (!font) fail(source, "A shape style font reference is missing.");
    attributes(font, ["idx"]); validateDrawingColorContainer(font, true);
    const role = value(font, "idx"); if (!role || !["none", "major", "minor"].includes(role)) fail(font, "A shape font reference is invalid.");
    if (role !== "none") {
      for (const script of ["ascii", "h-ansi", "east-asia", "complex-script"]) css[`--office-font-${script}`] = `var(--office-${role}-${script})`;
      css["font-family"] = "var(--office-font-ascii),var(--office-font-east-asia)";
    }
    if (elements(font, undefined, A).length) {
      const color = onlyColor(font); if (color.css.includes("--office-placeholder-color")) fail(font, "A shape font color cannot refer to itself.");
      paint(css, "--office-text", color); css.color = color.css;
    }
  }
  const direct = styles.objectStyle(css); if (direct) classes.push(direct);
  return classes.join(" ");
}
