import { check, elements, fail, S, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { attributes, boolean, integer, number, one, value } from "./values.js";
import { colorDeclarations } from "./colors.js";
import type { SpreadsheetColors } from "./colors.js";

export const cellAlignmentChoices = {
  horizontal: ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"],
  vertical: ["top", "center", "bottom", "justify", "distributed"],
  readingOrder: ["context", "ltr", "rtl"],
} as const;

export function alignmentDeclarations(node: XmlNode | undefined, full: boolean): Css {
  const css: Css = {};
  if (node) {
    attributes(node, ["horizontal", "vertical", "textRotation", "wrapText", "shrinkToFit", "indent", "relativeIndent", "justifyLastLine", "readingOrder"]); check(node, [], S);
    const rotation = integer(node, "textRotation", 0, 255, 0);
    if (rotation > 180 && rotation !== 255) fail(node, "Cell text rotation must be 0–180 or the vertical-text value 255.");
  }
  if (value(node, "textRotation") !== undefined || full) css["--office-cell-text-rotation"] = String(node ? integer(node, "textRotation", 0, 255, 0) : 0);
  for (const [key, choices, fallback] of [
    ["horizontal", cellAlignmentChoices.horizontal, "general"],
    ["vertical", cellAlignmentChoices.vertical, "bottom"],
  ] as const) {
    const setting = value(node, key);
    if (setting !== undefined && !choices.some(choice => choice === setting)) fail(node!, "A native cell alignment is invalid.");
    if (setting !== undefined || full) css[`--office-cell-${key}`] = setting ?? fallback;
  }
  for (const [key, property] of [["wrapText", "--office-cell-wrap"], ["shrinkToFit", "--office-cell-shrink-to-fit"], ["justifyLastLine", "--office-cell-justify-last-line"]] as const) {
    if (value(node, key) !== undefined || full) css[property] = String(boolean(node, key));
  }
  if (node && value(node, "indent") !== undefined || full) css["--office-cell-indent"] = node ? String(integer(node, "indent", 0, 255, 0)) : "0";
  if (node && value(node, "relativeIndent") !== undefined) css["--office-cell-relative-indent"] = String(integer(node, "relativeIndent", -15, 15));
  if (node && value(node, "readingOrder") !== undefined || full) css["--office-cell-reading-order"] = cellAlignmentChoices.readingOrder[node ? integer(node, "readingOrder", 0, 2, 0) : 0]!;
  return css;
}

export const cellFillPatterns: ReadonlySet<string> = new Set(["none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"]);
function gradientDeclarations(node: XmlNode, colors: SpreadsheetColors): Css {
  attributes(node, ["type", "degree", "left", "right", "top", "bottom"]); check(node, ["stop"], S);
  const kind = value(node, "type") ?? "linear";
  if (kind !== "linear" && kind !== "path") fail(node, "A cell gradient has an invalid type.");
  const css: Css = { "--office-cell-fill-kind": "gradient", "--office-cell-gradient-kind": kind, "--office-cell-gradient-degree": String(number(node, "degree", 0)) };
  for (const side of ["left", "right", "top", "bottom"]) {
    const amount = number(node, side, 0);
    if (amount < 0 || amount > 1) fail(node, "A cell gradient convergence is outside its fractional range.");
    css[`--office-cell-gradient-${side}`] = String(amount);
  }
  const stops = elements(node, "stop", S); css["--office-cell-gradient-stops"] = String(stops.length);
  for (const [index, stop] of stops.entries()) {
    attributes(stop, ["position"]); check(stop, ["color"], S);
    const position = number(stop, "position"), color = one(stop, "color"), prefix = `--office-cell-gradient-stop-${index}`;
    if (position < 0 || position > 1 || !color) fail(stop, "A gradient stop requires a fractional position and one color.");
    css[prefix + "-position"] = String(position); Object.assign(css, colorDeclarations(prefix, colors.read(color)));
  }
  return css;
}
export function fillDeclarations(node: XmlNode, colors: SpreadsheetColors): Css {
  attributes(node, []); check(node, ["patternFill", "gradientFill"], S);
  const pattern = one(node, "patternFill"), gradient = one(node, "gradientFill");
  if (pattern && gradient || !pattern && !gradient) fail(node, "A fill requires one definition.");
  if (gradient) return gradientDeclarations(gradient, colors);
  attributes(pattern!, ["patternType"]); check(pattern, ["fgColor", "bgColor"], S);
  const kind = value(pattern, "patternType") ?? "none", foreground = one(pattern, "fgColor"), background = one(pattern, "bgColor"), css: Css = { "--office-cell-fill-kind": kind };
  if (!cellFillPatterns.has(kind)) fail(pattern!, "A native cell fill pattern is invalid.");
  if (foreground) Object.assign(css, colorDeclarations("--office-cell-fill", colors.read(foreground)));
  if (background) Object.assign(css, colorDeclarations("--office-cell-fill-background", colors.read(background)));
  if (kind === "solid" && !foreground) fail(pattern!, "A solid cell fill requires a foreground color.");
  return css;
}
export const cellBorderKinds: ReadonlySet<string> = new Set(["none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"]);
export function borderDeclarations(node: XmlNode, colors: SpreadsheetColors, full: boolean): Css {
  attributes(node, ["diagonalUp", "diagonalDown", "outline"]); check(node, ["left", "right", "top", "bottom", "start", "end", "diagonal", "vertical", "horizontal"], S);
  const css: Css = {};
  for (const [key, property, fallback] of [["diagonalUp", "--office-cell-border-diagonal-up", false], ["diagonalDown", "--office-cell-border-diagonal-down", false], ["outline", "--office-cell-border-outline", true]] as const) {
    if (full || value(node, key) !== undefined) css[property] = String(boolean(node, key, fallback));
  }
  for (const side of ["left", "right", "top", "bottom", "start", "end", "diagonal", "vertical", "horizontal"]) {
    const border = one(node, side);
    if (!border) { if (full && ["left", "right", "top", "bottom", "diagonal"].includes(side)) css[`--office-cell-border-${side}-kind`] = "none"; continue; }
    attributes(border, ["style"]); check(border, ["color"], S);
    const kind = value(border, "style") ?? "none", color = one(border, "color");
    if (!cellBorderKinds.has(kind)) fail(border, "A native border style is invalid.");
    if (full && ["vertical", "horizontal"].includes(side)) {
      if (kind !== "none") unsupported(border, "This internal cell border is not implemented.");
      if (color) colors.read(color); continue;
    }
    css[`--office-cell-border-${side}-kind`] = kind;
    Object.assign(css, colorDeclarations(`--office-cell-border-${side}`, color ? colors.read(color) : { css: "var(--office-system-foreground)", tint: 0 }));
  }
  return css;
}
