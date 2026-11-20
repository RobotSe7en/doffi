import { parseDrawingColorTransforms } from "@doffi/core";
import type { Css } from "./css.js";
import { attributes, empty } from "./xml.js";
import { emu, integer, rgba, points, unsupported } from "./drawing-values.js";
const roles = new Set([
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
  "bg1",
  "bg2",
  "tx1",
  "tx2",
]);
export const drawingAmount = (css: Css, name: string, fallback: number): number => {
  const raw = css[name];
  if (raw === undefined) return fallback;
  const value = Number(raw);
  if (!Number.isFinite(value)) unsupported("A native paint constant is invalid.");
  return value;
};
export function colorXml(
  css: Css,
  prefix: string,
  raw: Css = {},
  placeholder = false,
  fallback = "#000000",
): string {
  const value = css[prefix + "-color"] ?? fallback,
    authored = raw[prefix + "-color"] ?? value,
    reference = /^var\(\s*--office-theme-([a-zA-Z\d]+)(?:\s*[,)]|\s*$)/.exec(authored)?.[1],
    isPlaceholder = placeholder && css[prefix + "-source"] === "reference";
  const parsed =
      isPlaceholder || (reference && roles.has(reference))
        ? { rgb: "000000", alpha: 1 }
        : rgba(value, css.color ?? "#000000"),
    kind = isPlaceholder || (reference && roles.has(reference)) ? "schemeClr" : "srgbClr",
    val = isPlaceholder ? "phClr" : reference && roles.has(reference) ? reference : parsed.rgb;
  let result = "<a:" + kind + attributes({ val }) + ">";
  const sequence = parseDrawingColorTransforms(css[prefix + "-color-transforms"] ?? "");
  for (const [name, amount] of sequence) result += empty("a:" + name, { val: amount });
  for (const [name, key, defaultValue] of [
    ["lumMod", "lum-mod", 1],
    ["lumOff", "lum-off", 0],
    ["shade", "shade", 1],
    ["tint", "tint", 0],
    ["alpha", "alpha", 1],
  ] as const) {
    if (sequence.length && name !== "alpha") continue;
    const value =
      drawingAmount(css, prefix + "-" + key, defaultValue) * (name === "alpha" ? parsed.alpha : 1);
    if (value < 0 || value > 1) unsupported("A color transform is outside the PPTX profile.");
    if (value !== defaultValue)
      result += empty("a:" + name, { val: integer(value, 100000, 0, 100000) });
  }
  return result + "</a:" + kind + ">";
}
export function fillXml(
  css: Css,
  prefix = "--office-fill",
  raw: Css = {},
  placeholder = false,
  fallback = "none",
): string {
  const kind = css[prefix + "-kind"],
    paint = css[prefix + "-color"];
  if (kind === "none" || (!kind && !paint && fallback === "none")) return "<a:noFill/>";
  if (kind !== "linear")
    return (
      "<a:solidFill>" +
      colorXml(css, prefix, raw, placeholder, fallback === "none" ? "#000000" : fallback) +
      "</a:solidFill>"
    );
  const count = drawingAmount(css, prefix + "-stop-count", 0);
  if (!Number.isSafeInteger(count) || count < 2 || count > 1024)
    unsupported("A gradient stop list is outside its native range.");
  let result =
    "<a:gradFill" +
    attributes({ rotWithShape: css[prefix + "-rotate-with-shape"] === "false" ? 0 : 1 }) +
    "><a:gsLst>";
  for (let index = 0; index < count; index++) {
    const stop = prefix + "-stop-" + index;
    result +=
      "<a:gs" +
      attributes({ pos: integer(drawingAmount(css, stop + "-position", 0), 100000, 0, 100000) }) +
      ">" +
      colorXml(css, stop, raw, placeholder) +
      "</a:gs>";
  }
  const angle = drawingAmount(css, prefix + "-angle", 0);
  return (
    result +
    "</a:gsLst>" +
    empty("a:lin", {
      ang: integer(((angle % 360) + 360) % 360, 60000, 0, 21600000),
      scaled: css[prefix + "-scaled"] === "true" ? 1 : 0,
    }) +
    "</a:gradFill>"
  );
}
export function lineXml(
  css: Css,
  prefix = "--office-line",
  raw: Css = {},
  placeholder = false,
  fallback = "none",
  tag = "a:ln",
): string {
  const width = points(css[prefix + "-width"], css, 0.75),
    kind = css[prefix + "-kind"],
    color = css[prefix + "-color"];
  let result =
    "<" +
    tag +
    attributes({
      w: emu(width, 0, 2147483647),
      cap: css[prefix + "-cap"] ?? "flat",
      cmpd: css[prefix + "-cmpd"] ?? "sng",
      algn: css[prefix + "-algn"] ?? "ctr",
    }) +
    ">";
  if (kind === "linear") result += fillXml(css, prefix, raw, placeholder);
  else if (kind === "none" || (!kind && !color && fallback === "none")) result += "<a:noFill/>";
  else
    result +=
      "<a:solidFill>" +
      colorXml(css, prefix, raw, placeholder, fallback === "none" ? "#000000" : fallback) +
      "</a:solidFill>";
  if (css[prefix + "-dash"]) result += empty("a:prstDash", { val: css[prefix + "-dash"] });
  const join = css[prefix + "-join"];
  if (join === "miter")
    result += empty("a:miter", {
      lim: integer(drawingAmount(css, prefix + "-miter-limit", 8), 100000, 0, 2147483647),
    });
  else if (join === "round" || join === "bevel") result += empty("a:" + join);
  for (const [side, name] of [
    ["start", "headEnd"],
    ["end", "tailEnd"],
  ] as const)
    if (css[prefix + "-" + side + "-arrow"])
      result += empty("a:" + name, {
        type: css[prefix + "-" + side + "-arrow"],
        w: css[prefix + "-" + side + "-w"] ?? "med",
        len: css[prefix + "-" + side + "-len"] ?? "med",
      });
  return result + "</" + tag + ">";
}
export function effectXml(css: Css, raw: Css = {}, placeholder = false): string {
  if (css["--office-shadow-kind"] !== "outer") return "<a:effectLst/>";
  const distance = points(css["--office-shadow-distance"], css, 0),
    direction = drawingAmount(css, "--office-shadow-direction", 0);
  return (
    "<a:effectLst><a:outerShdw" +
    attributes({
      blurRad: emu(points(css["--office-shadow-blur"], css, 0), 0),
      dist: emu(distance, 0),
      dir: integer(((direction % 360) + 360) % 360, 60000, 0, 21600000),
      algn: css["--office-shadow-align"] ?? "b",
      rotWithShape: css["--office-shadow-rotate-with-shape"] === "false" ? 0 : 1,
    }) +
    ">" +
    colorXml(css, "--office-shadow", raw, placeholder) +
    "</a:outerShdw></a:effectLst>"
  );
}
