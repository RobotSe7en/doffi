import { DocumentError } from "@doffi/core";
import { svgElement } from "./svg.js";
import { drawingGradientLine } from "../drawing-gradient.js";
import type { DrawingGradientBounds, DrawingGradientOrientation } from "../drawing-gradient.js";

export interface DrawingPaintStyle { getPropertyValue(name: string): string; readonly color: string }
export const property = (style: Pick<DrawingPaintStyle, "getPropertyValue">, name: string, fallback = ""): string => style.getPropertyValue(name).trim() || fallback;
export function cssStringValue(value: string): string {
  if (!/^["']/.test(value)) return value;
  return value.slice(1, -1).replace(/\\([0-9a-f]{1,6})(?:\r\n|[ \t\n\r\f])?|\\([\s\S])/gi, (_match, hex: string | undefined, escaped: string | undefined) => hex ? String.fromCodePoint(Math.max(1, Math.min(0x10ffff, parseInt(hex, 16)))) : escaped ?? "");
}
export const numeric = (style: Pick<DrawingPaintStyle, "getPropertyValue">, name: string, fallback: number): number => {
  const raw = property(style, name); const value = Number.parseFloat(raw); return Number.isFinite(value) ? value : fallback;
};
const colorContexts = new WeakMap<Document, CanvasRenderingContext2D>();
export function resolvedColor(document: Document, style: DrawingPaintStyle, prefix: string, fallback = "#000000", depth = 0): string {
  if (depth > 4) throw new DocumentError("InvalidDocument", "A presentation color reference is cyclic.");
  let context = colorContexts.get(document); if (!context) { context = document.createElement("canvas").getContext("2d")!; colorContexts.set(document, context); }
  let value = property(style, prefix + "-color", fallback);
  if (value.toLowerCase() === "currentcolor") value = style.color;
  if (property(style, prefix + "-source") === "reference" && property(style,prefix+"-color")) {
    const reference = property(style, prefix + "-reference-prefix");
    if (!["--office-fill-reference", "--office-line-reference", "--office-background-reference", "--office-effect-reference"].includes(reference)) throw new DocumentError("InvalidDocument", "A presentation color reference is unknown.");
    value = resolvedColor(document, style, reference, fallback, depth + 1);
  }
  context.fillStyle = fallback; context.fillStyle = value;
  const raw = context.fillStyle;
  const nativeAlpha = raw.startsWith("#") ? raw.length === 9 ? parseInt(raw.slice(7, 9), 16) / 255 : 1 : Number((raw.match(/[\d.]+/g) ?? [])[3] ?? 1);
  const rgb = raw.startsWith("#") ? [1, 3, 5].map(n => parseInt(raw.slice(n, n + 2), 16) / 255) : (raw.match(/[\d.]+/g) ?? []).slice(0, 3).map(v => Number(v) / 255);
  if (rgb.length !== 3 || rgb.some(n => !Number.isFinite(n))) return raw;
  const [r = 0, g = 0, b = 0] = rgb, max = Math.max(r, g, b), min = Math.min(r, g, b), delta = max - min;
  const light = (max + min) / 2;
  let hue = 0;
  if (delta) { hue = max === r ? ((g - b) / delta) % 6 : max === g ? (b - r) / delta + 2 : (r - g) / delta + 4; hue = (hue / 6 + 1) % 1; }
  const saturation = delta ? delta / (1 - Math.abs(2 * light - 1)) : 0;
  const lum = Math.max(0, Math.min(1, light * numeric(style, prefix + "-lum-mod", 1) + numeric(style, prefix + "-lum-off", 0)));
  const chroma = (1 - Math.abs(2 * lum - 1)) * saturation, x = chroma * (1 - Math.abs((hue * 6) % 2 - 1)), m = lum - chroma / 2;
  const channels = hue < 1 / 6 ? [chroma, x, 0] : hue < 2 / 6 ? [x, chroma, 0] : hue < 3 / 6 ? [0, chroma, x] : hue < 4 / 6 ? [0, x, chroma] : hue < 5 / 6 ? [x, 0, chroma] : [chroma, 0, x];
  const tint = numeric(style, prefix + "-tint", 0), shade = numeric(style, prefix + "-shade", 1);
  const values = channels.map(c => Math.round(Math.max(0, Math.min(1, ((c + m) * shade) * (1 - tint) + tint)) * 255));
  const alpha = Math.max(0, Math.min(1, nativeAlpha * numeric(style, prefix + "-alpha", 1)));
  return `rgba(${values.join(",")},${alpha})`;
}
export function svgFill(document: Document, style: DrawingPaintStyle, svg: SVGSVGElement, id: string, bounds: DrawingGradientBounds, prefix = "--office-fill", fallback = "none", account?: (amount: number) => void, orientation?: DrawingGradientOrientation): string {
  const kind = property(style, prefix + "-kind");
  if (kind === "none") return "none";
  if (kind !== "linear") return kind === "solid" || property(style, prefix + "-color") ? resolvedColor(document, style, prefix) : fallback === "none" ? "none" : resolvedColor(document,style,prefix,fallback);
  const count = numeric(style, prefix + "-stop-count", 0);
  if (!Number.isSafeInteger(count) || count < 2 || count > 1024) throw new DocumentError("ResourceLimitExceeded", "A presentation gradient exceeds its stop budget.");
  account?.(count + 2);
  const scaled = property(style, prefix + "-scaled", "false");
  if (scaled !== "true" && scaled !== "false") throw new DocumentError("InvalidDocument", "A drawing gradient scale flag must be a boolean.");
  const angle = property(style, prefix + "-angle", "0");
  if (!/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(angle)) throw new DocumentError("InvalidDocument", "A drawing gradient angle must be a source number.");
  const rotates = property(style, prefix + "-rotate-with-shape", "true");
  if (rotates !== "true" && rotates !== "false") throw new DocumentError("InvalidDocument", "A drawing gradient rotation flag must be a boolean.");
  const line = drawingGradientLine(bounds, Number(angle), scaled === "true", rotates === "false" ? orientation : undefined);
  const gradient = svgElement(document, "linearGradient", { id, gradientUnits: "userSpaceOnUse", "color-interpolation": "sRGB", ...line });
  let previous = -1;
  for (let n = 0; n < count; n++) {
    const raw = property(style, prefix + "-stop-" + n + "-position"), position = Number(raw);
    if (!raw || !Number.isFinite(position) || position < 0 || position > 1 || position < previous) throw new DocumentError("InvalidDocument", "Drawing gradient stops require ordered positions in [0,1].");
    previous = position; gradient.append(svgElement(document, "stop", { offset: position, "stop-color": resolvedColor(document, style, prefix + "-stop-" + n) }));
  }
  const defs = svgElement(document, "defs"); defs.append(gradient); svg.append(defs); return `url(#${id})`;
}
export function strokeAttributes(document: Document, style: DrawingPaintStyle, prefix = "--office-line", fallback = "none"): Record<string, string | number> {
  const kind = property(style, prefix + "-kind"), color = kind === "none" ? "none" : property(style, prefix + "-color") ? resolvedColor(document, style, prefix) : fallback === "none" ? "none" : resolvedColor(document,style,prefix,fallback);
  const width = numeric(style, prefix + "-width", .75), dash = property(style, prefix + "-dash", "solid");
  const patterns: Record<string, readonly number[]> = { dash: [4, 3], dot: [1, 3], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1] };
  const cap = property(style, prefix + "-cap", "flat");
  return { stroke: color, "stroke-width": width, "stroke-linecap": cap === "rnd" ? "round" : cap === "sq" ? "square" : "butt", "stroke-linejoin": property(style, prefix + "-join", "round"), "stroke-miterlimit": numeric(style, prefix + "-miter-limit", 8), ...(patterns[dash] ? { "stroke-dasharray": patterns[dash]!.map(n => n * width).join(" ") } : {}) };
}
export function addArrows(document: Document, svg: SVGSVGElement, path: SVGPathElement, style: CSSStyleDeclaration, id: string, account?: (amount: number) => void): void {
  for (const side of ["start", "end"] as const) {
    const prefix = "--office-line-" + side, type = property(style, prefix + "-arrow", "none"); if (type === "none") continue;
    account?.(3); // defs, marker and its one geometry primitive.
    const size = (value: string): number => value === "sm" ? 2 : value === "lg" ? 5 : 3;
    const width = size(property(style, prefix + "-w")), length = size(property(style, prefix + "-len"));
    const marker = svgElement(document, "marker", { id: id + "-" + side, markerWidth: length, markerHeight: width, refX: length, refY: width / 2, orient: "auto-start-reverse", viewBox: `0 0 ${length} ${width}`, overflow: "visible" });
    const color = resolvedColor(document, style, "--office-line");
    if (type === "oval") marker.append(svgElement(document, "ellipse", { cx: length / 2, cy: width / 2, rx: length / 2, ry: width / 2, fill: color }));
    else marker.append(svgElement(document, "path", { d: type === "diamond" ? `M0 ${width / 2}L${length / 2} 0L${length} ${width / 2}L${length / 2} ${width}Z` : `M0 0L${length} ${width / 2}L0 ${width}` + (type === "stealth" ? `L${length / 3} ${width / 2}Z` : type === "arrow" ? "" : "Z"), fill: type === "arrow" ? "none" : color, stroke: color, "stroke-width": 0.5 }));
    const defs = svgElement(document, "defs"); defs.append(marker); svg.append(defs); path.setAttribute("marker-" + side, `url(#${id}-${side})`);
  }
}
