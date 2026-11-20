import { DocumentError } from "@doffi/core";
import type { PresentationPresetGeometry } from "@doffi/core";
import { addArrows, strokeAttributes, svgFill } from "./presentation/appearance.js";
import { svgElement, svgPath } from "./presentation/svg.js";
import { markViewNode } from "./view-nodes.js";
import type { DrawingGradientOrientation } from "./drawing-gradient.js";

export interface DrawingGeometryViewInput {
  readonly preset: string; readonly geometry: PresentationPresetGeometry;
  readonly widthPt: number; readonly heightPt: number; readonly id: string;
  readonly fillFallback?: string; readonly strokeFallback?: string; readonly maxNodes?: number;
  readonly paint?: "fill" | "stroke" | "both";
  readonly orientation?: DrawingGradientOrientation;
}
/** Shared SVG projection of first-party preset paths. Layout, source ownership,
 * text rectangles and object transforms remain with each format's compositor. */
export function createDrawingGeometryView(document: Document, style: CSSStyleDeclaration, input: DrawingGeometryViewInput): SVGSVGElement {
  const { widthPt, heightPt, id } = input, maxNodes = input.maxNodes ?? 50_000;
  if (![widthPt, heightPt].every(value => Number.isFinite(value) && value >= 0) || !Number.isSafeInteger(maxNodes) || maxNodes < 0 || !/^[a-z][a-z0-9-]*$/i.test(id)) throw new DocumentError("InvalidOperation", "Drawing geometry requires finite extents, a node budget and an isolated SVG identity.");
  let count = 0;
  const account = (amount: number): void => { count += amount; if (count > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing geometry exceeds its node budget."); };
  account(1);
  const svg = markViewNode(svgElement(document, "svg", { width: "100%", height: "100%", viewBox: `0 0 ${Math.max(widthPt, .001)} ${Math.max(heightPt, .001)}`, "aria-hidden": "true", "data-sdk-preset": input.preset }));
  svg.style.cssText = `position:absolute;left:0;top:0;width:${Math.max(widthPt, .001)}pt;height:${Math.max(heightPt, .001)}pt;overflow:visible;pointer-events:none`;
  const fill = input.paint === "stroke" ? "none" : svgFill(document, style, svg, id + "-fill", { x: 0, y: 0, width: widthPt, height: heightPt }, "--office-fill", input.fillFallback ?? "none", account, input.orientation);
  const stroke = strokeAttributes(document, style, "--office-line", input.strokeFallback ?? "none");
  if (input.paint !== "fill" && input.geometry.paths.some(path => path.stroke) && style.getPropertyValue("--office-line-kind").trim() === "linear")
    stroke.stroke = svgFill(document, style, svg, id + "-line", { x: 0, y: 0, width: widthPt, height: heightPt }, "--office-line", "none", account, input.orientation);
  for (const [index, path] of input.geometry.paths.entries()) {
    account(1);
    const drawing = svgElement(document, "path", { d: svgPath(path), fill: path.fill ? fill : "none", ...stroke });
    if (!path.stroke || input.paint === "fill") drawing.setAttribute("stroke", "none");
    // Multi-path shapes must not publish duplicate marker IDs in one SVG.
    if (path.stroke && input.paint !== "fill") addArrows(document, svg, drawing, style, `${id}-path-${index}`, account); svg.append(drawing);
  }
  return svg;
}
