import { DocumentError } from "@doffi/core";
import type { BasicChartScene, BasicChartRect } from "./chart.js";
import type { DrawingGradientOrientation } from "./drawing-gradient.js";
import type { DrawingPaintStyle } from "./presentation/appearance.js";
import { property, strokeAttributes, svgFill } from "./presentation/appearance.js";
import { svgElement } from "./presentation/svg.js";
import { chartTextCarrier, ensureChartTextCarrier } from "./chart-text-decoration.js";
import { MAX_CHART_PAINT_EXTENT } from "./chart-clipping.js";

export interface ChartLabelFrame { readonly node: SVGTextElement; readonly style: DrawingPaintStyle }
/** Frames use the final searchable text/decorations' measured local rectangle.
 * Their carrier participates in the same label placement, clipping and skips. */
export function prepareChartLabelFrames(scene: BasicChartScene, account: (amount: number) => void, orientation?: DrawingGradientOrientation): void {
  let serial = 0;
  const rect = (bounds: BasicChartRect, style: DrawingPaintStyle, kind: string) => {
    if (![bounds.x, bounds.y, bounds.width, bounds.height].every(value => Number.isFinite(value) && Math.abs(value) <= MAX_CHART_PAINT_EXTENT) || bounds.width < 0 || bounds.height < 0)
      throw new DocumentError("ResourceLimitExceeded", "Chart label frames exceed finite measured bounds.");
    if (!bounds.width || !bounds.height) return undefined;
    account(1); return svgElement(scene.svg.ownerDocument, "rect", { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height, "aria-hidden": "true", ["data-sdk-chart-" + kind + "-frame"]: "",
      fill: svgFill(scene.svg.ownerDocument, style, scene.svg, scene.serial + "-label-frame-" + ++serial, bounds, "--office-fill", "none", account, orientation),
      ...strokeAttributes(scene.svg.ownerDocument, style, "--office-line", property(style, "--office-line-kind") === "solid" ? "#898989" : "none") });
  };
  for (const entry of scene.labelFrames ?? []) {
    const bounds = chartTextCarrier(entry.node).getBBox(), frame = rect(bounds, entry.style, "label");
    if (frame) { const carrier = ensureChartTextCarrier(entry.node, account); carrier.insertBefore(frame, carrier.firstChild); }
  }
  if (scene.legendFrame) {
    const groups = [...scene.svg.querySelectorAll<SVGGElement>("[data-sdk-chart-legend-entry]")];
    if (!groups.length) return;
    let bounds = scene.legend;
    if (!bounds) {
      const boxes = groups.map(group => group.getBBox()), x = Math.min(...boxes.map(box => box.x)), y = Math.min(...boxes.map(box => box.y));
      bounds = { x, y, width: Math.max(...boxes.map(box => box.x + box.width)) - x, height: Math.max(...boxes.map(box => box.y + box.height)) - y };
    }
    const frame = rect(bounds, scene.legendFrame, "legend"); if (frame) scene.svg.insertBefore(frame, groups[0]!);
  }
}
