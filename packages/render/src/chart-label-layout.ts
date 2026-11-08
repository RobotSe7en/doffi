import { DocumentError } from "@doffi/core";
import type { BasicChartScene } from "./chart.js";
import { chartTextCarrier } from "./chart-text-decoration.js";
import { svgElement } from "./presentation/svg.js";

export function chartLegendPlan(sizes: readonly { readonly width: number; readonly height: number }[], limit: number, horizontal: boolean): { readonly width: number; readonly height: number; readonly entries: readonly { readonly x: number; readonly y: number }[] } {
  if (!Number.isFinite(limit) || limit <= 0 || sizes.some(size => ![size.width, size.height].every(value => Number.isFinite(value) && value >= 0))) throw new DocumentError("ResourceLimitExceeded", "Chart legend dimensions exceed finite layout geometry.");
  let x = 0, y = 0, band = 0, width = 0, height = 0;
  const entries: { x: number; y: number }[] = [];
  for (const size of sizes) {
    if (horizontal ? x > 0 && x + size.width > limit : y > 0 && y + size.height > limit) {
      if (horizontal) { x = 0; y += band + 2; } else { y = 0; x += band + 8; }
      band = 0;
    }
    entries.push(Object.freeze({ x, y })); width = Math.max(width, x + size.width); height = Math.max(height, y + size.height);
    if (horizontal) { x += size.width + 8; band = Math.max(band, size.height); } else { y += size.height + 4; band = Math.max(band, size.width); }
  }
  if (![width, height].every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "Chart legend placement exceeds finite coordinates.");
  return Object.freeze({ width, height, entries: Object.freeze(entries) });
}

/** Measure final local fonts, then pack derived labels/legend entries. Source
 * text and source point identities remain untouched. All geometry is in the
 * owning chart's point-based viewBox, independent of its display transform. */
export function layoutBasicChartLabels(scene: BasicChartScene, account: (amount: number) => void): number {
  const { svg, width, height, plot } = scene;
  if (!svg.isConnected) throw new DocumentError("InvalidOperation", "Chart labels require their connected isolated measurement surface.");
  const translations = new Map<SVGGraphicsElement, { x: number; y: number }>();
  const overflow = new Set<SVGGraphicsElement>();
  const clipLabel = (node: SVGGraphicsElement): void => {
    overflow.add(node); node.dataset.sdkChartLabelOverflow = "true";
  };
  const box = (node: SVGGraphicsElement) => {
    const b = chartTextCarrier(node).getBBox(), shift = translations.get(node) ?? { x: 0, y: 0 };
    if (![b.x, b.y, b.width, b.height].every(Number.isFinite) || b.width < 0 || b.height < 0) throw new DocumentError("ResourceLimitExceeded", "Chart label bounds exceed finite layout geometry.");
    return { x: b.x + shift.x, y: b.y + shift.y, width: b.width, height: b.height };
  };
  const move = (node: SVGGraphicsElement, x: number, y: number): void => {
    const previous = translations.get(node) ?? { x: 0, y: 0 }, next = { x: previous.x + x, y: previous.y + y };
    if (![next.x, next.y].every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "Chart label placement exceeds finite coordinates.");
    translations.set(node, next); chartTextCarrier(node).setAttribute("transform", `translate(${next.x} ${next.y})`);
  };
  for (const label of scene.labels) {
    const bounds = box(label.node);
    move(label.node, bounds.width > width ? -bounds.x : Math.max(0, -bounds.x) + Math.min(0, width - bounds.x - bounds.width), bounds.height > height ? -bounds.y : Math.max(0, -bounds.y) + Math.min(0, height - bounds.y - bounds.height));
    if (bounds.width > width || bounds.height > height) clipLabel(label.node);
  }
  const center = plot.x + plot.width / 2;
  for (const side of [-1, 1]) {
    const labels = scene.labels.filter(label => label.outsidePie && (label.anchorX < center ? -1 : 1) === side).map(label => ({ label, bounds: box(label.node) })).sort((a, b) => a.bounds.y - b.bounds.y);
    const total = labels.reduce((sum, entry) => sum + entry.bounds.height, Math.max(0, labels.length - 1) * 2);
    if (total > plot.height) throw new DocumentError("ResourceLimitExceeded", "Pie labels exceed their nonoverlapping layout budget.");
    let bottom = plot.y - 2;
    for (const entry of labels) { const top = Math.max(entry.bounds.y, bottom + 2); move(entry.label.node, 0, top - entry.bounds.y); bottom = top + entry.bounds.height; }
    let top = plot.y + plot.height + 2;
    for (const entry of [...labels].reverse()) { const bounds = box(entry.label.node), y = Math.min(bounds.y, top - 2 - bounds.height); move(entry.label.node, 0, y - bounds.y); top = y; }
  }
  if (scene.automaticCategoryLabels) {
    const labels = [...svg.querySelectorAll<SVGTextElement>('[data-sdk-chart-label="category-axis"]')].map(node => ({ node, bounds: box(node) })).sort((a, b) => a.bounds.x - b.bounds.x);
    let right = -Infinity;
    for (const entry of labels) {
      if (entry.bounds.x < right + 2) { entry.node.setAttribute("visibility", "hidden"); chartTextCarrier(entry.node).setAttribute("visibility", "hidden"); entry.node.dataset.sdkChartAutomaticSkip = "true"; }
      else right = entry.bounds.x + entry.bounds.width;
    }
  }
  for (const node of svg.querySelectorAll<SVGTextElement>('[data-sdk-chart-label="value-axis"],[data-sdk-chart-label="category-axis"]')) {
    const bounds = box(node);
    if (bounds.x < 0 || bounds.x + bounds.width > width || bounds.y < 0 || bounds.y + bounds.height > height) clipLabel(node);
  }
  if (scene.legend) {
    const rect = scene.legend, entries = [...svg.querySelectorAll<SVGGElement>('[data-sdk-chart-legend-entry]')].map(node => ({ node, bounds: box(node) }));
    const horizontal = scene.legendPosition === "top" || scene.legendPosition === "bottom", plan = chartLegendPlan(entries.map(entry => entry.bounds), horizontal ? rect.width : rect.height, horizontal);
    if (plan.width > rect.width || plan.height > rect.height) throw new DocumentError("ResourceLimitExceeded", "Chart legend entries exceed their reserved layout band.");
    entries.forEach((entry, index) => move(entry.node, rect.x + (rect.width - plan.width) / 2 + plan.entries[index]!.x - entry.bounds.x, rect.y + (rect.height - plan.height) / 2 + plan.entries[index]!.y - entry.bounds.y));
  }
  for (const label of scene.labels) {
    if (!label.leader || label.leader.stroke === "none") continue;
    const bounds = box(label.node), x = Math.max(bounds.x, Math.min(bounds.x + bounds.width, label.anchorX)), y = Math.max(bounds.y, Math.min(bounds.y + bounds.height, label.anchorY));
    if (Math.hypot(x - label.anchorX, y - label.anchorY) < .01) continue;
    account(1); const line = svgElement(svg.ownerDocument, "line", { x1: label.anchorX, y1: label.anchorY, x2: x, y2: y, "data-sdk-chart-leader-line": "", ...label.leader });
    svg.insertBefore(line, chartTextCarrier(label.node));
  }
  let clipSerial = 0;
  for (const node of overflow) {
    account(3); const defs = svgElement(svg.ownerDocument, "defs", {}), id = scene.serial + "-label-clip-" + ++clipSerial;
    const clip = svgElement(svg.ownerDocument, "clipPath", { id, clipPathUnits: "userSpaceOnUse" });
    clip.append(svgElement(svg.ownerDocument, "rect", { x: 0, y: 0, width, height })); defs.append(clip); svg.append(defs);
    const shift = translations.get(node);
    if (shift) clip.setAttribute("transform", `translate(${-shift.x} ${-shift.y})`);
    chartTextCarrier(node).setAttribute("clip-path", `url(#${id})`);
  }
  return overflow.size;
}
