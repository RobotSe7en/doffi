import { DocumentError } from "@doffi/core";
import type { BasicChartScene } from "./chart.js";
import { markViewNode } from "./view-nodes.js";
import { svgElement } from "./presentation/svg.js";
import { MAX_CHART_PAINT_EXTENT } from "./chart-clipping.js";

export interface ChartTextDecoration { readonly node: SVGTextElement; readonly line: string; readonly style: string; readonly color: string; readonly size: number }
const carriers = new WeakMap<SVGGraphicsElement, SVGGraphicsElement>();
export function chartTextCarrier(node: SVGGraphicsElement): SVGGraphicsElement { return carriers.get(node) ?? node; }
/** SVG decoration longhands differ between engines. Measured font baselines
 * and bounded first-party strokes keep one searchable text node per label. */
export function prepareChartTextDecorations(scene: BasicChartScene, account: (amount: number) => void, accountText?: (units: number) => void): void {
  const document = scene.svg.ownerDocument;
  let generatedUnits = 0;
  const create: typeof svgElement = (owner, tag, attributes) => { account(1); return svgElement(owner, tag, attributes); };
  const finite = (...values: number[]) => { if (!values.every(value => Number.isFinite(value) && Math.abs(value) <= MAX_CHART_PAINT_EXTENT)) throw new DocumentError("ResourceLimitExceeded", "Chart font decoration exceeds finite glyph geometry."); };
  for (const decoration of scene.textDecorations ?? []) {
    const node = decoration.node, kinds = decoration.line.split(/\s+/).filter(value => value !== "none");
    if (!kinds.length || !node.getNumberOfChars() || !decoration.size) continue;
    if (!kinds.every(kind => ["underline", "line-through", "overline"].includes(kind)) || !["solid", "double", "dotted", "dashed", "wavy"].includes(decoration.style)) throw new DocumentError("UnsupportedFeature", "This chart font decoration is outside the selected style profile.");
    if (!node.parentElement || !node.isConnected) throw new DocumentError("InvalidOperation", "Chart font decoration requires its connected measurement surface.");
    const lines = [...node.querySelectorAll<SVGTSpanElement>("[data-sdk-chart-text-line]")], originals: SVGTextContentElement[] = lines.length ? lines : [node];
    const walker = document.createTreeWalker(node, 0xffff_ffff); let temporary = 1, units = 0;
    do { if (walker.currentNode !== node) temporary++; if (walker.currentNode.nodeType === 3) units += walker.currentNode.textContent!.length; } while (walker.nextNode());
    account(temporary); accountText?.(units);
    const probe = node.cloneNode(true) as SVGTextElement;
    probe.style.setProperty("dominant-baseline", "alphabetic", "important"); probe.style.setProperty("visibility", "hidden", "important"); node.parentElement.append(probe);
    const measured: { x: number; width: number; baseline: number; top: number; size: number }[] = [];
    try {
      const choices = [...probe.querySelectorAll<SVGTSpanElement>("[data-sdk-chart-text-line]")], comparisons: SVGTextContentElement[] = choices.length ? choices : [probe];
      for (const [index, line] of originals.entries()) {
        if (!line.getNumberOfChars()) continue;
        const bounds = line.getBBox(), comparison = comparisons[index]!.getBBox(), start = line.getStartPositionOfChar(0), size = parseFloat(document.defaultView!.getComputedStyle(line).fontSize);
        // Comparing identical text with an alphabetic baseline cancels font
        // ascent/descent differences, mixed script faces and dominant-baseline.
        const baseline = start.y + bounds.y - comparison.y;
        finite(bounds.x, bounds.y, bounds.width, bounds.height, baseline, size);
        if (size < 0 || bounds.width < 0) throw new DocumentError("InvalidDocument", "Chart font decoration has invalid glyph dimensions.");
        if (size && bounds.width) measured.push({ x: bounds.x, width: bounds.width, baseline, top: bounds.y, size });
      }
    } finally { probe.remove(); account(-temporary); }
    if (!measured.length) continue;
    const group = markViewNode(create(document, "g", { "data-sdk-chart-text-decoration": "" })), paint = create(document, "g", { "data-sdk-chart-decoration-paint": "", "aria-hidden": "true" });
    node.parentElement.insertBefore(group, node); group.append(node, paint); carriers.set(node, group);
    for (const line of measured) for (const kind of kinds) {
      const double = decoration.style === "double", thickness = line.size * (double ? .035 : .055);
      const y = kind === "underline" ? line.baseline + line.size * (double ? .035 : .08) : kind === "line-through" ? line.baseline - line.size * .28 : line.top + thickness / 2;
      for (let index = 0; index < (double ? 2 : 1); index++) {
        const at = y + index * line.size * .12; finite(line.x + line.width, at, thickness);
        if (decoration.style === "wavy") {
          const period = thickness * 8, steps = Math.ceil(line.width / period);
          if (!Number.isSafeInteger(steps) || steps > 1_000_000) throw new DocumentError("ResourceLimitExceeded", "Wavy chart decorations exceed their curve budget.");
          let path = `M${line.x} ${at}`;
          for (let step = 0; step < steps; step++) {
            const start = line.x + step * period, middle = Math.min(line.x + line.width, start + period / 2), end = Math.min(line.x + line.width, start + period);
            const segment = `Q${(start + middle) / 2} ${at - thickness} ${middle} ${at}Q${(middle + end) / 2} ${at + thickness} ${end} ${at}`;
            generatedUnits += segment.length; if (generatedUnits > 8 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "Chart decoration paths exceed their code-unit budget."); accountText?.(segment.length); path += segment;
          }
          paint.append(create(document, "path", { d: path, fill: "none", stroke: decoration.color, "stroke-width": thickness, "stroke-linecap": "butt" }));
        } else paint.append(create(document, "line", { x1: line.x, x2: line.x + line.width, y1: at, y2: at, stroke: decoration.color, "stroke-width": thickness,
          "stroke-linecap": decoration.style === "dotted" ? "round" : "butt", ...(decoration.style === "dotted" ? { "stroke-dasharray": `0 ${thickness * 3}` } : decoration.style === "dashed" ? { "stroke-dasharray": `${thickness * 4} ${thickness * 2}` } : {}) }));
      }
    }
  }
}
