import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookChart, WorkbookDisplayFormat, WorkbookNumberFormatter } from "@doffi/core";
import type { BasicChartLayout, BasicChartRect, BasicChartText } from "../chart.js";
import type { FontResolution } from "../contracts.js";
import { cssStringValue, property } from "../presentation/appearance.js";
import { chartAxisPlacement, chartValueScale, chartAxisTicks } from "../chart-axis.js";
import { chartLegendPlan } from "../chart-label-layout.js";
import { createWorkbookChartView } from "./chart-view.js";
import type { WorkbookChartBinding, WorkbookChartDataDiagnostic } from "./chart-values.js";
import type { WorkbookChartViewDiagnostic } from "./chart-view.js";
import { layoutWorkbookChartTitleContents } from "./drawing-text-layout.js";
import type { FontProbeCache } from "../font-probe-cache.js";
import type { WorkbookChartTitleContentInput, WorkbookDrawingTextLayout } from "./drawing-text-layout.js";
import type { WorkbookChartTextRole } from "./drawing-text.js";
import { measureNumberDisplay } from "./number-layout.js";

export interface WorkbookChartLayoutLimits { readonly maxNodes?: number; readonly maxPoints?: number; readonly maxCodeUnits?: number }
export interface WorkbookChartLayout {
  readonly element: HTMLElement; readonly nodes: number; readonly layout: BasicChartLayout;
  readonly texts: readonly WorkbookDrawingTextLayout[]; readonly fonts: readonly FontResolution[];
  readonly usedFallbackLineMetrics: boolean;
  readonly diagnostics: readonly (WorkbookChartDataDiagnostic | WorkbookChartViewDiagnostic)[];
  readonly overflowingTitles: readonly WorkbookChartTextRole[];
}

/** Compose current graphics and both kinds of titles before publishing a
 * candidate. Titles retain their measured local boxes; axis-title strips and
 * legend bands occupy separate space. Object transforms, frame effects and
 * worksheet pane clipping remain the drawing compositor's responsibility. */
export function layoutWorkbookChart(scope: HTMLElement, source: HTMLElement, chart: DeepReadonly<WorkbookChart>,
  binding: Extract<WorkbookChartBinding, { kind: "available" }>, widthPt: number, heightPt: number, id: string,
  formatter: WorkbookNumberFormatter, limits: WorkbookChartLayoutLimits = {}, probes?: FontProbeCache): WorkbookChartLayout {
  const document = scope.ownerDocument, view = document.defaultView, maxNodes = limits.maxNodes ?? 50_000, maxUnits = limits.maxCodeUnits ?? 8 * 1024 * 1024;
  if (!scope.isConnected || !source.isConnected || source.ownerDocument !== document || !view || binding.chartId !== chart.chartId
    || source.dataset.nodeId !== chart.nodeId || source.dataset.officeKind !== "chart"
    || ![maxNodes, maxUnits, limits.maxPoints ?? 50_000].every(value => Number.isSafeInteger(value) && value > 0)
    || ![widthPt, heightPt].every(value => Number.isFinite(value) && value > 0 && value <= 1_000_000))
    throw new DocumentError("InvalidOperation", "Chart composition requires its own current binding, connected surface and finite budgets.");
  const pie = chart.kind === "pie";
  let measureContext: CanvasRenderingContext2D | undefined;
  const placement = () => {
    if (maxNodes < 4) throw new DocumentError("ResourceLimitExceeded", "Chart composition exceeds its temporary node budget.");
    const scope = document.createElement("span"), category = document.createElement("span"), value = document.createElement("span"), legend = document.createElement("span");
    scope.className = chart.textCssClass ?? ""; scope.style.cssText = "position:absolute;display:none";
    category.className = chart.categoryAxisCssClass ?? ""; value.className = chart.valueAxisCssClass ?? ""; legend.className = chart.legendCssClass ?? ""; scope.append(category, value, legend); source.append(scope);
    try {
      const scale = binding.data.series.some(series => !series.hidden) ? chartValueScale(chart, binding.data.series) : { minimum: 0, maximum: 1, unit: 1 };
      const categoryStyle = view.getComputedStyle(category), valueStyle = view.getComputedStyle(value), legendStyle = view.getComputedStyle(legend), context = document.createElement("canvas").getContext("2d")!;
      const textSize = (text: BasicChartText, style: CSSStyleDeclaration): { width: number; height: number } => {
        context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
        const factor = Number(property(style, "--office-text-font-scale", "1"));
        const spacing = parseFloat(style.letterSpacing) || 0;
        const measure = (line: string) => (context.measureText(line).width + Math.max(0, [...line].length - 1) * spacing) * .75 * factor;
        const lines = typeof text === "string" ? text.split("\n") : [""];
        return { width: typeof text === "string" ? Math.max(0, ...lines.map(measure)) : measureNumberDisplay(text.parts, text.availableWidth, measure).width,
          height: lines.length * (parseFloat(style.lineHeight) || parseFloat(style.fontSize) * 1.2) * .75 * factor };
      };
      const valueLabels = pie || chart.showValueAxis === false || property(valueStyle, "--office-axis-label-position") === "none" ? [] : chartAxisTicks(scale.minimum, scale.maximum, scale.unit, 1024, .001)
        .map(value => textSize(format(value, linked(valueStyle, "axis") ? binding.valueAxisFormat : explicit(valueStyle, "axis", chart.numberFormat), valueStyle, widthPt / 4), valueStyle));
      const categoryLabels = pie || chart.showCategoryAxis === false || property(categoryStyle, "--office-axis-label-position") === "none" ? [] : binding.data.categories.map((name, index) => {
        const numeric = binding.categoryNumbers[index];
        return textSize(numeric === null || numeric === undefined ? name : format(numeric, linked(categoryStyle, "axis") ? binding.categoryFormats?.[index] : explicit(categoryStyle, "axis"), categoryStyle, widthPt / Math.max(1, binding.data.categories.length)), categoryStyle);
      });
      const visible = binding.data.series.filter(series => !series.hidden), vary = (chart.varyColors ?? pie) && chart.series.length === 1;
      const items = pie || vary && visible.length === 1 ? binding.data.categories.map((name, index) => ({ name, series: visible[0], index })) : visible.map(series => ({ name: series.name, series, index: -1 }));
      const definitions = new Map(chart.seriesStyles?.map(series => [series.seriesId, series])), points = new Map(chart.pointStyles?.map(point => [JSON.stringify([point.seriesId, point.index]), point]));
      const legendEntries = chart.legend === "none" ? [] : items.map(item => {
        const text = textSize(item.name, legendStyle), series = item.series ? definitions.get(item.series.seriesId) : undefined;
        const point = item.series && item.index >= 0 ? points.get(JSON.stringify([item.series.seriesId, binding.data.sourceIndexes?.[item.index] ?? item.index])) : undefined;
        const marker = chart.kind === "line" && (point?.marker ?? series?.marker ?? "none") !== "none" ? point?.markerSize ?? series?.markerSize ?? 5 : 0;
        return { width: text.width + (chart.kind === "line" ? Math.max(22, marker) : 7) + 11, height: Math.max(7, marker, text.height) };
      });
      return { ...(pie ? { categoryTop: false, valueRight: false } : chartAxisPlacement(chart, categoryStyle, valueStyle, scale.minimum, scale.maximum, Math.max(1, binding.data.categories.length))),
        valueBand: Math.min(widthPt * .3, Math.max(28, ...valueLabels.map(label => label.width + 12))), categoryBand: Math.max(26, ...categoryLabels.map(label => label.height / 2 + 10 + 10 * Number(property(categoryStyle, "--office-axis-label-offset", "100")) / 100)), legendEntries };

    } finally { scope.remove(); }
  };
  const measured = placement(), axes = pie ? undefined : measured, categoryTop = axes?.categoryTop ?? false, valueRight = axes?.valueRight ?? false;
  const hasTitle = chart.titleVisible !== false && !!(chart.titleNodeId || chart.titleReference);
  const categoryTitle = !pie && chart.showCategoryAxis !== false && !!chart.categoryAxisTitle, valueTitle = !pie && chart.showValueAxis !== false && !!chart.valueAxisTitle;
  const titleHeight = hasTitle ? Math.min(54, heightPt * .2) : 12, axisBand = 24;
  const valueBand = measured.valueBand, categoryBand = measured.categoryBand;
  const ordinaryLeft = pie || valueRight ? 12 : valueBand + (valueTitle ? axisBand : 0), ordinaryRight = pie ? 12 : valueRight ? valueBand + (valueTitle ? axisBand : 0) : 16;
  let legendBand = 0;
  if (chart.legend === "left" || chart.legend === "right") {
    const height = heightPt - titleHeight - (pie ? 23 : 18 + (categoryTop ? categoryBand + (categoryTitle ? axisBand : 0) + 12 : categoryBand + (categoryTitle ? axisBand : 0)));
    legendBand = Math.max(Math.min(widthPt * .25, 100), chartLegendPlan(measured.legendEntries, height, false).width + 8);
  }
  else if (chart.legend !== "none") {
    legendBand = Math.max(28, chartLegendPlan(measured.legendEntries, widthPt - ordinaryLeft - ordinaryRight, true).height + 4);
  }
  const left = ordinaryLeft + (chart.legend === "left" ? legendBand : 0), right = widthPt - ordinaryRight - (chart.legend === "right" ? legendBand : 0);
  const top = titleHeight + (pie ? 15 : 18) + (chart.legend === "top" ? legendBand : 0) + (categoryTop ? categoryBand + (categoryTitle ? axisBand : 0) : 0);
  const bottom = heightPt - (pie ? 8 : categoryTop ? 12 : categoryBand + (categoryTitle ? axisBand : 0)) - (chart.legend === "bottom" ? legendBand : 0);
  if (![left, right, top, bottom].every(Number.isFinite) || right <= left || bottom <= top || widthPt <= 12) throw new DocumentError("ResourceLimitExceeded", "This chart canvas has no space for its plot, titles and legend.");
  const plot = Object.freeze({ x: left, y: top, width: right - left, height: bottom - top });
  const title = hasTitle ? Object.freeze({ x: 6, y: 6, width: widthPt - 12, height: titleHeight }) : undefined;
  const legend: BasicChartRect | undefined = chart.legend === "none" ? undefined
    : chart.legend === "left" || chart.legend === "right" ? { x: chart.legend === "left" ? 0 : widthPt - legendBand, y: top, width: legendBand, height: plot.height }
      : { x: left, y: chart.legend === "top" ? titleHeight + 12 : heightPt - legendBand, width: plot.width, height: legendBand };
  const layout: BasicChartLayout = Object.freeze({ plot, ...(title ? { title } : {}), ...(legend ? { legend: Object.freeze(legend) } : {}) });
  const boxes: { role: WorkbookChartTextRole; rect: BasicChartRect; rotation: number }[] = [];
  if (title) boxes.push({ role: "title", rect: title, rotation: 0 });
  if (categoryTitle) boxes.push({ role: "category-axis-title", rect: { x: left, y: categoryTop ? top - categoryBand - axisBand : bottom + categoryBand, width: plot.width, height: axisBand }, rotation: 0 });
  if (valueTitle) boxes.push({ role: "value-axis-title", rect: { x: valueRight ? right + valueBand : left - valueBand - axisBand, y: bottom, width: plot.height, height: axisBand }, rotation: -90 });
  const inputs: WorkbookChartTitleContentInput[] = boxes.map(box => {
    const definition = box.role === "title" ? { reference: chart.titleReference } : box.role === "category-axis-title" ? chart.categoryAxisTitle! : chart.valueAxisTitle!;
    if (definition.reference && binding.titles[box.role] === undefined) throw new DocumentError("InvalidOperation", "A chart reference title has no current scalar display.");
    return { source, chart, role: box.role, widthPt: box.rect.width, heightPt: box.rect.height, ...(definition.reference ? { value: binding.titles[box.role]! } : {}) };
  });
  const texts = layoutWorkbookChartTitleContents(scope, inputs, { maxNodes, maxCodeUnits: maxUnits }, probes);
  let axisNodes = 0, axisUnits = 0;
  for (const [index, object] of texts.objects.entries()) {
    if (object.widthPt > boxes[index]!.rect.width || object.heightPt > boxes[index]!.rect.height) throw new DocumentError("ResourceLimitExceeded", "Auto-sized chart titles exceed their reserved text bands.");
    if (boxes[index]!.role !== "title") { axisNodes += object.nodes; axisUnits += object.element.textContent?.length ?? 0; }
  }
  if (axisNodes >= maxNodes || axisUnits >= maxUnits) throw new DocumentError("ResourceLimitExceeded", "Chart titles exhaust the graphic budget.");
  const main = boxes.findIndex(box => box.role === "title");
  const seriesIndexes = new Map(binding.data.series.map((series, index) => [series.seriesId, index]));
  function format(value: number, definition?: WorkbookDisplayFormat, style?: CSSStyleDeclaration, availableWidth = widthPt / 4): BasicChartText {
    const display = formatter.format(value, definition, style ? { availableWidth, measure(text) {
      measureContext ??= document.createElement("canvas").getContext("2d")!;
      measureContext.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
      return (measureContext.measureText(text).width + Math.max(0, [...text].length - 1) * (parseFloat(style.letterSpacing) || 0)) * .75 * Number(property(style, "--office-text-font-scale", "1"));
    } } : undefined);
    if (display.kind === "overflow") return "####";
    if (display.parts.some(part => part.kind !== "text")) return { parts: display.parts, availableWidth };
    return display.parts.map(part => part.kind === "text" ? part.text : "").join("");
  };
  function linked(style: CSSStyleDeclaration, kind: "axis" | "label"): boolean {
    const flag = property(style, `--office-${kind}-number-format-source-linked`, "true");
    if (flag !== "true" && flag !== "false") throw new DocumentError("InvalidDocument", "Chart source-linked number-format flags are invalid.");
    return flag === "true";
  };
  function explicit(style: CSSStyleDeclaration, kind: "axis" | "label", fallback?: string): WorkbookDisplayFormat | undefined {
    const raw = property(style, `--office-${kind}-number-format`), code = raw ? cssStringValue(raw) : fallback;
    return code === undefined ? undefined : { code };
  };
  const graphics = createWorkbookChartView(source, chart, binding.data, widthPt, heightPt, id, {
    ...limits, maxNodes: maxNodes - axisNodes, maxCodeUnits: maxUnits - axisUnits, layout, ...(probes ? { probes } : {}),
    ...(main < 0 ? {} : { title: texts.objects[main]!.element }),
    formatNumber(value, code, context) {
      if (!context) return format(value, code === undefined ? undefined : { code });
      if (context.kind === "value-axis") return format(value, linked(context.style, "axis") ? binding.valueAxisFormat : explicit(context.style, "axis", code), context.style, valueBand - 12);
      const sourceFormat = context.content === "percent" ? { code: "0%" } : binding.valueFormats[seriesIndexes.get(context.seriesId)!]?.[context.index];
      return format(value, linked(context.style, "label") ? sourceFormat : explicit(context.style, "label", context.content === "percent" ? "0%" : "General"), context.style, Math.max(1, plot.width / Math.max(1, binding.data.categories.length)));
    },
    formatCategory(index, style) {
      const value = binding.categoryNumbers[index];
      return value === null || value === undefined ? binding.data.categories[index]! : format(value, linked(style, "axis") ? binding.categoryFormats?.[index] : explicit(style, "axis"), style, plot.width / Math.max(1, binding.data.categories.length));
    },
  });
  for (const [index, object] of texts.objects.entries()) {
    const box = boxes[index]!;
    if (box.role === "title") continue;
    Object.assign(object.element.style, { position: "absolute", left: box.rect.x + "pt", top: box.rect.y + "pt", transformOrigin: "0 0", transform: `rotate(${box.rotation}deg)` });
    graphics.element.append(object.element);
  }
  return Object.freeze({ element: graphics.element, nodes: graphics.nodes + axisNodes, layout, texts: texts.objects, fonts: Object.freeze([...texts.fonts, ...graphics.fonts]),
    usedFallbackLineMetrics: texts.usedFallbackLineMetrics, diagnostics: Object.freeze([...binding.diagnostics, ...graphics.diagnostics]),
    overflowingTitles: Object.freeze(boxes.filter((_, index) => texts.objects[index]!.overflow).map(box => box.role)) });
}
