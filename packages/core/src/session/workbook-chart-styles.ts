import { ident, parse, walk } from "css-tree";
import type { WorkbookChart } from "../contracts/workbook-charts.js";
import type { WorkbookDrawingOperation } from "../contracts/workbook-drawing-operations.js";
import type { DeepReadonly } from "../runtime.js";
import type { SourceStyleNode } from "../document/source-css.js";
import { invalid } from "../document/validation.js";
import type { SourceStyleEdit } from "./source-styles.js";
import type { WorkbookContentEdit } from "./workbook-content.js";

type Operation = Extract<WorkbookDrawingOperation, { op: "setWorkbookChartStyle" }>;
const isSeries = (target: Operation["target"]) => ["series", "point", "seriesLabel", "pointLabel"].includes(target);
const isPoint = (target: Operation["target"]) => target === "point" || target === "pointLabel";
export function validateWorkbookChartStyleTarget(chart: DeepReadonly<WorkbookChart> | undefined, operation: Operation): void {
  if (!chart || isSeries(operation.target) && !chart.series.some(series => series.seriesId === operation.seriesId)
    || !isSeries(operation.target) && operation.seriesId !== undefined || isPoint(operation.target) !== (operation.index !== undefined))
    invalid("Chart appearance requires an existing chart/series and an original point index only for point targets.");
  if ((operation.marker !== undefined || operation.markerSize !== undefined) && (operation.target !== "series" && operation.target !== "point"))
    invalid("Markers belong to line-chart series and points.");
}
/** Isolate metadata-owned style slots. Cloning aliases preserves contextual CSS
 * and unrelated consumers; a local important override wins the authored cascade. */
export function editWorkbookChartStyle(chart: WorkbookChart, operation: Operation, styles: SourceStyleEdit, content: WorkbookContentEdit, allocate: () => string): void {
  validateWorkbookChartStyleTarget(chart, operation);
  if ((operation.marker !== undefined || operation.markerSize !== undefined) && chart.kind !== "line") invalid("Markers require the current line-chart kind.");
  for (const property of Object.keys(operation.properties)) if (!/^(?:--office-(?:fill|line|text|font|marker|axis|label|invert)-[a-z0-9-]+|color|background-color|font-(?:size|family|weight|style)|text-decoration(?:-(?:line|style|color))?)$/.test(property))
    invalid("The chart appearance property is outside the source style subset.");
  const properties = styles.drawingProperties(operation.properties), point = isPoint(operation.target), seriesTarget = isSeries(operation.target), label = operation.target === "seriesLabel" || operation.target === "pointLabel";
  const key = seriesTarget ? undefined : (operation.target + "CssClass") as "plotCssClass" | "legendCssClass" | "textCssClass" | "labelCssClass" | "categoryAxisCssClass" | "valueAxisCssClass" | "gridCssClass" | "categoryGridCssClass";
  const allocateClass = (): string => {
    const names = new Set<string>(); for (const sheet of styles.snapshot()) walk(parse(sheet.css, { context: "stylesheet" }), node => { if (node.type === "ClassSelector") names.add(ident.decode(node.name)); });
    let name: string; do { name = "workbook-chart-" + allocate(); } while (names.has(name)); return name;
  };
  let entry: NonNullable<WorkbookChart["seriesStyles"]>[number] & { index?: number } | undefined = seriesTarget ? point ? chart.pointStyles?.find(style => style.seriesId === operation.seriesId && style.index === operation.index) : chart.seriesStyles?.find(style => style.seriesId === operation.seriesId) : undefined;
  const current = key ? chart[key] : label ? entry?.labelCssClass : entry?.cssClass;
  let cssClass = current;
  if (Object.keys(properties).length) {
    const scope: SourceStyleNode = { tag: "span", attrs: chart.textCssClass ? { class: chart.textCssClass } : {}, parent: content.styleNode(chart.nodeId) };
    const probe: SourceStyleNode = { tag: "span", attrs: current ? { class: current } : {}, parent: scope };
    const specified = styles.resolver({ combineDecorations: false }).specified(probe);
    if (Object.entries(properties).some(([name, value]) => specified[name] !== value)) {
      cssClass = allocateClass();
      if (current) styles.aliasClass(current, cssClass, new Set(Object.keys(properties)));
      styles.overrideLocalClass(cssClass, properties, {}, "span");
    }
  }
  const markerChanged = operation.marker !== undefined && (operation.marker ?? undefined) !== entry?.marker
    || operation.markerSize !== undefined && (operation.markerSize ?? undefined) !== entry?.markerSize;
  if (cssClass === current && !markerChanged) return;
  if (key) { if (cssClass) chart[key] = cssClass; return; }
  if (!entry) {
    // An empty class supplies the required metadata slot for marker-only or
    // label-only edits. It contributes no separate style or chart data source.
    let ownClass = cssClass;
    if (!ownClass || label) { ownClass = allocateClass(); styles.addClass(ownClass, {}); }
    entry = { seriesId: operation.seriesId!, cssClass: ownClass, ...(point ? { index: operation.index! } : {}) };
    if (point) (chart.pointStyles ??= []).push(entry as NonNullable<WorkbookChart["pointStyles"]>[number]); else (chart.seriesStyles ??= []).push(entry);
  }
  if (cssClass) { if (label) entry.labelCssClass = cssClass; else entry.cssClass = cssClass; }
  for (const name of ["marker", "markerSize"] as const) { const value = operation[name]; if (value === null) delete entry[name]; else if (value !== undefined) Reflect.set(entry, name, value); }
}
