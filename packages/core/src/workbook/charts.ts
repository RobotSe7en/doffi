import type { WorkbookChartVector, WorkbookMetadata, WorkbookObject } from "../contracts/workbook.js";
import { WorkbookErrorValueSchema } from "../contracts/workbook.js";
import type { ParsedHtml } from "../document/html.js";
import { checkLimit, invalid } from "../document/validation.js";
import { parseFormula } from "./formula.js";
import type { FormulaExpression } from "./formula.js";

export function chartReferenceExpression(expression: string): FormulaExpression {
  const result = parseFormula(expression, { row: 0, column: 0 });
  const reference = (node: FormulaExpression): boolean => ["cell", "range", "axisRange", "name", "structured"].includes(node.kind)
    || node.kind === "error" && node.value === "#REF!"
    || node.kind === "binary" && [",", " ", ":"].includes(node.operator) && reference(node.left) && reference(node.right)
    || node.kind === "call" && ["INDEX", "OFFSET", "INDIRECT"].includes(node.name);
  if (!reference(result)) invalid("A workbook chart formula must produce a cell or named reference.");
  return result;
}
export function validateWorkbookCharts(workbook: WorkbookMetadata, objects: ReadonlyMap<string, WorkbookObject>, html: ParsedHtml | undefined, classes: ReadonlySet<string>, maxItems: number): number {
  const charts = new Map<string, string>(), sheets = new Set(workbook.sheets.map(sheet => sheet.sheetId)); let items = 0;
  const vector = (source: WorkbookChartVector): void => {
    if ("expression" in source) {
      chartReferenceExpression(source.expression);
      if (source.originSheetId !== undefined && !sheets.has(source.originSheetId)) invalid("A chart expression origin requires an existing worksheet.");
    }
    checkLimit(++items, maxItems);
    const cache = "expression" in source ? source.cache : source; if (!cache) return;
    items += cache.count; checkLimit(items, maxItems);
    if (cache.points.length > cache.count || source.valueType === "text" && cache.numberFormat !== undefined) invalid("A chart vector has an inconsistent point count or number format.");
    let previous = -1;
    for (const point of cache.points) {
      if (point.index <= previous || point.index >= cache.count) invalid("Chart vector indexes must be unique, increasing and within their declared vector."); previous = point.index;
      if (source.valueType === "text" && (typeof point.value !== "string" || point.numberFormat !== undefined) || source.valueType === "number" && typeof point.value === "string" && !WorkbookErrorValueSchema.safeParse(point.value).success) invalid("A chart vector point does not match its declared value type.");
    }
  };
  const style = (name: string | undefined): void => { if (name !== undefined && !classes.has(name)) invalid("A chart style refers to an undefined CSS class."); };
  for (const chart of workbook.charts ?? []) {
    const object = objects.get(chart.nodeId), node = html?.peekNode(chart.nodeId);
    if (charts.has(chart.chartId) || object?.kind !== "chart" || object.chartId !== chart.chartId || !node) invalid("A workbook chart must uniquely match its drawing object.");
    charts.set(chart.chartId, chart.nodeId); items++;
    if (chart.kind === "pie" && chart.series.length !== 1 || chart.kind !== "pie" && chart.firstSliceAngle !== undefined || chart.kind !== "column" && (chart.gapWidth !== undefined || chart.overlap !== undefined)) invalid("Chart options do not match the selected basic chart kind.");
    if (chart.minimum !== undefined && chart.maximum !== undefined && chart.minimum >= chart.maximum) invalid("Chart axis bounds must be increasing.");
    const series = new Set<string>(), native = new Set<number>();
    if (chart.categories) vector(chart.categories);
    for (const item of chart.series) {
      if (series.has(item.seriesId) || item.nativeIndex !== undefined && native.has(item.nativeIndex)) invalid("Chart series identities must be unique."); series.add(item.seriesId); if (item.nativeIndex !== undefined) native.add(item.nativeIndex);
      if (item.values.valueType !== "number") invalid("A chart series requires numeric source values.");
      vector(item.values); if (item.name?.kind === "reference") vector(item.name.reference);
    }
    const titleIds = new Set<string>();
    const titles = [{ nodeId: chart.titleNodeId, reference: chart.titleReference, cssClass: chart.titleCssClass, text: chart.titleText }, chart.categoryAxisTitle, chart.valueAxisTitle];
    if (chart.kind === "pie" && (chart.categoryAxisTitle || chart.valueAxisTitle)) invalid("A pie chart has no axis titles.");
    for (const title of titles) {
      if (!title) continue;
      if (title.reference) vector(title.reference);
      if (title.reference && title.nodeId) invalid("A chart title cannot have competing rich-text and cell-reference sources.");
      if (title.text && !title.nodeId) invalid("Chart title text-body properties require a rich source title.");
      if (title !== titles[0] && !title.reference && !title.nodeId) invalid("An axis title requires a text or reference source.");
      style(title.cssClass);
      if (title.nodeId) {
        const body = html?.peekNode(title.nodeId);
        if (titleIds.has(title.nodeId) || !body || body.tagName !== "div" || body.parentNodeId !== chart.nodeId || html!.directText(title.nodeId)?.trim() || body.childNodeIds.some(id => html!.peekNode(id)?.tagName !== "p")) invalid("A rich chart title requires its own paragraph container inside its chart.");
        titleIds.add(title.nodeId);
      }
    }
    if (node.childNodeIds.length !== titleIds.size || node.childNodeIds.some(id => !titleIds.has(id))) invalid("A chart drawing may contain only its registered source titles.");
    for (const key of ["plotCssClass", "legendCssClass", "textCssClass", "labelCssClass", "categoryAxisCssClass", "valueAxisCssClass", "gridCssClass", "categoryGridCssClass", "titleCssClass"] as const) style(chart[key]);
    const seriesStyles = new Set<string>(), points = new Set<string>();
    for (const entry of chart.seriesStyles ?? []) {
      if (!series.has(entry.seriesId) || seriesStyles.has(entry.seriesId)) invalid("Chart series styles require unique existing series."); seriesStyles.add(entry.seriesId); style(entry.cssClass); style(entry.labelCssClass); items++;
    }
    for (const entry of chart.pointStyles ?? []) {
      const key = `${entry.seriesId}:${entry.index}`;
      if (!series.has(entry.seriesId) || points.has(key) || entry.index > 0xffff_ffff) invalid("Chart point styles require unique points in existing series."); points.add(key); style(entry.cssClass); style(entry.labelCssClass); items++;
    }
    if (chart.labelContents && (new Set(chart.labelContents).size !== chart.labelContents.length || chart.showLabels !== !!chart.labelContents.length || chart.kind !== "pie" && chart.labelContents.includes("percent"))) invalid("Chart label contents are inconsistent with its kind or visibility.");
    checkLimit(items, maxItems);
  }
  for (const object of objects.values()) if (object.kind === "chart" && charts.get(object.chartId) !== object.nodeId) invalid("A chart object has no unique chart definition.");
  return items;
}
