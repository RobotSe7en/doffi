import { parseWorkbookFormula, printWorkbookFormula } from "@doffi/core";
import type { DeepReadonly, WorkbookCalculatedCell, WorkbookChart, WorkbookChartCache, WorkbookChartReference, WorkbookChartVector, WorkbookFormulaExpression, WorkbookModel } from "@doffi/core";
import type { OutputOpc } from "../export/opc.js";
import type { SourceElement, SourceTree } from "../export/source.js";
import { declaration, empty } from "../export/xml.js";
import { A, C, R } from "../ooxml/xml.js";
import { drawingString, shapeAppearance, WorkbookDrawingStyles } from "./drawing-style.js";
import type { XlsxExportLoss } from "./options.js";
import { cellKey } from "./sheet.js";
import { bool, element, formulaText, textElement, xstring } from "./xml.js";

type Chart = DeepReadonly<WorkbookChart>;
type Vector = DeepReadonly<WorkbookChartVector>;
type Cache = DeepReadonly<Omit<WorkbookChartCache, "state">>;
type Style = ReturnType<WorkbookDrawingStyles["classStyle"]>;
type Marker = { readonly marker?: string | undefined; readonly markerSize?: number | undefined };
const val = (tag: string, value: string | number | boolean | undefined): string =>
  value === undefined ? "" : empty("c:" + tag, { val: typeof value === "boolean" ? bool(value) : value });
const wrap = (tag: string, body: string): string => element("c:" + tag, body);
const appearance = (style: Style): string =>
  Object.keys(style.raw).some(key => /^--office-(fill|line|shadow)-/.test(key)) ? wrap("spPr", shapeAppearance(style)) : "";
const scope = { row: 0, column: 0 };

/** The owning sheet supplies native context for unqualified chart formulas.
 * Rebind references only when a move preserved another source sheet. */
function qualify(expression: string, sheet: string): string {
  const visit = (node: WorkbookFormulaExpression): WorkbookFormulaExpression => {
    switch (node.kind) {
      case "cell": case "name": case "axisRange":
        return node.sheet === undefined || node.sheet === "" ? { ...node, sheet } : node;
      case "range": return node.from.sheet === undefined || node.from.sheet === "" ?
        { ...node, from: { ...node.from, sheet }, to: { ...node.to, sheet } } : node;
      case "binary": return { ...node, left: visit(node.left), right: visit(node.right) };
      case "call": return { ...node, arguments: node.arguments.map(item => item && visit(item)) };
      default: return node;
    }
  };
  return printWorkbookFormula(visit(parseWorkbookFormula(expression, scope)), scope);
}
function cacheXml(cache: Cache, type: "num" | "str", tag: "Cache" | "Lit"): string {
  const code = type === "num" && cache.numberFormat !== undefined ? textElement("c:formatCode", xstring(cache.numberFormat)) : "";
  const points = cache.points.filter(point => point.value !== null || type === "num").map(point => element("c:pt",
    textElement("c:v", xstring(point.value === null ? "" : String(point.value))),
    { idx: point.index, formatCode: point.numberFormat === undefined ? undefined : xstring(point.numberFormat) })).join("");
  return wrap(type + tag, code + val("ptCount", cache.count) + points);
}

export class WorkbookChartsWriter {
  private count = 0;
  private points = 0;
  private readonly sheets: ReadonlyMap<string, string>;
  constructor(private readonly model: WorkbookModel, private readonly opc: OutputOpc,
    private readonly styles: WorkbookDrawingStyles, private readonly tree: SourceTree,
    private readonly calculated: ReadonlyMap<string, WorkbookCalculatedCell>, private readonly losses: XlsxExportLoss[],
    private readonly maxPoints: number) {
    this.sheets = new Map(model.snapshot().metadata.workbook.sheets.map(sheet => [sheet.sheetId, sheet.name]));
  }
  private cache(source: DeepReadonly<WorkbookChartReference>, sheetId: string, nodeId: string): Cache | undefined {
    const origin = { sheetId: source.originSheetId ?? sheetId, ...scope };
    const resolved = this.model.resolveReference(source.expression, origin);
    const fail = (): undefined => {
      this.losses.push({ kind: "metadata", sheetId, nodeId, message: "A current chart cache could not be produced. The worksheet reference is preserved." });
      return undefined;
    };
    if (resolved.kind !== "ranges") return fail();
    const size = resolved.areas.reduce((sum, area) => sum + (area.to.column - area.from.column + 1) * (area.to.row - area.from.row + 1), 0);
    if (size > this.maxPoints - this.points) return fail();
    this.points += size;
    const points: { index: number; value: string | number }[] = [];
    let index = 0, unavailable = false;
    for (const area of resolved.areas)
      for (let row = area.from.row; row <= area.to.row; row++)
        for (let column = area.from.column; column <= area.to.column; column++, index++) {
          const cell = this.model.getCell(area.sheetId, row, column);
          const computed = cell?.formula ? this.calculated.get(cellKey(area.sheetId, row, column)) : undefined;
          const value = cell?.formula ? computed?.state === "valid" ? computed.value : undefined : cell?.value;
          if (cell?.formula && !value) unavailable = true;
          if (!value || value.type === "blank") continue;
          if (source.valueType === "text") points.push({ index, value: value.type === "boolean" ? value.value ? "TRUE" : "FALSE" : String(value.value) });
          else if (value.type === "number" || value.type === "error") points.push({ index, value: value.value });
          else if (value.type === "boolean") points.push({ index, value: Number(value.value) });
        }
    return unavailable ? fail() : { count: size, points, ...(source.valueType === "number" ? { numberFormat: source.cache?.numberFormat ?? "General" } : {}) };
  }
  private vector(source: Vector, sheetId: string, nodeId: string, label = false): string {
    const type = source.valueType === "number" && !label ? "num" : "str";
    if (!("expression" in source)) return cacheXml(source, type, "Lit");
    const cache = this.cache(label && source.valueType !== "text" ? { ...source, valueType: "text" } : source, sheetId, nodeId);
    const expression = source.originSheetId && source.originSheetId !== sheetId ? qualify(source.expression, this.sheets.get(source.originSheetId)!) : source.expression;
    return wrap(type + "Ref", element("c:f", formulaText(expression)) + (cache ? cacheXml(cache, type, "Cache") : ""));
  }
  private marker(source: Marker | undefined, style: Style): string {
    if (!source?.marker && source?.markerSize === undefined && !Object.keys(style.raw).some(key => key.startsWith("--office-marker-"))) return "";
    const remap = (css: Style["raw"]): Style["raw"] => Object.fromEntries(Object.entries(css)
      .filter(([key]) => key.startsWith("--office-marker-")).map(([key, value]) => [key.replace("--office-marker-", "--office-"), value]));
    return wrap("marker", val("symbol", source?.marker) + val("size", source?.markerSize === undefined ? undefined : Math.max(2, Math.min(72, Math.round(source.markerSize)))) +
      appearance({ css: remap(style.css), raw: remap(style.raw) }));
  }
  private labels(chart: Chart, name: string | undefined, points: NonNullable<Chart["pointStyles"]> = []): string {
    const contents = chart.showLabels ? chart.labelContents ?? ["value"] : [];
    const options = (styleName: string | undefined): string => {
      const style = this.styles.classStyle(styleName), { css } = style, format = drawingString(css["--office-label-number-format"]);
      const separator = drawingString(css["--office-label-separator"]);
      return (format === undefined ? "" : empty("c:numFmt", { formatCode: xstring(format), sourceLinked: bool(css["--office-label-number-format-source-linked"] !== "false") })) +
        appearance(style) + this.styles.textStyle(styleName) + val("dLblPos", css["--office-label-position"]) +
        val("showVal", contents.includes("value")) + val("showCatName", contents.includes("category")) +
        val("showSerName", contents.includes("series")) + val("showPercent", contents.includes("percent")) +
        (separator === undefined ? "" : textElement("c:separator", separator));
    };
    const style = this.styles.classStyle(name);
    const leader = style.css["--office-label-show-leader-lines"];
    const leaderStyle = (css: Style["css"]): Style["css"] => Object.fromEntries(Object.entries(css)
      .filter(([key]) => key.startsWith("--office-label-leader-")).map(([key, value]) => [key.replace("--office-label-leader-", "--office-"), value]));
    const paint = appearance({ css: leaderStyle(style.css), raw: leaderStyle(style.raw) });
    return wrap("dLbls", points.map(point => point.labelCssClass ? wrap("dLbl", val("idx", point.index) + options(point.labelCssClass)) : "").join("") +
      options(name) + val("showLeaderLines", leader === undefined ? undefined : leader === "true") + (paint ? wrap("leaderLines", paint) : ""));
  }
  private title(source: Chart["categoryAxisTitle"], sheetId: string, nodeId: string, part: string): string {
    if (!source?.nodeId && !source?.reference) return "";
    const body = source.reference ? this.vector(source.reference, sheetId, nodeId, true) :
      this.styles.richBody("c:rich", this.tree.byId.get(source.nodeId!)!, part, source.text);
    return wrap("title", wrap("tx", body) + val("overlay", false) +
      appearance(this.styles.classStyle(source.cssClass)) + this.styles.textStyle(source.cssClass));
  }
  private axis(chart: Chart, sheetId: string, part: string, category: boolean): string {
    const name = category ? chart.categoryAxisCssClass : chart.valueAxisCssClass;
    const style = this.styles.classStyle(name), { css } = style, get = (key: string): string | undefined => css["--office-axis-" + key];
    const grid = this.styles.classStyle(category ? chart.categoryGridCssClass : chart.gridCssClass);
    const format = drawingString(get("number-format")) ?? (category ? undefined : chart.numberFormat);
    const title = this.title(category ? chart.categoryAxisTitle : chart.valueAxisTitle, sheetId, chart.nodeId, part);
    const axis = val("axId", category ? 1 : 2) + wrap("scaling", val("orientation", (category ? chart.reverseCategories : chart.reverseValues) ? "maxMin" : "minMax") +
      (category ? "" : val("max", chart.maximum) + val("min", chart.minimum))) +
      val("delete", (category ? chart.showCategoryAxis : chart.showValueAxis) === false) + val("axPos", get("position") ?? (category ? "b" : "l")) +
      (Object.keys(grid.raw).length ? wrap("majorGridlines", appearance(grid)) : "") + title +
      (format === undefined ? "" : empty("c:numFmt", { formatCode: xstring(format), sourceLinked: bool(get("number-format-source-linked") !== "false") })) +
      val("majorTickMark", get("major-tick")) + val("minorTickMark", get("minor-tick")) + val("tickLblPos", get("label-position")) +
      appearance(style) + this.styles.textStyle(name) + val("crossAx", category ? 2 : 1) +
      (get("crosses-at") === undefined ? val("crosses", get("crosses") ?? "autoZero") : val("crossesAt", get("crosses-at")));
    const option = (key: string): boolean | undefined => get(key) === undefined ? undefined : get(key) === "true";
    return category ? wrap("catAx", axis + val("auto", option("auto")) + val("lblAlgn", get("label-align")) + val("lblOffset", get("label-offset")) +
      val("tickLblSkip", get("label-skip")) + val("tickMarkSkip", get("tick-skip")) + val("noMultiLvlLbl", option("noMultiLvlLbl"))) :
      wrap("valAx", axis + val("crossBetween", get("cross-between") ?? "between") + val("majorUnit", chart.majorUnit) +
        val("minorUnit", get("minor-unit") === "auto" ? undefined : get("minor-unit")));
  }

  write(chart: Chart, node: SourceElement, sheetId: string, drawingPart: string): string {
    const name = "chart" + (++this.count) + ".xml", part = "xl/charts/" + name;
    const out = this.opc.part(part, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    const relationship = this.opc.relate(drawingPart, "chart", "../charts/" + name);
    const title = this.title({ ...(chart.titleNodeId ? { nodeId: chart.titleNodeId } : {}), ...(chart.titleReference ? { reference: chart.titleReference } : {}),
      ...(chart.titleCssClass ? { cssClass: chart.titleCssClass } : {}), ...(chart.titleText ? { text: chart.titleText } : {}) }, sheetId, chart.nodeId, part);
    out.add(declaration + '<c:chartSpace xmlns:c="' + C + '" xmlns:a="' + A + '" xmlns:r="' + R + '">' +
      val("date1904", chart.dateSystem === undefined ? undefined : chart.dateSystem === "1904") + val("lang", chart.language) + "<c:chart>" + title +
      val("autoTitleDeleted", chart.titleVisible === false || !title) + "<c:plotArea>");
    const kind = chart.kind === "column" ? "barChart" : chart.kind === "line" ? "lineChart" : "pieChart";
    out.add("<c:" + kind + ">" + (kind === "barChart" ? val("barDir", "col") + val("grouping", "clustered") :
      kind === "lineChart" ? val("grouping", "standard") : "") + val("varyColors", chart.varyColors));
    const categories = chart.categories ? wrap("cat", this.vector(chart.categories, sheetId, chart.nodeId)) : "";
    const nativeIndices = new Set(chart.series.flatMap(series => series.nativeIndex === undefined ? [] : [series.nativeIndex]));
    let next = 0;
    for (const [index, series] of chart.series.entries()) {
      while (nativeIndices.has(next)) next++;
      const nativeIndex = series.nativeIndex ?? next++;
      const entry = chart.seriesStyles?.find(style => style.seriesId === series.seriesId);
      const style = this.styles.classStyle(entry?.cssClass);
      const points = chart.pointStyles?.filter(point => point.seriesId === series.seriesId) ?? [];
      const seriesName = series.name?.kind === "text" ? textElement("c:v", xstring(series.name.text)) :
        series.name ? this.vector(series.name.reference, sheetId, chart.nodeId, true) : "";
      const invert = style.css["--office-invert-negative"];
      out.add("<c:ser>" + val("idx", nativeIndex) + val("order", index) + (seriesName ? wrap("tx", seriesName) : "") + appearance(style) +
        (kind === "lineChart" ? this.marker(entry, style) : kind === "barChart" && invert !== undefined ? val("invertIfNegative", invert === "true") : ""));
      for (const point of points) {
        const paint = this.styles.classStyle(point.cssClass), invertPoint = paint.css["--office-invert-negative"];
        out.add(wrap("dPt", val("idx", point.index) + val("invertIfNegative", invertPoint === undefined ? undefined : invertPoint === "true") +
          this.marker(point, paint) + appearance(paint)));
      }
      if (entry?.labelCssClass || points.some(point => point.labelCssClass)) out.add(this.labels(chart, entry?.labelCssClass ?? chart.labelCssClass, points));
      out.add(categories + wrap("val", this.vector(series.values, sheetId, chart.nodeId)) + (kind === "lineChart" ? val("smooth", false) : "") + "</c:ser>");
    }
    if (chart.showLabels || chart.labelCssClass) out.add(this.labels(chart, chart.labelCssClass));
    if (kind === "lineChart") out.add(val("marker", chart.nativeShowMarkers) + val("smooth", false));
    if (kind === "barChart") out.add(val("gapWidth", chart.gapWidth === undefined ? undefined : Math.round(chart.gapWidth)) +
      val("overlap", chart.overlap === undefined ? undefined : Math.round(chart.overlap)));
    out.add((kind === "pieChart" ? val("firstSliceAng", chart.firstSliceAngle === undefined ? undefined : Math.round(chart.firstSliceAngle)) : val("axId", 1) + val("axId", 2)) + "</c:" + kind + ">");
    if (kind !== "pieChart") out.add(this.axis(chart, sheetId, part, true) + this.axis(chart, sheetId, part, false));
    out.add(appearance(this.styles.classStyle(chart.plotCssClass)) + "</c:plotArea>");
    if (chart.legend !== "none") out.add(wrap("legend", val("legendPos", chart.legend[0]) + val("overlay", false) +
      appearance(this.styles.classStyle(chart.legendCssClass)) + this.styles.textStyle(chart.legendCssClass)));
    out.add(val("plotVisOnly", chart.visibleCellsOnly) + val("dispBlanksAs", chart.blankAs) + val("showDLblsOverMax", chart.showLabelsOverMaximum) +
      "</c:chart>" + appearance(this.styles.style(node)) + this.styles.textStyle(chart.textCssClass) + "</c:chartSpace>");
    out.finish();
    return empty("c:chart", { "r:id": relationship });
  }
}
