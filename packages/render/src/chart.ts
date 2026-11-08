import type { ChartAppearance, DeepReadonly, WorkbookNumberDisplayPart } from "@doffi/core";
import { DocumentError } from "@doffi/core";
import type { ChartTextDecoration } from "./chart-text-decoration.js";
import { chartFontDecoration } from "./chart-style-probe.js";
import { markViewNode, ownViewAttribute } from "./view-nodes.js";
import { cssStringValue, numeric, property, resolvedColor, strokeAttributes, svgFill } from "./presentation/appearance.js";
import { svgElement as createSvgElement } from "./presentation/svg.js";
import type { DrawingGradientBounds, DrawingGradientOrientation } from "./drawing-gradient.js";
import { appendChartMarker } from "./chart-marker.js";
import { chartAxisTicks, chartCategoryAxis, chartAxisPlacement, chartValueScale, chartValueFraction } from "./chart-axis.js";
import { officeFontSegments } from "./font-scripts.js";
import type { OfficeFontRole } from "./font-scripts.js";
import { chartStrokeOutset, clipChartValueSegment, MAX_CHART_PAINT_EXTENT } from "./chart-clipping.js";

/** Derived plot values. The owning format adapter supplies its authoritative
 * source binding, visibility, coercion and explicit number-display policy. */
export interface BasicChartData {
  readonly categories: readonly string[];
  /** Original source positions survive hidden-point compaction. */
  readonly sourceIndexes?: readonly number[];
  readonly series: readonly {
    readonly seriesId: string; readonly name: string; readonly values: readonly (number | null)[];
    readonly hidden?: boolean;
    /** Explicit missing-point behavior overrides the ordinary blank policy. */
    readonly missing?: readonly ("gap" | "span" | undefined)[];
  }[];
}
export type BasicChartText = string | { readonly parts: readonly WorkbookNumberDisplayPart[]; readonly availableWidth: number };
export interface BasicChartOptions {
  /** The adapter may resolve local fonts on private connected probes after
   * rendering. These runs are derived labels, without canonical text offsets. */
  readonly onTextRun?: (run: BasicChartTextRun) => void;
  /** Resolve styles against the source ancestry, which may be outside the view.
   * The adapter owns any temporary probes and releases them even on failure. */
  readonly style: (classes?: string | readonly string[]) => CSSStyleDeclaration;
  readonly formatNumber: (value: number, format?: string, context?: BasicChartNumberContext) => BasicChartText;
  readonly formatCategory?: (index: number, style: CSSStyleDeclaration) => BasicChartText;
  readonly accountText?: (codeUnits: number) => void;
  readonly title?: HTMLElement;
  /** The owning compositor may reserve measured text boxes around the plot.
   * Presentation callers retain the established automatic layout. */
  readonly layout?: BasicChartLayout;
  readonly orientation?: DrawingGradientOrientation;
  readonly negativeColumns?: "white";
  readonly showLabelsOverMaximum?: boolean;
  /** The owning format selects negative-value semantics explicitly. Value
   * labels keep their signed source numbers under the magnitude policy. */
  readonly pieValues?: "nonnegative" | "magnitude";
  /** The owning adapter applies native single-series/default semantics. */
  readonly varyColors?: boolean;
}
export interface BasicChartTextRun {
  readonly node: SVGTextElement | SVGTSpanElement; readonly text: string;
  readonly role: OfficeFontRole; readonly family: string; readonly style: CSSStyleDeclaration;
}
export type BasicChartNumberContext = { readonly kind: "value-axis"; readonly style: CSSStyleDeclaration }
  | { readonly kind: "data-label"; readonly content: "value" | "percent"; readonly seriesId: string; readonly index: number; readonly style: CSSStyleDeclaration };
export interface BasicChartRect { readonly x: number; readonly y: number; readonly width: number; readonly height: number }
export interface BasicChartLayout { readonly plot: BasicChartRect; readonly title?: BasicChartRect; readonly legend?: BasicChartRect }
export interface BasicChartScene {
  readonly svg: SVGSVGElement; readonly serial: string; readonly width: number; readonly height: number;
  readonly plot: BasicChartRect; readonly legend?: BasicChartRect;
  readonly legendPosition: ChartAppearance["legend"];
  readonly labels: readonly { readonly node: SVGTextElement; readonly outsidePie: boolean; readonly anchorX: number; readonly anchorY: number; readonly leader?: Readonly<Record<string, string | number>> }[];
  readonly automaticCategoryLabels: boolean;
  readonly textDecorations?: readonly ChartTextDecoration[];
  readonly numberTexts: readonly { readonly parts: readonly WorkbookNumberDisplayPart[]; readonly availableWidth: number; readonly fields: readonly SVGTSpanElement[] }[];
}
const defaultPalette = ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"];
// Native auto delegates its shape to the application. Keep one stable shape
// per original series, even when other series or source points are hidden.
const automaticMarkers = ["circle", "square", "diamond", "triangle", "x", "plus", "star", "dash", "dot"] as const;
export function renderBasicChart(node: HTMLElement, chart: DeepReadonly<ChartAppearance>, data: BasicChartData, width: number, height: number, serial: string, account: (amount: number) => void, options: BasicChartOptions): BasicChartScene {
  const svgElement: typeof createSvgElement = (document, tag, attributes) => { account(1); return createSvgElement(document, tag, attributes); };
  const document = node.ownerDocument, formatNumber = options.formatNumber, styles = new Map<string, CSSStyleDeclaration>();
  // Validated source CSS has no structural pseudo-classes. One probe per
  // role chain is enough for this immutable render, including repeated labels.
  const style = (classes?: string | readonly string[]): CSSStyleDeclaration => {
    const key = JSON.stringify(classes ?? ""); let computed = styles.get(key);
    if (!computed) { computed = options.style(classes); styles.set(key, computed); }
    return computed;
  };
  const seriesDefinitions = new Map((chart.seriesStyles ?? []).map(series => [series.seriesId, series]));
  const pointKey = (seriesId: string, index: number): string => JSON.stringify([seriesId, index]);
  const pointDefinitions = new Map((chart.pointStyles ?? []).map(point => [pointKey(point.seriesId, point.index), point]));
  const seriesIndexes = new Map(data.series.map((series, index) => [series.seriesId, index]));
  const sourceIndex = (index: number): number => data.sourceIndexes?.[index] ?? index;
  data = { ...data, series: data.series.filter(series => !series.hidden) };
  let paintSerial = 0;
  let plotLayer: SVGElement;
  const svg = markViewNode(svgElement(document, "svg", { width: "100%", height: "100%", viewBox: `0 0 ${width} ${height}`, role: "img", "aria-label": chart.kind + " chart", "data-sdk-chart": chart.kind }));
  svg.style.cssText = "position:absolute;inset:0;overflow:visible;pointer-events:none";
  node.prepend(svg); plotLayer = svg;
  const base = style(chart.textCssClass), legendStyle = style(chart.legendCssClass);
  const palette = defaultPalette.map((fallback, index) => {
    const name = "--office-theme-accent" + (index + 1);
    return resolvedColor(document, base, name, property(base, name, fallback));
  });
  const varyColors = options.varyColors ?? chart.kind === "pie";
  const fontScale = (s: CSSStyleDeclaration): number => {
    const value = Number(property(s, "--office-text-font-scale", "1"));
    if (!Number.isFinite(value) || value < 0 || value > 1) throw new DocumentError("InvalidDocument", "A chart text font scale must be in [0,1].");
    return value;
  };
  const font = (s: CSSStyleDeclaration): Record<string, string> => ({ "font-size": String(parseFloat(s.fontSize) * 0.75 * fontScale(s)), "font-weight": s.fontWeight, "font-style": s.fontStyle,
    "letter-spacing": String((parseFloat(s.letterSpacing) || 0) * .75 * fontScale(s)), "word-spacing": String((parseFloat(s.wordSpacing) || 0) * .75 * fontScale(s)),
    fill: property(s, "--office-text-color") ? resolvedColor(document, s, "--office-text") : s.color });
  const numberTexts: BasicChartScene["numberTexts"][number][] = [], textDecorations: ChartTextDecoration[] = [];
  const text = (value: BasicChartText, x: number, y: number, s = base, anchor = "middle"): SVGTextElement => {
    options.accountText?.(typeof value === "string" ? value.length : value.parts.reduce((sum, part) => sum + (part.kind === "text" ? part.text.length : part.character.length), 0));
    const attributes = font(s), label = svgElement(document, "text", { x, y, "text-anchor": anchor, "dominant-baseline": "middle", ...attributes });
    const decoration = chartFontDecoration(s, attributes.fill!);
    label.style.textDecorationLine = "none";
    label.dataset.sdkChartDecorationLine = decoration["text-decoration-line"]!; label.dataset.sdkChartDecorationStyle = decoration["text-decoration-style"]!; label.dataset.sdkChartDecorationColor = decoration["text-decoration-color"]!;
    if (decoration["text-decoration-line"] !== "none") textDecorations.push({ node: label, line: decoration["text-decoration-line"]!, style: decoration["text-decoration-style"]!, color: decoration["text-decoration-color"]!, size: Number(attributes["font-size"]) });
    const run = (target: SVGTextElement | SVGTSpanElement, value: string, role: OfficeFontRole): void => {
      const family = property(s, "--office-font-" + role, property(s, "--office-minor-" + role, s.fontFamily));
      if (!document.defaultView!.CSS.supports("font-family", family)) throw new DocumentError("InvalidDocument", "Chart text has an invalid source font family.");
      target.dataset.officeScript = role; target.setAttribute("font-family", family);
      if (value) { account(1); target.textContent = value; options.onTextRun?.({ node: target, text: value, role, family, style: s }); }
    };
    const line = (target: SVGTextElement | SVGTSpanElement, value: string): void => {
      const segments = officeFontSegments(value);
      if (segments.length <= 1) run(target, value, segments[0]?.role ?? "ascii");
      else for (const segment of segments) { const span = svgElement(document, "tspan", {}); target.append(span); run(span, segment.text, segment.role); }
    };
    if (typeof value !== "string") {
      const fields = value.parts.map(part => {
        const span = svgElement(document, "tspan", { "xml:space": "preserve" }); label.append(span);
        span.style.whiteSpace = "pre";
        const content = part.kind === "text" ? part.text : part.character;
        line(span, content);
        if (part.kind !== "text") span.setAttribute("visibility", "hidden");
        return span;
      });
      numberTexts.push({ parts: value.parts, availableWidth: value.availableWidth, fields });
      svg.append(label); return label;
    }
    const lines = value.split("\n");
    if (lines.length === 1) line(label, value);
    else {
      const height = (parseFloat(s.lineHeight) || parseFloat(s.fontSize) * 1.2) * .75 * fontScale(s) * (1 - numeric(s, "--office-text-line-spacing-reduction", 0));
      lines.forEach((value, index) => { const span = svgElement(document, "tspan", { x, dy: String(index ? height : -(lines.length - 1) * height / 2), "data-sdk-chart-text-line": "" }); label.append(span); line(span, value); });
    }
    svg.append(label); return label;
  };
  const catStyle = style(chart.categoryAxisCssClass), valStyle = style(chart.valueAxisCssClass), gridStyle = style(chart.gridCssClass);
  const scale = data.series.length ? chartValueScale(chart, data.series) : { minimum: 0, maximum: 1, unit: 1 };
  const axes = chart.kind === "pie" ? undefined : chartAxisPlacement(chart, catStyle, valStyle, scale.minimum, scale.maximum, Math.max(1, data.categories.length));
  const categoryTop = axes?.categoryTop ?? false, valueRight = axes?.valueRight ?? false;
  const plotStyle = style(chart.plotCssClass);
  const title = options.title;
  const titleHeight = title && chart.titleVisible !== false ? Math.min(54, height * 0.2) : 12;
  if (title) { ownViewAttribute(title, "data-sdk-chart-title"); title.setAttribute("data-sdk-chart-title", ""); }
  if (title) {
    const rect = options.layout?.title;
    title.style.cssText += `;position:absolute;${rect ? `left:${rect.x}pt;top:${rect.y}pt;width:${rect.width}pt;height:${rect.height}pt` : `left:6pt;right:6pt;top:20pt;height:${titleHeight}pt`};margin:0;text-align:center;${chart.titleVisible === false ? "display:none" : ""}`;
  }
  const entries = chart.kind === "pie" || varyColors && data.series.length === 1 ? data.categories.map((name, i) => ({ name, series: data.series[0]!, index: i })) : data.series.map(series => ({ name: series.name, series, index: -1 }));
  const legendBand = chart.legend === "none" ? 0 : chart.legend === "bottom" || chart.legend === "top" ? 28 : Math.min(width * 0.25, 100);
  const left = options.layout?.plot.x ?? (chart.kind === "pie" || valueRight ? 12 : 28) + (chart.legend === "left" ? legendBand : 0);
  const top = options.layout?.plot.y ?? titleHeight + (chart.kind === "pie" ? 15 : 18) + (chart.legend === "top" ? legendBand : 0) + (categoryTop ? 16 : 0);
  const right = options.layout ? left + options.layout.plot.width : width - (chart.kind === "pie" ? 12 : valueRight ? 28 : 16) - (chart.legend === "right" ? legendBand : 0), bottom = options.layout ? top + options.layout.plot.height : height - (chart.kind === "pie" ? 8 : categoryTop ? 12 : 26) - (chart.legend === "bottom" ? legendBand : 0);
  const w = Math.max(1, right - left), h = Math.max(1, bottom - top);
  const labels: BasicChartScene["labels"][number][] = [];
  const scene: BasicChartScene = { svg, serial, width, height, plot: { x: left, y: top, width: w, height: h }, legendPosition: chart.legend, ...(options.layout?.legend ? { legend: options.layout.legend } : {}), labels,
    automaticCategoryLabels: !property(catStyle, "--office-axis-label-skip"), numberTexts, textDecorations };
  svg.append(svgElement(document, "rect", { x: left, y: top, width: w, height: h, fill: svgFill(document, plotStyle, svg, serial + "-plot", { x: left, y: top, width: w, height: h }, "--office-fill", "none", account, options.orientation), ...strokeAttributes(document, plotStyle) }));
  if (!data.series.length) return scene;
  const seriesStyle = new Map(data.series.map(series => [series.seriesId, style(seriesDefinitions.get(series.seriesId)?.cssClass)]));
  const pointStyle = (seriesId: string, index: number): CSSStyleDeclaration => {
    const series = seriesDefinitions.get(seriesId), point = pointDefinitions.get(pointKey(seriesId, sourceIndex(index)));
    return point ? style([series?.cssClass, point.cssClass].filter((name): name is string => !!name)) : seriesStyle.get(seriesId)!;
  };
  const fill = (seriesId: string, index: number, s: CSSStyleDeclaration, bounds: DrawingGradientBounds, negativeColumn = false): string => {
    const invert = property(s, "--office-invert-negative", "false");
    if (invert !== "true" && invert !== "false") throw new DocumentError("InvalidDocument", "A chart negative-fill flag must be a boolean.");
    if (negativeColumn && options.negativeColumns === "white" && invert === "true" && property(s, "--office-fill-kind") !== "none") return "#ffffff";
    return svgFill(document, s, svg, serial + "-paint-" + ++paintSerial, bounds, "--office-fill", palette[(varyColors ? sourceIndex(index) : seriesIndexes.get(seriesId)!) % palette.length]!, account, options.orientation);
  };
  const marker = (seriesId: string, index: number, x: number, y: number | undefined, role: "point" | "legend"): void => {
    const definition = seriesDefinitions.get(seriesId), override = index < 0 ? undefined : pointDefinitions.get(pointKey(seriesId, sourceIndex(index)));
    const requested = override?.marker ?? definition?.marker ?? "none", size = override?.markerSize ?? definition?.markerSize ?? 5;
    const symbol = requested === "auto" ? automaticMarkers[seriesIndexes.get(seriesId)! % automaticMarkers.length]! : requested;
    if (symbol === "none") return;
    if (!Number.isFinite(size) || size <= 0 || size > MAX_CHART_PAINT_EXTENT) throw new DocumentError("ResourceLimitExceeded", "A chart marker exceeds its finite coordinate budget.");
    const point = index < 0 ? seriesStyle.get(seriesId)! : pointStyle(seriesId, index), position = Math.max(0, index);
    const automatic = palette[(varyColors ? sourceIndex(position) : seriesIndexes.get(seriesId)!) % palette.length]!;
    const stroke = strokeAttributes(document, point, "--office-marker-line", automatic), outset = chartStrokeOutset(stroke);
    if (y === undefined || role === "point" && (y + size / 2 + outset < top || y - size / 2 - outset > bottom)) return;
    const bounds = { x: x - size / 2, y: y - size / 2, width: size, height: size };
    const color = property(point, "--office-marker-fill-kind") || property(point, "--office-marker-fill-color")
      ? svgFill(document, point, svg, serial + "-marker-" + ++paintSerial, bounds, "--office-marker-fill", automatic, account, options.orientation)
      : fill(seriesId, position, point, bounds);
    const group = appendChartMarker(role === "point" ? plotLayer : svg, symbol, size, x, y, color, stroke, svgElement)!;
    group.dataset.sdkChartMarkerRole = role; group.dataset.seriesId = seriesId;
    if (requested === "auto") group.dataset.sdkChartMarkerAuto = "true";
    if (index >= 0) group.dataset.sdkChartSourceIndex = String(sourceIndex(index));
  };
  const label = (series: BasicChartData["series"][number], index: number, value: number, x: number, y: number, percent?: number, baseline?: number, pieAnchor?: readonly [number, number]): void => {
    if (!chart.showLabels) return;
    if (chart.kind !== "pie" && options.showLabelsOverMaximum === false && (value < scale.minimum || value > scale.maximum)) return;
    if (chart.kind !== "pie" && options.showLabelsOverMaximum === true) y = Math.max(top, Math.min(bottom, y));
    const anchorX = pieAnchor?.[0] ?? x, anchorY = pieAnchor?.[1] ?? y;
    const seriesDefinition = seriesDefinitions.get(series.seriesId), point = pointDefinitions.get(pointKey(series.seriesId, sourceIndex(index)));
    const s = style([chart.labelCssClass, seriesDefinition?.labelCssClass, point?.labelCssClass].filter((name): name is string => !!name));
    const contents = chart.labelContents ?? ["value"], separator = cssStringValue(property(s, "--office-label-separator", ", "));
    const values = contents.map(kind => kind === "category" ? data.categories[index]! : kind === "series" ? series.name : kind === "percent" ? formatNumber(percent ?? 0, "0%", { kind: "data-label", content: "percent", seriesId: series.seriesId, index, style: s })
      : formatNumber(value, cssStringValue(property(s, "--office-label-number-format", chart.numberFormat ?? "General")), { kind: "data-label", content: "value", seriesId: series.seriesId, index, style: s }));
    const position = property(s, "--office-label-position", chart.kind === "pie" ? "ctr" : "outEnd");
    let anchor = "middle";
    if (chart.kind === "column" && baseline !== undefined) {
      const origin = Math.max(scale.minimum, Math.min(scale.maximum, 0));
      const direction = (value < origin ? -1 : 1) * (chart.reverseValues ? -1 : 1);
      y = position === "ctr" ? (y + baseline) / 2 : position === "inBase" ? baseline - direction * 8 : position === "inEnd" ? y + direction * 8 : y - direction * 8;
    } else if (chart.kind === "line") {
      if (position === "l" || position === "r") { x += position === "l" ? -8 : 8; anchor = position === "l" ? "end" : "start"; }
      else if (position !== "ctr") y += position === "b" ? 8 : -8;
    } else if (chart.kind === "pie" && ["outEnd", "inEnd"].includes(position)) {
      const cx = left + w / 2, cy = top + h / 2, factor = position === "outEnd" ? 1.7 : 1.3; x = cx + (x - cx) * factor; y = cy + (y - cy) * factor;
    }
    const content: BasicChartText = values.every(value => typeof value === "string") ? values.join(separator) : {
      parts: values.flatMap((value, index): WorkbookNumberDisplayPart[] => [
        ...(index ? [{ kind: "text" as const, text: separator }] : []),
        ...(typeof value === "string" ? [{ kind: "text" as const, text: value }] : value.parts),
      ]), availableWidth: Math.max(0, ...values.map(value => typeof value === "string" ? 0 : value.availableWidth)),
    };
    const displayed = text(content, x, y, s, anchor); displayed.dataset.sdkChartLabel = "data-label";
    displayed.dataset.seriesId = series.seriesId; displayed.dataset.pointIndex = String(sourceIndex(index));
    const leaders = property(s, "--office-label-show-leader-lines", "false");
    if (leaders !== "true" && leaders !== "false") throw new DocumentError("InvalidDocument", "A chart leader-line flag must be a boolean.");
    labels.push({ node: displayed, outsidePie: chart.kind === "pie" && position === "outEnd", anchorX, anchorY,
      ...(leaders === "true" ? { leader: strokeAttributes(document, s, "--office-label-leader-line", "#898989") } : {}) });
  };
  if (chart.kind === "pie") {
    const series = data.series[0]!;
    const magnitude = (value: number | null): number => {
      if (value === null) return 0;
      if (value < 0 && options.pieValues !== "magnitude") throw new DocumentError("InvalidDocument", "Negative pie values require an explicit magnitude policy.");
      return Math.abs(value);
    };
    const normalizer = series.values.reduce<number>((max, value) => Math.max(max, magnitude(value)), 0) || 1;
    const sum = series.values.reduce<number>((total, value) => total + magnitude(value) / normalizer, 0), cx = left + w / 2, cy = top + h / 2, radius = Math.min(w, h) / 2;
    let angle = (chart.firstSliceAngle ?? 0) * Math.PI / 180 - Math.PI / 2;
    series.values.forEach((value, index) => {
      if (value === null || value === 0 || !sum) return;
      const fraction = magnitude(value) / normalizer / sum, sweep = fraction * Math.PI * 2, end = angle + sweep, middle = angle + sweep / 2;
      const points = [[cx, cy], [cx + radius * Math.cos(angle), cy + radius * Math.sin(angle)], [cx + radius * Math.cos(end), cy + radius * Math.sin(end)]];
      for (let quadrant = Math.ceil(angle / (Math.PI / 2)); quadrant * Math.PI / 2 <= end; quadrant++) points.push([cx + radius * Math.cos(quadrant * Math.PI / 2), cy + radius * Math.sin(quadrant * Math.PI / 2)]);
      const xs = points.map(point => point[0]!), ys = points.map(point => point[1]!);
      const bounds = { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) };
      const s = pointStyle(series.seriesId, index), appearance = { fill: fill(series.seriesId, index, s, bounds), ...strokeAttributes(document, s) };
      if (sweep >= Math.PI * 2 - 1e-9) svg.append(svgElement(document, "circle", { cx, cy, r: radius, ...appearance }));
      else svg.append(svgElement(document, "path", { d: `M${cx} ${cy}L${cx + radius * Math.cos(angle)} ${cy + radius * Math.sin(angle)}A${radius} ${radius} 0 ${sweep > Math.PI ? 1 : 0} 1 ${cx + radius * Math.cos(end)} ${cy + radius * Math.sin(end)}Z`, ...appearance }));
      label(series, index, value, cx + radius * 0.65 * Math.cos(middle), cy + radius * 0.65 * Math.sin(middle), fraction, undefined, [cx + radius * Math.cos(middle), cy + radius * Math.sin(middle)]); angle = end;
    });
  } else {
    const { minimum, maximum, unit } = scale;
    const majorTicks = chartAxisTicks(minimum, maximum, unit, 1024, 0.001);
    const y = (v: number): number => {
      const fraction = chartValueFraction(v, minimum, maximum, !chart.reverseValues), position = top + fraction * h;
      if (!Number.isFinite(position)) throw new DocumentError("ResourceLimitExceeded", "Chart values exceed the finite preview coordinate scale.");
      return position;
    };
    const axisX = left + axes!.x * w;
    const tick = (x: number, y: number, vertical: boolean, s: CSSStyleDeclaration, minor = false): void => {
      const kind = property(s, "--office-axis-" + (minor ? "minor" : "major") + "-tick", "cross"); if (kind === "none") return;
      const length = minor ? 2 : 4, outside = vertical ? axes!.categoryOutside : axes!.valueOutside;
      const start = (kind === "out" ? 0 : -length) * outside, end = (kind === "in" ? 0 : length) * outside;
      svg.append(svgElement(document, "line", { x1: x + (vertical ? 0 : start), x2: x + (vertical ? 0 : end), y1: y + (vertical ? start : 0), y2: y + (vertical ? end : 0), ...strokeAttributes(document, s, "--office-line", "#898989") }));
    };
    for (const value of majorTicks) {
      const position = y(value);
      if (chart.gridCssClass) svg.append(svgElement(document, "line", { x1: left, y1: position, x2: right, y2: position, "data-sdk-chart-grid": "value", ...strokeAttributes(document, gridStyle, "--office-line", "#d9d9d9") }));
      if (chart.showValueAxis !== false) {
        tick(axisX, position, false, valStyle);
        if (property(valStyle, "--office-axis-label-position") !== "none") text(formatNumber(value, chart.numberFormat, { kind: "value-axis", style: valStyle }), left + axes!.valueLabel * w + (valueRight ? 7 : -7), position, valStyle, valueRight ? "start" : "end").dataset.sdkChartLabel = "value-axis";
      }
    }
    const minorUnit = numeric(valStyle, "--office-axis-minor-unit", unit / 5);
    if (minorUnit > 0 && chart.showValueAxis !== false) {
      for (const [index, value] of chartAxisTicks(minimum, maximum, minorUnit, 4096).entries()) {
        const offset = index === 0 ? 0 : index * (minorUnit / unit);
        if (Math.abs(offset - Math.round(offset)) > 0.001) tick(axisX, y(value), false, valStyle, true);
      }
    }
    const baseline = y(Math.max(minimum, Math.min(maximum, 0))), axisY = top + axes!.y * h;
    if (chart.showCategoryAxis !== false) svg.append(svgElement(document, "line", { x1: left, y1: axisY, x2: right, y2: axisY, "data-sdk-chart-axis": "category", ...strokeAttributes(document, catStyle, "--office-line", "#b7b7b7") }));
    if (chart.showValueAxis !== false) svg.append(svgElement(document, "line", { x1: axisX, y1: top, x2: axisX, y2: bottom, "data-sdk-chart-axis": "value", ...strokeAttributes(document, valStyle, "--office-line", "#b7b7b7") }));
    const count = Math.max(1, data.categories.length), category = chartCategoryAxis(count, left, w, property(valStyle, "--office-axis-cross-between", "between") === "between", chart.reverseCategories ?? false);
    const group = category.interval, x = category.position;
    if (chart.categoryGridCssClass) {
      const grid = strokeAttributes(document, style(chart.categoryGridCssClass), "--office-line", "#d9d9d9");
      for (let index = 0; index < category.tickCount; index++) { const position = category.tick(index); svg.append(svgElement(document, "line", { x1: position, y1: top, x2: position, y2: bottom, "data-sdk-chart-grid": "category", ...grid })); }
    }
    if (chart.showCategoryAxis !== false) {
      const skip = Math.max(1, numeric(catStyle, "--office-axis-label-skip", 1)), tickSkip = Math.max(1, numeric(catStyle, "--office-axis-tick-skip", 1));
      const offset = 6 + 10 * numeric(catStyle, "--office-axis-label-offset", 100) / 100;
      const alignment = property(catStyle, "--office-axis-label-align", "ctr"), anchor = alignment === "l" ? "start" : alignment === "r" ? "end" : "middle";
      data.categories.forEach((category, index) => { if (index % skip === 0 && property(catStyle, "--office-axis-label-position") !== "none") text(options.formatCategory?.(index, catStyle) ?? category, x(index), top + axes!.categoryLabel * h + (categoryTop ? -offset : offset), catStyle, anchor).dataset.sdkChartLabel = "category-axis"; });
      for (let index = 0; index < category.tickCount; index++) {
        const position = category.tick(index);
        if (index % tickSkip === 0) tick(position, axisY, true, catStyle);
        if (index + 1 < category.tickCount) tick((position + category.tick(index + 1)) / 2, axisY, true, catStyle, true);
      }
    }
    const defs = svgElement(document, "defs"), clip = svgElement(document, "clipPath", { id: serial + "-plot-clip" });
    clip.append(svgElement(document, "rect", { x: left, y: top, width: w, height: h })); defs.append(clip); svg.append(defs);
    plotLayer = svgElement(document, "g", { "clip-path": `url(#${serial}-plot-clip)`, "data-sdk-chart-plot": "" }); svg.append(plotLayer);
    data.series.forEach((series, seriesIndex) => {
      const s = seriesStyle.get(series.seriesId)!, segments: string[] = [];
      const lineStroke = strokeAttributes(document, s, "--office-line", chart.kind === "line" ? palette[seriesIndexes.get(series.seriesId)! % palette.length]! : "none"), lineOutset = chart.kind === "line" ? chartStrokeOutset(lineStroke) + 1 : 0;
      const overlap = (chart.overlap ?? 0) / 100, slots = data.series.length - (data.series.length - 1) * overlap, bar = group / (slots + (chart.gapWidth ?? 150) / 100), total = slots * bar;
      const beforeSeries = plotLayer.lastChild;
      let previous: { x: number; value: number; y: number | undefined } | undefined, pen: { x: number; y: number } | undefined;
      series.values.forEach((raw, index) => {
        if (raw === null && series.missing?.[index]) { if (series.missing[index] === "gap") { previous = undefined; pen = undefined; } return; }
        const value = raw ?? (chart.blankAs === "zero" ? 0 : null);
        if (value === null) { if (chart.blankAs !== "span") { previous = undefined; pen = undefined; } return; }
        const px = x(index), fraction = chartValueFraction(value, minimum, maximum), position = top + (chart.reverseValues ? fraction : 1 - fraction) * h;
        const py = Number.isFinite(position) && position >= top - MAX_CHART_PAINT_EXTENT && position <= bottom + MAX_CHART_PAINT_EXTENT ? position : undefined, point = pointStyle(series.seriesId, index);
        if (chart.kind === "column") {
          const stroke = strokeAttributes(document, point), outset = chartStrokeOutset(stroke) + 1;
          if (py === undefined && property(point, "--office-fill-kind") === "linear" && !(value < 0 && options.negativeColumns === "white" && property(point, "--office-invert-negative", "false") === "true")) {
            const angle = numeric(point, "--office-fill-angle", 0), fixed = property(point, "--office-fill-rotate-with-shape", "true") === "false";
            if (angle % 180 !== 0 || fixed && (options.orientation?.rotationDegrees ?? 0) % 180 !== 0) throw new DocumentError("ResourceLimitExceeded", "A chart gradient exceeds its finite original fill-coordinate budget.");
          }
          const outsideBelow = (fraction < 0) !== Boolean(chart.reverseValues), bounded = py ?? (outsideBelow ? bottom + outset : top - outset);
          const bounds = { x: px - total / 2 + seriesIndex * bar * (1 - overlap), y: Math.min(bounded, baseline), width: bar, height: Math.abs(bounded - baseline) };
          plotLayer.append(svgElement(document, "rect", { ...bounds, "data-sdk-chart-point": "column", fill: fill(series.seriesId, index, point, bounds, value < 0), ...stroke }));
        }
        else {
          if (previous) {
            const segment = previous.y === undefined || py === undefined ? clipChartValueSegment(previous, { x: px, value }, minimum, maximum, -lineOutset / h, 1 + lineOutset / h) : undefined;
            const first = segment ? { x: segment[0].x, y: top + (chart.reverseValues ? segment[0].fraction : 1 - segment[0].fraction) * h } : previous.y === undefined || py === undefined ? undefined : { x: previous.x, y: previous.y };
            const last = segment ? { x: segment[1].x, y: top + (chart.reverseValues ? segment[1].fraction : 1 - segment[1].fraction) * h } : previous.y === undefined || py === undefined ? undefined : { x: px, y: py };
            if (first && last) { if (!pen || pen.x !== first.x || pen.y !== first.y) segments.push(`M${first.x} ${first.y}`); segments.push(`L${last.x} ${last.y}`); pen = last; } else pen = undefined;
          } else if (py !== undefined) { segments.push(`M${px} ${py}`); pen = { x: px, y: py }; }
          previous = { x: px, value, y: py };
          marker(series.seriesId, index, px, Number.isFinite(position) ? position : undefined, "point");
        }
        if (py !== undefined || options.showLabelsOverMaximum === true) label(series, index, value, chart.kind === "column" ? px - total / 2 + seriesIndex * bar * (1 - overlap) + bar / 2 : px, py ?? y(Math.max(minimum, Math.min(maximum, value))), undefined, baseline);
      });
      if (chart.kind === "line") plotLayer.insertBefore(svgElement(document, "path", { d: segments.join(" "), "data-sdk-chart-line": "", fill: "none", ...lineStroke }), beforeSeries ? beforeSeries.nextSibling : plotLayer.firstChild);
    });
  }
  if (chart.legend !== "none") {
    const horizontal = chart.legend === "top" || chart.legend === "bottom", slot = 18;
    const context = document.createElement("canvas").getContext("2d")!; context.font = legendStyle.font || legendStyle.fontSize + " " + legendStyle.fontFamily;
    const keys = entries.map(entry => {
      if (chart.kind !== "line") return 7;
      const series = seriesDefinitions.get(entry.series.seriesId), point = entry.index < 0 ? undefined : pointDefinitions.get(pointKey(entry.series.seriesId, sourceIndex(entry.index)));
      return Math.max(22, (point?.marker ?? series?.marker ?? "none") === "none" ? 0 : point?.markerSize ?? series?.markerSize ?? 5);
    });
    const widths = entries.map((entry, index) => context.measureText(entry.name).width * 0.75 + keys[index]! + 8), total = widths.reduce((sum, n) => sum + n, 0);
    const legend = options.layout?.legend;
    let cursor = legend ? legend.x + (legend.width - total) / 2 : left + (w - total) / 2;
    entries.forEach((entry, index) => {
      const previous = svg.lastChild;
      const keyWidth = keys[index]!;
      const x = horizontal ? cursor : legend ? legend.x + 8 : chart.legend === "left" ? 12 : right + 12; cursor += widths[index]!;
      const y = horizontal ? legend ? legend.y + legend.height / 2 : chart.legend === "top" ? top - 14 : height - 14
        : (legend ? legend.y + legend.height / 2 : top + h / 2) + (index - (entries.length - 1) / 2) * slot;
      const s = entry.index >= 0 ? pointStyle(entry.series.seriesId, entry.index) : seriesStyle.get(entry.series.seriesId)!;
      if (chart.kind === "line") { svg.append(svgElement(document, "line", { x1: x, y1: y, x2: x + keyWidth, y2: y, ...strokeAttributes(document, s, "--office-line", palette[(varyColors && entry.index >= 0 ? sourceIndex(entry.index) : seriesIndexes.get(entry.series.seriesId)!) % palette.length]!) })); marker(entry.series.seriesId, entry.index, x + keyWidth / 2, y, "legend"); }
      else svg.append(svgElement(document, "rect", { x, y: y - 3, width: keyWidth, height: 7, fill: fill(entry.series.seriesId, Math.max(0, entry.index), s, { x, y: y - 3, width: keyWidth, height: 7 }) }));
      text(entry.name, x + keyWidth + 3, y + 1, legendStyle, "start");
      const group = svgElement(document, "g", { "data-sdk-chart-legend-entry": "" });
      let child = previous?.nextSibling ?? svg.firstChild;
      while (child) { const next = child.nextSibling; group.append(child); child = next; }
      svg.append(group);
    });
  }
  return scene;
}
