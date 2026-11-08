import type { ChartAppearance, DeepReadonly, PresentationChart, PresentationObject } from "@doffi/core";
import type { Css } from "../export/css.js";
import type { SourceElement } from "../export/source.js";
import { attributes, declaration, empty, escapeXml } from "../export/xml.js";
import { effectXml, fillXml, lineXml } from "./appearance.js";
import { relativePart } from "./context.js";
import type { NativeSurface, PresentationExportContext } from "./context.js";
import { PresentationTextWriter, runProperties } from "./text.js";
import { cssString, integer, unsupported } from "./values.js";
import { exportConstantWorkbook } from "./workbook.js";
import type { ExportedConstantWorkbook } from "./workbook.js";
const cv = (name: string, value: string | number | boolean | undefined): string => value === undefined ? "" : empty("c:" + name, { val: typeof value === "boolean" ? value ? 1 : 0 : value });
function crossingXml(css: Css, category: boolean): string {
    const raw = css["--office-axis-crosses-at"];
    if (raw !== undefined) {
        const at = raw.trim(), number = Number(at);
        if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(at) || !Number.isFinite(number) || !category && !Number.isInteger(number)) unsupported("Chart crossing values exceed the selected Office axis profile.");
        return cv("crossesAt", number);
    }
    const kind = css["--office-axis-crosses"] ?? "autoZero";
    if (!["autoZero", "min", "max"].includes(kind)) unsupported("A chart axis crossing enum is invalid.");
    return cv("crosses", kind);
}
export class PresentationChartWriter {
    private readonly books = new Map<string, {
        part: string;
        book: ExportedConstantWorkbook;
    }>();
    private count = 0;
    constructor(private readonly ctx: PresentationExportContext, private readonly text: PresentationTextWriter) { }
    private style(name: string | readonly string[] | undefined, node: SourceElement, surface: NativeSurface): {
        css: Css;
        raw: Css;
    } { return this.ctx.styles.probeClass(name, node, surface); }
    private appearance(style: {
        css: Css;
        raw: Css;
    }, fallback?: string): string {
        const css = fallback && !style.css["--office-fill-kind"] ? { ...style.css, "--office-fill-kind": "solid", "--office-fill-color": fallback } : style.css;
        return "<c:spPr>" + (css["--office-fill-kind"] || css["--office-fill-color"] ? fillXml(css, "--office-fill", style.raw) : "") + (css["--office-line-kind"] || css["--office-line-color"] ? lineXml(css, "--office-line", style.raw) : "") + (css["--office-shadow-kind"] ? effectXml(css, style.raw) : "") + "</c:spPr>";
    }
    private marker(definition: { marker?: NonNullable<ChartAppearance["seriesStyles"]>[number]["marker"]; markerSize?: number | undefined } | undefined,
        style: { css: Css; raw: Css }, series = false): string {
        const remap = (css: Css): Css => Object.fromEntries(Object.entries(css).filter(([key]) => key.startsWith("--office-marker-")).map(([key, value]) => [key.replace("--office-marker-", "--office-"), value]));
        const css = remap(style.css), raw = remap(style.raw);
        if (!series && definition?.marker === undefined && definition?.markerSize === undefined && !Object.keys(css).length) return "";
        return "<c:marker>" + cv("symbol", definition?.marker ?? (series ? "none" : undefined))
            + cv("size", definition?.markerSize === undefined ? undefined : integer(definition.markerSize, 1, 2, 72))
            + (Object.keys(css).length ? this.appearance({ css, raw }) : "") + "</c:marker>";
    }
    private textStyle(style: {
        css: Css;
        raw: Css;
    }): string { const own = Object.fromEntries(Object.keys(style.raw).filter(key => ["font-size", "font-family", "font-weight", "font-style", "color", "text-align", "text-decoration-line", "text-decoration-style", "text-decoration-color", "--office-underline", "--office-strike"].includes(key) || key.startsWith("--office-font-") || key.startsWith("--office-text-")).map(key => [key, style.raw[key]!])); return '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr' + runProperties(style.css, own, style.raw) + "</a:defRPr></a:pPr><a:endParaRPr/></a:p></c:txPr>"; }
    private labels(chart: DeepReadonly<PresentationChart>, styleName: string | undefined, node: SourceElement, surface: NativeSurface, points: readonly {
        index: number;
        labelCssClass?: string | undefined;
    }[] = []): string {
        const contents = chart.showLabels ? (chart.labelContents ?? ["value"]) : [], names = [...new Set([chart.textCssClass, chart.labelCssClass, styleName].filter((name): name is string => !!name))];
        const properties = (name: string | undefined, styled: boolean): string => {
            const style = this.style([...names, ...(name && !names.includes(name) ? [name] : [])], node, surface), css = style.css;
            let result = "";
            if (css["--office-label-number-format"] !== undefined)
                result += empty("c:numFmt", { formatCode: cssString(css["--office-label-number-format"]!), sourceLinked: 0 });
            if (styled)
                result += this.appearance(style) + this.textStyle(style);
            if (css["--office-label-position"])
                result += cv("dLblPos", css["--office-label-position"]);
            result += cv("showLegendKey", false);
            for (const [native, kind] of [["showVal", "value"], ["showCatName", "category"], ["showSerName", "series"], ["showPercent", "percent"]] as const)
                result += cv(native, contents.includes(kind));
            result += cv("showBubbleSize", false);
            if (css["--office-label-separator"] !== undefined)
                result += "<c:separator>" + escapeXml(cssString(css["--office-label-separator"]!)) + "</c:separator>";
            return result;
        };
        let result = "<c:dLbls>";
        for (const point of points)
            if (point.labelCssClass)
                result += "<c:dLbl>" + cv("idx", point.index) + properties(point.labelCssClass, true) + "</c:dLbl>";
        return result + properties(undefined, Boolean(styleName)) + cv("showLeaderLines", false) + "</c:dLbls>";
    }
    write(object: DeepReadonly<Extract<PresentationObject, {
        kind: "chart";
    }>>, node: SourceElement, surface: NativeSurface): string {
        const ctx = this.ctx, chart = ctx.p.charts.find(chart => chart.chartId === object.chartId)!, data = ctx.p.dataSets.find(data => data.dataSetId === chart.dataSetId)!;
        let embedded = this.books.get(data.dataSetId);
        if (!embedded) {
            embedded = { part: "ppt/embeddings/data" + (this.books.size + 1) + ".xlsx", book: exportConstantWorkbook(data, ctx.limits, ctx.budget, ctx.xmlBudget, ctx.options) };
            this.books.set(data.dataSetId, embedded);
            ctx.opc.binary(embedded.part, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", embedded.book.bytes);
        }
        const part = "ppt/charts/chart" + (++this.count) + ".xml", out = ctx.opc.part(part, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"), relationship = ctx.opc.relate(surface.part, "chart", relativePart(surface.part, part)), bookRelation = ctx.opc.relate(part, "package", relativePart(part, embedded.part));
        const vector = (values: readonly (string | number | null)[], numeric: boolean, formula?: string): string => { const type = numeric ? "num" : "str", cache = "<c:" + type + "Cache>" + (numeric ? "<c:formatCode>" + escapeXml(chart.numberFormat ?? "General") + "</c:formatCode>" : "") + cv("ptCount", values.length) + values.flatMap((value, index) => value === null ? [] : ["<c:pt" + attributes({ idx: index }) + "><c:v>" + escapeXml(String(value)) + "</c:v></c:pt>"]).join("") + "</c:" + type + "Cache>"; return formula ? "<c:" + type + "Ref><c:f>" + escapeXml(formula) + "</c:f>" + cache + "</c:" + type + "Ref>" : cache.replaceAll(type + "Cache", type + "Lit"); };
        out.add(declaration + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>');
        if (chart.titleNodeId) {
            const title = ctx.tree.byId.get(chart.titleNodeId)!;
            out.add("<c:title><c:tx>" + this.text.body(title, surface, "c:rich", part) + "</c:tx><c:overlay val=\"0\"/>" + this.appearance({ css: ctx.styles.get(title, surface), raw: ctx.styles.raw(title, surface) }) + "</c:title>");
        }
        out.add(cv("autoTitleDeleted", chart.titleVisible === false || !chart.titleNodeId));
        const kind = chart.kind === "column" ? "barChart" : chart.kind === "line" ? "lineChart" : "pieChart";
        out.add("<c:plotArea><c:layout/><c:" + kind + ">");
        if (chart.kind === "column")
            out.add(cv("barDir", "col") + cv("grouping", "clustered"));
        else if (chart.kind === "line")
            out.add(cv("grouping", "standard"));
        out.add(cv("varyColors", ctx.styles.get(node, surface)["--office-chart-vary-colors"] === "true"));
        for (const [index, series] of data.series.entries()) {
            const style = chart.seriesStyles?.find(style => style.seriesId === series.seriesId), seriesStyle = this.style(style?.cssClass, node, surface), ranges = embedded.book.series[index]!, points = chart.pointStyles?.filter(point => point.seriesId === series.seriesId) ?? [];
            const fallback = "var(--office-theme-accent" + (index % 6 + 1) + ")", palette = ctx.styles.get(node, surface)["--office-theme-accent" + (index % 6 + 1)] ?? "#4472c4";
            const s = seriesStyle.css["--office-fill-kind"] ? seriesStyle : { css: { ...seriesStyle.css, "--office-fill-kind": "solid", "--office-fill-color": palette }, raw: { ...seriesStyle.raw, "--office-fill-color": fallback } };
            out.add("<c:ser>" + cv("idx", index) + cv("order", index) + "<c:tx>" + vector([series.name], false, ranges.nameFormula) + "</c:tx>" + this.appearance(s));
            if (chart.kind === "line")
                out.add(this.marker(style, seriesStyle, true));
            else if (chart.kind === "column")
                out.add(cv("invertIfNegative", seriesStyle.css["--office-invert-negative"] === "true"));
            for (const point of points)
                out.add("<c:dPt>" + cv("idx", point.index) + this.marker(point, this.style(point.cssClass, node, surface)) + this.appearance(this.style([style?.cssClass, point.cssClass].filter(Boolean).join(" "), node, surface)) + "</c:dPt>");
            if (style?.labelCssClass || points.some(point => point.labelCssClass))
                out.add(this.labels(chart, style?.labelCssClass ?? chart.labelCssClass, node, surface, points));
            out.add("<c:cat>" + vector(data.categories, false, embedded.book.categoryFormula) + "</c:cat><c:val>" + vector(series.values, true, ranges.valueFormula) + "</c:val>" + (chart.kind === "line" ? cv("smooth", false) : "") + "</c:ser>");
        }
        out.add(this.labels(chart, chart.labelCssClass, node, surface));
        if (chart.kind === "line" && chart.nativeShowMarkers !== undefined)
            out.add(cv("marker", chart.nativeShowMarkers));
        if (chart.kind === "column")
            out.add(cv("gapWidth", chart.gapWidth) + cv("overlap", chart.overlap));
        if (chart.kind === "pie")
            out.add(cv("firstSliceAng", chart.firstSliceAngle));
        else
            out.add(cv("axId", 1) + cv("axId", 2));
        out.add("</c:" + kind + ">");
        if (chart.kind !== "pie")
            for (const category of [true, false]) {
                const name = category ? "catAx" : "valAx", style = this.style(category ? chart.categoryAxisCssClass : chart.valueAxisCssClass, node, surface), css = style.css;
                out.add("<c:" + name + ">" + cv("axId", category ? 1 : 2) + "<c:scaling>" + cv("orientation", (category ? chart.reverseCategories : chart.reverseValues) ? "maxMin" : "minMax") + (category ? "" : cv("max", chart.maximum) + cv("min", chart.minimum)) + "</c:scaling>" + cv("delete", (category ? chart.showCategoryAxis : chart.showValueAxis) === false) + cv("axPos", css["--office-axis-position"] ?? (category ? "b" : "l")));
                const gridClass = category ? chart.categoryGridCssClass : chart.gridCssClass;
                if (gridClass) out.add("<c:majorGridlines>" + this.appearance(this.style(gridClass, node, surface)) + "</c:majorGridlines>");
                if (!category)
                    out.add(empty("c:numFmt", { formatCode: chart.numberFormat ?? "General", sourceLinked: 0 }));
                for (const [native, key, fallback] of [["majorTickMark", "major-tick", "cross"], ["minorTickMark", "minor-tick", "cross"], ["tickLblPos", "label-position", "nextTo"]])
                    out.add(cv(native!, css["--office-axis-" + key] ?? fallback));
                out.add(this.appearance(style) + this.textStyle(style) + cv("crossAx", category ? 2 : 1) + crossingXml(css, category));
                if (category) {
                    out.add(cv("auto", true) + cv("lblAlgn", css["--office-axis-label-align"] ?? "ctr") + cv("lblOffset", css["--office-axis-label-offset"] ?? 100));
                    for (const [native, key] of [["tickLblSkip", "label-skip"], ["tickMarkSkip", "tick-skip"]])
                        out.add(cv(native!, css["--office-axis-" + key]));
                }
                else {
                    out.add(cv("crossBetween", css["--office-axis-cross-between"] ?? "between") + cv("majorUnit", chart.majorUnit));
                    const minor = css["--office-axis-minor-unit"];
                    if (minor && minor !== "auto")
                        out.add(cv("minorUnit", minor));
                }
                out.add("</c:" + name + ">");
            }
        out.add(this.appearance(this.style(chart.plotCssClass, node, surface)) + "</c:plotArea>");
        if (chart.legend !== "none")
            out.add("<c:legend>" + cv("legendPos", ({ left: "l", right: "r", top: "t", bottom: "b" }[chart.legend])) + '<c:layout/><c:overlay val="0"/>' + this.appearance(this.style([chart.textCssClass, chart.legendCssClass].filter(Boolean).join(" "), node, surface)) + this.textStyle(this.style([chart.textCssClass, chart.legendCssClass].filter(Boolean).join(" "), node, surface)) + "</c:legend>");
        out.add(cv("plotVisOnly", true) + cv("dispBlanksAs", chart.blankAs ?? "gap") + "</c:chart>" + this.appearance({ css: ctx.styles.get(node, surface), raw: ctx.styles.raw(node, surface) }) + this.textStyle(this.style(chart.textCssClass, node, surface)) + "<c:externalData" + attributes({ "r:id": bookRelation }) + '><c:autoUpdate val="0"/></c:externalData></c:chartSpace>');
        out.finish();
        return empty("c:chart", { "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart", "r:id": relationship });
    }
}
