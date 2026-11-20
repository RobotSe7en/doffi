import type { PresentationOperation } from "../contracts/presentation-operations.js";
import type { PresentationChart } from "../contracts/presentation.js";
import type { PresentationEditContext } from "./presentation-edit.js";
import { checkLimit, invalid } from "../document/validation.js";
import { geometryAttributes } from "./presentation-geometry.js";
import { allocatePresentationClass } from "./presentation-styles.js";
const escape = (value: string): string => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
export function chartOperation(operation: PresentationOperation): boolean { return ["insertPresentationChart", "setPresentationChart", "setPresentationChartData", "setPresentationChartStyle"].includes(operation.op); }
export function editPresentationChart(ctx: PresentationEditContext, operation: PresentationOperation): void {
    const { p, edit } = ctx;
    if (operation.op === "insertPresentationChart") {
        const html = ctx.parse(), parent = html.peekNode(operation.parentNodeId), kind = parent?.attributes["data-office-kind"];
        if (!["slide", "master", "layout", "notes-master", "notes", "group"].includes(kind ?? ""))
            invalid("Chart insertion requires a presentation surface or group.");
        if (operation.data.series.some(s => s.seriesId !== undefined))
            invalid("New chart series receive core-allocated identities.");
        const nodeId = edit.insertSource(operation.parentNodeId, "<div data-office-kind=\"chart\"></div>").rootNodeIds[0]!, dataSetId = ctx.allocate(), chartId = ctx.allocate();
        edit.attributes(nodeId, geometryAttributes(operation.geometry));
        checkLimit(operation.data.categories.length * operation.data.series.length, ctx.limits.maxNodes);
        p.dataSets.push({ dataSetId, categories: [...operation.data.categories], series: operation.data.series.map(s => ({ seriesId: ctx.allocate(), name: s.name, values: [...s.values] })) });
        const chart: PresentationChart = { chartId, nodeId, dataSetId, kind: operation.kind, legend: "bottom", showLabels: false };
        p.charts.push(chart);
        p.objects.push({ kind: "chart", nodeId, chartId });
        if (operation.title !== undefined)
            chart.titleNodeId = edit.insertSource(nodeId, "<p>" + escape(operation.title) + "</p>").rootNodeIds[0]!;
        return;
    }
    if (!("nodeId" in operation))
        invalid("Chart editing requires a chart object.");
    const chart = p.charts.find(c => c.nodeId === operation.nodeId) ?? invalid("The chart is missing."), data = p.dataSets.find(d => d.dataSetId === chart.dataSetId)!;
    if (operation.op === "setPresentationChart") {
        const format = operation.changes.numberFormat;
        if (format !== undefined && format !== "General" && format !== "" && (format.split(";").length > 4 || format.split(";").some(section => section !== "" && !/^\(?[#,0]+(?:\.[0#]{1,20})?%?\)?$/.test(section))))
            invalid("The chart number format is outside the supported basic numeric subset.");
        Object.assign(chart, operation.changes);
        for (const key of operation.clear ?? [])
            delete chart[key];
        if (operation.changes.showLabels === false && operation.changes.labelContents === undefined)
            delete chart.labelContents;
        if (operation.changes.labelContents !== undefined && operation.changes.showLabels === undefined)
            chart.showLabels = !!operation.changes.labelContents.length;
        if (operation.title === null) {
            if (chart.titleNodeId)
                edit.removeSource(chart.titleNodeId);
            delete chart.titleNodeId;
        }
        else if (operation.title !== undefined) {
            if (chart.titleNodeId) {
                const source = ctx.parse();
                if (source.getNode(chart.titleNodeId)!.text !== operation.title) {
                    edit.removeSource(chart.titleNodeId, true);
                    edit.insertSource(chart.titleNodeId, escape(operation.title));
                }
            }
            else
                chart.titleNodeId = edit.insertSource(chart.nodeId, "<p>" + escape(operation.title) + "</p>").rootNodeIds[0]!;
        }
        return;
    }
    if (operation.op === "setPresentationChartData") {
        const base = ctx.base.snapshot.metadata.presentation!.dataSets.find(d => d.dataSetId === data.dataSetId)!, ids = new Set(base.series.map(s => s.seriesId)), provided = operation.data.series.flatMap(s => s.seriesId ? [s.seriesId] : []);
        if (new Set(provided).size !== provided.length || provided.some(id => !ids.has(id)))
            invalid("Retained chart series must be distinct identities in the base data set.");
        checkLimit(operation.data.categories.length * operation.data.series.length, ctx.limits.maxNodes);
        data.categories = [...operation.data.categories];
        data.series = operation.data.series.map(s => ({ seriesId: s.seriesId ?? ctx.allocate(), name: s.name, values: [...s.values] }));
        const nextIds = new Set(data.series.map(s => s.seriesId));
        for (const consumer of p.charts.filter(c => c.dataSetId === data.dataSetId)) {
            if (consumer.seriesStyles)
                consumer.seriesStyles = consumer.seriesStyles.filter(s => nextIds.has(s.seriesId));
            if (consumer.pointStyles)
                consumer.pointStyles = consumer.pointStyles.filter(s => nextIds.has(s.seriesId) && s.index < data.categories.length);
        }
        return;
    }
    if (operation.op === "setPresentationChartStyle") {
        if (!Object.keys(operation.properties).length)
            return;
        const series = operation.target === "series" || operation.target === "point";
        if (series && !data.series.some(s => s.seriesId === operation.seriesId) || !series && (operation.seriesId !== undefined || operation.index !== undefined) || operation.target === "point" && (operation.index === undefined || operation.index >= data.categories.length) || operation.target !== "point" && operation.index !== undefined)
            invalid("Chart styles must target an existing series or point.");
        for (const property of Object.keys(operation.properties))
            if (!/^(?:--office-(?:fill|line|text|font)-[a-z0-9-]+|color|font-(?:size|family|weight|style)|text-decoration)$/.test(property))
                invalid("The chart appearance property is outside the supported style subset.");
        let current: string | undefined;
        const key = operation.target === "series" || operation.target === "point" ? undefined : (operation.target + "CssClass") as "plotCssClass" | "legendCssClass" | "textCssClass" | "labelCssClass" | "categoryAxisCssClass" | "valueAxisCssClass" | "gridCssClass";
        const entry = operation.target === "series" ? chart.seriesStyles?.find(s => s.seriesId === operation.seriesId) : operation.target === "point" ? chart.pointStyles?.find(s => s.seriesId === operation.seriesId && s.index === operation.index) : undefined;
        current = key ? chart[key] : entry?.cssClass;
        const html = ctx.parse();
        const reused = current && (!current.startsWith("chart-style-") || p.charts.some(c => c !== chart && JSON.stringify(c).includes('"' + current + '"')) || html.nodeIds().some(id => (html.peekNode(id)!.attributes.class ?? "").split(/\s+/).includes(current!)));
        const cssClass = !current || reused ? allocatePresentationClass(ctx, "chart-style-") : current;
        if (current && cssClass !== current)
            ctx.styles.cloneClass(current, cssClass);
        const properties = { ...operation.properties };
        for (const prefix of ["--office-fill", "--office-line", "--office-text"])
            if (properties[prefix + "-color"] !== undefined)
                for (const [name, value] of Object.entries({ source: "direct", "color-transforms": "none", alpha: "1", "lum-mod": "1", "lum-off": "0", tint: "0", shade: "1" }))
                    if (properties[prefix + "-" + name] === undefined)
                        properties[prefix + "-" + name] = value;
        ctx.styles.setClass(cssClass, properties);
        if (key)
            chart[key] = cssClass;
        else if (entry)
            entry.cssClass = cssClass;
        else if (operation.target === "series")
            (chart.seriesStyles ??= []).push({ seriesId: operation.seriesId!, cssClass });
        else
            (chart.pointStyles ??= []).push({ seriesId: operation.seriesId!, index: operation.index!, cssClass });
        return;
    }
    invalid("Unsupported chart operation.");
}
