import type { WorkbookChart, WorkbookDrawingAnchor, WorkbookMetadata, WorkbookObject } from "../contracts/workbook.js";
import { WorkbookChartSchema, WorkbookChartVectorSchema } from "../contracts/workbook-charts.js";
import { WorkbookDrawingOperationSchema } from "../contracts/workbook-drawing-operations.js";
import type { WorkbookChartDataInput, WorkbookDrawingOperation } from "../contracts/workbook-drawing-operations.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { DeepReadonly, WorkbookModel } from "../runtime.js";
import { invalid } from "../document/validation.js";
import { WorkbookContentEdit } from "./workbook-content.js";
import { escapeWorkbookText } from "./workbook-cells.js";

const copy = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T;
const names = new Set<string>(WorkbookDrawingOperationSchema.options.map(schema => schema.shape.op.value));
export function isWorkbookDrawingOperation(operation: WorkbookOperation): operation is WorkbookDrawingOperation { return names.has(operation.op); }
export function validateWorkbookDrawingTarget(base: WorkbookModel, operation: WorkbookDrawingOperation): void {
  const objects = new Set(base.snapshot().metadata.workbook.objects?.map(object => object.nodeId));
  const ids = "nodeIds" in operation ? operation.nodeIds : "nodeId" in operation ? [operation.nodeId] : [];
  if (new Set(ids).size !== ids.length || ids.some(id => !objects.has(id))) invalid("Drawing targets must be distinct object identities in the base snapshot.");
}
function withoutCaches<T>(input: T): T {
  const value = copy(input);
  function visit(input: unknown): void {
    if (!input || typeof input !== "object") return;
    if (Array.isArray(input)) { input.forEach(visit); return; }
    const item = input as Record<string, unknown>;
    if ("expression" in item) delete item.cache;
    Object.values(item).forEach(visit);
  }
  visit(value); return value;
}
function data(input: WorkbookChartDataInput, allocate: () => string, original?: DeepReadonly<WorkbookChart>): Pick<WorkbookChart, "categories" | "series"> {
  const previous = original && WorkbookChartSchema.parse(original);
  const allowed = new Set(original?.series.map(series => series.seriesId)), retained = input.series.flatMap(series => series.seriesId ? [series.seriesId] : []);
  if (new Set(retained).size !== retained.length || retained.some(id => !allowed.has(id))) invalid("Retained series must be distinct identities in the base chart.");
  const preserve = (value: WorkbookChartDataInput["series"][number]["values"], old: DeepReadonly<WorkbookChartDataInput["series"][number]["values"]> | undefined) =>
    "expression" in value && old && "expression" in old && value.expression === old.expression && value.valueType === old.valueType
      ? WorkbookChartVectorSchema.parse(old) : withoutCaches(value);
  return { ...(input.categories ? { categories: preserve(input.categories, original?.categories) } : {}), series: input.series.map(series => {
    const old = previous?.series.find(item => item.seriesId === series.seriesId);
    const name = series.name?.kind === "reference" && old?.name?.kind === "reference" && series.name.reference.expression === old.name.reference.expression && series.name.reference.valueType === old.name.reference.valueType
      ? { kind: "reference" as const, reference: WorkbookChartVectorSchema.options[0].parse(old.name.reference) } : series.name && withoutCaches(series.name);
    const item = old ? { ...old } : { seriesId: series.seriesId ?? allocate(), values: withoutCaches(series.values) };
    item.values = preserve(series.values, old?.values); if (name) item.name = name; else delete item.name; return item;
  }) };
}
function offset(anchor: WorkbookDrawingAnchor, x: number, y: number): WorkbookDrawingAnchor {
  if (anchor.kind === "absolute") return { ...anchor, xPt: anchor.xPt + x, yPt: anchor.yPt + y };
  const shift = (marker: typeof anchor.from) => ({ ...marker, offsetXPt: marker.offsetXPt + x, offsetYPt: marker.offsetYPt + y });
  return { ...anchor, from: shift(anchor.from), ...(anchor.kind === "two-cell" ? { to: shift(anchor.to) } : {}) };
}
export function editWorkbookDrawing(workbook: WorkbookMetadata, content: WorkbookContentEdit, base: WorkbookModel, operation: WorkbookDrawingOperation, allocate: () => string): void {
  const edit = () => content.edit(), parse = () => content.parse(), object = (id: string) => workbook.objects?.find(object => object.nodeId === id) ?? invalid("An earlier operation removed this drawing.");
  const chartFor = (id: string) => { if (object(id).kind !== "chart") invalid("Chart editing requires a chart drawing."); return workbook.charts?.find(chart => chart.nodeId === id) ?? invalid("The chart definition is missing."); };
  function appearance(nodeId: string, properties: Readonly<Record<string, string>>): void {
    for (const [property, value] of Object.entries(properties)) {
      if (!/^(?:color|background-color|font-(?:family|size|weight|style)|text-(?:align|decoration)|line-height|letter-spacing|padding(?:-(?:top|right|bottom|left))?|--office-(?:fill|line|text|font|effect)-[a-z0-9-]+)$/.test(property)) invalid("The drawing appearance property is outside its source subset.");
      edit().sourceDrawingStyle(nodeId, property, value);
    }
  }
  const append = (sheetId: string, anchor: WorkbookDrawingAnchor, kind: WorkbookObject["kind"], body: string, name?: string) => {
    if (!workbook.sheets.some(sheet => sheet.sheetId === sheetId)) invalid("The drawing worksheet no longer exists.");
    const nodeId = edit().insertSource(content.root(), '<div data-office-kind="' + kind + '">' + body + '</div>').rootNodeIds[0]!;
    return { nodeId, sheetId, anchor: copy(anchor), locked: true, printable: true, ...(name !== undefined ? { name } : {}) };
  };
  if (operation.op === "insertWorkbookObject") {
    const input = operation.object, body = input.kind === "image" ? '<img src="' + escapeWorkbookText(content.relative(input.assetPath)) + '" alt="' + escapeWorkbookText(input.alt) + '">' : input.html ?? "<p></p>";
    const common = append(operation.sheetId, operation.anchor, input.kind, body, operation.name);
    const added: WorkbookObject = input.kind === "image" ? { ...common, kind: "image" } : input.kind === "text" ? { ...common, kind: "text", ...(input.text ? { text: copy(input.text) } : {}) }
      : { ...common, kind: "geometry", preset: input.preset, ...(input.text ? { text: copy(input.text) } : {}) };
    (workbook.objects ??= []).push(added);
    if (operation.properties) appearance(common.nodeId, operation.properties); return;
  }
  if (operation.op === "insertWorkbookChart") {
    const values = data(operation.data, allocate), common = append(operation.sheetId, operation.anchor, "chart", "", operation.name), { nodeId } = common, chartId = allocate();
    (workbook.objects ??= []).push({ ...common, kind: "chart", chartId });
    const chart: WorkbookChart = { chartId, nodeId, kind: operation.kind, legend: "bottom", showLabels: false, ...values };
    (workbook.charts ??= []).push(chart);
    if (operation.title !== undefined) chart.titleNodeId = edit().insertSource(nodeId, "<div>" + operation.title + "</div>").rootNodeIds[0]!;
    return;
  }
  if (operation.op === "duplicateWorkbookObjects") {
    const selected = operation.nodeIds.map(object);
    for (const original of selected) {
      const cloned = edit().cloneSource(original.nodeId, content.root()), nodeId = cloned.rootNodeIds[0]!, value = copy(original);
      value.nodeId = nodeId; value.anchor = offset(value.anchor, operation.offset.xPt, operation.offset.yPt); delete value.nativeId;
      if (value.kind === "chart") {
        const chart = copy(chartFor(original.nodeId)), series = new Map(chart.series.map(item => [item.seriesId, allocate()]));
        chart.nodeId = nodeId; chart.chartId = allocate(); value.chartId = chart.chartId;
        for (const item of chart.series) { item.seriesId = series.get(item.seriesId)!; delete item.nativeIndex; }
        for (const item of [...chart.seriesStyles ?? [], ...chart.pointStyles ?? []]) item.seriesId = series.get(item.seriesId)!;
        if (chart.titleNodeId) chart.titleNodeId = cloned.identities.get(chart.titleNodeId)!;
        for (const title of [chart.categoryAxisTitle, chart.valueAxisTitle]) if (title?.nodeId) title.nodeId = cloned.identities.get(title.nodeId)!;
        (workbook.charts ??= []).push(chart);
      }
      workbook.objects!.push(value);
      for (const link of [...workbook.objectHyperlinks ?? []]) if (cloned.identities.has(link.nodeId)) workbook.objectHyperlinks!.push({ ...copy(link), nodeId: cloned.identities.get(link.nodeId)! });
    }
    return;
  }
  if (operation.op === "reorderWorkbookObjects") {
    const selected = operation.nodeIds.map(object), sheets = new Set(selected.map(value => value.sheetId)); if (sheets.size !== 1) invalid("A drawing order selection must belong to one worksheet.");
    const all = workbook.objects!, chosen = new Set(operation.nodeIds), local = all.filter(value => sheets.has(value.sheetId)), ordered = local.filter(value => chosen.has(value.nodeId));
    let result = [...local];
    if (operation.position === "front" || operation.position === "back") result = operation.position === "front" ? [...local.filter(value => !chosen.has(value.nodeId)), ...ordered] : [...ordered, ...local.filter(value => !chosen.has(value.nodeId))];
    else if (operation.position === "forward") for (let index = result.length - 2; index >= 0; index--) { if (chosen.has(result[index]!.nodeId) && !chosen.has(result[index + 1]!.nodeId)) [result[index], result[index + 1]] = [result[index + 1]!, result[index]!]; }
    else for (let index = 1; index < result.length; index++) if (chosen.has(result[index]!.nodeId) && !chosen.has(result[index - 1]!.nodeId)) [result[index], result[index - 1]] = [result[index - 1]!, result[index]!];
    let index = 0; workbook.objects = all.map(value => sheets.has(value.sheetId) ? result[index++]! : value); return;
  }
  const current = object(operation.nodeId), source = parse();
  if (operation.op === "removeWorkbookObject") {
    const descendants = new Set<string>(), pending = [current.nodeId];
    while (pending.length) { const id = pending.pop()!; descendants.add(id); pending.push(...source.peekNode(id)!.childNodeIds); }
    edit().removeSource(current.nodeId); workbook.objects = workbook.objects!.filter(value => value !== current);
    if (workbook.charts) workbook.charts = workbook.charts.filter(chart => chart.nodeId !== current.nodeId);
    if (workbook.objectHyperlinks) workbook.objectHyperlinks = workbook.objectHyperlinks.filter(link => !descendants.has(link.nodeId)); return;
  }
  if (operation.op === "updateWorkbookObject") {
    const { rotation, flipH, flipV, hidden, preset, ...changes } = operation.changes;
    if (preset !== undefined) { if (current.kind !== "geometry") invalid("Only a shape can change its preset."); current.preset = preset; }
    for (const [key, value] of Object.entries(changes)) { if (value === null) Reflect.deleteProperty(current, key); else if (value !== undefined) Reflect.set(current, key, copy(value)); }
    for (const [key, value] of [["rotation", rotation], ["flip-h", flipH], ["flip-v", flipV], ["hidden", hidden]] as const) if (value !== undefined) edit().attributes(current.nodeId, { ["data-office-" + key]: String(value) }); return;
  }
  if (operation.op === "setWorkbookObjectStyle") { appearance(current.nodeId, operation.properties); return; }
  if (operation.op === "setWorkbookObjectText") {
    if (current.kind !== "text" && current.kind !== "geometry") invalid("Text editing requires a text box or shape.");
    edit().replaceSourceContent(current.nodeId, operation.html);
    if (operation.text === null) delete current.text; else if (operation.text !== undefined) current.text = copy(operation.text); return;
  }
  if (operation.op === "replaceWorkbookImage") {
    if (current.kind !== "image") invalid("Image replacement requires an image drawing.");
    edit().attributes(source.peekNode(current.nodeId)!.childNodeIds[0]!, { src: content.relative(operation.assetPath), alt: operation.alt }); return;
  }
  if (operation.op === "setWorkbookImageCrop") {
    if (current.kind !== "image" || operation.crop.left + operation.crop.right >= 1 || operation.crop.top + operation.crop.bottom >= 1) invalid("Image cropping must retain a visible area.");
    for (const [side, value] of Object.entries(operation.crop)) edit().attributes(current.nodeId, { ["data-office-crop-" + side]: String(value) }); return;
  }
  const chart = chartFor(current.nodeId);
  if (operation.op === "setWorkbookChart") {
    Object.assign(chart, copy(operation.changes)); for (const key of operation.clear ?? []) delete chart[key];
    if (operation.changes.showLabels === false && operation.changes.labelContents === undefined) delete chart.labelContents;
    if (operation.changes.showLabels === true && operation.changes.labelContents === undefined && chart.labelContents?.length === 0) chart.labelContents = ["value"];
    if (operation.changes.labelContents !== undefined && operation.changes.showLabels === undefined) chart.showLabels = !!operation.changes.labelContents.length;
    return;
  }
  if (operation.op === "setWorkbookChartData") {
    const original = base.snapshot().metadata.workbook.charts?.find(value => value.nodeId === current.nodeId), updated = data(operation.data, allocate, original);
    delete chart.categories; Object.assign(chart, updated);
    const retained = new Set(chart.series.map(series => series.seriesId));
    if (chart.seriesStyles) chart.seriesStyles = chart.seriesStyles.filter(style => retained.has(style.seriesId));
    if (chart.pointStyles) chart.pointStyles = chart.pointStyles.filter(style => retained.has(style.seriesId)); return;
  }
  if (operation.op === "setWorkbookChartTitle") {
    if (operation.target !== "title" && chart.kind === "pie" && operation.source !== null) invalid("A pie chart has no axis titles.");
    const old = operation.target === "title" ? { nodeId: chart.titleNodeId } : operation.target === "categoryAxis" ? chart.categoryAxisTitle : chart.valueAxisTitle;
    const rich = operation.source !== null && "html" in operation.source ? operation.source : undefined;
    if (old?.nodeId) { if (rich) edit().replaceSourceContent(old.nodeId, rich.html); else edit().removeSource(old.nodeId); }
    const value = operation.source === null ? undefined : rich ? { nodeId: old?.nodeId ?? edit().insertSource(current.nodeId, "<div>" + rich.html + "</div>").rootNodeIds[0]!, ...(rich.text ? { text: copy(rich.text) } : {}), ...(rich.cssClass ? { cssClass: rich.cssClass } : {}) } : "reference" in operation.source ? withoutCaches(operation.source) : undefined;
    if (operation.target === "title") {
      delete chart.titleNodeId; delete chart.titleReference; delete chart.titleText; delete chart.titleCssClass;
      if (value) { if ("nodeId" in value) { chart.titleNodeId = value.nodeId; if (value.text) chart.titleText = value.text; } else chart.titleReference = value.reference; if (value.cssClass) chart.titleCssClass = value.cssClass; }
    } else { const key = operation.target === "categoryAxis" ? "categoryAxisTitle" : "valueAxisTitle"; if (value) chart[key] = value; else delete chart[key]; }
  }
}
