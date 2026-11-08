import { DocumentError } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, WorkbookChart } from "@doffi/core";
import { createWorkbookChartDetailFields } from "./chart-detail-fields.js";

type Operation = Extract<DocumentOperation, { op: "setWorkbookChartStyle" }>;
export type WorkbookChartStyleDraftGroup = `chart-style:${string}`;
interface Host {
  readonly signal: AbortSignal;
  current(): DeepReadonly<WorkbookChart>;
  perform(operations: readonly DocumentOperation[], group: WorkbookChartStyleDraftGroup): Promise<void>;
  run(action: () => Promise<void>): void;
  changed(group: WorkbookChartStyleDraftGroup): void;
}
interface Controls { load(chart: DeepReadonly<WorkbookChart> | undefined): void; refresh(disabled: boolean, composing?: HTMLElement, stale?: boolean): void }
/** Unsubmitted appearance values stay local to their original part/series/point.
 * Navigating these drafts never publishes a revision or a derived chart cache. */
export function createWorkbookChartStyleControls(document: Document, parent: HTMLElement, host: Host): Controls {
  const root = document.createElement("fieldset"), legend = document.createElement("legend"); legend.textContent = "图表部件与数据点样式"; root.append(legend); parent.append(root);
  const fields: (HTMLSelectElement | HTMLInputElement)[] = [], values = new Map<string, HTMLSelectElement | HTMLInputElement>(), drafts = new Map<string, Record<string, string>>();
  let available = false, kind: WorkbookChart["kind"] = "column", active = "", blocked = false, sourceStale = false;
  let activeComposition: HTMLElement | undefined;
  const label = <T extends HTMLElement>(name: string, field: T): T => { const wrapper = document.createElement("label"); field.setAttribute("aria-label", name); wrapper.append(name, field); root.append(wrapper); return field; };
  const choice = (name: string, options: readonly (readonly [string, string])[]) => { const field = label(name, document.createElement("select")); for (const [value, text] of options) { const option = document.createElement("option"); option.value = value; option.textContent = text; field.append(option); } fields.push(field); return field; };
  const input = (name: string, type = "text") => { const field = label(name, document.createElement("input")); field.type = type; fields.push(field); return field; };
  const target = choice("图表样式部件", [["plot", "绘图区"], ["legend", "图例"], ["text", "默认文字"], ["label", "所有数据标签"], ["categoryAxis", "分类轴"], ["valueAxis", "数值轴"], ["grid", "数值轴网格线"], ["categoryGrid", "分类轴网格线"], ["series", "系列"], ["point", "数据点"], ["seriesLabel", "系列数据标签"], ["pointLabel", "单点数据标签"]]);
  const series = choice("图表样式系列", []), index = input("图表样式原始数据位置（从 1 开始）", "number"); index.min = "1"; index.max = "4294967296"; index.step = "1"; index.value = "1";
  const hasText = () => ["text", "legend", "label", "categoryAxis", "valueAxis", "seriesLabel", "pointLabel"].includes(target.value);
  const hasFill = () => ["plot", "legend", "label", "series", "point", "seriesLabel", "pointLabel"].includes(target.value), hasLine = () => target.value !== "text";
  const hasSeries = () => ["series", "point", "seriesLabel", "pointLabel"].includes(target.value), hasPoint = () => target.value === "point" || target.value === "pointLabel";
  const key = () => JSON.stringify([target.value, hasSeries() ? series.value : "", hasPoint() ? index.value : ""]);
  const group = (): WorkbookChartStyleDraftGroup => `chart-style:${active}`;
  const value = (key: string, field: HTMLInputElement | HTMLSelectElement) => { values.set(key, field); field.addEventListener("input", () => { drafts.set(active, Object.fromEntries([...values].map(([key, field]) => [key, field.value]))); host.changed(group()); refresh(blocked, activeComposition, sourceStale); }, { signal: host.signal }); return field; };
  value("fill", input("图表部件填充颜色")); value("fillKind", choice("图表部件填充", [["", "保持原设置"], ["solid", "纯色"], ["none", "无填充"]]));
  value("line", input("图表部件线条颜色")); value("lineWidth", input("图表部件线宽（磅）", "number"));
  value("lineKind", choice("图表部件线条", [["", "保持原设置"], ["solid", "实线"], ["none", "无线条"]]));
  value("dash", choice("图表部件线型", [["", "保持原设置"], ["solid", "实线"], ["dash", "虚线"], ["dot", "点线"], ["dashDot", "点划线"]]));
  value("color", input("图表部件文字颜色")); value("latin", input("图表部件西文字体")); value("eastAsia", input("图表部件中文字体")); value("fontSize", input("图表部件字号（磅）", "number"));
  value("weight", choice("图表部件字重", [["", "保持原设置"], ["bold", "加粗"], ["normal", "常规"]])); value("italic", choice("图表部件斜体", [["", "保持原设置"], ["italic", "斜体"], ["normal", "常规"]]));
  value("decoration", choice("图表部件文字装饰", [["", "保持原设置"], ["none", "清除装饰线"], ["underline", "下划线"], ["line-through", "删除线"], ["underline line-through", "下划线与删除线"]]));
  value("decorationStyle", choice("图表部件装饰线样式", [["", "保持原设置"], ["solid", "单线"], ["double", "双线"], ["dotted", "点线"], ["dashed", "虚线"], ["wavy", "波浪线"]]));
  const marker = value("marker", choice("图表数据标记", [["", "保持原设置"], ["default", "恢复默认"], ["none", "无标记"], ["auto", "自动"], ["circle", "圆"], ["square", "方形"], ["diamond", "菱形"], ["triangle", "三角形"], ["x", "交叉"], ["plus", "加号"], ["star", "星形"], ["dot", "点"], ["dash", "短线"]]));
  const markerSize = value("markerSize", input("图表标记大小（2 至 72 磅，0 恢复默认）", "number"));
  const details = createWorkbookChartDetailFields({ input, choice, value });
  const read = (name: string) => values.get(name)!.value.trim();
  const number = (name: string, minimum: number, maximum: number): number => { const field = values.get(name) as HTMLInputElement, result = Number(field.value); if (field.validity.badInput || !field.value || !Number.isFinite(result) || result < minimum || result > maximum) throw new DocumentError("InvalidOperation", "请输入有效的图表样式数值。"); return result; };
  const apply = document.createElement("button"); apply.type = "button"; apply.textContent = "应用当前图表部件样式"; root.append(apply);
  apply.addEventListener("click", () => host.run(async () => {
    const chart = host.current(), properties: Record<string, string> = details.properties(target.value as Operation["target"], kind), font = (name: string) => { const raw = read(name); if (/[\u0000-\u001f\u007f]/.test(raw)) throw new DocumentError("InvalidOperation", "字体名称不能包含控制字符。"); return '"' + raw.replaceAll("\\", "\\\\").replaceAll('"', '\\"') + '"'; };
    for (const [field, property] of [["fill", "--office-fill-color"], ["line", "--office-line-color"], ["color", "color"], ["weight", "font-weight"], ["italic", "font-style"], ["dash", "--office-line-dash"], ["decoration", "text-decoration-line"], ["decorationStyle", "text-decoration-style"]] as const) if (read(field) && (field === "fill" ? hasFill() : ["line", "dash"].includes(field) ? hasLine() : hasText())) properties[property] = read(field);
    if (hasFill() && (read("fillKind") || read("fill"))) properties["--office-fill-kind"] = read("fillKind") || "solid";
    if (hasLine() && (read("lineKind") || read("line") || read("lineWidth") || read("dash"))) properties["--office-line-kind"] = read("lineKind") || "solid";
    if (hasLine() && (read("lineWidth") || (values.get("lineWidth") as HTMLInputElement).validity.badInput)) properties["--office-line-width"] = number("lineWidth", 0, 1584) + "pt";
    if (hasText() && (read("fontSize") || (values.get("fontSize") as HTMLInputElement).validity.badInput)) properties["font-size"] = number("fontSize", .1, 4000) + "pt";
    if (hasText() && read("latin")) for (const role of ["ascii", "h-ansi"]) properties["--office-font-" + role] = font("latin");
    if (hasText() && read("eastAsia")) properties["--office-font-east-asia"] = font("eastAsia");
    const position = Number(index.value) - 1;
    if (hasPoint() && (index.validity.badInput || !Number.isInteger(position) || position < 0 || position > 0xffff_ffff)) throw new DocumentError("InvalidOperation", "原始数据位置须在 1 至 4294967296 之间。");
    const operation: Operation = { op: "setWorkbookChartStyle", nodeId: chart.nodeId, target: target.value as Operation["target"], properties,
      ...(hasSeries() ? { seriesId: series.value } : {}), ...(hasPoint() ? { index: position } : {}) };
    if (!marker.disabled && marker.value) operation.marker = marker.value === "default" ? null : marker.value as NonNullable<Operation["marker"]>;
    if (!markerSize.disabled && (markerSize.value || (markerSize as HTMLInputElement).validity.badInput)) { const size = number("markerSize", 0, 72); if (size && (size < 2 || !Number.isInteger(size))) throw new DocumentError("InvalidOperation", "标记大小须为 2 至 72 的整数，或用 0 恢复默认。"); operation.markerSize = size || null; }
    if (!Object.keys(properties).length && operation.marker === undefined && operation.markerSize === undefined) throw new DocumentError("InvalidOperation", "请先指定当前图表部件的样式。");
    await host.perform([operation], group());
  }), { signal: host.signal });
  function show(): void { active = key(); const draft = drafts.get(active); for (const [name, field] of values) field.value = draft?.[name] ?? ""; refresh(blocked, undefined, sourceStale); }
  for (const field of [target, series, index]) field.addEventListener(field === index ? "input" : "change", show, { signal: host.signal });
  function load(chart: DeepReadonly<WorkbookChart> | undefined): void {
    available = !!chart; root.hidden = !available; drafts.clear(); sourceStale = false; if (!chart) return; kind = chart.kind;
    const previous = series.value; series.replaceChildren(...chart.series.map((item, i) => { const option = document.createElement("option"); option.value = item.seriesId; option.textContent = item.name?.kind === "text" ? item.name.text : "系列 " + (i + 1); return option; }));
    if (chart.series.some(item => item.seriesId === previous)) series.value = previous; show();
  }
  function refresh(disabled: boolean, composing?: HTMLElement, stale = false): void {
    blocked = disabled; sourceStale = stale; activeComposition = composing; root.hidden = !available;
    for (const field of fields) field.disabled = (disabled || !available) && field !== composing;
    series.disabled ||= !hasSeries(); index.disabled ||= !hasPoint();
    for (const name of ["color", "latin", "eastAsia", "fontSize", "weight", "italic", "decoration", "decorationStyle"]) { const field = values.get(name)!; field.disabled ||= !hasText() && field !== composing; }
    for (const name of ["fill", "fillKind", "line", "lineKind", "lineWidth", "dash"]) { const field = values.get(name)!; field.disabled ||= !(name.startsWith("fill") ? hasFill() : hasLine()) && field !== composing; }
    const markersAvailable = kind === "line" && (target.value === "series" || target.value === "point");
    marker.disabled ||= !markersAvailable; markerSize.disabled ||= !markersAvailable;
    details.refresh(target.value as Operation["target"], kind, composing);
    apply.disabled = disabled || stale || !available;
  }
  return { load, refresh };
}
