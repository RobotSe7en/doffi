import { createWorkbookChartStyleControls } from "./chart-style-controls.js";
import type { WorkbookChartStyleDraftGroup } from "./chart-style-controls.js";
import { DocumentError } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, WorkbookChart, WorkbookChartDataInput, WorkbookSnapshot } from "@doffi/core";
import { createWorkbookChartVectorControls, readWorkbookChartVectorDraft, workbookChartVectorDraft } from "./chart-vector-controls.js";
import type { WorkbookChartVectorDraft } from "./chart-vector-controls.js";

interface Host {
  readonly signal: AbortSignal;
  current(): DeepReadonly<WorkbookChart>;
  perform(operations: readonly DocumentOperation[], group: ChartDraftGroup): Promise<void>;
  run(action: () => Promise<void>): void;
  changed(group: ChartDraftGroup): void;
}
export type ChartDraftGroup = WorkbookChartStyleDraftGroup | "chart-data" | "chart-axis" | "chart-title" | "chart-categoryAxis" | "chart-valueAxis";
type DataSeries = { seriesId?: string; nameKind: "none" | "text" | "reference"; name: string; nameValueType: "text" | "number"; values: WorkbookChartVectorDraft };
type Appearance = Extract<DocumentOperation, { op: "setWorkbookChart" }>;
interface Controls { load(chart: DeepReadonly<WorkbookChart> | undefined, snapshot: WorkbookSnapshot): void; refresh(disabled: boolean, available?: boolean, composing?: HTMLElement, stale?: boolean): void }
/** A source editor: references and literals are distinct inputs, with no copied
 * cache values, independently writable data set or partial source submission. */
export function createWorkbookChartControls(document: Document, parent: HTMLElement, host: Host): Controls {
  const root = document.createElement("fieldset"), legend = document.createElement("legend"); legend.textContent = "图表数据、坐标轴与标题"; root.append(legend); parent.append(root);
  const styleControls = createWorkbookChartStyleControls(document, root, host);
  let series: DataSeries[] = [], selected = 0, available = false, blocked = true, sourceStale = false, kind: WorkbookChart["kind"] = "column";
  const fields: (HTMLInputElement | HTMLSelectElement)[] = [], buttons: HTMLButtonElement[] = [];
  const label = <T extends HTMLElement>(name: string, field: T, container = root): T => { const wrapper = document.createElement("label"); field.setAttribute("aria-label", name); wrapper.append(name, field); container.append(wrapper); return field; };
  const input = (name: string, type = "text", container = root) => { const field = label(name, document.createElement("input"), container); field.type = type; fields.push(field); return field; };
  const select = (name: string, choices: readonly (readonly [string, string])[], container = root) => { const field = label(name, document.createElement("select"), container); for (const [value, text] of choices) { const option = document.createElement("option"); option.value = value; option.textContent = text; field.append(option); } fields.push(field); return field; };
  const button = (name: string, action: () => Promise<void> | void, container = root) => { const field = document.createElement("button"); field.type = "button"; field.textContent = name; container.append(field); buttons.push(field); field.addEventListener("click", () => host.run(async () => { host.current(); await action(); }), { signal: host.signal }); return field; };
  const categoryMode = select("图表分类", [["automatic", "使用默认序号"], ["source", "指定分类来源"]]);
  let categories = workbookChartVectorDraft();
  const categoryContainer = document.createElement("div"); root.append(categoryContainer);
  const dataChanged = () => host.changed("chart-data");
  const categoryControls = createWorkbookChartVectorControls(document, categoryContainer, "分类数据", false, host.signal, dataChanged);
  const seriesChoice = select("编辑图表系列", []), nameKind = select("系列名称来源", [["none", "使用默认名称"], ["text", "文字"], ["reference", "单元格或名称引用"]]), name = input("系列名称或引用"), nameValueType = select("系列名称引用类型", [["text", "文字"], ["number", "数值"]]);
  const valueControls = createWorkbookChartVectorControls(document, root, "系列数值", true, host.signal, dataChanged);
  function showSeries(): void {
    seriesChoice.replaceChildren(...series.map((item, index) => { const option = document.createElement("option"); option.value = String(index); option.textContent = item.nameKind === "text" && item.name ? item.name : "系列 " + (index + 1); return option; }));
    selected = Math.max(0, Math.min(selected, series.length - 1)); seriesChoice.value = String(selected);
    const item = series[selected]; nameKind.value = item?.nameKind ?? "none"; name.value = item?.name ?? ""; nameValueType.value = item?.nameValueType ?? "text";
    if (item) valueControls.load(item.values);
    refresh(blocked, available);
  }
  seriesChoice.addEventListener("change", () => { selected = Number(seriesChoice.value); showSeries(); }, { signal: host.signal });
  nameKind.addEventListener("change", () => { const item = series[selected]; if (item) item.nameKind = nameKind.value as DataSeries["nameKind"]; dataChanged(); refresh(blocked, available); }, { signal: host.signal });
  nameValueType.addEventListener("change", () => { const item = series[selected]; if (item) item.nameValueType = nameValueType.value as DataSeries["nameValueType"]; dataChanged(); }, { signal: host.signal });
  name.addEventListener("input", () => { const item = series[selected]; if (item) { item.name = name.value; const option = seriesChoice.options[selected]; if (option) option.textContent = item.nameKind === "text" && item.name ? item.name : "系列 " + (selected + 1); } dataChanged(); }, { signal: host.signal });
  categoryMode.addEventListener("change", () => { categoryContainer.hidden = categoryMode.value !== "source"; dataChanged(); }, { signal: host.signal });
  const add = button("新增图表系列", () => { series.push({ nameKind: "text", name: "新系列", nameValueType: "text", values: { ...workbookChartVectorDraft(), mode: "literal" } }); selected = series.length - 1; dataChanged(); showSeries(); });
  const remove = button("删除当前图表系列", () => { if (series.length === 1) throw new DocumentError("InvalidOperation", "图表至少需要一个系列。"); series.splice(selected, 1); dataChanged(); showSeries(); });
  const up = button("图表系列上移", () => move(-1)), down = button("图表系列下移", () => move(1));
  function move(offset: number): void { const next = selected + offset; if (next < 0 || next >= series.length) return; [series[selected], series[next]] = [series[next]!, series[selected]!]; selected = next; dataChanged(); showSeries(); }
  button("应用图表数据", async () => {
    const chart = host.current();
    const data: WorkbookChartDataInput = { ...(categoryMode.value === "source" ? { categories: readWorkbookChartVectorDraft(categories) } : {}), series: series.map(item => ({
      ...(item.seriesId ? { seriesId: item.seriesId } : {}), ...(item.nameKind === "none" ? {} : item.nameKind === "text" ? { name: { kind: "text" as const, text: item.name } } : { name: { kind: "reference" as const, reference: { expression: item.name, valueType: item.nameValueType } } }), values: readWorkbookChartVectorDraft(item.values),
    })) };
    await host.perform([{ op: "setWorkbookChartData", nodeId: chart.nodeId, data }], "chart-data");
  });

  const appearance = document.createElement("fieldset"), appearanceLegend = document.createElement("legend"); appearanceLegend.textContent = "坐标轴、标签与空白"; appearance.append(appearanceLegend); root.append(appearance);
  const edited = new Set<string>(), settings = new Map<string, HTMLInputElement | HTMLSelectElement>();
  const setting = (key: string, field: HTMLInputElement | HTMLSelectElement) => { settings.set(key, field); field.addEventListener("input", () => { edited.add(key); host.changed("chart-axis"); }, { signal: host.signal }); return field; };
  for (const [key, text] of [["minimum", "数值轴最小值"], ["maximum", "数值轴最大值"], ["majorUnit", "数值轴主刻度"], ["gapWidth", "柱间距百分比"], ["overlap", "柱重叠百分比"], ["firstSliceAngle", "饼图起始角度"]] as const) setting(key, input(text, "number", appearance));
  setting("numberFormat", input("图表数值格式", "text", appearance));
  for (const [key, text] of [["showCategoryAxis", "显示分类轴"], ["showValueAxis", "显示数值轴"], ["reverseCategories", "分类轴逆序"], ["reverseValues", "数值轴逆序"], ["visibleCellsOnly", "只绘制可见单元格"], ["varyColors", "按数据点分色"], ["showLabelsOverMaximum", "显示超出最大值的标签"]] as const) setting(key, input(text, "checkbox", appearance));
  setting("blankAs", select("图表空白处理", [["gap", "留空"], ["zero", "作为零"], ["span", "连接数据点"]], appearance));
  const labels = new Map(([ ["value", "数据标签数值"], ["category", "数据标签分类"], ["series", "数据标签系列"], ["percent", "数据标签百分比"] ] as const).map(([key, text]) => [key, setting("labels", input(text, "checkbox", appearance)) as HTMLInputElement]));
  button("应用坐标轴与标签", async () => {
    const chart = host.current(), changes: Appearance["changes"] = {}, clear: NonNullable<Appearance["clear"]> = [];
    for (const key of ["minimum", "maximum", "majorUnit", "gapWidth", "overlap", "firstSliceAngle"] as const) if (edited.has(key)) {
      const field = settings.get(key)!;
      if (field instanceof document.defaultView!.HTMLInputElement && field.validity.badInput) throw new DocumentError("InvalidOperation", "请输入完整的图表数值。");
      if (!field.value.trim()) clear.push(key);
      else { const number = Number(field.value); if (!Number.isFinite(number)) throw new DocumentError("InvalidOperation", "请输入有限的图表数值。"); changes[key] = number; }
    }
    if (edited.has("numberFormat")) { const value = settings.get("numberFormat")!.value; if (value) changes.numberFormat = value; else clear.push("numberFormat"); }
    for (const key of ["showCategoryAxis", "showValueAxis", "reverseCategories", "reverseValues", "visibleCellsOnly", "varyColors", "showLabelsOverMaximum"] as const) if (edited.has(key)) changes[key] = (settings.get(key) as HTMLInputElement).checked;
    if (edited.has("blankAs")) changes.blankAs = settings.get("blankAs")!.value as "gap" | "zero" | "span";
    if (edited.has("labels")) { changes.labelContents = [...labels].filter(([, field]) => field.checked).map(([key]) => key); changes.showLabels = !!changes.labelContents.length; }
    await host.perform([{ op: "setWorkbookChart", nodeId: chart.nodeId, changes, ...(clear.length ? { clear } : {}) }], "chart-axis");
  }, appearance);

  const titleDrafts = ([ ["title", "主标题"], ["categoryAxis", "分类轴标题"], ["valueAxis", "数值轴标题"] ] as const).map(([target, text]) => {
    const mode = select(text + "来源", [["none", "无标题"], ["text", "文字"], ["reference", "单元格或名称引用"]]), value = input(text + "文字或引用"), valueType = select(text + "引用类型", [["text", "文字"], ["number", "数值"]]);
    let changed = false;
    for (const field of [mode, value, valueType]) field.addEventListener("input", () => { changed = true; host.changed("chart-" + target as ChartDraftGroup); }, { signal: host.signal });
    button("应用" + text, async () => {
      if (!changed) return;
      const chart = host.current(), prior = target === "title" ? { cssClass: chart.titleCssClass, text: chart.titleText } : target === "categoryAxis" ? chart.categoryAxisTitle : chart.valueAxisTitle;
      const source = mode.value === "none" ? null : mode.value === "reference" ? { reference: { expression: value.value, valueType: valueType.value as "text" | "number" }, ...(prior?.cssClass ? { cssClass: prior.cssClass } : {}) }
        : { html: "<p>" + value.value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;").replace(/\n/g, "<br>") + "</p>", ...(prior?.text ? { text: prior.text } : {}), ...(prior?.cssClass ? { cssClass: prior.cssClass } : {}) };
      await host.perform([{ op: "setWorkbookChartTitle", nodeId: chart.nodeId, target, source }], "chart-" + target as ChartDraftGroup);
    });
    return { target, mode, value, valueType, reset() { changed = false; } };
  });
  function load(chart: DeepReadonly<WorkbookChart> | undefined, snapshot: WorkbookSnapshot): void {
    styleControls.load(chart);
    available = !!chart; edited.clear(); if (!chart) { refresh(blocked, false); return; } kind = chart.kind;
    categories = workbookChartVectorDraft(chart.categories); if (!chart.categories) categories.valueType = "text"; categoryControls.load(categories); categoryMode.value = chart.categories ? "source" : "automatic"; categoryContainer.hidden = !chart.categories;
    series = chart.series.map(item => ({ seriesId: item.seriesId, nameKind: item.name?.kind ?? "none", name: item.name?.kind === "text" ? item.name.text : item.name?.reference.expression ?? "", nameValueType: item.name?.kind === "reference" ? item.name.reference.valueType : "text", values: workbookChartVectorDraft(item.values) }));
    showSeries();
    for (const [key, field] of settings) {
      if (key === "labels") continue;
      const value = chart[key as keyof WorkbookChart];
      if (field instanceof document.defaultView!.HTMLInputElement && field.type === "checkbox") field.checked = value === undefined ? ["showCategoryAxis", "showValueAxis", "visibleCellsOnly"].includes(key) || key === "varyColors" && kind === "pie" : value === true;
      else field.value = value === undefined ? key === "blankAs" ? "gap" : "" : String(value);
    }
    for (const [key, field] of labels) field.checked = chart.labelContents?.includes(key) ?? (chart.showLabels && key === "value");
    const template = document.createElement("template"); template.innerHTML = snapshot.content?.html ?? "";
    for (const item of titleDrafts) {
      const source = item.target === "title" ? { nodeId: chart.titleNodeId, reference: chart.titleReference } : item.target === "categoryAxis" ? chart.categoryAxisTitle : chart.valueAxisTitle;
      const node = source?.nodeId ? [...template.content.querySelectorAll<HTMLElement>("[data-node-id]")].find(node => node.getAttribute("data-node-id") === source.nodeId) : undefined;
      if (node) for (const br of node.querySelectorAll("br")) br.replaceWith("\n");
      item.mode.value = source?.reference ? "reference" : source?.nodeId ? "text" : "none"; item.value.value = source?.reference?.expression ?? node?.textContent ?? ""; item.valueType.value = source?.reference?.valueType ?? "text"; item.reset();
    }
  }
  function refresh(disabled: boolean, isAvailable = available, composing?: HTMLElement, stale = sourceStale): void {
    styleControls.refresh(disabled, composing, stale);
    available = isAvailable; blocked = disabled; sourceStale = stale; root.hidden = !available;
    for (const field of fields) field.disabled = (disabled || !available) && field !== composing;
    for (const field of buttons) field.disabled = disabled || stale || !available;
    labels.get("percent")!.disabled ||= kind !== "pie";
    for (const key of ["gapWidth", "overlap", "firstSliceAngle"] as const) settings.get(key)!.disabled ||= key === "firstSliceAngle" ? kind !== "pie" : kind !== "column";
    name.disabled ||= nameKind.value === "none"; nameValueType.disabled ||= nameKind.value !== "reference"; remove.disabled ||= series.length < 2; up.disabled ||= selected === 0; down.disabled ||= selected >= series.length - 1;
    add.disabled ||= !available; categoryControls.refresh(disabled || !available, composing); valueControls.refresh(disabled || !available, composing);
  }
  return { load, refresh };
}
