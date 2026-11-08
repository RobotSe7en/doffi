import { createAssetDescriptor, createOpaqueId, DocumentError, WorkbookPresetSchema } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, TransactionPrepareOptions, WorkbookDrawingAnchor, WorkbookObject, WorkbookSnapshot } from "@doffi/core";
import type { WorkbookControlsHost, WorkbookControlsIdentity } from "./controls-host.js";
import { createWorkbookChartControls } from "./chart-controls.js";
import { createWorkbookObjectTextControls } from "./object-text-controls.js";
/** The fields are transient input. Only the owning transaction stores an
 * object, resource, anchor or chart; refresh never writes form values back. */
export function createWorkbookDrawingControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void } {
  const panel = document.createElement("details"), summary = document.createElement("summary"), fields: (HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement)[] = [];
  const defaults = new Map<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement, string>();
  summary.textContent = "图片、形状与图表"; panel.append(summary); parent.append(panel);
  const label = <T extends HTMLElement>(name: string, field: T): T => { const wrapper = document.createElement("label"); field.setAttribute("aria-label", name); wrapper.append(name, field); panel.append(wrapper); return field; };
  const input = (name: string, value = "", type = "text") => { const field = label(name, document.createElement("input")); field.type = type; field.value = value; fields.push(field); defaults.set(field, value); return field; };
  const select = (name: string, values: readonly (readonly [string, string])[]) => { const field = label(name, document.createElement("select")); for (const [value, text] of values) { const option = document.createElement("option"); option.value = value; option.textContent = text; field.append(option); } fields.push(field); defaults.set(field, field.value); return field; };
  const objects = select("选择绘图对象", []), name = input("对象名称"), kind = select("对象锚定方式", [["one-cell", "随起始单元格移动"], ["two-cell", "双单元格锚定"], ["absolute", "绝对位置"]]);
  const start = input("对象起始单元格", "A1"), end = input("对象结束单元格", "F9"), x = input("对象水平位置或起点偏移", "0", "number"), y = input("对象垂直位置或起点偏移", "0", "number");
  const width = input("对象宽度（磅）", "240", "number"), height = input("对象高度（磅）", "120", "number"), endX = input("对象终点水平偏移", "0", "number"), endY = input("对象终点垂直偏移", "0", "number");
  const behavior = select("双锚定移动和大小", [["both", "随单元格移动并调整大小"], ["move", "随单元格移动，保持大小"], ["fixed", "保持位置和大小"]]);
  const rotation = input("对象旋转角度", "0", "number"), flipH = input("对象水平翻转", "", "checkbox"), flipV = input("对象垂直翻转", "", "checkbox");
  const preset = select("形状预设", WorkbookPresetSchema.options.map(value => [value, value] as const));
  const text = label("对象文字", document.createElement("textarea")); text.maxLength = 32767; fields.push(text);
  defaults.set(text, "");
  const fill = input("对象填充颜色"), line = input("对象线条颜色"), font = input("对象文字颜色");
  const image = input("选择本地图片", "", "file"); image.accept = "image/png,image/jpeg,image/webp";
  const alt = input("图片替代文字"), crop = ["左", "上", "右", "下"].map(side => input("图片" + side + "侧裁剪比例", "0", "number"));
  for (const field of crop) { field.min = "0"; field.max = "0.999999"; field.step = "0.01"; }
  const chartKind = select("图表类型", [["column", "簇状柱状图"], ["line", "普通折线图"], ["pie", "单系列饼图"]]), legend = select("图例位置", [["bottom", "底部"], ["top", "顶部"], ["left", "左侧"], ["right", "右侧"], ["none", "隐藏"]]);
  const title = input("图表标题"), showLabels = input("显示图表值标签", "", "checkbox");
  const message = document.createElement("p"); message.setAttribute("role", "status"); message.setAttribute("aria-label", "绘图编辑状态"); panel.append(message);
  let selected: string | undefined, boundRevision = -1, dirty = false, observed = "", composing: HTMLInputElement | HTMLTextAreaElement | undefined;
  const draftGroups = new Set<string>();
  const buttons: { element: HTMLButtonElement; needsObject: boolean; action?: () => boolean }[] = [];
  const number = (field: HTMLInputElement): number => { if (field.validity.badInput || !field.value.trim() || !Number.isFinite(Number(field.value))) throw new DocumentError("InvalidOperation", "请输入完整的有限数值。"); return Number(field.value); };
  const point = (field: HTMLInputElement) => {
    const match = /^([A-Za-z]{1,3})([1-9]\d*)$/.exec(field.value.trim()); if (!match) throw new DocumentError("InvalidOperation", "请输入有效的单元格地址。");
    const column = [...match[1]!.toUpperCase()].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0) - 1, row = Number(match[2]) - 1;
    if (column >= 16384 || row >= 1048576) throw new DocumentError("InvalidOperation", "对象锚定超出工作表范围。"); return { row, column };
  };
  const address = (at: { row: number; column: number }) => { let column = at.column + 1, text = ""; while (column) { column--; text = String.fromCharCode(65 + column % 26) + text; column = Math.floor(column / 26); } return text + (at.row + 1); };
  const anchor = (): WorkbookDrawingAnchor => kind.value === "absolute" ? { kind: "absolute", xPt: number(x), yPt: number(y), widthPt: number(width), heightPt: number(height) }
    : kind.value === "two-cell" ? { kind: "two-cell", from: { ...point(start), offsetXPt: number(x), offsetYPt: number(y) }, to: { ...point(end), offsetXPt: number(endX), offsetYPt: number(endY) }, moveWithCells: behavior.value !== "fixed", sizeWithCells: behavior.value === "both" }
      : { kind: "one-cell", from: { ...point(start), offsetXPt: number(x), offsetYPt: number(y) }, widthPt: number(width), heightPt: number(height) };
  function current(): DeepReadonly<WorkbookObject> {
    const snapshot = host.snapshot(), object = snapshot?.metadata.workbook.objects?.find(object => object.nodeId === selected);
    if (!snapshot || !object) throw new DocumentError("InvalidOperation", "请先选择绘图对象。");
    if (snapshot.manifest.revision !== boundRevision) throw new DocumentError("RevisionConflict", "对象表单属于旧修订，请先重新读取对象。"); return object;
  }
  function body(snapshot: WorkbookSnapshot, id: string): HTMLElement | undefined {
    const template = document.createElement("template"); template.innerHTML = snapshot.content?.html ?? "";
    return [...template.content.querySelectorAll<HTMLElement>("[data-node-id]")].find(node => node.getAttribute("data-node-id") === id);
  }
  function load(): void {
    const snapshot = host.snapshot(), object = snapshot?.metadata.workbook.objects?.find(object => object.nodeId === objects.value); selected = object?.nodeId; dirty = false; draftGroups.clear(); message.textContent = "";
    if (!snapshot || !object) {
      boundRevision = snapshot?.manifest.revision ?? -1; for (const [field, value] of defaults) if (field !== objects) { if (field instanceof document.defaultView!.HTMLInputElement && field.type === "checkbox") field.checked = false; else field.value = value; }
      if (snapshot) { chartControls.load(undefined, snapshot); objectTextControls.load(undefined, snapshot); } refresh(); return;
    }
    boundRevision = snapshot.manifest.revision; name.value = object.name ?? ""; kind.value = object.anchor.kind;
    const a = object.anchor;
    x.value = String(a.kind === "absolute" ? a.xPt : a.from.offsetXPt); y.value = String(a.kind === "absolute" ? a.yPt : a.from.offsetYPt);
    if (a.kind !== "absolute") start.value = address(a.from);
    if (a.kind === "two-cell") { end.value = address(a.to); endX.value = String(a.to.offsetXPt); endY.value = String(a.to.offsetYPt); behavior.value = a.moveWithCells ? a.sizeWithCells ? "both" : "move" : "fixed"; }
    else { width.value = String(a.widthPt); height.value = String(a.heightPt); }
    const node = body(snapshot, object.nodeId); rotation.value = node?.getAttribute("data-office-rotation") ?? "0";
    flipH.checked = node?.getAttribute("data-office-flip-h") === "true"; flipV.checked = node?.getAttribute("data-office-flip-v") === "true";
    text.value = node ? [...node.children].filter(child => child.tagName.toLowerCase() === "p").map(child => {
      const clone = child.cloneNode(true) as HTMLElement; for (const br of clone.querySelectorAll("br")) br.replaceWith("\n"); return clone.textContent ?? "";
    }).join("\n") : "";
    if (object.kind === "geometry") preset.value = object.preset;
    if (object.kind === "image") { alt.value = node?.querySelector("img")?.getAttribute("alt") ?? ""; ["left", "top", "right", "bottom"].forEach((side, index) => { crop[index]!.value = node?.getAttribute("data-office-crop-" + side) ?? "0"; }); }
    const chart = snapshot.metadata.workbook.charts?.find(chart => chart.nodeId === object.nodeId);
    if (chart) { chartKind.value = chart.kind; legend.value = chart.legend; showLabels.checked = chart.showLabels; title.value = chart.titleNodeId ? body(snapshot, chart.titleNodeId)?.textContent ?? "" : ""; }
    chartControls.load(chart, snapshot);
    objectTextControls.load(object, snapshot);
    fill.value = line.value = font.value = ""; refresh();
  }
  const html = (value: string) => "<p>" + value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;").replace(/\n/g, "<br>") + "</p>";
  async function apply(operations: readonly DocumentOperation[], options?: TransactionPrepareOptions, expected: WorkbookControlsIdentity | undefined = host.snapshot()?.manifest, submitted?: readonly string[]): Promise<void> {
    const groups = submitted ?? operations.flatMap(operation => operation.op === "updateWorkbookObject" ? ["geometry"] : operation.op === "setWorkbookObjectText" ? ["text"]
      : operation.op === "setWorkbookObjectStyle" ? ["appearance"] : operation.op === "replaceWorkbookImage" ? ["image"] : operation.op === "setWorkbookImageCrop" ? ["crop"]
      : operation.op === "setWorkbookChart" ? ["chart-appearance"] : operation.op === "setWorkbookChartTitle" ? operation.target === "title" ? ["new-chart-title"] : [] : operation.op === "insertWorkbookChart" ? ["new-chart-title"]
      : operation.op === "insertWorkbookObject" ? operation.object.kind === "image" ? ["image"] : ["geometry", "text"] : []);
    await host.perform(operations, options, expected);
    for (const group of groups) draftGroups.delete(group);
    dirty = draftGroups.size > 0; if (!dirty) boundRevision = -1;
  }
  function button(label: string, action: () => Promise<void>, needsObject = true, available?: () => boolean) {
    const element = document.createElement("button"); element.type = "button"; element.textContent = label; panel.append(element); buttons.push({ element, needsObject, ...(available ? { action: available } : {}) });
    element.addEventListener("click", () => run(action), { signal: host.signal });
  }
  function run(action: () => Promise<void>): void { host.run(async () => { if (!host.idle() || composing) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。"); await action(); refresh(); }); }
  const chartControls = createWorkbookChartControls(document, panel, { signal: host.signal, current: () => {
    const object = current(), chart = host.snapshot()!.metadata.workbook.charts?.find(chart => chart.nodeId === object.nodeId);
    if (!chart) throw new DocumentError("InvalidOperation", "请先选择图表。"); return chart;
  }, perform: (operations, group) => apply(operations, undefined, undefined, [group]), run, changed: group => { dirty = true; draftGroups.add(group); } });
  const objectTextControls = createWorkbookObjectTextControls(document, panel, { signal: host.signal, current, snapshot: () => {
    const snapshot = host.snapshot(); if (!snapshot) throw new DocumentError("InvalidOperation", "绘图来源尚未加载。"); return snapshot;
  }, perform: (operations, groups) => apply(operations, undefined, undefined, groups), run, changed: group => { dirty = true; draftGroups.add(group); } });
  const target = () => host.selection() ?? (() => { throw new DocumentError("InvalidOperation", "请先选择工作表单元格。"); })();
  button("插入文本框", async () => { const at = target(); await apply([{ op: "insertWorkbookObject", sheetId: at.sheetId, anchor: anchor(), name: name.value, object: { kind: "text", html: html(text.value) } }]); }, false);
  button("插入基础形状", async () => { const at = target(); await apply([{ op: "insertWorkbookObject", sheetId: at.sheetId, anchor: anchor(), name: name.value, object: { kind: "geometry", preset: WorkbookPresetSchema.parse(preset.value), html: html(text.value) } }]); }, false);
  async function imageInput(replace: boolean): Promise<void> {
    const expected = host.snapshot()?.manifest, object = replace ? current() : undefined, sheet = target(), file = image.files?.[0], description = alt.value, position = anchor();
    if (!file || !["image/png", "image/jpeg", "image/webp"].includes(file.type)) throw new DocumentError("InvalidOperation", "请选择 PNG、JPEG 或 WebP 图片。");
    if (file.size > 32 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "图片超过单次输入预算。");
    const bytes = new Uint8Array(await file.arrayBuffer()), now = host.snapshot()?.manifest;
    if (!expected || now?.documentId !== expected.documentId || now.revision !== expected.revision) throw new DocumentError("RevisionConflict", "读取图片期间文档已变化。");
    const path = "assets/" + createOpaqueId() + ({ "image/png": ".png", "image/jpeg": ".jpg", "image/webp": ".webp" }[file.type] ?? ".png"), descriptor = createAssetDescriptor(path, file.type, bytes);
    await apply([object ? { op: "replaceWorkbookImage", nodeId: object.nodeId, assetPath: path, alt: description } : { op: "insertWorkbookObject", sheetId: sheet.sheetId, anchor: position, object: { kind: "image", assetPath: path, alt: description } }], { assets: [{ descriptor, bytes }] }, expected);
  }
  button("插入本地图片", () => imageInput(false), false); button("替换所选图片", () => imageInput(true), true, () => selectedObject()?.kind === "image");
  button("应用对象位置与属性", async () => { const object = current(); await apply([{ op: "updateWorkbookObject", nodeId: object.nodeId, changes: { anchor: anchor(), name: name.value, rotation: number(rotation), flipH: flipH.checked, flipV: flipV.checked, ...(object.kind === "geometry" ? { preset: WorkbookPresetSchema.parse(preset.value) } : {}) } }]); });
  button("替换对象文字", async () => { const object = current(); await apply([{ op: "setWorkbookObjectText", nodeId: object.nodeId, html: html(text.value) }]); }, true, () => ["text", "geometry"].includes(selectedObject()?.kind ?? ""));
  button("应用对象颜色", async () => { const object = current(), properties: Record<string, string> = {}; if (fill.value) { properties["--office-fill-color"] = fill.value; properties["--office-fill-kind"] = "solid"; } if (line.value) { properties["--office-line-color"] = line.value; properties["--office-line-kind"] = "solid"; } if (font.value) properties.color = font.value; await apply([{ op: "setWorkbookObjectStyle", nodeId: object.nodeId, properties }]); });
  button("应用图片裁剪", async () => { const object = current(); await apply([{ op: "setWorkbookImageCrop", nodeId: object.nodeId, crop: { left: number(crop[0]!), top: number(crop[1]!), right: number(crop[2]!), bottom: number(crop[3]!) } }]); }, true, () => selectedObject()?.kind === "image");
  button("从选区插入图表", async () => {
    const at = target(), r = at.range, sheet = host.snapshot()!.metadata.workbook.sheets.find(sheet => sheet.sheetId === at.sheetId)!;
    const columns = r.to.column - r.from.column + 1; if (columns < 2 || chartKind.value === "pie" && columns !== 2) throw new DocumentError("InvalidOperation", "请选择分类列及至少一列数值；饼图只需一列数值。");
    const reference = (column: number) => "='" + sheet.name.replace(/'/g, "''") + "'!" + address({ row: r.from.row, column }) + ":" + address({ row: r.to.row, column });
    await apply([{ op: "insertWorkbookChart", sheetId: at.sheetId, anchor: anchor(), kind: chartKind.value as "column", title: html(title.value), data: { categories: { expression: reference(r.from.column), valueType: "text" }, series: Array.from({ length: columns - 1 }, (_, index) => ({ values: { expression: reference(r.from.column + 1 + index), valueType: "number" } })) } }]);
  }, false);
  button("应用图表外观", async () => {
    const object = current(), chart = host.snapshot()!.metadata.workbook.charts!.find(chart => chart.nodeId === object.nodeId)!, operations: DocumentOperation[] = [];
    const next = chartKind.value as "column" | "line" | "pie", clear: ("gapWidth" | "overlap" | "firstSliceAngle")[] = next === "column" ? ["firstSliceAngle"] : next === "pie" ? ["gapWidth", "overlap"] : ["gapWidth", "overlap", "firstSliceAngle"];
    if (next === "pie" && chart.kind !== "pie") for (const target of ["categoryAxis", "valueAxis"] as const) if (target === "categoryAxis" ? chart.categoryAxisTitle : chart.valueAxisTitle) operations.push({ op: "setWorkbookChartTitle", nodeId: object.nodeId, target, source: null });
    const contents = chart.labelContents?.filter(value => next === "pie" || value !== "percent");
    operations.push({ op: "setWorkbookChart", nodeId: object.nodeId, changes: { kind: next, legend: legend.value as "bottom", showLabels: showLabels.checked,
      ...(next !== chart.kind && contents ? { labelContents: showLabels.checked ? contents.length ? [...contents] : ["value"] : [] } : {}) }, clear });
    await apply(operations);
  }, true, () => selectedObject()?.kind === "chart");
  button("替换图表标题", async () => { const object = current(); await apply([{ op: "setWorkbookChartTitle", nodeId: object.nodeId, target: "title", source: title.value ? { html: html(title.value) } : null }]); }, true, () => selectedObject()?.kind === "chart");
  button("复制绘图对象", async () => { const object = current(); await apply([{ op: "duplicateWorkbookObjects", nodeIds: [object.nodeId], offset: { xPt: 12, yPt: 12 } }]); });
  for (const [position, text] of [["front", "置于顶层"], ["back", "置于底层"], ["forward", "上移一层"], ["backward", "下移一层"]] as const) button(text, async () => { const object = current(); await apply([{ op: "reorderWorkbookObjects", nodeIds: [object.nodeId], position }]); });
  button("删除绘图对象", async () => { const object = current(); await apply([{ op: "removeWorkbookObject", nodeId: object.nodeId }]); });
  const reload = document.createElement("button"); reload.type = "button"; reload.textContent = "重新读取对象"; panel.append(reload); reload.addEventListener("click", () => { if (!composing) load(); }, { signal: host.signal });
  const discard = document.createElement("button"); discard.type = "button"; discard.textContent = "放弃绘图表单草稿"; panel.append(discard); discard.addEventListener("click", () => { if (!composing) load(); }, { signal: host.signal });
  objects.addEventListener("change", () => {
    if (dirty || composing) { objects.value = selected ?? ""; message.textContent = "请先应用或放弃当前表单草稿，再切换绘图对象。"; return; } load();
  }, { signal: host.signal });
  for (const field of fields.filter(field => field !== objects)) field.addEventListener("input", () => {
    dirty = true; draftGroups.add(field === text ? "text" : [fill, line, font].includes(field as HTMLInputElement) ? "appearance" : [image, alt].includes(field as HTMLInputElement) ? "image"
      : crop.includes(field as HTMLInputElement) ? "crop" : [chartKind, legend, showLabels].includes(field as HTMLSelectElement) ? "chart-appearance" : field === title ? "new-chart-title" : "geometry");
  }, { signal: host.signal });
  panel.addEventListener("compositionstart", event => { if (event.target instanceof document.defaultView!.HTMLInputElement || event.target instanceof document.defaultView!.HTMLTextAreaElement) composing = event.target; refresh(); }, { signal: host.signal });
  panel.addEventListener("compositionend", () => { composing = undefined; refresh(); }, { signal: host.signal });
  function selectedObject() { return host.snapshot()?.metadata.workbook.objects?.find(object => object.nodeId === selected); }
  function refresh(): void {
    const snapshot = host.snapshot(), sheetId = host.selection()?.sheetId, items = snapshot?.metadata.workbook.objects?.filter(object => object.sheetId === sheetId) ?? [];
    const key = JSON.stringify([snapshot?.manifest.documentId, sheetId, items.map(item => [item.nodeId, item.name])]);
    if (key !== observed) {
      observed = key; const previous = objects.value; objects.replaceChildren(); const blank = document.createElement("option"); blank.textContent = "请选择对象"; blank.value = ""; objects.append(blank);
      for (const [index, item] of items.entries()) { const option = document.createElement("option"); option.value = item.nodeId; option.textContent = item.name || ({ text: "文本框", geometry: "形状", image: "图片", chart: "图表" }[item.kind] + " " + (index + 1)); objects.append(option); }
      objects.value = items.some(item => item.nodeId === previous) ? previous : ""; if (!objects.value) selected = undefined;
    }
    if (selected && snapshot && boundRevision !== snapshot.manifest.revision && !dirty && !composing) { load(); return; }
    const idle = host.idle() && !composing, stale = !!selected && snapshot?.manifest.revision !== boundRevision;
    message.textContent = stale ? "文档已变化。表单草稿已保留，请重新读取对象后编辑。" : "";
    for (const field of fields) field.disabled = !idle && field !== composing; reload.disabled = !idle || !objects.value;
    discard.disabled = !idle || !dirty;
    for (const button of buttons) button.element.disabled = !idle || button.needsObject && (!selectedObject() || stale) || button.action?.() === false;
    chartControls.refresh(!idle, selectedObject()?.kind === "chart", composing, stale);
    objectTextControls.refresh(!idle, composing, stale || !selectedObject());
  }
  return { element: panel, refresh };
}
