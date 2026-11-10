import { DocumentError, WorkbookChartVectorSchema, WorkbookErrorValueSchema } from "@doffi/core";
import type { DeepReadonly, WorkbookChartVector } from "@doffi/core";

type NativePoint = Extract<WorkbookChartVector, { kind: "literal" }>["points"][number];
type Point = { position: string; value: string; numberFormat: string; original?: DeepReadonly<NativePoint> };
/** Input spellings remain private drafts. In particular, native omitted points,
 * errors, per-point formats and the declared vector length survive paging. */
export interface WorkbookChartVectorDraft {
  mode: "reference" | "literal"; valueType: "text" | "number"; expression: string;
  count: string; numberFormat: string; points: Point[]; original?: DeepReadonly<WorkbookChartVector>; originSheetId?: string;
}
export function workbookChartVectorDraft(value?: DeepReadonly<WorkbookChartVector>): WorkbookChartVectorDraft {
  const literal = value && "kind" in value ? value : undefined;
  return { mode: literal ? "literal" : "reference", valueType: value?.valueType ?? "number", expression: value && "expression" in value ? value.expression : "=",
    ...(value && "expression" in value && value.originSheetId !== undefined ? { originSheetId: value.originSheetId } : {}),
    count: String(literal?.count ?? 0), numberFormat: literal?.numberFormat ?? "", points: literal?.points.map(point => ({ position: String(point.index + 1), value: point.value === null ? "" : String(point.value), numberFormat: point.numberFormat ?? "", original: point })) ?? [], ...(value ? { original: value } : {}) };
}
export function readWorkbookChartVectorDraft(draft: WorkbookChartVectorDraft): WorkbookChartVector {
  const validate = (value: unknown): WorkbookChartVector => { const result = WorkbookChartVectorSchema.safeParse(value); if (!result.success) throw new DocumentError("InvalidOperation", "图表引用或常量数据不符合当前格式和长度限制。"); return result.data; };
  if (draft.mode === "reference") return validate({ expression: draft.expression, valueType: draft.valueType,
    ...(draft.originSheetId !== undefined ? { originSheetId: draft.originSheetId } : {}) });
  const integer = (value: string): number => { if (!/^\d+$/.test(value) || !Number.isSafeInteger(Number(value))) throw new DocumentError("InvalidOperation", "图表数据的位置和总长度须为整数。"); return Number(value); };
  const points = draft.points.map(point => {
    const position = integer(point.position); if (position < 1) throw new DocumentError("InvalidOperation", "图表数据位置从 1 开始。");
    const value = draft.valueType === "text" ? point.value : !point.value.trim() ? null : WorkbookErrorValueSchema.safeParse(point.value).success ? point.value : Number(point.value);
    if (typeof value === "number" && !Number.isFinite(value)) throw new DocumentError("InvalidOperation", "图表数值须为有限数字、错误值或留空。");
    return { index: position - 1, value, ...(draft.valueType === "number" && (point.numberFormat || point.original?.numberFormat === "") ? { numberFormat: point.numberFormat } : {}) };
  }).sort((a, b) => a.index - b.index);
  const count = integer(draft.count);
  if (points.some((point, index) => point.index >= count || index > 0 && points[index - 1]!.index === point.index)) throw new DocumentError("InvalidOperation", "图表数据位置须唯一且不超过总长度。");
  return validate({ kind: "literal", valueType: draft.valueType, count, points, ...(draft.valueType === "number" && (draft.numberFormat || draft.original && "kind" in draft.original && draft.original.numberFormat === "") ? { numberFormat: draft.numberFormat } : {}) });
}

type OriginSheet = { readonly sheetId: string; readonly name: string };
interface VectorControls { load(value: WorkbookChartVectorDraft, sheets?: readonly OriginSheet[]): void; refresh(disabled: boolean, composing?: HTMLElement): void }

/** Pages stored native points, never the potentially enormous implicit gaps. */
export function createWorkbookChartVectorControls(document: Document, parent: HTMLElement, name: string, numeric: boolean, signal: AbortSignal, changed: () => void): VectorControls {
  const root = document.createElement("fieldset"), legend = document.createElement("legend"); legend.textContent = name; root.append(legend); parent.append(root);
  let draft = workbookChartVectorDraft(), page = 0, disabled = false, rows = new AbortController();
  signal.addEventListener("abort", () => rows.abort(), { once: true });
  const label = <T extends HTMLElement>(text: string, field: T, container: HTMLElement = root): T => { const wrapper = document.createElement("label"); field.setAttribute("aria-label", text); wrapper.append(text, field); container.append(wrapper); return field; };
  const select = (text: string, options: readonly (readonly [string, string])[]) => { const field = label(text, document.createElement("select")); for (const [value, text] of options) { const option = document.createElement("option"); option.value = value; option.textContent = text; field.append(option); } return field; };
  const input = (text: string, container?: HTMLElement) => label(text, document.createElement("input"), container);
  const mode = select(name + "来源", [["reference", "单元格或名称引用"], ["literal", "图表内的常量"]]), valueType = select(name + "类型", [["text", "文字"], ["number", "数值"]]);
  const expression = input(name + "引用"), origin = select(name + "引用工作表", []), literals = document.createElement("div"); root.append(literals);
  const count = input(name + "总长度", literals), format = input(name + "数字格式", literals), grid = document.createElement("div"), pager = document.createElement("p"); pager.setAttribute("role", "status"); literals.append(grid, pager);
  const button = (text: string, action: () => void, container: HTMLElement = literals, eventSignal = signal) => { const field = document.createElement("button"); field.type = "button"; field.textContent = text; field.addEventListener("click", action, { signal: eventSignal }); container.append(field); return field; };
  const previous = button(name + "上一页", () => { page--; render(); }), next = button(name + "下一页", () => { page++; render(); });
  const add = button(name + "新增数据点", () => {
    const largest = draft.points.reduce((largest, point) => Math.max(largest, /^\d+$/.test(point.position) ? Number(point.position) : 0), 0), position = largest + 1;
    draft.points.push({ position: String(position), value: "", numberFormat: "" });
    if (/^\d+$/.test(draft.count) && Number(draft.count) < position) draft.count = String(position);
    page = Math.floor((draft.points.length - 1) / 100); changed(); render();
  });
  mode.addEventListener("change", () => { draft.mode = mode.value as WorkbookChartVectorDraft["mode"]; changed(); render(); }, { signal });
  valueType.addEventListener("change", () => { draft.valueType = valueType.value as WorkbookChartVectorDraft["valueType"]; changed(); render(); }, { signal });
  origin.addEventListener("change", () => { if (origin.value) draft.originSheetId = origin.value; else delete draft.originSheetId; changed(); }, { signal });
  for (const [field, key] of [[expression, "expression"], [count, "count"], [format, "numberFormat"]] as const) field.addEventListener("input", () => { draft[key] = field.value; changed(); }, { signal });
  function render(): void {
    mode.value = draft.mode; valueType.value = draft.valueType; valueType.hidden = numeric;
    expression.value = draft.expression; count.value = draft.count; format.value = draft.numberFormat; format.disabled = disabled || draft.valueType === "text";
    literals.hidden = draft.mode !== "literal"; expression.hidden = origin.hidden = draft.mode !== "reference"; origin.value = draft.originSheetId ?? "";
    rows.abort(); rows = new AbortController(); grid.replaceChildren();
    page = Math.max(0, Math.min(page, Math.max(0, Math.ceil(draft.points.length / 100) - 1)));
    for (const [offset, point] of draft.points.slice(page * 100, (page + 1) * 100).entries()) {
      const index = page * 100 + offset, row = document.createElement("div"); grid.append(row);
      for (const [title, key] of [["位置", "position"], ["内容", "value"], ["数字格式", "numberFormat"]] as const) {
        if (key === "numberFormat" && draft.valueType === "text") continue;
        const field = input(name + "数据点 " + (index + 1) + title, row); field.value = point[key]; field.disabled = disabled;
        field.addEventListener("input", () => { point[key] = field.value; changed(); }, { signal: rows.signal });
      }
      button(name + "删除数据点 " + (index + 1), () => { draft.points.splice(index, 1); changed(); render(); }, row, rows.signal).disabled = disabled;
    }
    pager.textContent = "已存储 " + draft.points.length + " 个数据点，第 " + (page + 1) + "/" + Math.max(1, Math.ceil(draft.points.length / 100)) + " 页；未列出的位置保持空缺。";
    previous.disabled = disabled || page === 0; next.disabled = disabled || (page + 1) * 100 >= draft.points.length; add.disabled = disabled;
    for (const field of [mode, valueType, expression, origin, count]) field.disabled = disabled;
  }
  return { load(value: WorkbookChartVectorDraft, sheets: readonly OriginSheet[] = []) {
    origin.replaceChildren(...[{ sheetId: "", name: "图表所在工作表" }, ...sheets].map(sheet => { const option = document.createElement("option"); option.value = sheet.sheetId; option.textContent = sheet.name; return option; }));
    draft = value; page = 0; render();
  }, refresh(value: boolean, composing?: HTMLElement) {
    disabled = value; for (const field of root.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLButtonElement>("input,select,button")) field.disabled = disabled && field !== composing;
    format.disabled = (disabled || draft.valueType === "text") && format !== composing; previous.disabled = disabled || page === 0; next.disabled = disabled || (page + 1) * 100 >= draft.points.length;
  } };
}
