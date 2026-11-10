import { DocumentError, WorkbookTableStyleSchema } from "@doffi/core";
import type { DocumentOperation, WorkbookTableStyle } from "@doffi/core";
import type { WorkbookControlsHost, WorkbookControlsIdentity } from "./controls-host.js";
import { createWorkbookRuleAppearanceControls } from "./rule-appearance-controls.js";
import { reconcileWorkbookInputText } from "./text-projection.js";

const regions = [
  ["wholeTable", "整个表"], ["headerRow", "标题行"], ["totalRow", "汇总行"], ["firstColumn", "首列"], ["lastColumn", "末列"],
  ["firstRowStripe", "第一行条纹"], ["secondRowStripe", "第二行条纹"], ["firstColumnStripe", "第一列条纹"], ["secondColumnStripe", "第二列条纹"],
  ["firstHeaderCell", "标题首格"], ["lastHeaderCell", "标题末格"], ["firstTotalCell", "汇总首格"], ["lastTotalCell", "汇总末格"],
] as const;
type Region = typeof regions[number][0];
interface StyleDraft { identity: WorkbookControlsIdentity; style: WorkbookTableStyle; existing: boolean; dirty: boolean }
export function createWorkbookTableStyleControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void; hasPendingInput(): boolean } {
  const panel = document.createElement("details"), summary = document.createElement("summary"); summary.textContent = "自定义智能表样式"; panel.append(summary); parent.append(panel);
  const fields: (HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement)[] = [];
  function field<T extends HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(name: string, input: T): T {
    const wrapper = document.createElement("label"); wrapper.style.cssText = "display:flex;flex-wrap:wrap;gap:4px;align-items:center;max-width:100%";
    input.setAttribute("aria-label", name); input.style.cssText = "max-width:100%;min-width:0"; wrapper.append(name, input); panel.append(wrapper); fields.push(input); return input;
  }
  const choice = field("选择自定义智能表样式", document.createElement("select")), name = field("自定义智能表样式名称", document.createElement("textarea")); name.maxLength = 255; name.rows = 1;
  const useDefault = field("新建智能表默认使用此样式", document.createElement("input")); useDefault.type = "checkbox";
  const region = field("选择智能表样式区域", document.createElement("select")), enabled = field("启用智能表样式区域", document.createElement("input")); enabled.type = "checkbox";
  const size = field("智能表样式条纹大小", document.createElement("input")); size.type = "number"; size.min = "1"; size.max = "9"; size.step = "1";
  const format = field("智能表区域数字格式", document.createElement("select"));
  const option = (select: HTMLSelectElement, value: string, label: string) => { const element = document.createElement("option"); element.value = value; element.textContent = label; select.append(element); };
  for (const [value, label] of regions) option(region, value, label);
  const appearance = createWorkbookRuleAppearanceControls(document, panel, host.signal, "表样式区域", true, () => changed(true));
  const status = document.createElement("p"); status.setAttribute("role", "status"); status.setAttribute("aria-label", "自定义智能表样式编辑状态"); panel.append(status);
  let draft: StyleDraft | undefined, selected = "", areaDirty = false, composing: HTMLElement | undefined, saving = false, serial = 0, listKey = "", formatKey = "";
  const styles = () => host.snapshot()?.metadata.workbook.tableStyles ?? [];
  const stale = () => !!draft && (host.snapshot()?.manifest.documentId !== draft.identity.documentId || host.snapshot()?.manifest.revision !== draft.identity.revision);
  const fail = (message: string): never => { throw new DocumentError("InvalidOperation", message); };
  function fresh(): string { const used = new Set(styles().map(style => style.styleId)); let id: string; do { id = "editor-table-style-" + ++serial; } while (used.has(id)); return id; }
  function loadArea(): void {
    const snapshot = host.snapshot(); if (!snapshot || !draft) return;
    const element = draft.style.elements.find(value => value.region === region.value); enabled.checked = !!element; size.value = String(element?.size ?? 1); format.value = "";
    appearance.load(snapshot, element?.style?.cssClass); areaDirty = false;
  }
  function load(): void {
    const snapshot = host.snapshot(); draft = undefined; if (!snapshot) return;
    const original = styles().find(style => style.styleId === selected);
    draft = { identity: { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision }, existing: !!original, dirty: false,
      style: original ? JSON.parse(JSON.stringify(original)) as WorkbookTableStyle : { styleId: fresh(), name: "", pivotCompatible: false, elements: [] } };
    name.value = draft.style.name; const defaultStyle = snapshot.metadata.workbook.defaultTableStyle;
    useDefault.checked = defaultStyle?.kind === "custom" && defaultStyle.styleId === draft.style.styleId; loadArea();
  }
  function changed(area = false): void { if (!draft) return; draft.dirty = true; areaDirty ||= area; refresh(); host.draftChanged?.(); }
  function defaultOperation(styleId: string): DocumentOperation[] {
    const previous = host.snapshot()!.metadata.workbook.defaultTableStyle;
    if (useDefault.checked) return [{ op: "setWorkbookSettings", changes: { defaultTableStyle: { kind: "custom", styleId } } }];
    return previous?.kind === "custom" && previous.styleId === styleId ? [{ op: "setWorkbookSettings", changes: { defaultTableStyle: null } }] : [];
  }
  function button(label: string, action: () => Promise<void>): HTMLButtonElement {
    const button = document.createElement("button"); button.type = "button"; button.textContent = label; panel.append(button);
    button.addEventListener("click", () => host.run(async () => {
      if (!host.idle() || composing || saving) fail("请先完成当前输入或操作。"); saving = true;
      try { const work = action(); refresh(); host.draftChanged?.(); await work; } finally { saving = false; refresh(); host.draftChanged?.(); }
    }), { signal: host.signal }); return button;
  }
  const save = button("保存自定义智能表样式", async () => {
    if (!draft || stale()) throw new DocumentError("RevisionConflict", "样式草稿属于旧修订，请重新读取。");
    const submitted = draft, style: WorkbookTableStyle = JSON.parse(JSON.stringify(submitted.style)) as WorkbookTableStyle;
    style.name = reconcileWorkbookInputText(style.name, name.value); let element = style.elements.find(element => element.region === region.value);
    if (!enabled.checked) style.elements = style.elements.filter(element => element.region !== region.value);
    else {
      if (!element) { element = { region: region.value as Region, size: 1 }; style.elements.push(element); }
      element.size = element.region.endsWith("Stripe") ? Number(size.value) : 1;
    }
    const operations: DocumentOperation[] = [{ op: "setWorkbookTableStyle", style: WorkbookTableStyleSchema.parse(style) }];
    if (enabled.checked) {
      const properties = appearance.properties(), numberFormatId = format.value === "clear" ? null : format.value.startsWith("format:") ? format.value.slice(7) : undefined;
      if (Object.keys(properties).length || numberFormatId !== undefined) operations.push({ op: "formatWorkbookTableStyleElement", styleId: style.styleId, region: region.value as Region, properties, ...(numberFormatId !== undefined ? { numberFormatId } : {}) });
    }
    operations.push(...defaultOperation(style.styleId));
    await host.perform(operations, undefined, submitted.identity); selected = style.styleId; draft = undefined; listKey = "";
  });
  const remove = button("删除自定义智能表样式", async () => {
    if (!draft?.existing || stale()) throw new DocumentError("RevisionConflict", "请重新读取要删除的样式。");
    const operations = defaultOperation(draft.style.styleId);
    await host.perform([...operations, { op: "removeWorkbookTableStyle", styleId: draft.style.styleId }], undefined, draft.identity); selected = ""; draft = undefined; listKey = "";
  });
  const reload = button("放弃表样式草稿并重新读取", async () => { selected = styles().some(style => style.styleId === selected) ? selected : ""; draft = undefined; listKey = ""; refresh(); });
  choice.addEventListener("change", () => { if (draft?.dirty || composing || saving) { choice.value = selected; return; } selected = choice.value; load(); refresh(); host.draftChanged?.(); }, { signal: host.signal });
  region.addEventListener("change", () => { loadArea(); refresh(); }, { signal: host.signal });
  for (const input of fields.filter(input => input !== choice && input !== region)) input.addEventListener("input", () => changed(input !== name && input !== useDefault), { signal: host.signal });
  panel.addEventListener("compositionstart", event => { if (event.target instanceof document.defaultView!.HTMLElement) { composing = event.target; changed(event.target !== name); } }, { signal: host.signal });
  panel.addEventListener("compositionend", () => { composing = undefined; refresh(); host.draftChanged?.(); }, { signal: host.signal });
  function refresh(): void {
    const snapshot = host.snapshot(), all = styles(), next = JSON.stringify([snapshot?.manifest.documentId, all.map(style => [style.styleId, style.name])]);
    if (listKey !== next) {
      listKey = next; choice.replaceChildren(); option(choice, "", "新建自定义智能表样式"); for (const value of all) option(choice, value.styleId, value.name);
      if (selected && !all.some(style => style.styleId === selected)) option(choice, selected, "已移除的样式 · 草稿保留"); choice.value = selected;
    }
    const formats = snapshot?.metadata.workbook.numberFormats ?? [], nextFormats = JSON.stringify(formats);
    if (formatKey !== nextFormats) {
      const previous = format.value; formatKey = nextFormats; format.replaceChildren(); option(format, "", "保留来源数字格式"); option(format, "clear", "清除区域数字格式");
      for (const value of formats) option(format, "format:" + value.formatId, "code" in value ? [...value.code.slice(0, 96)].slice(0, 48).join("") : "内置数字格式 " + value.builtinId); format.value = previous;
    }
    if (!saving && !composing && (!draft || !draft.dirty && stale())) load();
    const blocked = !host.idle() || saving || !!composing, old = stale();
    for (const input of fields) input.disabled = (blocked || !draft) && input !== composing;
    choice.disabled = blocked || !!draft?.dirty; region.disabled = blocked || areaDirty; size.disabled ||= !enabled.checked || !region.value.endsWith("Stripe"); format.disabled ||= !enabled.checked;
    appearance.refresh(blocked || !draft || !enabled.checked, composing); save.disabled = blocked || !draft || old;
    const references = snapshot?.metadata.workbook.sheets.reduce((count, sheet) => count + (sheet.tables ?? []).filter(table => table.style?.kind === "custom" && table.style.styleId === draft?.style.styleId).length, 0) ?? 0;
    remove.disabled = save.disabled || !draft?.existing || references > 0 || useDefault.checked; reload.disabled = blocked;
    status.textContent = old ? "文档已变化。样式草稿已保留，请重新读取后编辑。" : !draft ? "正在读取样式。" : `${references} 个智能表引用此样式。${references ? "删除前请先更换这些表的样式。" : ""}${areaDirty ? "保存或放弃当前区域修改后，可切换样式区域。" : ""}${draft.dirty ? "草稿尚未保存。" : ""}`;
  }
  refresh(); return { element: panel, refresh, hasPendingInput: () => !!draft?.dirty || !!composing || saving };
}
