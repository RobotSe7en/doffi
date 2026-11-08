import { DocumentError } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, DrawingTextProperties, WorkbookObject, WorkbookSnapshot } from "@doffi/core";
import { reconcileWorkbookInputText, workbookInputTextSourceRange, workbookInputTextView, workbookTextSplice } from "./text-projection.js";

export type WorkbookObjectTextDraftGroup = `object-text:${string}` | "object-font" | "object-paragraph" | "object-body";
interface Host {
  readonly signal: AbortSignal;
  current(): DeepReadonly<WorkbookObject>;
  snapshot(): WorkbookSnapshot;
  perform(operations: readonly DocumentOperation[], groups: readonly WorkbookObjectTextDraftGroup[]): Promise<void>;
  run(action: () => Promise<void>): void;
  changed(group: WorkbookObjectTextDraftGroup): void;
}
interface Controls { load(object: DeepReadonly<WorkbookObject> | undefined, snapshot: WorkbookSnapshot): void; refresh(disabled: boolean, composing?: HTMLElement, stale?: boolean): void }

/** Paragraph input is a private LF projection. Precise source edits and font
 * spans are prepared in the owning core session, including on the Worker path. */
export function createWorkbookObjectTextControls(document: Document, parent: HTMLElement, host: Host): Controls {
  const root = document.createElement("fieldset"), legend = document.createElement("legend"); legend.textContent = "编辑对象段落与选中文字"; root.append(legend); parent.append(root);
  let available = false, selected = "", raw = "", paragraphs = new Map<string, string>(), preferred: string | undefined;
  const drafts = new Map<string, string>();
  const fields: (HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement)[] = [], buttons: HTMLButtonElement[] = [];
  const label = <T extends HTMLElement>(name: string, field: T): T => { const wrapper = document.createElement("label"); field.setAttribute("aria-label", name); wrapper.append(name, field); root.append(wrapper); return field; };
  const input = (name: string, type = "text", group: WorkbookObjectTextDraftGroup = "object-font") => { const field = label(name, document.createElement("input")); field.type = type; fields.push(field); field.addEventListener("input", () => host.changed(group), { signal: host.signal }); return field; };
  const choice = (name: string, options: readonly (readonly [string, string])[], group?: WorkbookObjectTextDraftGroup) => { const field = label(name, document.createElement("select")); for (const [value, text] of options) { const option = document.createElement("option"); option.value = value; option.textContent = text; field.append(option); } fields.push(field); if (group) field.addEventListener("input", () => host.changed(group), { signal: host.signal }); return field; };
  const button = (name: string, action: () => Promise<void>) => { const field = document.createElement("button"); field.type = "button"; field.textContent = name; root.append(field); buttons.push(field); field.addEventListener("click", () => host.run(action), { signal: host.signal }); return field; };
  const paragraph = choice("对象文字段落", []), text = label("所选对象段落文字", document.createElement("textarea")); text.maxLength = 2 * 1024 * 1024; fields.push(text);
  const textGroup = (): WorkbookObjectTextDraftGroup => `object-text:${selected}`;
  text.addEventListener("input", () => { drafts.set(selected, text.value); host.changed(textGroup()); }, { signal: host.signal });
  const latin = input("对象段落西文字体"), eastAsia = input("对象段落中文字体"), size = input("对象段落字号（磅）", "number"), color = input("对象段落文字颜色");
  const weight = choice("对象段落字重", [["", "保持原设置"], ["700", "加粗"], ["400", "常规"]], "object-font"), italic = choice("对象段落斜体", [["", "保持原设置"], ["italic", "斜体"], ["normal", "常规"]], "object-font");
  const underline = choice("对象段落下划线", [["", "保持原设置"], ["underline", "单线"]], "object-font");
  const alignment = choice("对象段落对齐", [["", "保持原设置"], ["left", "左对齐"], ["center", "居中"], ["right", "右对齐"], ["justify", "两端对齐"]], "object-paragraph"), lineHeight = input("对象段落行距倍数", "number", "object-paragraph");
  const wrap = choice("对象文字自动换行", [["", "使用默认换行"], ["true", "自动换行"], ["false", "不换行"]], "object-body"), fit = choice("对象文字自适应", [["", "使用默认自适应"], ["none", "固定文本框"], ["normal", "溢出时缩小文字"], ["shape", "随文字扩展文本框"]], "object-body");
  const scale = input("对象文字缩放比例", "number", "object-body"), reduction = input("对象文字行距缩减比例", "number", "object-body"), bodyEdited = new Set<string>();
  for (const [field, key] of [[wrap, "wrap"], [fit, "autoFit"], [scale, "fontScale"], [reduction, "lineSpacingReduction"]] as const) field.addEventListener("input", () => bodyEdited.add(key), { signal: host.signal });
  const current = () => { const object = host.current(); if (object.kind !== "text" && object.kind !== "geometry" || !paragraphs.has(selected)) throw new DocumentError("InvalidOperation", "请选择当前对象的段落。"); return object; };
  const edit = (nodeId: string): DocumentOperation[] => { const splice = workbookTextSplice(raw, reconcileWorkbookInputText(raw, text.value)); return splice ? [{ op: "replaceWorkbookObjectText", nodeId, paragraphId: selected, ...splice }] : []; };
  button("保存对象段落文字", async () => { const object = current(); await host.perform(edit(object.nodeId), [textGroup()]); });
  const fontProperties = (): Record<string, string> => {
    const properties: Record<string, string> = {}, family = (value: string) => { if (/[\u0000-\u001f\u007f]/.test(value)) throw new DocumentError("InvalidOperation", "字体名称不能包含控制字符。"); return '"' + value.trim().replaceAll("\\", "\\\\").replaceAll('"', '\\"') + '"'; };
    if (latin.value.trim()) { properties["--office-font-ascii"] = family(latin.value); properties["--office-font-h-ansi"] = family(latin.value); }
    if (eastAsia.value.trim()) properties["--office-font-east-asia"] = family(eastAsia.value);
    if (size.validity.badInput) throw new DocumentError("InvalidOperation", "请输入完整字号。");
    if (size.value) { const value = Number(size.value); if (!Number.isFinite(value) || value <= 0 || value > 4000) throw new DocumentError("InvalidOperation", "字号须大于 0 且不超过 4000 磅。"); properties["font-size"] = value + "pt"; }
    for (const [field, key] of [[color, "color"], [weight, "font-weight"], [italic, "font-style"], [underline, "text-decoration-line"]] as const) if (field.value.trim()) properties[key] = field.value.trim();
    return properties;
  };
  button("应用对象选中文字字体", async () => {
    const object = current(), after = reconcileWorkbookInputText(raw, text.value), range = workbookInputTextSourceRange(after, text.selectionStart, text.selectionEnd), properties = fontProperties();
    if (!Object.keys(properties).length) throw new DocumentError("InvalidOperation", "请指定选中文字的字体设置。");
    await host.perform([...edit(object.nodeId), { op: "formatWorkbookObjectText", nodeId: object.nodeId, paragraphId: selected, ...range, expectedText: after.slice(range.from, range.to), properties }], [textGroup(), "object-font"]);
  });
  button("应用对象段落排版", async () => {
    const object = current(), properties: Record<string, string> = {};
    if (alignment.value) properties["text-align"] = alignment.value;
    if (lineHeight.validity.badInput) throw new DocumentError("InvalidOperation", "请输入完整行距。");
    if (lineHeight.value) { const value = Number(lineHeight.value); if (!Number.isFinite(value) || value <= 0 || value > 20) throw new DocumentError("InvalidOperation", "行距倍数须大于 0 且不超过 20。"); properties["line-height"] = String(value); }
    await host.perform([{ op: "setWorkbookObjectParagraphStyle", nodeId: object.nodeId, paragraphId: selected, properties }], ["object-paragraph"]);
  });
  button("应用对象文字换行与自适应", async () => {
    const object = current(), settings: DrawingTextProperties = { ...object.text };
    if (bodyEdited.has("wrap")) { if (!wrap.value) delete settings.wrap; else settings.wrap = wrap.value === "true"; }
    if (bodyEdited.has("autoFit")) { if (!fit.value) delete settings.autoFit; else settings.autoFit = fit.value as "none" | "normal" | "shape"; }
    for (const [field, key] of [[scale, "fontScale"], [reduction, "lineSpacingReduction"]] as const) if (bodyEdited.has(key)) {
      if (field.validity.badInput) throw new DocumentError("InvalidOperation", "请输入完整的文字比例。");
      if (!field.value) delete settings[key]; else { const number = Number(field.value); if (!Number.isFinite(number) || number < 0 || number > 1 || key === "fontScale" && !number) throw new DocumentError("InvalidOperation", "文字缩放须在 (0,1] 内，行距缩减须在 [0,1] 内。"); settings[key] = number; }
    }
    await host.perform([{ op: "updateWorkbookObject", nodeId: object.nodeId, changes: { text: Object.keys(settings).length ? settings : null } }], ["object-body"]);
  });
  const add = button("新增空对象段落", async () => {
    const object = host.current(), previous = new Set(paragraphs.keys());
    if (object.kind !== "text" && object.kind !== "geometry") throw new DocumentError("InvalidOperation", "请选择文本框或形状。");
    await host.perform([...(selected ? edit(object.nodeId) : []), { op: "insertWorkbookObjectParagraph", nodeId: object.nodeId, html: "<p></p>" }], [textGroup()]);
    const template = document.createElement("template"); template.innerHTML = host.snapshot().content?.html ?? "";
    const root = [...template.content.querySelectorAll<HTMLElement>("[data-node-id]")].find(node => node.getAttribute("data-node-id") === object.nodeId);
    preferred = [...root?.children ?? []].map(node => node.getAttribute("data-node-id")!).find(id => !previous.has(id));
  });
  button("删除当前对象段落", async () => { const object = current(); await host.perform([{ op: "removeWorkbookObjectParagraph", nodeId: object.nodeId, paragraphId: selected }], [textGroup(), "object-font", "object-paragraph"]); });
  function show(): void { selected = paragraph.value; raw = paragraphs.get(selected) ?? ""; text.value = drafts.get(selected) ?? workbookInputTextView(raw); }
  paragraph.addEventListener("change", show, { signal: host.signal });
  function load(object: DeepReadonly<WorkbookObject> | undefined, snapshot: WorkbookSnapshot): void {
    available = object?.kind === "text" || object?.kind === "geometry"; root.hidden = !available; if (!available || !object || object.kind !== "text" && object.kind !== "geometry") return; drafts.clear();
    const template = document.createElement("template"); template.innerHTML = snapshot.content?.html ?? "";
    const source = [...template.content.querySelectorAll<HTMLElement>("[data-node-id]")].find(node => node.getAttribute("data-node-id") === object.nodeId);
    paragraphs = new Map([...source?.children ?? []].map(node => {
      const values: string[] = [], pending: Node[] = [node];
      while (pending.length) { const item = pending.pop()!; if (item.nodeType === 3) values.push(item.textContent ?? ""); else if (item.nodeType === 1 && (item as Element).localName === "br") values.push("\n"); else if (item.nodeType === 1 && (item as Element).localName === "img") values.push("\ufffc"); else for (let index = item.childNodes.length - 1; index >= 0; index--) pending.push(item.childNodes[index]!); }
      return [node.getAttribute("data-node-id")!, values.join("")];
    }));
    paragraph.replaceChildren(...[...paragraphs.keys()].map((id, index) => { const option = document.createElement("option"); option.value = id; option.textContent = "段落 " + (index + 1); return option; }));
    paragraph.value = preferred && paragraphs.has(preferred) ? preferred : paragraphs.has(selected) ? selected : paragraphs.keys().next().value ?? ""; preferred = undefined; show();
    for (const field of [latin, eastAsia, size, color, weight, italic, underline, alignment, lineHeight]) field.value = "";
    wrap.value = object.text?.wrap === undefined ? "" : String(object.text.wrap); fit.value = object.text?.autoFit ?? ""; scale.value = object.text?.fontScale === undefined ? "" : String(object.text.fontScale); reduction.value = object.text?.lineSpacingReduction === undefined ? "" : String(object.text.lineSpacingReduction); bodyEdited.clear();
  }
  return { load, refresh(disabled: boolean, composing?: HTMLElement, stale = false): void {
    root.hidden = !available; for (const field of fields) field.disabled = (disabled || !selected) && field !== composing;
    for (const field of buttons) field.disabled = disabled || stale || !selected && field !== add;
  } };
}
