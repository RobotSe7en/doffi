import { DEFAULT_DOCUMENT_READ_LIMITS, resolveSourceCssColor, resolveSourceCssLength, SourceStyleResolver } from "@doffi/core";
import type { WorkbookSnapshot } from "@doffi/core";

/** Overrides are explicit. An untouched theme/variable/role declaration is not
 * rewritten because a native color or font input needs a display fallback. */
export function createWorkbookRuleAppearanceControls(document: Document, parent: ParentNode, signal: AbortSignal, label: string, fonts: boolean, changed: () => void): { readonly element: HTMLDivElement; load(snapshot: WorkbookSnapshot, cssClass: string | undefined, newFill?: string): void; properties(): Record<string, string>; refresh(blocked: boolean, composing?: HTMLElement): void } {
  const element = document.createElement("div"); parent.append(element);
  const fields: (HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement)[] = [], pairs: { enabled: HTMLInputElement; value: HTMLInputElement | HTMLTextAreaElement }[] = [];
  const field = <T extends HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(name: string, input: T): T => { const wrapper = document.createElement("label"); wrapper.style.cssText = "display:flex;flex-wrap:wrap;align-items:center;gap:4px;max-width:100%"; input.style.maxWidth = "100%"; input.style.minWidth = "0"; input.setAttribute("aria-label", label + name); wrapper.append(name, input); element.append(wrapper); fields.push(input); input.addEventListener("input", changed, { signal }); return input; };
  const pair = (name: string, type: "color" | "text" | "number") => {
    const enabled = field("修改" + name, document.createElement("input")); enabled.type = "checkbox";
    const value = type === "text" ? field(name, document.createElement("textarea")) : field(name, document.createElement("input"));
    if (value instanceof document.defaultView!.HTMLInputElement) value.type = type;
    else value.maxLength = 1024;
    pairs.push({ enabled, value }); return { enabled, value };
  };
  const fill = pair("填充颜色", "color"), foreground = fonts ? pair("文字颜色", "color") : undefined, latin = fonts ? pair("西文字体", "text") : undefined, eastAsian = fonts ? pair("中文字体", "text") : undefined, size = fonts ? pair("字号", "number") : undefined;
  if (size) { size.value.setAttribute("min", "1"); size.value.setAttribute("max", "409"); size.value.setAttribute("step", "0.5"); }
  const weight = fonts ? field("粗体", document.createElement("select")) : undefined;
  if (weight) for (const [value, label] of [["", "保留来源"], ["normal", "常规"], ["bold", "粗体"]]) { const option = document.createElement("option"); option.value = value!; option.textContent = label!; weight.append(option); }
  const styles = fonts ? [
    { property: "font-style", name: "斜体", options: [["normal", "常规"], ["italic", "斜体"]] },
    { property: "--office-font-strike", name: "删除线", options: [["false", "无"], ["true", "删除线"]] },
    { property: "--office-font-underline", name: "下划线", options: [["none", "无"], ["single", "单线"], ["double", "双线"], ["singleAccounting", "财务单线"], ["doubleAccounting", "财务双线"]] },
  ].map(source => { const input = field(source.name, document.createElement("select")); for (const [value, label] of [["", "保留来源"], ...source.options]) { const option = document.createElement("option"); option.value = value!; option.textContent = label!; input.append(option); } return { ...source, input }; }) : [];
  const color = (source: string | undefined, fallback: string) => { try { const value = source ? resolveSourceCssColor(source) : undefined; return value ? "#" + value : fallback; } catch { return fallback; } };
  function load(snapshot: WorkbookSnapshot, cssClass: string | undefined, newFill?: string): void {
    const style = cssClass ? new SourceStyleResolver(snapshot.styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set()).classStyle(cssClass) : {};
    for (const pair of pairs) pair.enabled.checked = false;
    fill.value.value = newFill ?? color(style[fonts ? "background-color" : "--office-scale-color"], "#ffeb9c"); fill.enabled.checked = newFill !== undefined;
    if (foreground) foreground.value.value = color(style.color, "#000000");
    if (latin) latin.value.value = style["--office-font-ascii"] ?? style["font-family"] ?? "";
    if (eastAsian) eastAsian.value.value = style["--office-font-east-asia"] ?? style["font-family"] ?? "";
    if (size) { let points = 12; try { if (style["font-size"]) points = resolveSourceCssLength(style["font-size"], { font: 12, percent: 12 }); } catch { /* Preserve the original unless explicitly overridden. */ } size.value.value = String(points); }
    if (weight) weight.value = "";
    for (const style of styles) style.input.value = "";
  }
  return { element, load, properties(): Record<string, string> {
    const properties: Record<string, string> = {};
    if (fill.enabled.checked) properties["background-color"] = fill.value.value;
    if (foreground?.enabled.checked) properties.color = foreground.value.value;
    if (latin?.enabled.checked) { properties["--office-font-ascii"] = latin.value.value; properties["--office-font-h-ansi"] = latin.value.value; }
    if (eastAsian?.enabled.checked) properties["--office-font-east-asia"] = eastAsian.value.value;
    if (size?.enabled.checked) properties["font-size"] = size.value.value + "pt";
    if (weight?.value) properties["font-weight"] = weight.value;
    for (const style of styles) if (style.input.value) properties[style.property] = style.input.value;
    return properties;
  }, refresh(blocked: boolean, composing?: HTMLElement): void {
    for (const field of fields) field.disabled = blocked && field !== composing;
    for (const pair of pairs) pair.value.disabled ||= !pair.enabled.checked;
  } };
}
