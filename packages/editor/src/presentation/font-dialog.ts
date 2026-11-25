import { attachOfficeFieldList } from "../office-field-list.js";
import { officeSizePresets } from "../react-office/font-presets.js";
import { createOfficeDialog } from "../office-dialog.js";
import type { TextFormatState } from "../text-format-state.js";

/** A font confirmation is one presentation input transaction. */
export function createPresentationFontDialog(parent: ShadowRoot, host: {
  state(): TextFormatState;
  fonts: readonly string[];
  apply(properties: ReadonlyMap<string, string>): void;
  attempt(action: () => void): void;
}): { open(): void; dispose(): void } {
  let active: HTMLDialogElement | undefined;
  const dispose = () => { active?.close(); active?.remove(); active = undefined; };
  return { dispose, open() {
    dispose();
    const document = parent.ownerDocument, ui = createOfficeDialog(document, "字体"), state = host.state();
    const events = new AbortController(); ui.dialog.addEventListener("close", () => events.abort(), { once: true });
    active = ui.dialog; ui.dialog.style.cssText = "width:574px;--office-accent:#b7472a"; ui.content.style.minHeight = "365px";
    const font = ui.tab("字体"), spacing = ui.tab("字符间距"), grid = document.createElement("div"); grid.className = "dialog-grid"; font.append(grid);
    const changed = new Map<string, string>(), names = document.createElement("datalist"); names.id = "ppt-font-" + crypto.randomUUID();
    for (const name of new Set(host.fonts)) { const option = document.createElement("option"); option.value = name; names.append(option); } ui.dialog.append(names);
    const field = (parent: HTMLElement, title: string, property: string, initial = "", type = "text") => {
      const label = document.createElement("label"), input = document.createElement("input"); label.append(title, input); input.type = type; input.value = initial; input.setAttribute("aria-label", title); parent.append(label);
      input.addEventListener("input", () => changed.set(property, input.value)); return input;
    };
    const face = /^\s*(?:"([^"]*)"|'([^']*)'|([^,]+))/.exec(state.fontFamily ?? "");
    const latin = field(grid, "西文字体:", "--office-font-ascii", face?.[1] ?? face?.[2] ?? face?.[3] ?? ""); latin.setAttribute("list", names.id);
    const shapeLabel = document.createElement("label"), shape = document.createElement("select"); shapeLabel.append("字形:", shape); shape.setAttribute("aria-label", "字形"); grid.append(shapeLabel);
    for (const [value, title] of [["normal", "常规"], ["bold", "加粗"], ["italic", "斜体"], ["boldItalic", "加粗斜体"]]) { const option = document.createElement("option"); option.value = value!; option.textContent = title!; shape.append(option); }
    shape.value = state.bold ? state.italic ? "boldItalic" : "bold" : state.italic ? "italic" : "normal";
    shape.addEventListener("change", () => { changed.set("font-weight", shape.value === "bold" || shape.value === "boldItalic" ? "700" : "400"); changed.set("font-style", shape.value === "italic" || shape.value === "boldItalic" ? "italic" : "normal"); });
    const size = field(grid, "字号:", "font-size", state.fontSize === undefined ? "" : String(state.fontSize), "number"); size.min = "1"; size.max = "4000"; size.step = ".1";
    const eastAsian = field(grid, "中文字体:", "--office-font-east-asia", latin.value); eastAsian.setAttribute("list", names.id); eastAsian.parentElement!.style.gridColumn = "1";
    for (const field of [latin, eastAsian]) attachOfficeFieldList(field, [...new Set(host.fonts)], events.signal, true);
    attachOfficeFieldList(size, officeSizePresets.presentation.map(String), events.signal);
    const colors = document.createElement("fieldset"), colorLegend = document.createElement("legend"); colorLegend.textContent = "颜色和下划线:"; colors.append(colorLegend); font.append(colors);
    const color = field(colors, "字体颜色:", "color", /^#[0-9a-f]{6}$/i.test(state.color ?? "") ? state.color : "#000000", "color"); color.style.width = "100%";
    const underLabel = document.createElement("label"), under = document.createElement("select"); underLabel.append("下划线样式:", under); under.setAttribute("aria-label", "下划线样式"); colors.append(underLabel);
    for (const [value, title] of [["none", "(无)"], ["solid", "单线"], ["double", "双线"], ["dotted", "点线"], ["dashed", "虚线"]]) { const option = document.createElement("option"); option.value = value!; option.textContent = title!; under.append(option); }
    under.value = state.underline ? "solid" : "none";
    let strikeEnabled = state.strikethrough ?? false;
    const decoration = () => changed.set("text-decoration-line", [under.value !== "none" ? "underline" : "", strikeEnabled ? "line-through" : ""].filter(Boolean).join(" ") || "none");
    under.addEventListener("change", () => { decoration(); changed.set("text-decoration-style", under.value === "none" ? "solid" : under.value); });
    const effects = document.createElement("fieldset"), effectsLegend = document.createElement("legend"); effectsLegend.textContent = "效果:"; effects.append(effectsLegend); font.append(effects);
    const effect = (title: string, property: string, yes: string, no: string, checked = false, disabled = false) => {
      const label = document.createElement("label"), input = document.createElement("input"); input.type = "checkbox"; input.checked = checked; input.disabled = disabled; label.append(input, title); effects.append(label);
      input.addEventListener("change", () => { if (property === "text-decoration-line") { strikeEnabled = input.checked; decoration(); } else changed.set(property, input.checked ? yes : no); });
    };
    effect("删除线", "text-decoration-line", "line-through", "none", state.strikethrough);
    effect("小型大写字母", "font-variant", "small-caps", "normal");
    effect("双删除线", "text-decoration-line", "line-through", "none", false, true);
    effect("全部大写字母", "text-transform", "uppercase", "none");
    effect("上标", "vertical-align", "super", "baseline", state.verticalAlign === "super");
    effect("等高字符", "font-variant", "normal", "normal", false, true);
    effect("下标", "vertical-align", "sub", "baseline", state.verticalAlign === "sub");
    const interval = field(spacing, "间距（磅）:", "letter-spacing", "0", "number"); interval.min = "-1000"; interval.max = "1000"; interval.step = ".1";
    const kerning = document.createElement("label"), kernBox = document.createElement("input"); kernBox.type = "checkbox"; kernBox.disabled = true; kerning.append(kernBox, "为字体调整字间距"); spacing.append(kerning);
    const cancel = document.createElement("button"), accept = document.createElement("button"); cancel.type = accept.type = "button"; cancel.textContent = "取消"; accept.textContent = "确定"; accept.className = "primary"; ui.footer.append(cancel, accept);
    cancel.addEventListener("click", dispose); ui.dialog.addEventListener("cancel", event => { event.preventDefault(); dispose(); });
    accept.addEventListener("click", () => host.attempt(() => {
      try {
      if (!size.reportValidity() || !interval.reportValidity()) return;
      const properties = new Map(changed);
      for (const name of ["--office-font-ascii", "--office-font-east-asia"]) if (properties.has(name)) properties.set(name, JSON.stringify(properties.get(name)!));
      for (const name of ["font-size", "letter-spacing"]) if (properties.has(name)) properties.set(name, Number(properties.get(name)) + "pt");
      if (properties.size) host.apply(properties);
      dispose();
      } catch (reason) { ui.error(reason); throw reason; }
    }));
    parent.append(ui.dialog); ui.show();
  } };
}
