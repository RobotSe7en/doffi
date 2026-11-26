import { attachOfficeFieldList } from "./office-field-list.js";
import { createOfficeDialog } from "./office-dialog.js";
import { officeSizeLabel, officeSizePresets, wordSizeNames } from "./react-office/font-presets.js";
import type { TextFormatState } from "./text-format-state.js";
/** Font roles and language are authored values; installed preview substitutes
 * are never used to rewrite the source faces. */
export function createWordFontControls(
  parent: ShadowRoot,
  before: HTMLElement,
  read: (paragraphMark?: boolean) => TextFormatState,
  apply: (properties: ReadonlyMap<string, string>, paragraphMark?: boolean) => void,
  onError: (error: unknown) => void,
  availableFonts: readonly string[] = [],
): {
  open(): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let panel: HTMLElement | undefined, controller: AbortController | undefined;
  const close = (): void => {
    controller?.abort();
    if (panel instanceof HTMLDialogElement && panel.open) panel.close();
    panel?.remove();
    panel = undefined;
  };
  function open(): void {
    close();
    controller = new AbortController();
    const signal = controller.signal,
      state = read();
    const ui = createOfficeDialog(document, "字体");
    ui.dialog.style.width = "675px";
    ui.content.style.minHeight = "574px";
    panel = ui.dialog;
    panel.addEventListener("cancel", event => { event.preventDefault(); close(); }, { signal });
    const fontTab = ui.tab("字体"), advancedTab = ui.tab("高级");
    const fontGrid = document.createElement("div"), effects = document.createElement("fieldset"), preview = document.createElement("div");
    fontGrid.className = "dialog-grid"; fontGrid.style.gridTemplateColumns = "minmax(120px,224px) 124px 70px"; effects.innerHTML = "<legend>效果</legend>";
    preview.className = "dialog-preview"; preview.textContent = "微软卓越　AaBbCc";
    fontTab.style.cssText = "display:flex;flex-direction:column;min-height:530px"; fontTab.append(fontGrid, effects, preview); advancedTab.className = "dialog-options";
    let target: HTMLElement = fontGrid;
    const changed = new Set<HTMLInputElement | HTMLSelectElement>();
    const allFields: (HTMLInputElement | HTMLSelectElement)[] = [];
    parent.insertBefore(panel, before);
    const button = (label: string, run: () => void): void => {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = label;
      button.addEventListener(
        "click",
        () => {
          try {
            run();
          } catch (error) {
            ui.error(error);
            onError(error);
          }
        },
        { signal },
      );
      ui.footer.append(button);
    };
    const field = (name: string, type = "text"): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.type = type;
      input.setAttribute("aria-label", name);
      input.placeholder = "";
      allFields.push(input);
      input.addEventListener("input", () => changed.add(input), { signal });
      input.addEventListener("change", () => changed.add(input), { signal });
      if (type === "number") input.step = ".5";
      label.append(input);
      target.append(label);
      return input;
    };
    const choose = (
      name: string,
      options: readonly (readonly [string, string])[],
    ): HTMLSelectElement => {
      const label = document.createElement("label"),
        input = document.createElement("select");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      allFields.push(input);
      input.addEventListener("change", () => changed.add(input), { signal });
      for (const [value, text] of [["", "保留"], ...options]) {
        const option = document.createElement("option");
        option.value = value!;
        option.textContent = text!;
        input.append(option);
      }
      label.append(input);
      target.append(label);
      return input;
    };
    button("取消", close);
    const family = field("统一字体"),
      size = field("字符字号（磅）"),
      roles = (["ascii", "hAnsi", "eastAsia", "cs"] as const).map((role, index) => ({
        role,
        input: field(["ASCII 字体", "高位 ANSI 字体", "东亚字体", "复杂文字字体"][index]!),
      }));
    family.parentElement!.firstChild!.textContent = "中文字体:";
    family.parentElement!.style.gridColumn = "1";
    roles[0]!.input.parentElement!.firstChild!.textContent = "西文字体:";
    roles[0]!.input.parentElement!.style.gridColumn = "1";
    for (const role of [roles[1]!, roles[3]!]) advancedTab.append(role.input.parentElement!);
    family.parentElement!.hidden = true;
    roles[2]!.input.parentElement!.firstChild!.textContent = "中文字体:";
    roles[2]!.input.parentElement!.style.cssText = "grid-column:1;grid-row:1";
    roles[0]!.input.parentElement!.style.cssText = "grid-column:1;grid-row:2"; roles[0]!.input.placeholder = "(使用中文字体)";
    size.parentElement!.firstChild!.textContent = "字号:";
    size.parentElement!.style.gridColumn = "3";
    size.parentElement!.style.gridRow = "1";
    const currentFamily = /^\s*(?:"([^"]*)"|'([^']*)'|([^,]+))/.exec(state.fontFamily ?? "");
    family.value = currentFamily?.[1] ?? currentFamily?.[2] ?? currentFamily?.[3] ?? "";
    roles[2]!.input.value = family.value;
    size.value = officeSizeLabel("document", state.fontSize);
    const families = [...new Set([...availableFonts, "宋体", "等线", "微软雅黑", "黑体", "仿宋", "楷体", "Arial", "Calibri", "Aptos", "Times New Roman", "Noto Serif CJK SC", "Noto Sans CJK SC"])];
    const fontList = document.createElement("datalist"); fontList.id = "office-font-" + crypto.randomUUID();
    for (const name of families) { const option = document.createElement("option"); option.value = name; fontList.append(option); }
    panel.append(fontList); family.setAttribute("list", fontList.id); for (const role of roles) role.input.setAttribute("list", fontList.id);
    const sizeList = document.createElement("datalist"); sizeList.id = "office-size-" + crypto.randomUUID();
    for (const value of [...Object.keys(wordSizeNames), ...officeSizePresets.document.map(String)]) { const option = document.createElement("option"); option.value = value; sizeList.append(option); }
    panel.append(sizeList); size.setAttribute("list", sizeList.id);
    for (const role of roles) attachOfficeFieldList(role.input, families, signal, true);
    attachOfficeFieldList(size, [...Object.keys(wordSizeNames), ...officeSizePresets.document.map(String)], signal);
    const weight = choose("字符字重", [
        ["400", "常规"],
        ["700", "加粗"],
      ]),
      italic = choose("字符倾斜", [
        ["normal", "正体"],
        ["italic", "斜体"],
      ]),
      underline = choose("字符下划线", [
        ["none", "无"],
        ["solid", "单线"],
        ["words", "仅字下划线"],
        ["double", "双线"],
        ["thick", "粗线"],
        ["dotted", "点线"],
        ["dottedHeavy", "粗点线"],
        ["dashed", "虚线"],
        ["dashedHeavy", "粗虚线"],
        ["dashLong", "长虚线"],
        ["dashLongHeavy", "粗长虚线"],
        ["dotDash", "点划线"],
        ["dashDotHeavy", "粗点划线"],
        ["dotDotDash", "双点划线"],
        ["dashDotDotHeavy", "粗双点划线"],
        ["wavy", "波浪线"],
        ["wavyHeavy", "粗波浪线"],
        ["wavyDouble", "双波浪线"],
      ]),
      strike = choose("字符删除线", [
        ["none", "无"],
        ["line-through", "删除线"],
      ]),
      vertical = choose("字符位置", [
        ["baseline", "基线"],
        ["super", "上标"],
        ["sub", "下标"],
      ]),
      caps = choose("字符大小写", [
        ["none", "正常"],
        ["uppercase", "全部大写"],
        ["small-caps", "小型大写"],
      ]);
    for (const field of [strike, vertical, caps, italic, weight]) field.parentElement!.hidden = true;
    const shapeLabel = document.createElement("label"), shape = document.createElement("select");
    shapeLabel.append("字形:", shape); shapeLabel.style.cssText = "grid-column:2;grid-row:1";
    for (const [value, name] of [["normal", "常规"], ["bold", "加粗"], ["italic", "斜体"], ["boldItalic", "加粗斜体"]]) { const option = document.createElement("option"); option.value = value!; option.textContent = name!; shape.append(option); }
    shape.value = state.bold ? state.italic ? "boldItalic" : "bold" : state.italic ? "italic" : "normal";
    shape.addEventListener("change", () => { weight.value = shape.value === "bold" || shape.value === "boldItalic" ? "700" : "400"; italic.value = shape.value === "italic" || shape.value === "boldItalic" ? "italic" : "normal"; changed.add(weight); changed.add(italic); }, { signal });
    fontGrid.append(shapeLabel);
    const native = state.characterProperties ?? {};
    const effect = (
      title: string,
      field: HTMLSelectElement,
      value: string,
      clear: string,
      checked = false,
      onToggle?: (isChecked: boolean) => void,
    ): HTMLInputElement => {
      const label = document.createElement("label"), checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = checked;
      checkbox.setAttribute("aria-label", title);
      label.append(checkbox, title);
      checkbox.addEventListener(
        "change",
        () => {
          field.value = checkbox.checked ? value : clear;
          changed.add(field);
          onToggle?.(checkbox.checked);
        },
        { signal },
      );
      effects.append(label);
      return checkbox;
    };
    effects.style.cssText = "grid-template-columns:1fr 1fr;gap:4px 40px;margin-top:22px";
    const doubleStrike = document.createElement("label"), doubleStrikeBox = document.createElement("input");
    doubleStrikeBox.type = "checkbox";
    doubleStrikeBox.setAttribute("aria-label", "双删除线");
    doubleStrikeBox.checked = state.doubleStrikethrough === true || native["--word-dstrike"] === "1";
    doubleStrike.append(doubleStrikeBox, "双删除线");
    const strikeBox = effect("删除线", strike, "line-through", "none", state.strikethrough, (checked) => {
      if (checked && doubleStrikeBox.checked) {
        doubleStrikeBox.checked = false;
        doubleStrikeBox.dataset.changed = "1";
      }
    });
    const smallCapsBox = effect("小型大写字母", caps, "small-caps", "none", native["font-variant"] === "small-caps");
    doubleStrikeBox.addEventListener(
      "change",
      () => {
        doubleStrikeBox.dataset.changed = "1";
        if (doubleStrikeBox.checked && strikeBox.checked) {
          strikeBox.checked = false;
          strike.value = "none";
          changed.add(strike);
        }
      },
      { signal },
    );
    effects.append(doubleStrike);
    const allCapsBox = effect("全部大写字母", caps, "uppercase", "none", native["text-transform"] === "uppercase");
    smallCapsBox.addEventListener(
      "change",
      () => {
        if (smallCapsBox.checked) allCapsBox.checked = false;
      },
      { signal },
    );
    allCapsBox.addEventListener(
      "change",
      () => {
        if (allCapsBox.checked) smallCapsBox.checked = false;
      },
      { signal },
    );
    const superBox = effect("上标", vertical, "super", "baseline", state.verticalAlign === "super");
    const hiddenLabel = document.createElement("label"), hiddenBox = document.createElement("input");
    hiddenBox.type = "checkbox";
    hiddenBox.setAttribute("aria-label", "隐藏");
    hiddenBox.checked = native["--word-run-native-vanish"] === "1";
    hiddenBox.addEventListener("change", () => { hiddenBox.dataset.changed = "1"; }, { signal });
    hiddenLabel.append(hiddenBox, "隐藏");
    effects.append(hiddenLabel);
    const subBox = effect("下标", vertical, "sub", "baseline", state.verticalAlign === "sub");
    superBox.addEventListener("change", () => { if (superBox.checked) subBox.checked = false; }, { signal });
    subBox.addEventListener("change", () => { if (subBox.checked) superBox.checked = false; }, { signal });
    weight.parentElement!.firstChild!.textContent = "字形:";
    weight.parentElement!.style.gridColumn = "2"; weight.parentElement!.style.gridRow = "1";
    weight.value = state.bold === undefined ? "" : state.bold ? "700" : "400";
    italic.value = state.italic === undefined ? "" : state.italic ? "italic" : "normal";
    italic.parentElement!.firstChild!.textContent = "倾斜:";
    const rawUnderline = native["--word-underline-style"];
    const normalizedUnderline = rawUnderline
      ? rawUnderline === "single"
        ? "solid"
        : rawUnderline === "dash"
          ? "dashed"
          : rawUnderline === "wave"
            ? "wavy"
            : rawUnderline
      : undefined;
    underline.value =
      state.underline === undefined
        ? ""
        : state.underline
          ? (normalizedUnderline && normalizedUnderline !== "none"
              ? normalizedUnderline
              : (native["text-decoration-style"] ?? "solid"))
          : "none";
    underline.parentElement!.firstChild!.textContent = "下划线线型:";
    target = advancedTab;
    const spacing = field("字符间距（磅）", "number"),
      color = field("字符颜色"),
      background = field("字符底纹"),
      complexSize = field("复杂文字字号（磅）", "number"),
      complexBold = choose("复杂文字字重", [
        ["0", "常规"],
        ["1", "加粗"],
      ]),
      complexItalic = choose("复杂文字倾斜", [
        ["0", "正体"],
        ["1", "斜体"],
      ]),
      kern = field("字距调整起始字号（磅）", "number");
    color.parentElement!.firstChild!.textContent = "字体颜色:";
    background.parentElement!.firstChild!.textContent = "文字底纹:";
    const colors = document.createElement("fieldset"), colorsLegend = document.createElement("legend"); colorsLegend.textContent = "颜色和下划线"; colors.style.cssText = "grid-template-columns:1fr 1fr 1fr 1fr;margin-top:20px";
    const underlineColorInput = field("下划线颜色"), underlineColor = underlineColorInput.parentElement!;
    underlineColorInput.placeholder = "自动";
    underlineColorInput.value = state.characterProperties?.["text-decoration-color"] ?? "";
    underlineColorInput.addEventListener("input", () => underlineColorInput.setCustomValidity(""), { signal });
    const emphasisInput = choose("着重号", [
      ["none", "(无)"],
      ["dot", "圆点（下）"],
      ["comma", "顿号（上）"],
      ["circle", "空心圆（上）"],
      ["underDot", "下圆点"],
    ]);
    emphasisInput.parentElement!.firstChild!.textContent = "着重号:";
    emphasisInput.value = native["--word-run-native-em"] ?? "";
    const emphasis = emphasisInput.parentElement!;
    const rgb = /^rgb\(\s*(\d+)[, ]+\s*(\d+)[, ]+\s*(\d+)\s*\)$/.exec(state.color ?? "");
    color.value = /^#[0-9a-f]{6}$/i.test(state.color ?? "") ? state.color! : rgb ? "#" + rgb.slice(1).map(value => Number(value).toString(16).padStart(2, "0")).join("") : "#000000"; color.type = "color"; color.style.width = "100%";
    colors.append(colorsLegend, color.parentElement!, underline.parentElement!, underlineColor, emphasis);
    fontTab.insertBefore(colors, effects);
    const previewLabel = document.createElement("div"); previewLabel.textContent = "预览"; previewLabel.style.cssText = "border-top:1px solid #ddd;padding-top:10px;margin-top:auto;font-weight:600";
    fontTab.insertBefore(previewLabel, preview);
    const updatePreview = () => {
      const cssUnderlineStyle = (value: string) =>
        value === "double"
          ? "double"
          : value === "dotted" || value === "dottedHeavy"
            ? "dotted"
            : value === "wavy" || value === "wavyHeavy" || value === "wavyDouble"
              ? "wavy"
              : [
                    "dashed",
                    "dashedHeavy",
                    "dashLong",
                    "dashLongHeavy",
                    "dotDash",
                    "dashDotHeavy",
                    "dotDotDash",
                    "dashDotDotHeavy",
                  ].includes(value)
                ? "dashed"
                : "solid";
      for (const box of [preview, advancedPreview]) {
        if (!box) continue;
        box.style.fontFamily = JSON.stringify(roles[2]!.input.value || roles[0]!.input.value);
        const pointSize = wordSizeNames[size.value] ?? Number(size.value);
        if (pointSize) box.style.fontSize = Math.min(pointSize, 40) + "pt";
        box.style.fontWeight = weight.value;
        box.style.fontStyle = italic.value;
        box.style.color = color.value;
        box.style.textTransform = allCapsBox.checked ? "uppercase" : "none";
        box.style.fontVariant = smallCapsBox.checked ? "small-caps" : "normal";
        box.style.textDecoration = [
          underline.value && underline.value !== "none" ? "underline" : "",
          strike.value === "line-through" ? "line-through" : "",
        ]
          .filter(Boolean)
          .join(" ");
        box.style.textDecorationStyle =
          underline.value && underline.value !== "none" ? cssUnderlineStyle(underline.value) : "solid";
        box.style.textDecorationColor = underlineColorInput.value.trim() || "currentColor";
        box.style.backgroundImage = doubleStrikeBox.checked
          ? "linear-gradient(to bottom, transparent 40%, currentColor 40%, currentColor 46%, transparent 46%, transparent 56%, currentColor 56%, currentColor 62%, transparent 62%)"
          : "";
        box.style.opacity = hiddenBox.checked ? "0.45" : "1";
        if (emphasisInput.value === "dot" || emphasisInput.value === "underDot") {
          box.style.setProperty("text-emphasis-style", "filled dot");
          box.style.setProperty("text-emphasis-position", "under right");
        } else if (emphasisInput.value === "comma") {
          box.style.setProperty("text-emphasis-style", "filled sesame");
          box.style.setProperty("text-emphasis-position", "over right");
        } else if (emphasisInput.value === "circle") {
          box.style.setProperty("text-emphasis-style", "open circle");
          box.style.setProperty("text-emphasis-position", "over right");
        } else {
          box.style.setProperty("text-emphasis-style", "none");
        }
      }
    };
    const mark = field("仅设置段落标记", "checkbox");
    complexSize.min = ".5";
    kern.min = "0";
    const languages = (["val", "eastAsia", "bidi"] as const).map((role, index) => ({
      role,
      input: field(["西文语言", "东亚语言", "复杂文字语言"][index]!),
    }));
    const readName = (value: string) => { const match = /^\s*(?:"([^"]*)"|'([^']*)'|([^,]+))/.exec(value); return match?.[1] ?? match?.[2] ?? match?.[3] ?? ""; };
    for (const { role, input } of roles) if (native["--word-font-" + role]) input.value = readName(native["--word-font-" + role]!);
    for (const { role, input } of languages) if (native["--word-language-" + role]) input.value = readName(native["--word-language-" + role]!);
    const points = (value: string | undefined) => value && Number.isFinite(parseFloat(value)) ? String(parseFloat(value) * (value.endsWith("px") ? .75 : 1)) : "";
    spacing.value = points(native["letter-spacing"]); kern.value = points(native["--word-kerning-threshold"]); complexSize.value = points(native["--word-font-size-cs"]);
    complexBold.value = native["--word-bold-cs"] ?? ""; complexItalic.value = native["--word-italic-cs"] ?? "";
    const extra = document.createElement("details"), extraTitle = document.createElement("summary"), extraFields = document.createElement("div");
    extraTitle.textContent = "其他语言设置"; extraFields.className = "dialog-options"; extra.append(extraTitle, extraFields);
    for (const node of [...advancedTab.children]) if (node !== spacing.parentElement && node !== kern.parentElement) extraFields.append(node);
    const intervalGroup = document.createElement("fieldset"), intervalTitle = document.createElement("legend"); intervalTitle.textContent = "字符间距"; intervalGroup.append(intervalTitle);
    const scale = field("字符缩放（百分比）", "number");
    scale.min = "1"; scale.max = "600"; scale.step = "1";
    scale.value = native["--word-run-native-w"] ?? (state.hasSelection ? "" : "100");
    scale.parentElement!.firstChild!.textContent = "缩放（%）:";
    intervalGroup.append(scale.parentElement!);
    const intervalLabel = document.createElement("label"), interval = document.createElement("select"); intervalLabel.append("间距:", interval); intervalGroup.append(intervalLabel);
    for (const [value, title] of [["normal", "标准"], ["expanded", "加宽"], ["condensed", "紧缩"]]) { const option = document.createElement("option"); option.value = value!; option.textContent = title!; interval.append(option); }
    spacing.parentElement!.firstChild!.textContent = "磅值:"; intervalGroup.append(spacing.parentElement!);
    interval.addEventListener("change", () => { spacing.value = interval.value === "normal" ? "0" : String((interval.value === "condensed" ? -1 : 1) * Math.abs(Number(spacing.value) || 1)); changed.add(spacing); }, { signal });
    const posMode = choose("字符位置偏移", [
      ["normal", "标准"],
      ["raised", "提升"],
      ["lowered", "降低"],
    ]);
    posMode.parentElement!.firstChild!.textContent = "位置:";
    const posAmount = field("位置磅值", "number");
    posAmount.min = "0";
    posAmount.max = "1585";
    posAmount.step = ".5";
    posAmount.parentElement!.firstChild!.textContent = "磅值:";
    const rawPos = Number(native["--word-run-native-position"] ?? "0");
    if (Number.isFinite(rawPos) && rawPos > 0) {
      posMode.value = "raised";
      posAmount.value = String(rawPos / 2);
    } else if (Number.isFinite(rawPos) && rawPos < 0) {
      posMode.value = "lowered";
      posAmount.value = String(Math.abs(rawPos) / 2);
    } else {
      posMode.value = native["--word-run-native-position"] !== undefined ? "normal" : "";
      posAmount.value = native["--word-run-native-position"] !== undefined ? "0" : "";
    }
    posMode.addEventListener(
      "change",
      () => {
        if (posMode.value === "normal") posAmount.value = "0";
        else if ((posMode.value === "raised" || posMode.value === "lowered") && !(Number(posAmount.value) > 0))
          posAmount.value = "3";
        changed.add(posAmount);
      },
      { signal },
    );
    posAmount.addEventListener(
      "input",
      () => {
        if (Number(posAmount.value) > 0 && (!posMode.value || posMode.value === "normal")) {
          posMode.value = "raised";
          changed.add(posMode);
        }
      },
      { signal },
    );
    intervalGroup.append(posMode.parentElement!, posAmount.parentElement!);
    const kernLabel = document.createElement("label"), kernBox = document.createElement("input"); kernBox.type = "checkbox"; kernBox.checked = Number(kern.value) > 0; kernLabel.append(kernBox, "为字体调整字间距:"); intervalGroup.append(kernLabel, kern.parentElement!);
    kern.parentElement!.firstChild!.textContent = "磅或更大"; kernBox.addEventListener("change", () => { kern.value = kernBox.checked ? "12" : "0"; changed.add(kern); }, { signal });
    const snapLabel = document.createElement("label"), snap = document.createElement("input"); snap.type = "checkbox"; snap.checked = native["--word-run-snap-to-grid"] !== "0"; snap.addEventListener("change", () => { snap.dataset.changed = "1"; }, { signal }); snapLabel.append(snap, "定义文档网格时对齐网格"); snapLabel.style.gridColumn = "1/-1"; intervalGroup.append(snapLabel);
    const layoutGroup = document.createElement("fieldset"), layoutTitle = document.createElement("legend"); layoutTitle.textContent = "高级版式"; layoutGroup.append(layoutTitle);
    const ligaturesOptions: [string, string][] = [
      ["none", "无"],
      ["standard", "仅标准"],
      ["standardContextual", "标准和上下文"],
      ["standardHistorical", "历史和自由"],
      ["all", "全部"],
    ];
    const currentLigature = native["--word-opentype-ligatures"] ?? "";
    if (currentLigature && !ligaturesOptions.some(([v]) => v === currentLigature))
      ligaturesOptions.push([currentLigature, currentLigature]);
    const ligaturesSelect = choose("连字", ligaturesOptions);
    ligaturesSelect.parentElement!.firstChild!.textContent = "连字:";
    ligaturesSelect.value = currentLigature;
    const numSpacingSelect = choose("数字间距", [
      ["default", "默认"],
      ["proportional", "比例"],
      ["tabular", "表格"],
    ]);
    numSpacingSelect.parentElement!.firstChild!.textContent = "数字间距:";
    numSpacingSelect.value = native["--word-opentype-num-spacing"] ?? "";
    const numFormSelect = choose("数字形式", [
      ["default", "默认"],
      ["lining", "齐线"],
      ["oldStyle", "旧式"],
    ]);
    numFormSelect.parentElement!.firstChild!.textContent = "数字形式:";
    numFormSelect.value = native["--word-opentype-num-form"] ?? "";
    const stylisticOptions: [string, string][] = [
      ["none", "默认"],
      ...Array.from({ length: 20 }, (_, i) => [String(i + 1), `样式集 ${i + 1}`] as [string, string]),
    ];
    const currentStylistic = native["--word-opentype-stylistic-set"] ?? "";
    if (currentStylistic && !stylisticOptions.some(([v]) => v === currentStylistic))
      stylisticOptions.push([currentStylistic, `样式集 ${currentStylistic}`]);
    const stylisticSelect = choose("样式集", stylisticOptions);
    stylisticSelect.parentElement!.firstChild!.textContent = "样式集:";
    stylisticSelect.value = currentStylistic;
    layoutGroup.append(
      ligaturesSelect.parentElement!,
      numSpacingSelect.parentElement!,
      numFormSelect.parentElement!,
      stylisticSelect.parentElement!,
    );
    const contextLabel = document.createElement("label"), context = document.createElement("input");
    context.type = "checkbox";
    context.setAttribute("aria-label", "使用上下文替换");
    context.checked = native["--word-opentype-contextual-alternates"] === "1";
    context.addEventListener("change", () => { context.dataset.changed = "1"; }, { signal });
    contextLabel.append(context, "使用上下文替换");
    layoutGroup.append(contextLabel);
    const advancedPreview = document.createElement("div"); advancedPreview.className = "dialog-preview"; advancedPreview.textContent = "微软卓越　AaBbCc";
    advancedTab.className = "dialog-stack"; interval.value = Number(spacing.value) < 0 ? "condensed" : Number(spacing.value) > 0 ? "expanded" : "normal";
    for (const type of ["input", "change"]) panel.addEventListener(type, updatePreview, { signal });
    updatePreview(); advancedTab.append(intervalGroup, layoutGroup, extra, advancedPreview);
    button("确定", () => {
      if (changed.has(underlineColorInput) && underlineColorInput.value.trim()
        && !CSS.supports("color", underlineColorInput.value.trim())) {
        underlineColorInput.setCustomValidity("请输入有效颜色，或留空使用自动颜色。");
      }
      if (!allFields.every(field => field.reportValidity())) return;
      if (changed.has(size) && size.value) {
        const points = wordSizeNames[size.value] ?? Number(size.value);
        if (!Number.isFinite(points) || points < .5 || points > 1638 || !Number.isInteger(points * 2)) { size.setCustomValidity("请输入 0.5～1638 磅的有效字号。"); size.reportValidity(); return; }
        size.value = String(points);
      }
      if ((changed.has(posMode) || changed.has(posAmount)) && posMode.value && posMode.value !== "normal") {
        const amount = Number(posAmount.value);
        if (!Number.isFinite(amount) || amount < 0 || amount > 1585 || !Number.isInteger(amount * 2)) {
          posAmount.setCustomValidity("请输入 0～1585 磅的有效偏移值（步长 0.5 磅）。");
          posAmount.reportValidity();
          return;
        }
      }
      for (const field of allFields) if (!changed.has(field) && field.type !== "checkbox") field.value = "";
      const basis = mark.checked ? read(true) : state;
      const properties = new Map<string, string>();
      if (family.value.trim()) {
        const font = JSON.stringify(family.value.trim());
        properties.set("font-family", font);
        for (const role of ["ascii", "hAnsi", "eastAsia", "cs"])
          properties.set("--word-font-" + role, font);
      }
      if (size.value) properties.set("font-size", size.value + "pt");
      for (const { role, input } of roles)
        if (input.value.trim())
          properties.set("--word-font-" + role, JSON.stringify(input.value.trim()));
      if (weight.value) properties.set("font-weight", weight.value);
      if (italic.value) properties.set("font-style", italic.value);
      if (vertical.value) properties.set("vertical-align", vertical.value);
      if (spacing.value) properties.set("letter-spacing", spacing.value + "pt");
      if (scale.value) properties.set("--word-run-native-w", scale.value);
      if (changed.has(underlineColorInput)) properties.set("text-decoration-color", underlineColorInput.value.trim() || "currentColor");
      if (changed.has(color) && color.value.trim()) properties.set("color", color.value.trim());
      if (background.value.trim()) properties.set("background-color", background.value.trim());
      if (underline.value || strike.value) {
        const lines = [
          ...(underline.value
            ? underline.value === "none"
              ? []
              : ["underline"]
            : basis.underline
              ? ["underline"]
              : []),
          ...(strike.value
            ? strike.value === "none"
              ? []
              : ["line-through"]
            : basis.strikethrough
              ? ["line-through"]
              : []),
        ];
        properties.set("text-decoration-line", lines.join(" ") || "none");
        if (underline.value) {
          const ooxmlUnderline =
            underline.value === "solid"
              ? "single"
              : underline.value === "dashed"
                ? "dash"
                : underline.value === "wavy"
                  ? "wave"
                  : underline.value;
          const cssStyle =
            underline.value === "double"
              ? "double"
              : underline.value === "dotted" || underline.value === "dottedHeavy"
                ? "dotted"
                : underline.value === "wavy" ||
                    underline.value === "wavyHeavy" ||
                    underline.value === "wavyDouble"
                  ? "wavy"
                  : [
                        "dashed",
                        "dashedHeavy",
                        "dashLong",
                        "dashLongHeavy",
                        "dotDash",
                        "dashDotHeavy",
                        "dotDotDash",
                        "dashDotDotHeavy",
                      ].includes(underline.value)
                    ? "dashed"
                    : "solid";
          properties.set("--word-underline-style", ooxmlUnderline);
          if (underline.value !== "none") properties.set("text-decoration-style", cssStyle);
        }
      }
      if (doubleStrikeBox.dataset.changed)
        properties.set("--word-dstrike", doubleStrikeBox.checked ? "1" : "0");
      if (hiddenBox.dataset.changed)
        properties.set("--word-run-native-vanish", hiddenBox.checked ? "1" : "0");
      if (emphasisInput.value) properties.set("--word-run-native-em", emphasisInput.value);
      if (posMode.value || posAmount.value) {
        const mode = posMode.value || "normal";
        const amount = Number(posAmount.value || "0");
        const halfPoints =
          mode === "normal" ? 0 : Math.round(amount * 2) * (mode === "lowered" ? -1 : 1);
        properties.set("--word-run-native-position", String(halfPoints));
      }
      if (ligaturesSelect.value) properties.set("--word-opentype-ligatures", ligaturesSelect.value);
      if (numSpacingSelect.value)
        properties.set("--word-opentype-num-spacing", numSpacingSelect.value);
      if (numFormSelect.value) properties.set("--word-opentype-num-form", numFormSelect.value);
      if (stylisticSelect.value)
        properties.set("--word-opentype-stylistic-set", stylisticSelect.value);
      if (context.dataset.changed)
        properties.set("--word-opentype-contextual-alternates", context.checked ? "1" : "0");
      if (caps.value) {
        properties.set("font-variant", caps.value === "small-caps" ? "small-caps" : "normal");
        properties.set("text-transform", caps.value === "uppercase" ? "uppercase" : "none");
      }
      if (complexSize.value) properties.set("--word-font-size-cs", complexSize.value + "pt");
      if (complexBold.value) properties.set("--word-bold-cs", complexBold.value);
      if (complexItalic.value) properties.set("--word-italic-cs", complexItalic.value);
      if (kern.value) properties.set("--word-kerning-threshold", kern.value + "pt");
      if (snap.dataset.changed) properties.set("--word-run-snap-to-grid", snap.checked ? "1" : "0");
      for (const { role, input } of languages)
        if (input.value.trim())
          properties.set("--word-language-" + role, JSON.stringify(input.value.trim()));
      if (properties.size) apply(properties, mark.checked);
      close();
    });
    ui.footer.lastElementChild!.classList.add("primary");
    const defaults = document.createElement("button"); defaults.type = "button"; defaults.textContent = "默认…"; defaults.disabled = true; defaults.style.marginRight = "auto"; ui.footer.prepend(defaults);
    size.addEventListener("input", () => size.setCustomValidity(""), { signal });
    ui.show();
  }
  return { open, dispose: close };
}
