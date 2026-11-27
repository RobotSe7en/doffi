import { DocumentError, WORD_ART_PAGE_BORDER_STYLES, getWordOutline } from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation, WordPageBorders } from "@doffi/core";
import { evaluateWordTableFormula } from "@doffi/render";
import { createOfficeDialog } from "./office-dialog.js";
import { createWordDockedPane, replaceSameSideDock } from "./word-pane-chrome.js";
import type { WordPasteMode } from "./clipboard-command.js";

const ART_BORDER_LABELS: readonly (readonly [string, string])[] = [
  ["", "（无艺术图案）"],
  ["apples", "苹果 (Apples)"],
  ["stars", "金星 (Stars)"],
  ["starsBlack", "黑星 (Black Stars)"],
  ["hearts", "红心 (Hearts)"],
  ["trees", "树木 (Trees)"],
  ["christmasTree", "圣诞树 (Christmas Tree)"],
  ["mapleLeaf", "枫叶 (Maple Leaf)"],
  ["palmsColor", "棕榈 (Palms)"],
  ["flowersBlockPrint", "木刻花朵 (Flowers)"],
  ["checkered", "棋盘格 (Checkered)"],
  ["classicalWave", "古典波浪 (Classical Wave)"],
  ["diamondsGray", "灰菱形 (Diamonds)"],
  ["doubleDiamonds", "双菱形 (Double Diamonds)"],
  ["basicBlackSquares", "黑方块 (Black Squares)"],
  ["basicBlackDots", "黑圆点 (Black Dots)"],
  ["confetti", "彩纸屑 (Confetti)"],
];
const ART_BORDER_SET = new Set<string>(WORD_ART_PAGE_BORDER_STYLES);

export interface WordParitySelection {
  readonly paragraphId?: string;
  readonly paragraphIds: readonly string[];
  readonly from: number;
  readonly to: number;
  readonly text: string;
  readonly rubyNodeId?: string;
  readonly rubyText?: string;
  readonly tableId?: string;
  readonly cellId?: string;
  readonly sectionRootId: string;
}

const SYMBOL_GROUPS: readonly (readonly [string, readonly string[]])[] = [
  ["常用标点", ["·", "—", "…", "§", "※", "★", "☆", "●", "○", "■", "□", "▲", "△", "◆", "◇", "✓", "✗", "✔", "✘", "☑", "☐"]],
  ["数学符号", ["±", "×", "÷", "≠", "≈", "≤", "≥", "∞", "√", "∑", "∏", "∫", "∂", "∆", "∇", "∈", "∉", "⊂", "⊃", "∪", "∩", "∀", "∃", "∴", "∵", "°", "′", "″", "‰", "‱"]],
  ["箭头", ["←", "→", "↑", "↓", "↔", "↕", "↖", "↗", "↘", "↙", "⇐", "⇒", "⇑", "⇓", "⇔", "➜", "➤", "➥", "↩", "↪"]],
  ["希腊字母", ["α", "β", "γ", "δ", "ε", "ζ", "η", "θ", "λ", "μ", "ν", "ξ", "π", "ρ", "σ", "τ", "φ", "χ", "ψ", "ω", "Γ", "Δ", "Θ", "Λ", "Π", "Σ", "Φ", "Ψ", "Ω"]],
  ["货币与法律", ["¥", "$", "€", "£", "¢", "₩", "₹", "₽", "©", "®", "™", "℃", "℉", "№", "℗", "℠"]],
  ["带圈数字", ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳", "⑴", "⑵", "⑶", "⑷", "⑸", "Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ"]],
];

/** First-party 1:1 Word parity dialogs: Page Borders, TOC, Ruby, Enclosed Chars,
 * Symbols, Table Formulas, Drop Cap, Chinese Layout, Properties, Protection, and Outline Arranger. */
export function createWordParityControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  readSelection: () => WordParitySelection,
  callbacks: {
    onUpdateFields?(): void;
    focusNode?(nodeId: string): void;
    onError(error: unknown): void;
  },
): {
  openPageBorders(): void;
  openTocDialog(): void;
  openRubyDialog(): void;
  openEnclosedCharDialog(): void;
  openSymbolsDialog(): void;
  openTableFormulaDialog(): void;
  openDropCapDialog(): void;
  openChineseLayoutDialog(mode: "combine" | "vert" | "fit"): void;
  openDocPropertiesDialog(): void;
  openProtectionDialog(): void;
  openOutlineView(): void;
  openPasteSpecialDialog(onApply: (mode: WordPasteMode | "image") => Promise<void> | void): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let activeDialog: ReturnType<typeof createOfficeDialog> | undefined;
  let activePanel: HTMLElement | undefined;
  let controller: AbortController | undefined;

  const close = (): void => {
    controller?.abort();
    controller = undefined;
    activeDialog?.close();
    activeDialog?.dialog.remove();
    activeDialog = undefined;
    activePanel?.remove();
    activePanel = undefined;
  };

  const commit = (operations: DocumentOperation[]): void => {
    const snapshot = session.snapshot();
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations,
      }),
    );
  };

  const documentRootId = (): string =>
    new DOMParser()
      .parseFromString(session.snapshot().html, "text/html")
      .querySelector<HTMLElement>('[data-office-kind="document"]')!.dataset.nodeId!;

  const showDialog = (title: string, width = "480px") => {
    close();
    controller = new AbortController();
    const ui = createOfficeDialog(document, title);
    ui.dialog.style.width = width;
    ui.content.style.minHeight = "0";
    activeDialog = ui;
    parent.insertBefore(ui.dialog, before);
    ui.dialog.addEventListener(
      "cancel",
      (event) => {
        event.preventDefault();
        close();
      },
      { signal: controller.signal },
    );
    const addButton = (label: string, run: () => void, primary = false) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = label;
      if (primary) btn.className = "primary";
      btn.addEventListener(
        "click",
        () => {
          try {
            run();
          } catch (error) {
            ui.error(error);
            callbacks.onError(error);
          }
        },
        { signal: controller!.signal },
      );
      ui.footer.append(btn);
      return btn;
    };
    return { ui, signal: controller.signal, addButton };
  };

  const makeSelect = (
    labelText: string,
    options: readonly (readonly [string, string])[],
    initial: string,
    ariaLabel = labelText,
  ) => {
    const label = document.createElement("label");
    label.textContent = labelText;
    const select = document.createElement("select");
    select.setAttribute("aria-label", ariaLabel);
    for (const [val, text] of options) {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = text;
      select.append(opt);
    }
    select.value = initial;
    label.append(select);
    return { label, select };
  };

  const makeInput = (labelText: string, value: string, type = "text", ariaLabel = labelText) => {
    const label = document.createElement("label");
    label.textContent = labelText;
    const input = document.createElement("input");
    input.type = type;
    input.value = value;
    input.setAttribute("aria-label", ariaLabel);
    label.append(input);
    return { label, input };
  };

  const makeCheckbox = (labelText: string, checked: boolean, ariaLabel = labelText) => {
    const label = document.createElement("label");
    const input = document.createElement("input");
    input.type = "checkbox";
    input.checked = checked;
    input.setAttribute("aria-label", ariaLabel);
    label.append(input, labelText);
    return { label, input };
  };

  function openPageBorders(): void {
    const sel = readSelection();
    const meta = session.snapshot().metadata;
    if (meta.presentation) throw new DocumentError("InvalidOperation", "页面边框仅适用于 Word 文档。");
    const section = meta.sections.find((s) => s.rootNodeId === sel.sectionRootId) ?? meta.sections[0]!;
    const current = section.pageBorders;
    const edge = current?.top ?? current?.left ?? current?.bottom ?? current?.right;

    const { ui, signal, addButton } = showDialog("页面边框", "540px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const preset = makeSelect(
      "设置",
      [
        ["none", "无"],
        ["box", "方框"],
        ["shadow", "阴影"],
        ["3d", "三维"],
      ],
      !edge || edge.style === "none" ? "none" : edge.shadow ? "shadow" : edge.frame ? "3d" : "box",
      "页面边框设置",
    );
    const initialIsArt = Boolean(edge?.style && ART_BORDER_SET.has(edge.style));
    const lineStyle = makeSelect(
      "线型",
      [
        ["single", "实线"],
        ["double", "双线"],
        ["triple", "三线"],
        ["dotted", "点线"],
        ["dashed", "虚线"],
        ["dashSmallGap", "短虚线"],
        ["dotDash", "点划线"],
        ["dotDotDash", "双点划线"],
        ["thick", "粗实线"],
        ["wave", "波浪线"],
        ["doubleWave", "双波浪线"],
        ["threeDEmboss", "三维浮雕"],
        ["threeDEngrave", "三维雕刻"],
      ],
      edge?.style && edge.style !== "none" && !initialIsArt ? edge.style : "single",
      "页面边框线型",
    );
    const artStyle = makeSelect(
      "艺术型",
      ART_BORDER_LABELS,
      initialIsArt ? edge!.style : "",
      "艺术型页面边框",
    );
    const color = makeInput(
      "颜色",
      "#" + (edge?.color && edge.color !== "auto" ? edge.color : "000000"),
      "color",
      "页面边框颜色",
    );
    const width = makeInput("宽度（磅）", String(edge?.width ?? (initialIsArt ? 14 : 1)), "number", "页面边框宽度");
    width.input.min = "0.25";
    width.input.max = "31";
    width.input.step = "0.25";
    artStyle.select.addEventListener(
      "change",
      () => {
        if (artStyle.select.value) {
          if (preset.select.value === "none") preset.select.value = "box";
          if (Number(width.input.value) < 8) width.input.value = "14";
        }
      },
      { signal },
    );
    const space = makeInput("距正文距离（磅）", String(edge?.space ?? 24), "number", "页面边框距离");
    space.input.min = "0";
    space.input.max = "31";
    const offsetFrom = makeSelect("度量基准", [["text", "文字"], ["page", "页边"]], current?.offsetFrom ?? "text", "度量基准");
    const display = makeSelect(
      "应用于页面",
      [
        ["allPages", "整节所有页"],
        ["firstPage", "仅本节首页"],
        ["notFirstPage", "除首页外所有页"],
      ],
      current?.display ?? "allPages",
      "应用于页面",
    );
    const scope = makeSelect("应用范围", [["all", "整篇文档"], ["section", "本节"]], "all", "应用范围");
    const sidesRow = document.createElement("div");
    sidesRow.style.cssText = "display:flex;gap:14px;flex-wrap:wrap";
    const sideBoxes = {
      top: makeCheckbox("上边框", current ? Boolean(current.top) : true, "上边框"),
      bottom: makeCheckbox("下边框", current ? Boolean(current.bottom) : true, "下边框"),
      left: makeCheckbox("左边框", current ? Boolean(current.left) : true, "左边框"),
      right: makeCheckbox("右边框", current ? Boolean(current.right) : true, "右边框"),
    };
    sidesRow.append(sideBoxes.top.label, sideBoxes.bottom.label, sideBoxes.left.label, sideBoxes.right.label);
    const surroundHeader = makeCheckbox("环绕页眉", !meta.word?.bordersDoNotSurroundHeader, "环绕页眉");
    const surroundFooter = makeCheckbox("环绕页脚", !meta.word?.bordersDoNotSurroundFooter, "环绕页脚");

    const preview = document.createElement("div");
    preview.className = "dialog-preview";
    preview.textContent = "页面边框预览";
    preview.style.fontSize = "14px";
    const syncPreview = () => {
      const none = preset.select.value === "none";
      const isArt = Boolean(artStyle.select.value);
      const cssStyle = isArt
        ? "double"
        : lineStyle.select.value === "double" || lineStyle.select.value === "triple"
          ? "double"
          : lineStyle.select.value === "dotted"
            ? "dotted"
            : lineStyle.select.value.includes("dash") || lineStyle.select.value.includes("Dash")
              ? "dashed"
              : lineStyle.select.value === "threeDEmboss"
                ? "ridge"
                : lineStyle.select.value === "threeDEngrave"
                  ? "groove"
                  : "solid";
      const w = Math.max(1, Math.min(12, Number(width.input.value) || 1));
      const spec = none ? "none" : `${w}px ${cssStyle} ${color.input.value}`;
      preview.textContent = isArt
        ? `艺术边框预览：${ART_BORDER_LABELS.find((x) => x[0] === artStyle.select.value)?.[1] ?? artStyle.select.value}`
        : "页面边框预览";
      preview.style.borderTop = sideBoxes.top.input.checked ? spec : "none";
      preview.style.borderBottom = sideBoxes.bottom.input.checked ? spec : "none";
      preview.style.borderLeft = sideBoxes.left.input.checked ? spec : "none";
      preview.style.borderRight = sideBoxes.right.input.checked ? spec : "none";
      preview.style.boxShadow = !none && preset.select.value === "shadow" ? "4px 4px 0 #333" : "none";
    };
    syncPreview();
    for (const el of [preset.select, lineStyle.select, artStyle.select, color.input, width.input, sideBoxes.top.input, sideBoxes.bottom.input, sideBoxes.left.input, sideBoxes.right.input]) {
      el.addEventListener("change", syncPreview, { signal });
      el.addEventListener("input", syncPreview, { signal });
    }

    const grid = document.createElement("div");
    grid.className = "dialog-options";
    grid.append(preset.label, lineStyle.label, artStyle.label, color.label, width.label, space.label, offsetFrom.label, display.label, scope.label);
    stack.append(grid, sidesRow, surroundHeader.label, surroundFooter.label, preview);

    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const latestMeta = session.snapshot().metadata;
        if (latestMeta.presentation) return;
        const targets = scope.select.value === "all" ? latestMeta.sections : [latestMeta.sections.find((s) => s.rootNodeId === section.rootNodeId) ?? latestMeta.sections[0]!];
        const none = preset.select.value === "none";
        const borders: WordPageBorders = {
          offsetFrom: offsetFrom.select.value as WordPageBorders["offsetFrom"],
          display: display.select.value as WordPageBorders["display"],
          zOrder: "front",
        };
        if (!none) {
          const chosenStyle = (artStyle.select.value || lineStyle.select.value) as NonNullable<WordPageBorders["top"]>["style"];
          const edgeValue = {
            style: chosenStyle,
            color: color.input.value.replace(/^#/, ""),
            width: Math.max(0.25, Math.min(31, Number(width.input.value) || 1)),
            space: Math.max(0, Math.min(31, Number(space.input.value) || 24)),
            shadow: preset.select.value === "shadow",
            frame: preset.select.value === "3d",
          };
          for (const side of ["top", "bottom", "left", "right"] as const) {
            if (sideBoxes[side].input.checked) borders[side] = edgeValue;
          }
        }
        commit([
          ...targets.map((target) => ({
            op: "setWordSection" as const,
            nodeId: target.rootNodeId,
            changes: { pageBorders: borders },
          })),
          {
            op: "setWordOptions",
            nodeId: documentRootId(),
            bordersDoNotSurroundHeader: !surroundHeader.input.checked,
            bordersDoNotSurroundFooter: !surroundFooter.input.checked,
          },
        ]);
        close();
      },
      true,
    );
    ui.show();
  }

  function openTocDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const { ui, addButton } = showDialog("目录", "480px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const kind = makeSelect(
      "目录类型",
      [
        ["toc-1-3", "自动目录（标题 1～3 级）"],
        ["toc-1-6", "完整目录（标题 1～6 级）"],
        ["tof-fig", "插图目录（题注：图）"],
        ["tof-tbl", "表格目录（题注：表）"],
      ],
      "toc-1-3",
      "目录类型",
    );
    const hyperlinks = makeCheckbox("使用超链接而不使用页码跳转（\\h）", true, "目录使用超链接");
    const showPages = makeCheckbox("显示页码", true, "显示页码");

    const info = document.createElement("p");
    info.style.cssText = "margin:0;font-size:12px;color:#555";
    const headings = getWordOutline(session);
    info.textContent = headings.length
      ? `当前文档检测到 ${headings.length} 个标题，插入后将自动按页码计算目录项。`
      : "当前文档尚未设置标题样式或大纲级别，插入后可在“更新目录”时自动刷新。";

    stack.append(kind.label, hyperlinks.label, showPages.label, info);

    addButton("更新现有目录", () => {
      callbacks.onUpdateFields?.();
      close();
    });
    addButton("取消", close);
    addButton(
      "插入目录",
      () => {
        const flags = [hyperlinks.input.checked ? "\\h" : "", "\\z", showPages.input.checked ? "" : "\\n"].filter(Boolean).join(" ");
        let instruction = ` TOC \\o "1-3" ${flags} \\u `;
        let cached = "目录";
        if (kind.select.value === "toc-1-6") {
          instruction = ` TOC \\o "1-6" ${flags} \\u `;
        } else if (kind.select.value === "tof-fig") {
          instruction = ` TOC ${flags} \\c "图" `;
          cached = "插图目录";
        } else if (kind.select.value === "tof-tbl") {
          instruction = ` TOC ${flags} \\c "表" `;
          cached = "表格目录";
        }
        commit([
          {
            op: "insertWordField",
            nodeId: sel.paragraphId!,
            offset: sel.from,
            instruction,
            cachedText: cached,
          },
        ]);
        queueMicrotask(() => {
          try {
            callbacks.onUpdateFields?.();
          } catch {
            /* Layout may refresh asynchronously. */
          }
        });
        close();
      },
      true,
    );
    ui.show();
  }

  function openRubyDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先选择要添加拼音指南的文字。");
    if (!sel.rubyNodeId && sel.from === sel.to)
      throw new DocumentError("InvalidOperation", "请先选中要注音的中文文字。");

    const { ui, signal, addButton } = showDialog("拼音指南", "460px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const baseText = sel.rubyNodeId ? (session.getNode(sel.rubyNodeId)?.text ?? sel.text) : sel.text;
    const baseField = makeInput("基准文字", baseText, "text", "基准文字");
    baseField.input.readOnly = true;
    const rubyInput = makeInput("拼音文字", sel.rubyText ?? "", "text", "拼音文字");
    rubyInput.input.placeholder = "例如：pīn yīn";
    const alignSelect = makeSelect(
      "对齐方式",
      [
        ["distributeSpace", "分散对齐（1-2-1）"],
        ["center", "居中"],
        ["distributeLetter", "分散对齐（0-1-0）"],
        ["left", "左对齐"],
        ["right", "右对齐"],
      ],
      "distributeSpace",
      "拼音对齐方式",
    );
    const sizeInput = makeInput("字号（磅）", "6", "number", "拼音字号");
    sizeInput.input.min = "4";
    sizeInput.input.max = "36";
    sizeInput.input.step = "0.5";
    const raiseInput = makeInput("偏移量（磅）", "11", "number", "拼音偏移量");
    raiseInput.input.min = "0";
    raiseInput.input.max = "72";

    const preview = document.createElement("div");
    preview.className = "dialog-preview";
    const syncPreview = () => {
      preview.replaceChildren();
      const wrap = document.createElement("span");
      wrap.style.cssText = "display:inline-flex;flex-direction:column;align-items:center;line-height:1.1";
      const rt = document.createElement("span");
      rt.style.fontSize = Math.max(8, (Number(sizeInput.input.value) || 6) * 1.4) + "px";
      rt.textContent = rubyInput.input.value || "pīn yīn";
      const rb = document.createElement("span");
      rb.textContent = baseText;
      wrap.append(rt, rb);
      preview.append(wrap);
    };
    syncPreview();
    rubyInput.input.addEventListener("input", syncPreview, { signal });
    sizeInput.input.addEventListener("input", syncPreview, { signal });

    const grid = document.createElement("div");
    grid.className = "dialog-options";
    grid.append(baseField.label, rubyInput.label, alignSelect.label, sizeInput.label, raiseInput.label);
    stack.append(grid, preview);

    if (sel.rubyNodeId) {
      addButton("清除拼音", () => {
        commit([{ op: "removeWordRuby", nodeId: sel.rubyNodeId! }]);
        close();
      });
    }
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const rubyText = rubyInput.input.value.trim();
        if (!rubyText) throw new DocumentError("InvalidOperation", "请输入拼音文字。");
        const ops: DocumentOperation[] = [];
        if (sel.rubyNodeId) {
          ops.push({ op: "removeWordRuby", nodeId: sel.rubyNodeId });
        }
        const paraText = session.getNode(sel.paragraphId!)?.text ?? "";
        const expected = sel.from < sel.to ? paraText.slice(sel.from, sel.to) : baseText;
        const from = sel.from < sel.to ? sel.from : Math.max(0, paraText.indexOf(baseText));
        const to = from + expected.length;
        ops.push({
          op: "insertWordRuby",
          nodeId: sel.paragraphId!,
          from,
          to,
          expectedText: expected,
          rubyText,
          align: alignSelect.select.value as "distributeSpace" | "center" | "distributeLetter" | "left" | "right",
          fontSize: Math.max(4, Number(sizeInput.input.value) || 6),
          raise: Math.max(0, Number(raiseInput.input.value) || 11),
          languageId: "zh-CN",
        });
        commit(ops);
        close();
      },
      true,
    );
    ui.show();
    queueMicrotask(() => rubyInput.input.focus());
  }

  function openEnclosedCharDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在段落中放置光标或选择字符。");
    const { ui, signal, addButton } = showDialog("带圈字符", "440px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const textInput = makeInput("文字（1～2 个字符）", sel.text.slice(0, 2) || "1", "text", "带圈文字");
    textInput.input.maxLength = 4;
    const enclosure = makeSelect(
      "圈号",
      [
        ["○", "○ 圆形"],
        ["□", "□ 方形"],
        ["△", "△ 三角形"],
        ["◇", "◇ 菱形"],
      ],
      "○",
      "圈号形状",
    );
    const mode = makeSelect(
      "样式",
      [
        ["shrink", "缩小文字"],
        ["enlarge", "增大圈号"],
      ],
      "shrink",
      "带圈样式",
    );
    const preview = document.createElement("div");
    preview.className = "dialog-preview";
    const syncPreview = () => {
      preview.replaceChildren();
      const box = document.createElement("span");
      const shape = enclosure.select.value;
      box.style.cssText =
        "display:inline-flex;align-items:center;justify-content:center;width:1.6em;height:1.6em;border:1.5px solid currentColor;font-size:" +
        (mode.select.value === "shrink" ? "16px" : "22px") +
        (shape === "○" ? ";border-radius:50%" : shape === "◇" ? ";transform:rotate(45deg)" : "");
      const inner = document.createElement("span");
      inner.textContent = textInput.input.value || "1";
      if (shape === "◇") inner.style.transform = "rotate(-45deg)";
      box.append(inner);
      preview.append(box);
    };
    syncPreview();
    for (const el of [textInput.input, enclosure.select, mode.select]) {
      el.addEventListener("input", syncPreview, { signal });
      el.addEventListener("change", syncPreview, { signal });
    }

    stack.append(textInput.label, enclosure.label, mode.label, preview);
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const chars = textInput.input.value.trim();
        if (!chars) throw new DocumentError("InvalidOperation", "请输入要加圈的文字。");
        const ops: DocumentOperation[] = [];
        if (sel.to > sel.from) {
          ops.push({
            op: "replaceText",
            nodeId: sel.paragraphId!,
            from: sel.from,
            to: sel.to,
            expectedText: sel.text,
            text: "",
          });
        }
        const shape = enclosure.select.value;
        const instruction =
          mode.select.value === "enlarge"
            ? ` EQ \\o\\ac(\\s\\up 0(${shape}),${chars}) `
            : ` EQ \\o\\ac(${shape},${chars}) `;
        ops.push({
          op: "insertWordField",
          nodeId: sel.paragraphId!,
          offset: sel.from,
          instruction,
          cachedText: chars,
        });
        commit(ops);
        close();
      },
      true,
    );
    ui.show();
  }

  function openSymbolsDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在段落中放置光标。");
    const { ui, signal, addButton } = showDialog("符号", "540px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const groupSelect = makeSelect(
      "子集",
      SYMBOL_GROUPS.map(([name]) => [name, name] as const),
      SYMBOL_GROUPS[0]![0],
      "符号子集",
    );
    const grid = document.createElement("div");
    grid.style.cssText = "display:grid;grid-template-columns:repeat(10,1fr);gap:4px;padding:6px;background:#fff;border:1px solid #d5d5d5;border-radius:4px;max-height:220px;overflow:auto";
    const status = document.createElement("p");
    status.style.cssText = "margin:0;font-size:12px;color:#555";
    let chosen = SYMBOL_GROUPS[0]![1][0]!;

    const renderGrid = () => {
      grid.replaceChildren();
      const group = SYMBOL_GROUPS.find(([name]) => name === groupSelect.select.value) ?? SYMBOL_GROUPS[0]!;
      chosen = group[1][0]!;
      updateStatus();
      for (const sym of group[1]) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.textContent = sym;
        btn.setAttribute("aria-label", "符号 " + sym);
        btn.style.cssText = "height:32px;font-size:17px;border:1px solid #e0e0e0;background:#fafafa;cursor:pointer;border-radius:3px";
        btn.addEventListener(
          "click",
          () => {
            chosen = sym;
            for (const other of grid.querySelectorAll<HTMLButtonElement>("button")) {
              other.style.background = other === btn ? "#dce8f8" : "#fafafa";
              other.style.borderColor = other === btn ? "#185abd" : "#e0e0e0";
            }
            updateStatus();
          },
          { signal },
        );
        btn.addEventListener(
          "dblclick",
          () => {
            chosen = sym;
            insertChosen();
          },
          { signal },
        );
        grid.append(btn);
      }
    };
    const updateStatus = () => {
      const cp = chosen.codePointAt(0)?.toString(16).toUpperCase().padStart(4, "0") ?? "";
      status.textContent = `当前符号：${chosen} （Unicode: U+${cp}）`;
    };
    const insertChosen = () => {
      const latest = readSelection();
      const targetId = latest.paragraphId ?? sel.paragraphId!;
      commit([
        {
          op: "replaceText",
          nodeId: targetId,
          from: latest.from,
          to: latest.to,
          expectedText: latest.text,
          text: chosen,
        },
      ]);
      close();
    };
    groupSelect.select.addEventListener("change", renderGrid, { signal });
    renderGrid();

    stack.append(groupSelect.label, grid, status);
    addButton("取消", close);
    addButton("插入", insertChosen, true);
    ui.show();
  }

  function openTableFormulaDialog(): void {
    const sel = readSelection();
    if (!sel.tableId || !sel.cellId || !sel.paragraphId)
      throw new DocumentError("InvalidOperation", "请先将光标放在表格单元格内。");

    const { ui, signal, addButton } = showDialog("公式", "460px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const formulaInput = makeInput("公式", "=SUM(ABOVE)", "text", "表格公式");
    const formatSelect = makeSelect(
      "编号格式",
      [
        ["", "默认"],
        ["#,##0", "#,##0"],
        ["#,##0.00", "#,##0.00"],
        ["0.00%", "0.00%"],
        ["¥#,##0.00", "¥#,##0.00"],
      ],
      "",
      "编号格式",
    );
    const fnSelect = makeSelect(
      "粘贴函数",
      [
        ["", "选择函数…"],
        ["SUM(ABOVE)", "SUM(ABOVE) — 对上方求和"],
        ["SUM(LEFT)", "SUM(LEFT) — 对左侧求和"],
        ["AVERAGE(ABOVE)", "AVERAGE(ABOVE) — 上方平均值"],
        ["AVERAGE(LEFT)", "AVERAGE(LEFT) — 左侧平均值"],
        ["COUNT(ABOVE)", "COUNT(ABOVE) — 上方计数"],
        ["MAX(ABOVE)", "MAX(ABOVE) — 上方最大值"],
        ["MIN(ABOVE)", "MIN(ABOVE) — 上方最小值"],
        ["PRODUCT(LEFT)", "PRODUCT(LEFT) — 左侧乘积"],
      ],
      "",
      "粘贴函数",
    );
    fnSelect.select.addEventListener(
      "change",
      () => {
        if (fnSelect.select.value) {
          formulaInput.input.value = "=" + fnSelect.select.value;
          syncResult();
        }
      },
      { signal },
    );
    const resultPreview = document.createElement("p");
    resultPreview.style.cssText = "margin:0;font-size:12px;color:#185abd;font-weight:600";

    const computeInstruction = () => {
      const expr = formulaInput.input.value.trim().startsWith("=")
        ? formulaInput.input.value.trim()
        : "=" + formulaInput.input.value.trim();
      const fmt = formatSelect.select.value ? ` \\# "${formatSelect.select.value}"` : "";
      return ` ${expr}${fmt} `;
    };
    const evaluatePreview = (): string => {
      const doc = new DOMParser().parseFromString(session.snapshot().html, "text/html");
      const cellEl = doc.querySelector<HTMLElement>(`[data-node-id="${CSS.escape(sel.cellId!)}"]`);
      if (!cellEl) return "0";
      return evaluateWordTableFormula(computeInstruction(), cellEl) ?? "0";
    };
    const syncResult = () => {
      resultPreview.textContent = "实时计算结果：" + evaluatePreview();
    };
    syncResult();
    formulaInput.input.addEventListener("input", syncResult, { signal });
    formatSelect.select.addEventListener("change", syncResult, { signal });

    stack.append(formulaInput.label, formatSelect.label, fnSelect.label, resultPreview);
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const instruction = computeInstruction();
        const cachedText = evaluatePreview();
        commit([
          {
            op: "insertWordField",
            nodeId: sel.paragraphId!,
            offset: sel.from,
            instruction,
            cachedText,
          },
        ]);
        close();
      },
      true,
    );
    ui.show();
  }

  function openDropCapDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先将光标放在正文段落中。");
    const { ui, addButton } = showDialog("首字下沉", "420px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const pos = makeSelect(
      "位置",
      [
        ["none", "无"],
        ["drop", "下沉"],
        ["margin", "悬挂"],
      ],
      "drop",
      "首字下沉位置",
    );
    const lines = makeInput("下沉行数", "3", "number", "下沉行数");
    lines.input.min = "1";
    lines.input.max = "10";
    const distance = makeInput("距正文距离（磅）", "0", "number", "距正文距离");
    distance.input.min = "0";
    distance.input.max = "72";
    distance.input.step = "0.5";

    stack.append(pos.label, lines.label, distance.label);
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        commit([
          {
            op: "setWordParagraph",
            nodeId: sel.paragraphId!,
            properties:
              pos.select.value === "none"
                ? { "--word-drop-cap": "none" }
                : {
                    "--word-drop-cap": pos.select.value,
                    "--word-drop-cap-lines": String(Math.max(1, Math.min(10, Math.round(Number(lines.input.value) || 3)))),
                    "--word-drop-cap-space": Math.max(0, Number(distance.input.value) || 0) + "pt",
                  },
          },
        ]);
        close();
      },
      true,
    );
    ui.show();
  }

  function openChineseLayoutDialog(mode: "combine" | "vert" | "fit"): void {
    const sel = readSelection();
    if (!sel.paragraphId || sel.from === sel.to)
      throw new DocumentError("InvalidOperation", "请先选择要设置中文版式的文字。");
    const titles = { combine: "双行合一", vert: "纵横混排", fit: "调整宽度" };
    const { ui, addButton } = showDialog(titles[mode], "420px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    let bracketSelect: HTMLSelectElement | undefined;
    let widthInput: HTMLInputElement | undefined;
    let compressCheck: HTMLInputElement | undefined;

    if (mode === "combine") {
      const b = makeSelect(
        "带括号",
        [
          ["none", "无"],
          ["round", "( ) 圆括号"],
          ["square", "[ ] 方括号"],
          ["angle", "< > 尖括号"],
          ["curly", "{ } 花括号"],
        ],
        "none",
        "双行合一括号",
      );
      bracketSelect = b.select;
      stack.append(b.label);
    } else if (mode === "vert") {
      const c = makeCheckbox("适应行宽（压缩纵横混排字符）", true, "适应行宽");
      compressCheck = c.input;
      stack.append(c.label);
    } else {
      const w = makeInput("新文字宽度（磅）", "36", "number", "调整宽度");
      w.input.min = "6";
      w.input.max = "400";
      widthInput = w.input;
      stack.append(w.label);
    }

    addButton("清除效果", () => {
      const props: Record<string, string> =
        mode === "combine"
          ? { "--word-ea-combine": "none" }
          : mode === "vert"
            ? { "--word-ea-vert": "none" }
            : { "--word-fit-text": "none" };
      commit([
        {
          op: "formatWordText",
          nodeId: sel.paragraphId!,
          from: sel.from,
          to: sel.to,
          expectedText: sel.text,
          properties: props,
        },
      ]);
      close();
    });
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const props: Record<string, string> =
          mode === "combine"
            ? { "--word-ea-combine": bracketSelect!.value }
            : mode === "vert"
              ? { "--word-ea-vert": compressCheck!.checked ? "compress" : "1" }
              : { "--word-fit-text": Math.max(6, Number(widthInput!.value) || 36) + "pt" };
        commit([
          {
            op: "formatWordText",
            nodeId: sel.paragraphId!,
            from: sel.from,
            to: sel.to,
            expectedText: sel.text,
            properties: props,
          },
        ]);
        close();
      },
      true,
    );
    ui.show();
  }

  function openDocPropertiesDialog(): void {
    const meta = session.snapshot().metadata;
    const props = meta.word?.properties ?? {};
    const { ui, addButton } = showDialog("文档属性", "480px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const title = makeInput("标题", meta.title ?? "", "text", "文档标题");
    const creator = makeInput("作者", props.creator ?? "", "text", "文档作者");
    const subject = makeInput("主题", props.subject ?? "", "text", "文档主题");
    const keywords = makeInput("关键词", props.keywords ?? "", "text", "文档关键词");
    const category = makeInput("类别", props.category ?? "", "text", "文档类别");
    const desc = makeInput("备注", props.description ?? "", "text", "文档备注");

    stack.append(title.label, creator.label, subject.label, keywords.label, category.label, desc.label);
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        commit([
          {
            op: "setWordOptions",
            nodeId: documentRootId(),
            title: title.input.value.trim(),
            properties: {
              ...(creator.input.value.trim() ? { creator: creator.input.value.trim() } : {}),
              ...(subject.input.value.trim() ? { subject: subject.input.value.trim() } : {}),
              ...(keywords.input.value.trim() ? { keywords: keywords.input.value.trim() } : {}),
              ...(category.input.value.trim() ? { category: category.input.value.trim() } : {}),
              ...(desc.input.value.trim() ? { description: desc.input.value.trim() } : {}),
            },
          },
        ]);
        close();
      },
      true,
    );
    ui.show();
  }

  function openProtectionDialog(): void {
    const current = session.snapshot().metadata.word?.protection;
    const { ui, addButton } = showDialog("限制编辑 / 文档保护", "460px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const mode = makeSelect(
      "编辑限制",
      [
        ["none", "无限制（自由编辑）"],
        ["readOnly", "不允许任何更改（只读）"],
        ["trackedChanges", "仅允许修订（强制记录修订）"],
        ["comments", "仅允许批注"],
        ["forms", "仅允许填写窗体（内容控件）"],
      ],
      current?.edit ?? "none",
      "编辑限制",
    );
    const enforce = makeCheckbox("立即启动强制保护", current?.enforcement ?? false, "启动强制保护");
    const note = document.createElement("p");
    note.style.cssText = "margin:0;font-size:12px;color:#555";
    note.textContent = "保护设置会写入 OOXML settings.xml 的 w:documentProtection，与 Microsoft Word 互通。";

    stack.append(mode.label, enforce.label, note);
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const editMode = mode.select.value as "none" | "readOnly" | "comments" | "trackedChanges" | "forms";
        const ops: DocumentOperation[] = [
          {
            op: "setWordOptions",
            nodeId: documentRootId(),
            protection: editMode === "none" ? null : { edit: editMode, enforcement: enforce.input.checked },
          },
        ];
        if (editMode === "trackedChanges" && enforce.input.checked) {
          ops.push({
            op: "setWordReview",
            nodeId: documentRootId(),
            trackChanges: true,
            author: session.snapshot().metadata.word?.trackingAuthor ?? "审阅者",
          });
        }
        commit(ops);
        close();
      },
      true,
    );
    ui.show();
  }

  function openOutlineView(): void {
    close();
    replaceSameSideDock(parent, "left");
    controller = new AbortController();
    const signal = controller.signal;
    const chrome = createWordDockedPane(document, {
      ariaLabel: "大纲编排视图",
      title: "大纲视图",
      dock: "left",
      closeLabel: "关闭大纲视图",
      onClose: close,
      signal,
      width: "320px",
    });
    const panel = chrome.panel;
    const body = chrome.body;
    activePanel = panel;

    const collapsed = new Set<string>();

    const renderList = () => {
      body.replaceChildren();
      const subtitle = document.createElement("p");
      subtitle.className = "word-pane-muted";
      subtitle.textContent = "支持标题升降级、折叠展开与整节上下移动，全部操作走统一撤销。";
      body.append(subtitle);

      const snapshot = session.snapshot();
      const mainStory = snapshot.metadata.word?.stories.find((s) => s.kind === "main");
      if (!mainStory) return;
      const sections = snapshot.metadata.sections ?? [];
      const outlineMap = new Map(getWordOutline(session).map((h) => [h.nodeId, h.level]));

      interface OutlineEntry {
        nodeId: string;
        parentId: string;
        level: number | null; // 1..9 or null for body
        text: string;
      }
      const entries: OutlineEntry[] = [];
      for (const sec of sections) {
        const secNode = session.getNode(sec.rootNodeId);
        if (!secNode) continue;
        for (const childId of secNode.childNodeIds) {
          const node = session.getNode(childId);
          if (!node || !/^p$|^h[1-6]$/.test(node.tagName)) continue;
          const level = outlineMap.get(childId) ?? null;
          entries.push({ nodeId: childId, parentId: sec.rootNodeId, level, text: node.text.trim() || "（空段落）" });
        }
      }

      const list = document.createElement("div");
      list.style.cssText = "display:flex;flex-direction:column;gap:4px";
      let hiddenBelowLevel: number | undefined;
      for (let i = 0; i < entries.length; i++) {
        const item = entries[i]!;
        if (hiddenBelowLevel !== undefined) {
          if (item.level !== null && item.level <= hiddenBelowLevel) hiddenBelowLevel = undefined;
          else continue;
        }
        const row = document.createElement("div");
        const indent = item.level !== null ? (item.level - 1) * 18 : 24;
        row.style.cssText = `display:flex;align-items:center;gap:6px;padding:3px 6px 3px ${indent + 6}px;background:#fff;border:1px solid #e5e9f0;border-radius:4px`;

        const hasSubordinates =
          item.level !== null &&
          i + 1 < entries.length &&
          (entries[i + 1]!.level === null || entries[i + 1]!.level! > item.level);

        const toggleBtn = document.createElement("button");
        toggleBtn.type = "button";
        toggleBtn.style.cssText = "width:22px;padding:1px;border:0;background:transparent;cursor:pointer";
        toggleBtn.textContent = item.level === null ? "•" : hasSubordinates ? (collapsed.has(item.nodeId) ? "▸" : "▾") : "◦";
        toggleBtn.setAttribute("aria-label", collapsed.has(item.nodeId) ? "展开章节" : "折叠章节");
        if (hasSubordinates) {
          toggleBtn.addEventListener(
            "click",
            () => {
              if (collapsed.has(item.nodeId)) collapsed.delete(item.nodeId);
              else collapsed.add(item.nodeId);
              renderList();
            },
            { signal },
          );
        }

        const labelBtn = document.createElement("button");
        labelBtn.type = "button";
        labelBtn.style.cssText =
          "flex:1;text-align:left;border:0;background:transparent;cursor:pointer;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:" +
          (item.level !== null ? "600" : "400");
        labelBtn.textContent = (item.level !== null ? `[H${item.level}] ` : "[正文] ") + item.text.slice(0, 60);
        labelBtn.addEventListener("click", () => callbacks.focusNode?.(item.nodeId), { signal });

        const smallBtn = (text: string, titleText: string, disabled: boolean, run: () => void) => {
          const b = document.createElement("button");
          b.type = "button";
          b.textContent = text;
          b.title = titleText;
          b.setAttribute("aria-label", titleText + " " + item.text.slice(0, 12));
          b.disabled = disabled;
          b.style.cssText = "padding:2px 6px;font-size:11px";
          b.addEventListener(
            "click",
            () => {
              try {
                run();
                renderList();
              } catch (err) {
                callbacks.onError(err);
              }
            },
            { signal },
          );
          return b;
        };

        const setLevel = (nextOutline: number | null) => {
          commit([
            {
              op: "setWordParagraph",
              nodeId: item.nodeId,
              outlineLevel: nextOutline,
              ...(nextOutline !== null && nextOutline <= 5 ? { styleId: "Heading" + (nextOutline + 1) } : nextOutline === null ? { styleId: "Normal" } : {}),
            },
          ]);
        };

        // Compute the subtree slice of this item within its parent section.
        const parentNode = session.getNode(item.parentId)!;
        const siblings = [...parentNode.childNodeIds];
        const blockStart = siblings.indexOf(item.nodeId);
        let blockEnd = blockStart + 1;
        if (item.level !== null) {
          while (blockEnd < siblings.length) {
            const sibLvl = outlineMap.get(siblings[blockEnd]!);
            if (sibLvl !== undefined && sibLvl <= item.level) break;
            blockEnd++;
          }
        }
        const moveSection = (direction: -1 | 1) => {
          const slice = siblings.slice(blockStart, blockEnd);
          if (direction === -1) {
            if (blockStart === 0) return;
            // Find previous peer or preceding node in same parent.
            let targetIdx = blockStart - 1;
            if (item.level !== null) {
              while (targetIdx > 0) {
                const lvl = outlineMap.get(siblings[targetIdx]!);
                if (lvl !== undefined && lvl <= item.level) break;
                targetIdx--;
              }
            }
            const nextOrder = [
              ...siblings.slice(0, targetIdx),
              ...slice,
              ...siblings.slice(targetIdx, blockStart),
              ...siblings.slice(blockEnd),
            ];
            commit([{ op: "reorderWordNodes", nodeId: item.parentId, childNodeIds: nextOrder }]);
          } else {
            if (blockEnd >= siblings.length) return;
            let afterEnd = blockEnd + 1;
            const nextLvl = outlineMap.get(siblings[blockEnd]!);
            if (item.level !== null && nextLvl !== undefined) {
              while (afterEnd < siblings.length) {
                const lvl = outlineMap.get(siblings[afterEnd]!);
                if (lvl !== undefined && lvl <= nextLvl) break;
                afterEnd++;
              }
            }
            const nextOrder = [
              ...siblings.slice(0, blockStart),
              ...siblings.slice(blockEnd, afterEnd),
              ...slice,
              ...siblings.slice(afterEnd),
            ];
            commit([{ op: "reorderWordNodes", nodeId: item.parentId, childNodeIds: nextOrder }]);
          }
        };

        row.append(
          toggleBtn,
          labelBtn,
          smallBtn("←", "升级", item.level === 1, () => setLevel(item.level === null ? 0 : Math.max(0, item.level - 2))),
          smallBtn("→", "降级", item.level === 9, () => setLevel(item.level === null ? 0 : Math.min(8, item.level))),
          smallBtn("正文", "降为正文", item.level === null, () => setLevel(null)),
          smallBtn("↑", "上移", blockStart === 0, () => moveSection(-1)),
          smallBtn("↓", "下移", blockEnd >= siblings.length, () => moveSection(1)),
        );
        list.append(row);

        if (hasSubordinates && collapsed.has(item.nodeId)) {
          hiddenBelowLevel = item.level!;
        }
      }
      if (!entries.length) {
        const empty = document.createElement("p");
        empty.className = "word-pane-muted";
        empty.textContent = "文档中暂无段落。";
        list.append(empty);
      }
      body.append(list);
    };

    renderList();
    parent.insertBefore(panel, before);
  }

  function openPasteSpecialDialog(
    onApply: (mode: WordPasteMode | "image") => Promise<void> | void,
  ): void {
    const { ui, signal, addButton } = showDialog("选择性粘贴", "480px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const formats: readonly {
      readonly id: string;
      readonly mode: WordPasteMode | "image";
      readonly label: string;
      readonly desc: string;
    }[] = [
      {
        id: "html",
        mode: "keepSource",
        label: "HTML 格式（保留源格式）",
        desc: "以 HTML 格式插入剪贴板内容，完整保留原有字体、颜色、表格、列表与结构。",
      },
      {
        id: "merge",
        mode: "mergeFormatting",
        label: "带格式文本（合并格式）",
        desc: "保留强调（加粗、斜体、超链接等），其余外观自动匹配目标段落的字体与样式。",
      },
      {
        id: "plain",
        mode: "plainText",
        label: "无格式文本",
        desc: "以无任何格式的纯文本插入剪贴板内容。",
      },
      {
        id: "unicode",
        mode: "plainText",
        label: "无格式 Unicode 文本",
        desc: "以无格式 Unicode 纯文本插入剪贴板内容，完整保留中日韩及多语种字符。",
      },
      {
        id: "image",
        mode: "image",
        label: "图片 (PNG / JPEG)",
        desc: "将剪贴板中的位图图像作为嵌入式 Word 图片对象插入当前光标处。",
      },
    ];

    const listLabel = document.createElement("div");
    listLabel.style.cssText = "font-size:12px;font-weight:600";
    listLabel.textContent = "形式(A)：";

    const select = document.createElement("select");
    select.setAttribute("aria-label", "选择性粘贴形式");
    select.size = 5;
    select.style.cssText = "width:100%;font-size:13px;padding:4px";
    for (const f of formats) {
      const opt = document.createElement("option");
      opt.value = f.id;
      opt.textContent = f.label;
      select.append(opt);
    }
    select.value = "html";

    const descBox = document.createElement("div");
    descBox.style.cssText =
      "font-size:12px;color:#445;background:#f7f9fc;border:1px solid #d5dce6;border-radius:4px;padding:10px;min-height:48px";
    const syncDesc = () => {
      const current = formats.find((f) => f.id === select.value) ?? formats[0]!;
      descBox.textContent = "结果：" + current.desc;
    };
    syncDesc();
    select.addEventListener("change", syncDesc, { signal });

    stack.append(listLabel, select, descBox);
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const chosen = formats.find((f) => f.id === select.value) ?? formats[0]!;
        close();
        void Promise.resolve(onApply(chosen.mode)).catch(callbacks.onError);
      },
      true,
    );
    ui.show();
  }

  return {
    openPageBorders,
    openTocDialog,
    openRubyDialog,
    openEnclosedCharDialog,
    openSymbolsDialog,
    openTableFormulaDialog,
    openDropCapDialog,
    openChineseLayoutDialog,
    openDocPropertiesDialog,
    openProtectionDialog,
    openOutlineView,
    openPasteSpecialDialog,
    dispose: close,
  };
}
