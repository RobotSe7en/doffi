import {
  createOpaqueId,
  WordStyleSchema,
  DocumentError,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
} from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation, WordMetadata } from "@doffi/core";
import { createWordDockedPane, replaceSameSideDock } from "./word-pane-chrome.js";

type Style = WordMetadata["styles"][number];

const STYLE_PREVIEW_META: Readonly<Record<string, { label: string; css: string }>> = {
  Normal: { label: "正文", css: "font-size:13px;font-weight:400;color:#222" },
  NoSpacing: { label: "无间隔", css: "font-size:13px;font-weight:400;color:#222" },
  Heading1: { label: "标题 1", css: "font-size:15px;font-weight:700;color:#2f5496" },
  Heading2: { label: "标题 2", css: "font-size:13.5px;font-weight:700;color:#2f5496" },
  Heading3: { label: "标题 3", css: "font-size:12.5px;font-weight:700;color:#1f3763" },
  Heading4: { label: "标题 4", css: "font-size:12px;font-weight:600;font-style:italic;color:#2f5496" },
  Title: { label: "标题", css: "font-size:16px;font-weight:700;color:#1f1f1f" },
  Subtitle: { label: "副标题", css: "font-size:12.5px;font-style:italic;color:#595959" },
  Quote: { label: "引用", css: "font-size:12.5px;font-style:italic;color:#404040" },
  IntenseQuote: { label: "明显引用", css: "font-size:12.5px;font-style:italic;font-weight:600;color:#2f5496" },
  Strong: { label: "要点", css: "font-size:13px;font-weight:700;color:#111" },
  Emphasis: { label: "强调", css: "font-size:13px;font-style:italic;color:#222" },
};

export function createWordStyleControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  applyStyle: (style: Style) => void,
  onError: (error: unknown) => void,
): {
  open(): void;
  toggle(): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let panel: HTMLElement | undefined,
    controller: AbortController | undefined,
    selected: string | undefined,
    tableRegion: NonNullable<Style["conditions"]>[number]["region"] | "" = "";

  const close = (): void => {
    controller?.abort();
    panel?.remove();
    panel = undefined;
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

  function toggle(): void {
    if (panel?.isConnected) close();
    else open();
  }

  function open(): void {
    close();
    replaceSameSideDock(parent, "right");
    controller = new AbortController();
    const signal = controller.signal,
      snapshot = session.snapshot(),
      word = snapshot.metadata.word,
      styles = word?.styles ?? [],
      current = selected?.startsWith("new-")
        ? undefined
        : (styles.find((style) => style.styleId === selected) ?? styles[0]),
      root = new DOMParser()
        .parseFromString(snapshot.html, "text/html")
        .querySelector<HTMLElement>("article")!.dataset.nodeId!;

    const chrome = createWordDockedPane(document, {
      ariaLabel: "样式管理",
      title: "样式",
      dock: "right",
      closeLabel: "关闭样式管理",
      onClose: close,
      signal,
      width: "320px",
    });
    panel = chrome.panel;
    const body = chrome.body;
    parent.insertBefore(panel, before);

    const makeButton = (
      label: string,
      run: () => void,
      container: HTMLElement,
      variant: "default" | "primary" | "danger" = "default",
    ): HTMLButtonElement => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = label;
      if (variant === "primary") btn.className = "word-btn-primary";
      else if (variant === "danger") btn.className = "word-btn-danger";
      btn.addEventListener(
        "click",
        () => {
          try {
            run();
          } catch (error) {
            onError(error);
          }
        },
        { signal },
      );
      container.append(btn);
      return btn;
    };

    const makeField = (
      name: string,
      value: string,
      container: HTMLElement,
      type = "text",
      wide = false,
    ): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      if (wide) label.className = "wide";
      input.setAttribute("aria-label", name);
      input.type = type;
      input.value = value;
      if (type === "checkbox") {
        label.append(input, name);
      } else {
        label.textContent = name;
        label.append(input);
      }
      container.append(label);
      return input;
    };

    const makeSelect = (
      name: string,
      options: readonly (readonly [string, string])[],
      value: string,
      container: HTMLElement,
      wide = false,
    ): HTMLSelectElement => {
      const label = document.createElement("label"),
        input = document.createElement("select");
      if (wide) label.className = "wide";
      label.textContent = name;
      input.setAttribute("aria-label", name);
      for (const [val, text] of options) {
        const option = document.createElement("option");
        option.value = val;
        option.textContent = text;
        input.append(option);
      }
      input.value = value;
      label.append(input);
      container.append(label);
      return input;
    };

    // 1. Word 1:1 Quick Styles List (live preview rows, click to select, dblclick to apply)
    const listSection = document.createElement("div");
    listSection.className = "word-pane-section";
    const listHeader = document.createElement("div");
    listHeader.className = "word-pane-row";
    listHeader.style.justifyContent = "space-between";
    const listTitle = document.createElement("span");
    listTitle.className = "word-pane-section-title";
    listTitle.textContent = "选择要应用的样式";
    const topActions = document.createElement("div");
    topActions.className = "word-pane-row";
    makeButton("新建样式", () => {
      selected = "new-" + createOpaqueId();
      tableRegion = "";
      open();
    }, topActions);
    listHeader.append(listTitle, topActions);
    listSection.append(listHeader);

    const visualList = document.createElement("div");
    visualList.className = "word-pane-list";
    visualList.setAttribute("role", "listbox");
    visualList.setAttribute("aria-label", "文档样式列表");
    const kindBadge = { paragraph: "¶", character: "a", table: "⊞" };
    const existingIds = new Set(styles.map((s) => s.styleId));
    const displayStyles = [
      ...styles,
      ...Object.entries(STYLE_PREVIEW_META)
        .filter(([styleId]) => !existingIds.has(styleId))
        .map(([styleId, meta]) =>
          WordStyleSchema.parse({
            styleId,
            kind: "paragraph",
            name: meta.label,
            cssClass: "word-style-" + styleId.toLowerCase(),
          }),
        ),
    ];
    for (const st of displayStyles) {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "word-style-item";
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", String(st.styleId === (current?.styleId ?? selected)));
      const preview = document.createElement("span");
      const meta = STYLE_PREVIEW_META[st.styleId];
      preview.textContent = meta?.label || st.name || st.styleId;
      if (meta) preview.style.cssText = meta.css;
      else preview.style.cssText = "font-size:12.5px;color:#222";
      preview.style.overflow = "hidden";
      preview.style.textOverflow = "ellipsis";
      preview.style.whiteSpace = "nowrap";
      const badge = document.createElement("span");
      badge.className = "word-style-badge";
      badge.textContent = (kindBadge as Record<string, string>)[st.kind] ?? "¶";
      badge.title = st.kind === "paragraph" ? "段落样式" : st.kind === "character" ? "字符样式" : st.kind === "numbering" ? "编号样式" : "表格样式";
      item.append(preview, badge);
      item.addEventListener(
        "click",
        () => {
          selected = st.styleId;
          tableRegion = "";
          try {
            applyStyle(WordStyleSchema.parse(st));
          } catch (error) {
            onError(error);
          }
          open();
        },
        { signal },
      );
      visualList.append(item);
    }
    listSection.append(visualList);

    const existing = makeSelect(
      "已有样式",
      styles.map((style) => [style.styleId, style.name ?? style.styleId]),
      current?.styleId ?? "",
      listSection,
      true,
    );
    existing.addEventListener(
      "change",
      () => {
        selected = existing.value;
        tableRegion = "";
        open();
      },
      { signal },
    );

    const quickApplyRow = document.createElement("div");
    quickApplyRow.className = "word-pane-row";
    if (current) {
      makeButton("应用到选区", () => applyStyle(WordStyleSchema.parse(current)), quickApplyRow, "primary");
    }
    listSection.append(quickApplyRow);
    body.append(listSection);

    // 2. Modify Style Core Properties Card (always accessible for editing & existing tests)
    const editSection = document.createElement("div");
    editSection.className = "word-pane-section";
    const editTitle = document.createElement("span");
    editTitle.className = "word-pane-section-title";
    editTitle.textContent = current ? `修改样式：${current.name ?? current.styleId}` : "新建样式属性";
    editSection.append(editTitle);

    const coreGrid = document.createElement("div");
    coreGrid.className = "word-pane-grid";
    editSection.append(coreGrid);

    const name = makeField("样式显示名称", current?.name ?? current?.styleId ?? "新样式", coreGrid, "text", true),
      kind = makeSelect(
        "样式类型",
        [
          ["paragraph", "段落"],
          ["character", "字符"],
          ["table", "表格"],
        ],
        current?.kind ?? "paragraph",
        coreGrid,
      );
    if (current) kind.disabled = true;
    const peers = styles.filter(
        (style) =>
          style.kind === (current?.kind ?? "paragraph") && style.styleId !== current?.styleId,
      ),
      base = makeSelect(
        "基准样式",
        [
          ["", "无"],
          ...peers.map((style) => [style.styleId, style.name ?? style.styleId] as const),
        ],
        current?.basedOn ?? "",
        coreGrid,
      ),
      next = makeSelect(
        "后续段落样式",
        [
          ["", "当前样式"],
          ...styles
            .filter((style) => style.kind === "paragraph")
            .map((style) => [style.styleId, style.name ?? style.styleId] as const),
        ],
        current?.nextStyleId ?? "",
        coreGrid,
      ),
      outline = makeSelect(
        "大纲级别",
        [
          ["", "继承"],
          ["9", "正文"],
          ...Array.from({ length: 9 }, (_, index) => [String(index), String(index + 1)] as const),
        ],
        current?.outlineLevel === undefined ? "" : String(current.outlineLevel),
        coreGrid,
      );
    kind.addEventListener(
      "change",
      () => {
        base.replaceChildren();
        for (const style of [
          { styleId: "", name: "无" },
          ...styles.filter((style) => style.kind === kind.value),
        ]) {
          const option = document.createElement("option");
          option.value = style.styleId;
          option.textContent = style.name ?? style.styleId;
          base.append(option);
        }
      },
      { signal },
    );

    const hint = document.createElement("p");
    hint.className = "word-pane-muted";
    hint.textContent = "以下格式留空保留当前值；在高级选项中可恢复继承。";
    editSection.append(hint);

    const fmtGrid = document.createElement("div");
    fmtGrid.className = "word-pane-grid";
    editSection.append(fmtGrid);

    const font = makeField("样式字体", "", fmtGrid, "text"),
      size = makeField("样式字号（磅）", "", fmtGrid, "number"),
      color = makeField("样式文字颜色", "", fmtGrid, "text"),
      weight = makeSelect(
        "样式字重",
        [
          ["", "保留"],
          ["400", "常规"],
          ["700", "加粗"],
        ],
        "",
        fmtGrid,
      ),
      alignment = makeSelect(
        "样式对齐",
        [
          ["", "保留"],
          ["left", "左"],
          ["center", "中"],
          ["right", "右"],
          ["justify", "两端"],
        ],
        "",
        fmtGrid,
      ),
      space = makeField("样式段后（磅）", "", fmtGrid, "number");

    const saveRow = document.createElement("div");
    saveRow.className = "word-pane-row";
    saveRow.style.justifyContent = "flex-end";
    editSection.append(saveRow);
    body.append(editSection);

    // 3. Advanced Options (cleanly collapsed so Styles Pane stays uncluttered like Word!)
    const advDetails = document.createElement("details");
    const advSummary = document.createElement("summary");
    advSummary.textContent = "高级选项（样式标志 / 语言 / 继承 / 表格）";
    const advGrid = document.createElement("div");
    advGrid.className = "word-pane-grid";
    advDetails.append(advSummary, advGrid);
    body.append(advDetails);

    const linked = makeSelect(
      "链接样式",
      [
        ["", "无"],
        ...styles
          .filter(
            (style) => style.kind === (kind.value === "paragraph" ? "character" : "paragraph"),
          )
          .map((style) => [style.styleId, style.name ?? style.styleId] as const),
      ],
      current?.linkedStyleId ?? "",
      advGrid,
      true,
    );
    const defaults = makeField("设为默认", "", advGrid, "checkbox");
    defaults.checked = current?.default ?? false;
    const gallery = makeField("显示在样式库", "", advGrid, "checkbox");
    gallery.checked = current?.qFormat ?? true;
    const flags = (["hidden", "semiHidden", "unhideWhenUsed", "locked"] as const).map(
      (key, index) => {
        const input = makeField(
          ["隐藏样式", "默认隐藏", "使用时显示", "锁定样式"][index]!,
          "",
          advGrid,
          "checkbox",
        );
        input.checked = current?.[key] ?? false;
        return { key, input };
      },
    );
    const aliases = makeField("样式别名", current?.aliases ?? "", advGrid, "text"),
      priority = makeField("显示顺序", String(current?.uiPriority ?? 99), advGrid, "number");
    priority.min = "0";
    priority.max = "65535";
    const languages = (["val", "eastAsia", "bidi"] as const).map((role, index) => ({
      role,
      input: makeField(["样式西文语言", "样式东亚语言", "样式复杂文字语言"][index]!, "", advGrid, "text"),
    }));
    const resetFont = makeField("字体格式恢复继承", "", advGrid, "checkbox"),
      resetParagraph = makeField("段落格式恢复继承", "", advGrid, "checkbox");

    const cellFill = makeField("表格样式单元格底纹", "", advGrid, "text"),
      cellBorder = makeField("表格样式单元格边框", "", advGrid, "text"),
      cellPadding = makeField("表格样式单元格边距（磅）", "", advGrid, "number"),
      tableBorder = makeField("表格样式外边框", "", advGrid, "text");
    const region = makeSelect(
      "表格样式区域",
      [
        ["", "整张表"],
        ["firstRow", "首行"],
        ["lastRow", "末行"],
        ["firstCol", "首列"],
        ["lastCol", "末列"],
        ["band1Horz", "奇数行条带"],
        ["band2Horz", "偶数行条带"],
        ["band1Vert", "奇数列条带"],
        ["band2Vert", "偶数列条带"],
        ["nwCell", "左上角"],
        ["neCell", "右上角"],
        ["swCell", "左下角"],
        ["seCell", "右下角"],
      ],
      tableRegion,
      advGrid,
      true,
    );
    region.addEventListener(
      "change",
      () => {
        tableRegion = region.value as typeof tableRegion;
        open();
      },
      { signal },
    );
    const rowSplit = makeSelect(
        "样式跨页断行",
        [
          ["", "保留"],
          ["allow", "允许"],
          ["avoid", "禁止"],
        ],
        "",
        advGrid,
      ),
      rowHeader = makeSelect(
        "样式重复表头",
        [
          ["", "保留"],
          ["1", "重复"],
          ["0", "不重复"],
        ],
        "",
        advGrid,
      );
    const resetTable = makeField("表格格式恢复继承", "", advGrid, "checkbox"),
      rowBand = makeField("样式行条带大小", String(current?.rowBandSize ?? ""), advGrid, "number"),
      columnBand = makeField("样式列条带大小", String(current?.columnBandSize ?? ""), advGrid, "number");
    rowBand.min = columnBand.min = "0";

    const tableFields = [
      cellFill,
      cellBorder,
      cellPadding,
      tableBorder,
      region,
      rowSplit,
      rowHeader,
      resetTable,
      rowBand,
      columnBand,
    ];
    const updateTableFields = () => {
      for (const input of tableFields) input.parentElement!.hidden = kind.value !== "table";
    };
    kind.addEventListener("change", updateTableFields, { signal });
    updateTableFields();

    makeButton("保存样式", () => {
      const id = current?.styleId ?? selected ?? "style-" + createOpaqueId(),
        definition: Style = current
          ? WordStyleSchema.parse(current)
          : {
              styleId: id,
              kind: kind.value as Style["kind"],
              cssClass: "word-style-" + createOpaqueId().replaceAll("-", ""),
              customStyle: true,
            };
      if (linked.value) definition.linkedStyleId = linked.value;
      else delete definition.linkedStyleId;
      for (const { key, input } of flags) definition[key] = input.checked;
      definition.name = name.value;
      definition.aliases = aliases.value;
      definition.default = defaults.checked;
      definition.qFormat = gallery.checked;
      definition.uiPriority = Number(priority.value);
      if (base.value) definition.basedOn = base.value;
      else delete definition.basedOn;
      if (next.value && kind.value === "paragraph") definition.nextStyleId = next.value;
      else delete definition.nextStyleId;
      if (outline.value !== "" && kind.value === "paragraph")
        definition.outlineLevel = Number(outline.value);
      else delete definition.outlineLevel;
      const run: Record<string, string> = {},
        paragraph: Record<string, string> = {};
      if (font.value.trim()) {
        run["font-family"] = font.value;
        for (const role of ["ascii", "hAnsi", "eastAsia", "cs"])
          run["--word-font-" + role] = font.value;
      }
      if (size.value) run["font-size"] = size.value + "pt";
      if (color.value) run.color = color.value;
      for (const { role, input } of languages)
        if (input.value.trim())
          run["--word-language-" + role] = JSON.stringify(input.value.trim());
      if (weight.value) run["font-weight"] = weight.value;
      if (alignment.value) paragraph["text-align"] = alignment.value;
      if (space.value) paragraph["margin-bottom"] = space.value + "pt";
      const cell: Record<string, string> = {},
        table: Record<string, string> = {};
      if (cellFill.value) cell["background-color"] = cellFill.value;
      if (cellBorder.value) cell.border = cellBorder.value;
      if (cellPadding.value) cell.padding = cellPadding.value + "pt";
      if (tableBorder.value) table.border = tableBorder.value;
      const row: Record<string, string> = {};
      if (rowSplit.value) row["break-inside"] = rowSplit.value === "avoid" ? "avoid" : "auto";
      if (rowHeader.value) row["--word-row-repeat-header"] = rowHeader.value;
      if (kind.value === "table") {
        if (rowBand.value !== "") definition.rowBandSize = Number(rowBand.value);
        if (columnBand.value !== "") definition.columnBandSize = Number(columnBand.value);
      }
      const resolver = new SourceStyleResolver(
          snapshot.styles.map((sheet) => sheet.css),
          DEFAULT_DOCUMENT_READ_LIMITS,
        ),
        condition = current?.conditions?.find((condition) => condition.region === tableRegion),
        components = tableRegion ? condition?.components : current?.components;
      const keys = (scope: "p" | "r" | "table" | "cell" | "row") =>
        components?.[scope] ? [...resolver.classKeys(components[scope]!)] : [];
      const patch = {
        runProperties: run,
        paragraphProperties: paragraph,
        ...(kind.value === "table"
          ? { cellProperties: cell, tableProperties: table, rowProperties: row }
          : {}),
        ...(resetFont.checked
          ? { removeRunProperties: tableRegion ? keys("r") : [...(current?.runProperties ?? [])] }
          : {}),
        ...(resetParagraph.checked
          ? {
              removeParagraphProperties: tableRegion
                ? keys("p")
                : [...(current?.paragraphProperties ?? [])],
            }
          : {}),
        ...(resetTable.checked
          ? {
              removeTableProperties: keys("table"),
              removeCellProperties: keys("cell"),
              removeRowProperties: keys("row"),
            }
          : {}),
      };
      commit([
        {
          op: "setWordStyle",
          nodeId: root,
          style: definition,
          ...(kind.value === "table" && tableRegion
            ? { tableCondition: { region: tableRegion, ...patch } }
            : patch),
        },
      ]);
      selected = id;
      open();
    }, saveRow, "primary");

    if (current?.kind === "table" && tableRegion) {
      makeButton("删除当前条件格式", () => {
        commit([
          {
            op: "setWordStyle",
            nodeId: root,
            style: WordStyleSchema.parse(current),
            tableCondition: {
              region: tableRegion as NonNullable<Style["conditions"]>[number]["region"],
              clear: true,
            },
          },
        ]);
        open();
      }, saveRow, "danger");
    }

    // 4. Delete Style section
    if (current) {
      const delSection = document.createElement("div");
      delSection.className = "word-pane-section";
      const replacement = makeSelect(
        "删除后替换为",
        peers.map((style) => [style.styleId, style.name ?? style.styleId]),
        peers[0]?.styleId ?? "",
        delSection,
        true,
      );
      const delRow = document.createElement("div");
      delRow.className = "word-pane-row";
      makeButton("删除样式", () => {
        if (!replacement.value)
          throw new DocumentError("InvalidOperation", "请先选择同类型替代样式。");
        commit([
          {
            op: "deleteWordStyle",
            nodeId: root,
            styleId: current.styleId,
            replacementStyleId: replacement.value,
          },
        ]);
        selected = replacement.value;
        open();
      }, delRow, "danger");
      delSection.append(delRow);
      body.append(delSection);
    }
  }

  return { open, toggle, dispose: close };
}
