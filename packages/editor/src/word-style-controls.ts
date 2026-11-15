import {
  createOpaqueId,
  WordStyleSchema,
  DocumentError,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
} from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation, WordMetadata } from "@doffi/core";
type Style = WordMetadata["styles"][number];
export function createWordStyleControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  applyStyle: (style: Style) => void,
  onError: (error: unknown) => void,
): {
  open(): void;
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
  function open(): void {
    close();
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
    panel = document.createElement("section");
    panel.setAttribute("role", "region");
    panel.setAttribute("aria-label", "样式管理");
    panel.style.cssText =
      "display:flex;flex-wrap:wrap;gap:8px;padding:12px;background:white;border:1px solid #dbe3ef";
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
            onError(error);
          }
        },
        { signal },
      );
      panel!.append(button);
    };
    const field = (name: string, value: string, type = "text"): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      input.type = type;
      input.value = value;
      input.style.width = "110px";
      label.append(input);
      panel!.append(label);
      return input;
    };
    const choose = (
      name: string,
      options: readonly (readonly [string, string])[],
      value: string,
    ): HTMLSelectElement => {
      const label = document.createElement("label"),
        input = document.createElement("select");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      for (const [value, text] of options) {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = text;
        input.append(option);
      }
      input.value = value;
      label.append(input);
      panel!.append(label);
      return input;
    };
    button("关闭样式管理", close);
    const existing = choose(
      "已有样式",
      styles.map((style) => [style.styleId, style.name ?? style.styleId]),
      current?.styleId ?? "",
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
    button("新建样式", () => {
      selected = "new-" + createOpaqueId();
      tableRegion = "";
      open();
    });
    openDefinition(current);
    function openDefinition(current: (typeof styles)[number] | undefined): void {
      if (!panel) {
        panel = document.createElement("section");
        panel.setAttribute("role", "region");
        panel.setAttribute("aria-label", "样式管理");
        panel.style.cssText =
          "display:flex;flex-wrap:wrap;gap:8px;padding:12px;background:white;border:1px solid #dbe3ef";
        parent.insertBefore(panel, before);
        button("关闭样式管理", close);
      }
      const name = field("样式显示名称", current?.name ?? current?.styleId ?? "新样式"),
        kind = choose(
          "样式类型",
          [
            ["paragraph", "段落"],
            ["character", "字符"],
            ["table", "表格"],
          ],
          current?.kind ?? "paragraph",
        );
      if (current) kind.disabled = true;
      const peers = styles.filter(
          (style) =>
            style.kind === (current?.kind ?? "paragraph") && style.styleId !== current?.styleId,
        ),
        base = choose(
          "基准样式",
          [
            ["", "无"],
            ...peers.map((style) => [style.styleId, style.name ?? style.styleId] as const),
          ],
          current?.basedOn ?? "",
        ),
        next = choose(
          "后续段落样式",
          [
            ["", "当前样式"],
            ...styles
              .filter((style) => style.kind === "paragraph")
              .map((style) => [style.styleId, style.name ?? style.styleId] as const),
          ],
          current?.nextStyleId ?? "",
        ),
        outline = choose(
          "大纲级别",
          [
            ["", "继承"],
            ["9", "正文"],
            ...Array.from({ length: 9 }, (_, index) => [String(index), String(index + 1)] as const),
          ],
          current?.outlineLevel === undefined ? "" : String(current.outlineLevel),
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
      const linked = choose(
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
      );
      const defaults = field("设为默认", "", "checkbox");
      defaults.checked = current?.default ?? false;
      const gallery = field("显示在样式库", "", "checkbox");
      gallery.checked = current?.qFormat ?? true;
      const flags = (["hidden", "semiHidden", "unhideWhenUsed", "locked"] as const).map(
        (key, index) => {
          const input = field(
            ["隐藏样式", "默认隐藏", "使用时显示", "锁定样式"][index]!,
            "",
            "checkbox",
          );
          input.checked = current?.[key] ?? false;
          return { key, input };
        },
      );
      const aliases = field("样式别名", current?.aliases ?? ""),
        priority = field("显示顺序", String(current?.uiPriority ?? 99), "number");
      priority.min = "0";
      priority.max = "65535";
      const hint = document.createElement("span");
      hint.textContent = "以下留空保留当前设置；勾选恢复继承可取消本样式覆盖。";
      panel!.append(hint);
      const font = field("样式字体", ""),
        size = field("样式字号（磅）", "", "number"),
        color = field("样式文字颜色", ""),
        weight = choose(
          "样式字重",
          [
            ["", "保留"],
            ["400", "常规"],
            ["700", "加粗"],
          ],
          "",
        );
      const languages = (["val", "eastAsia", "bidi"] as const).map((role, index) => ({
        role,
        input: field(["样式西文语言", "样式东亚语言", "样式复杂文字语言"][index]!, ""),
      }));
      const alignment = choose(
        "样式对齐",
        [
          ["", "保留"],
          ["left", "左"],
          ["center", "中"],
          ["right", "右"],
          ["justify", "两端"],
        ],
        "",
      );
      const space = field("样式段后（磅）", "", "number"),
        resetFont = field("字体格式恢复继承", "", "checkbox"),
        resetParagraph = field("段落格式恢复继承", "", "checkbox");
      const cellFill = field("表格样式单元格底纹", ""),
        cellBorder = field("表格样式单元格边框", ""),
        cellPadding = field("表格样式单元格边距（磅）", "", "number"),
        tableBorder = field("表格样式外边框", "");
      const region = choose(
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
      );
      region.addEventListener(
        "change",
        () => {
          tableRegion = region.value as typeof tableRegion;
          open();
        },
        { signal },
      );
      const rowSplit = choose(
          "样式跨页断行",
          [
            ["", "保留"],
            ["allow", "允许"],
            ["avoid", "禁止"],
          ],
          "",
        ),
        rowHeader = choose(
          "样式重复表头",
          [
            ["", "保留"],
            ["1", "重复"],
            ["0", "不重复"],
          ],
          "",
        );
      const resetTable = field("表格格式恢复继承", "", "checkbox"),
        rowBand = field("样式行条带大小", String(current?.rowBandSize ?? ""), "number"),
        columnBand = field("样式列条带大小", String(current?.columnBandSize ?? ""), "number");
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
      button("保存样式", () => {
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
      });
      if (current?.kind === "table" && tableRegion)
        button("删除当前条件格式", () => {
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
        });
      if (current) {
        button("应用到选区", () => applyStyle(WordStyleSchema.parse(current)));
        const replacement = choose(
          "删除后替换为",
          peers.map((style) => [style.styleId, style.name ?? style.styleId]),
          peers[0]?.styleId ?? "",
        );
        button("删除样式", () => {
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
        });
      }
    }
  }
  return { open, dispose: close };
}
