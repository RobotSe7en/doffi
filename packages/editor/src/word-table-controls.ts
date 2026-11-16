import { wordRowExceptionProperties } from "@doffi/core";
import {
  DocumentError,
  resolveSourceCssLength,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
} from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation } from "@doffi/core";
export interface WordTableSelection {
  readonly paragraphId?: string;
  readonly tableId?: string;
  readonly cellId?: string;
}
/** Transient Word table forms delegate every write to the shared source transaction. */
export function createWordTableControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  selection: () => WordTableSelection,
  onError: (error: unknown) => void,
): { open(insert: boolean): void; dispose(): void } {
  const document = before.ownerDocument;
  let panel: HTMLElement | undefined, controller: AbortController | undefined;
  const close = (): void => {
    controller?.abort();
    controller = undefined;
    panel?.remove();
    panel = undefined;
  };
  function open(insert: boolean, fixed?: WordTableSelection): void {
    close();
    const selected = fixed ?? selection();
    if (insert && !selected.paragraphId)
      throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    if (!insert && !selected.tableId)
      throw new DocumentError("InvalidOperation", "请先选择表格中的单元格。");
    controller = new AbortController();
    const signal = controller.signal;
    panel = document.createElement("section");
    panel.setAttribute("role", "region");
    panel.setAttribute("aria-label", insert ? "插入表格" : "表格设置");
    panel.style.cssText =
      "padding:12px;background:white;border:1px solid #dbe3ef;max-height:300px;overflow:auto;display:flex;flex-wrap:wrap;align-items:center;gap:8px";
    parent.insertBefore(panel, before);
    const apply = (operations: DocumentOperation[]): void => {
      const snapshot = session.snapshot();
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations,
        }),
      );
      if (!insert)
        queueMicrotask(() => {
          if (panel && session.getNode(selected.tableId!)) open(false, selected);
        });
    };
    const button = (name: string, run: () => void): void => {
      const node = document.createElement("button");
      node.type = "button";
      node.textContent = name;
      node.addEventListener(
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
      panel!.append(node);
    };
    const field = (name: string, value: number | string, type = "number"): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      input.type = type;
      input.value = String(value);
      input.style.width = type === "number" ? "70px" : "120px";
      if (type === "number") {
        input.min = "0";
        input.step = "1";
      }
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
      for (const [key, text] of options) {
        const option = document.createElement("option");
        option.value = key;
        option.textContent = text;
        input.append(option);
      }
      input.value = value;
      label.append(input);
      panel!.append(label);
      return input;
    };
    button("关闭表格设置", close);
    if (insert) {
      const rows = field("表格行数", 3),
        columns = field("表格列数", 3);
      rows.min = columns.min = "1";
      button("插入表格", () => {
        apply([
          {
            op: "insertWordTable",
            nodeId: selected.paragraphId!,
            position: "after",
            rows: Number(rows.value),
            columns: Number(columns.value),
          },
        ]);
        close();
      });
      return;
    }
    const tableId = selected.tableId!,
      table = session.getNode(tableId)!;
    const rowIds = table.childNodeIds.flatMap((id) => {
      const node = session.getNode(id)!;
      return node.tagName === "tr" ? [id] : node.childNodeIds;
    });
    const grid = session
      .snapshot()
      .metadata.word?.tables.find((table) => table.nodeId === tableId)?.grid;
    const cells: { id: string; row: number; column: number; rows: number; columns: number }[] = [],
      occupied: string[][] = Array.from({ length: rowIds.length }, () => []);
    let columnCount = 0;
    for (const [r, id] of rowIds.entries())
      for (const cellId of session.getNode(id)!.childNodeIds) {
        const cell = session.getNode(cellId)!;
        let c = 0;
        while (occupied[r]![c]) c++;
        const rows = Number(cell.attributes.rowspan ?? 1) || rowIds.length - r,
          columns = Number(cell.attributes.colspan ?? 1);
        cells.push({ id: cellId, row: r, column: c, rows, columns });
        for (let y = r; y < r + rows; y++)
          for (let x = c; x < c + columns; x++) occupied[y]![x] = cellId;
        columnCount = Math.max(columnCount, c + columns);
      }
    const active = cells.find((cell) => cell.id === selected.cellId) ?? cells[0]!;
    const row = field("起始行", active.row + 1),
      column = field("起始列", active.column + 1),
      rows = field("选区行数", active.rows),
      columns = field("选区列数", active.columns);
    row.min = column.min = rows.min = columns.min = "1";
    const change = (value: Extract<DocumentOperation, { op: "editWordTable" }>["change"]): void => {
      apply([{ op: "editWordTable", nodeId: tableId, change: value }]);
    };
    const mergeNotice = document.createElement("p");
    mergeNotice.textContent =
      "记录修订时，合并或拆分单元格不会记录为修订；可通过撤销恢复。若已有相关修订导致冲突，请先接受或拒绝。";
    mergeNotice.style.cssText = "width:100%;margin:0;color:#715100";
    panel.append(mergeNotice);
    button("在上方插入行", () =>
      change({ kind: "insertRows", at: Number(row.value) - 1, count: Number(rows.value) }),
    );
    button("在下方插入行", () =>
      change({ kind: "insertRows", at: Number(row.value), count: Number(rows.value) }),
    );
    button("删除表格行", () =>
      change({ kind: "deleteRows", at: Number(row.value) - 1, count: Number(rows.value) }),
    );
    button("在左侧插入列", () =>
      change({ kind: "insertColumns", at: Number(column.value) - 1, count: Number(columns.value) }),
    );
    button("在右侧插入列", () =>
      change({ kind: "insertColumns", at: Number(column.value), count: Number(columns.value) }),
    );
    button("删除表格列", () =>
      change({ kind: "deleteColumns", at: Number(column.value) - 1, count: Number(columns.value) }),
    );
    button("合并单元格", () =>
      change({
        kind: "mergeCells",
        row: Number(row.value) - 1,
        column: Number(column.value) - 1,
        rows: Number(rows.value),
        columns: Number(columns.value),
      }),
    );
    button("拆分单元格", () =>
      change({ kind: "splitCell", row: Number(row.value) - 1, column: Number(column.value) - 1 }),
    );
    const splitRows = field("拆分后的行数", 2),
      splitColumns = field("拆分后的列数", 2);
    splitRows.min = splitColumns.min = "1";
    button("按指定行列拆分", () =>
      change({
        kind: "splitCell",
        row: Number(row.value) - 1,
        column: Number(column.value) - 1,
        rows: Number(splitRows.value),
        columns: Number(splitColumns.value),
      }),
    );
    const widths = field(
      "各列宽度（磅，以逗号分隔）",
      (grid ?? Array.from({ length: columnCount }, () => 120)).join(","),
      "text",
    );
    button("设置列宽", () => change({ kind: "resize", grid: widths.value.split(",").map(Number) }));
    button("平均分布各列", () => {
      const latest =
        session.snapshot().metadata.word?.tables.find((table) => table.nodeId === tableId)?.grid ??
        grid;
      if (!latest) return;
      change({
        kind: "resize",
        grid: latest.map(() => latest.reduce((a, b) => a + b, 0) / latest.length),
      });
    });
    const heightRule = choose(
        "行高规则",
        [
          ["atLeast", "最小值"],
          ["exact", "固定值"],
          ["auto", "自动"],
        ],
        /--word-row-height-rule\s*:\s*(auto|atLeast|exact)/.exec(
          session.getNode(rowIds[0]!)!.attributes.style ?? "",
        )?.[1] ?? "atLeast",
      ),
      height = field("行高（磅）", 20);
    button("设置行高", () => {
      const latest = session.getNode(tableId)!;
      const count = latest.childNodeIds.flatMap((id) => {
        const node = session.getNode(id)!;
        return node.tagName === "tr" ? [id] : node.childNodeIds;
      }).length;
      apply([
        {
          op: "editWordTable",
          nodeId: tableId,
          change: {
            kind: "resize",
            rowHeights: Array.from({ length: count }, () => Number(height.value)),
          },
        },
        {
          op: "formatWordTable",
          nodeId: tableId,
          rowHeightRule: heightRule.value as "atLeast" | "exact" | "auto",
        },
      ]);
    });
    const rowStyleResolver = new SourceStyleResolver(
        session.snapshot().styles.map((sheet) => sheet.css),
        DEFAULT_DOCUMENT_READ_LIMITS,
      ),
      projectRow = (id: string): Parameters<typeof rowStyleResolver.get>[0] => {
        const node = session.getNode(id)!;
        return {
          tag: node.tagName,
          attrs: node.attributes,
          parent: node.parentNodeId ? projectRow(node.parentNodeId) : undefined,
        };
      };
    let headerCount = 0;
    for (const id of rowIds) {
      const row = session.getNode(id)!,
        direct = row.attributes["data-word-repeat-header"];
      if (
        direct !== undefined
          ? direct !== "true"
          : rowStyleResolver.get(projectRow(id))["--word-row-repeat-header"] !== "1"
      )
        break;
      headerCount++;
    }
    const tableDirection = choose(
      "表格排列方向",
      [
        ["0", "从左到右"],
        ["1", "从右到左"],
      ],
      rowStyleResolver.get(projectRow(tableId))["--word-table-bidi"] ?? "0",
    );
    button("应用表格方向", () =>
      apply([
        {
          op: "formatWordTable",
          nodeId: tableId,
          properties: { "--word-table-bidi": tableDirection.value },
        },
      ]),
    );
    const position = rowStyleResolver.get(projectRow(tableId)),
      floating = choose(
        "表格文字环绕",
        [
          ["0", "无"],
          ["1", "环绕"],
        ],
        position["--word-table-floating"] ?? "0",
      ),
      anchorOptions = [
        ["text", "正文"],
        ["margin", "页边距"],
        ["page", "页面"],
      ] as const,
      hAnchor = choose(
        "表格水平参照",
        anchorOptions,
        position["--word-table-position-h-anchor"] ?? "text",
      ),
      vAnchor = choose(
        "表格垂直参照",
        anchorOptions,
        position["--word-table-position-v-anchor"] ?? "margin",
      ),
      hAlign = choose(
        "表格水平定位",
        [
          ["offset", "指定距离"],
          ["left", "左侧"],
          ["center", "居中"],
          ["right", "右侧"],
          ["inside", "内侧"],
          ["outside", "外侧"],
        ],
        position["--word-table-position-x-align"] ?? "offset",
      ),
      vAlign = choose(
        "表格垂直定位",
        [
          ["offset", "指定距离"],
          ["inline", "行内"],
          ["top", "顶部"],
          ["center", "居中"],
          ["bottom", "底部"],
          ["inside", "内侧"],
          ["outside", "外侧"],
        ],
        position["--word-table-position-y-align"] ?? "offset",
      );
    const coordinate = (key: string): number => {
        const amount = resolveSourceCssLength(position["--word-table-position-" + key] ?? "0pt", {
          font: 12,
          percent: 468,
        });
        return amount > 0 ? amount - 0.05 : amount;
      },
      px = field("表格水平位置（磅）", coordinate("x")),
      py = field("表格垂直位置（磅）", coordinate("y"));
    px.min = py.min = "-1638";
    px.step = py.step = ".05";
    const distances = (["left", "right", "top", "bottom"] as const).map((side, index) => ({
        side,
        input: field(
          "表格距正文" + ["左", "右", "上", "下"][index] + "（磅）",
          resolveSourceCssLength(position["--word-table-position-" + side + "-distance"] ?? "0pt", {
            font: 12,
            percent: 468,
          }),
        ),
      })),
      overlap = choose(
        "浮动表格重叠",
        [
          ["overlap", "允许"],
          ["never", "不允许"],
        ],
        position["--word-table-overlap"] ?? "overlap",
      );
    button("应用表格位置", () => {
      const properties: Record<string, string> = {
        "--word-table-floating": floating.value,
        "--word-table-position-h-anchor": hAnchor.value,
        "--word-table-position-v-anchor": vAnchor.value,
        "--word-table-position-x":
          (Number(px.value) >= 0 ? Number(px.value) + 0.05 : Number(px.value)) + "pt",
        "--word-table-position-y":
          (Number(py.value) >= 0 ? Number(py.value) + 0.05 : Number(py.value)) + "pt",
        "--word-table-position-x-align": hAlign.value === "offset" ? "none" : hAlign.value,
        "--word-table-position-y-align": vAlign.value === "offset" ? "none" : vAlign.value,
        "--word-table-overlap": overlap.value,
      };
      for (const { side, input } of distances) {
        const amount = Number(input.value);
        if (!Number.isFinite(amount) || amount < 0)
          throw new DocumentError("InvalidOperation", "表格与正文的距离必须为非负数。");
        properties["--word-table-position-" + side + "-distance"] = amount + "pt";
      }
      apply([{ op: "formatWordTable", nodeId: tableId, properties }]);
    });
    const cellSpacing = field(
      "表格单元格间距（磅）",
      parseFloat(rowStyleResolver.get(projectRow(tableId))["--word-table-cell-spacing"] ?? "0"),
    );
    cellSpacing.step = ".1";
    button("应用表格间距", () => {
      const amount = Number(cellSpacing.value);
      if (!Number.isFinite(amount) || amount < 0)
        throw new DocumentError("InvalidOperation", "单元格间距必须是非负数。");
      apply([
        {
          op: "formatWordTable",
          nodeId: tableId,
          properties: {
            "--word-table-cell-spacing": amount + "pt",
            "--word-table-cell-spacing-type": "dxa",
            "--word-table-cell-spacing-value": String(Math.round(amount * 20)),
            "border-spacing": amount + "pt",
            "border-collapse": amount > 0 ? "separate" : "collapse",
          },
        },
      ]);
    });
    const currentRow = session.getNode(rowIds[active.row]!)!,
      currentRowCss = rowStyleResolver.get(projectRow(currentRow.nodeId)),
      currentRowSpacing =
        currentRowCss["--word-row-cell-spacing"] ??
        currentRowCss["--word-row-exception-cell-spacing"] ??
        rowStyleResolver.get(projectRow(tableId))["--word-table-cell-spacing"],
      directRowSpacing = rowStyleResolver.specified({
        tag: "tr",
        attrs: { style: currentRow.attributes.style ?? "" },
        parent: undefined,
      })["--word-row-cell-spacing"];
    const rowSpacingMode = choose(
        "当前行间距来源",
        [
          ["inherit", "继承样式或表格"],
          ["custom", "指定间距"],
        ],
        directRowSpacing === undefined ? "inherit" : "custom",
      ),
      rowSpacingValue = field("当前行单元格间距（磅）", parseFloat(currentRowSpacing ?? "0"));
    rowSpacingValue.step = ".1";
    button("应用当前行间距", () =>
      apply([
        {
          op: "formatWordTable",
          nodeId: tableId,
          rowIndex: Number(row.value) - 1,
          rowCellSpacing: rowSpacingMode.value === "inherit" ? null : Number(rowSpacingValue.value),
        },
      ]),
    );
    const exception = wordRowExceptionProperties(currentRowCss),
      rowFill = field("当前行例外底纹", exception["background-color"] ?? "#ffffff", "color"),
      rowBorder = field("当前行例外边框颜色", exception["border-left-color"] ?? "#000000", "color"),
      rowBorderWidth = field(
        "当前行例外边框宽度（磅）",
        parseFloat(exception["border-left-width"] ?? "0.5"),
      ),
      rowIndent = field(
        "当前行例外缩进（磅）",
        parseFloat(exception.indent ?? exception["margin-left"] ?? "0") || 0,
      ),
      rowWidth = field(
        "当前行例外宽度（磅）",
        parseFloat(exception.width ?? String(grid?.reduce((a, b) => a + b, 0) ?? 468)),
      ),
      rowPadding = field(
        "当前行例外单元格边距（磅）",
        parseFloat(exception["padding-left"] ?? "5.4"),
      ),
      rowAlign = choose(
        "当前行例外对齐",
        [
          ["left", "左对齐"],
          ["center", "居中"],
          ["right", "右对齐"],
        ],
        exception["margin-left"] === "auto"
          ? exception["margin-right"] === "auto"
            ? "center"
            : "right"
          : "left",
      );
    rowBorderWidth.step = rowIndent.step = rowWidth.step = rowPadding.step = ".1";
    button("应用当前行表格例外", () =>
      apply([
        {
          op: "formatWordTable",
          nodeId: tableId,
          rowIndex: Number(row.value) - 1,
          rowExceptionProperties: {
            "background-color": rowFill.value,
            border: rowBorderWidth.value + "pt solid " + rowBorder.value,
            width: rowWidth.value + "pt",
            indent: rowIndent.value + "pt",
            alignment: rowAlign.value,
            "margin-left": rowAlign.value === "left" ? rowIndent.value + "pt" : "auto",
            "margin-right": rowAlign.value === "center" ? "auto" : "0pt",
            padding: rowPadding.value + "pt",
          },
        },
      ]),
    );
    button("清除当前行表格例外", () =>
      apply([
        {
          op: "formatWordTable",
          nodeId: tableId,
          rowIndex: Number(row.value) - 1,
          rowExceptionProperties: Object.fromEntries(
            Object.keys(exception).map((key) => [key, null]),
          ),
        },
      ]),
    );
    const headers = field("重复表头行数", headerCount);
    button("设置重复表头", () =>
      apply([{ op: "formatWordTable", nodeId: tableId, repeatHeaderRows: Number(headers.value) }]),
    );
    const split = choose(
      "跨页断行",
      [
        ["allow", "允许"],
        ["avoid", "禁止"],
      ],
      "allow",
    );
    button("设置跨页断行", () =>
      apply([{ op: "formatWordTable", nodeId: tableId, allowRowSplit: split.value === "allow" }]),
    );
    const alignment = choose(
        "单元格垂直对齐",
        [
          ["top", "顶端"],
          ["middle", "居中"],
          ["bottom", "底端"],
        ],
        "top",
      ),
      direction = choose(
        "单元格文字方向",
        [
          ["lrTb", "横排"],
          ["tbRl", "从上到下"],
          ["btLr", "从下到上"],
        ],
        "lrTb",
      );
    const fill = field("单元格底纹", "#ffffff", "color"),
      border = field("单元格边框颜色", "#000000", "color"),
      borderWidth = field("边框宽度（磅）", 0.5),
      padding = field("单元格边距（磅）", 4);
    borderWidth.step = padding.step = "0.1";
    button("应用单元格设置", () => {
      const selectedIds = new Set<string>();
      for (let r = Number(row.value) - 1; r < Number(row.value) - 1 + Number(rows.value); r++)
        for (
          let c = Number(column.value) - 1;
          c < Number(column.value) - 1 + Number(columns.value);
          c++
        ) {
          const id = occupied[r]?.[c];
          if (!id) throw new DocumentError("InvalidOperation", "单元格选区超出表格。");
          selectedIds.add(id);
        }
      apply(
        [...selectedIds].map((nodeId) => ({
          op: "formatWordCell",
          nodeId,
          properties: {
            "vertical-align": alignment.value,
            "background-color": fill.value,
            border: borderWidth.value + "pt solid " + border.value,
            padding: padding.value + "pt",
          },
          textDirection: direction.value as "lrTb" | "tbRl" | "btLr",
        })),
      );
    });
    const styles =
      session.snapshot().metadata.word?.styles.filter((style) => style.kind === "table") ?? [];
    if (styles.length) {
      const style = choose(
        "表格样式",
        [
          ["", "无样式"],
          ...styles.map((style) => [style.styleId, style.name ?? style.styleId] as const),
        ],
        table.attributes["data-word-style"] ?? "",
      );
      button("应用表格样式", () =>
        apply([{ op: "formatWordTable", nodeId: tableId, styleId: style.value || null }]),
      );
    }
    const setting = (
      key: "tableLookMask" | "rowBandSize" | "columnBandSize",
    ): number | undefined => {
      const definitions = session.snapshot().metadata.word?.styles ?? [];
      let style = definitions.find(
        (style) => style.styleId === table.attributes["data-word-style"],
      );
      const seen = new Set<string>();
      while (style && !seen.has(style.styleId)) {
        seen.add(style.styleId);
        if (style[key] !== undefined) return style[key];
        style = definitions.find((item) => item.styleId === style!.basedOn);
      }
      return undefined;
    };
    const tableInfo = session
        .snapshot()
        .metadata.word?.tables.find((info) => info.nodeId === tableId),
      mask = tableInfo?.lookMask ?? setting("tableLookMask") ?? 0;
    const lookFlags = (
      [
        ["首行样式", 32],
        ["末行样式", 64],
        ["首列样式", 128],
        ["末列样式", 256],
        ["行条带", 512],
        ["列条带", 1024],
      ] as const
    ).map(([name, bit]) => {
      const input = field(name, "", "checkbox");
      input.checked = bit >= 512 ? !(mask & bit) : !!(mask & bit);
      return { input, bit };
    });
    const rowBand = field("行条带大小", tableInfo?.rowBandSize ?? setting("rowBandSize") ?? 0),
      columnBand = field("列条带大小", tableInfo?.columnBandSize ?? setting("columnBandSize") ?? 0);
    rowBand.min = columnBand.min = "0";
    button("应用表格样式选项", () => {
      let next = mask;
      for (const { input, bit } of lookFlags) {
        if (bit >= 512 ? !input.checked : input.checked) next |= bit;
        else next &= ~bit;
      }
      apply([
        {
          op: "formatWordTable",
          nodeId: tableId,
          lookMask: next,
          rowBandSize: Number(rowBand.value),
          columnBandSize: Number(columnBand.value),
        },
      ]);
    });
    button("删除表格", () => {
      apply([{ op: "deleteWordTable", nodeId: tableId }]);
      close();
    });
  }
  return { open, dispose: close };
}
