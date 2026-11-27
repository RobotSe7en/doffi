import { DocumentError } from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation, WordChart } from "@doffi/core";
import { createWordDockedPane, replaceSameSideDock } from "./word-pane-chrome.js";
const kinds = [
  ["column", "柱状图"],
  ["line", "折线图"],
  ["pie", "饼图"],
  ["bar", "条形图"],
  ["area", "面积图"],
  ["doughnut", "圆环图"],
  ["scatter", "散点图"],
  ["combo", "组合图"],
] as const;
export function createWordChartControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  selection: () =>
    | {
        paragraphId: string;
        offset: number;
      }
    | undefined,
  onError: (error: unknown) => void,
): {
  open(insert: boolean): void;
  bind(root: HTMLElement, signal: AbortSignal): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let shell: HTMLElement | undefined,
    panel: HTMLElement | undefined,
    controller: AbortController | undefined,
    selected: string | undefined;
  const close = (): void => {
    controller?.abort();
    shell?.remove();
    shell = undefined;
    panel = undefined;
  };
  const attempt = (run: () => void): void => {
    try {
      run();
    } catch (error) {
      onError(error);
    }
  };
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
  };
  function open(insert: boolean): void {
    close();
    const location = selection(),
      baseRevision = session.snapshot().manifest.revision,
      drawings = (session.snapshot().metadata.word?.drawings ?? []).filter(
        (drawing) => drawing.kind === "chart",
      ),
      drawing = drawings.find((drawing) => drawing.nodeId === selected) ?? drawings[0];
    if (insert && !location) throw new DocumentError("InvalidOperation", "请先在段落中放置光标。");
    if (!insert && !drawing)
      throw new DocumentError("InvalidOperation", "文档中还没有可编辑图表。");
    let draft: WordChart = insert
      ? {
          kind: "column",
          title: "图表标题",
          legend: "bottom",
          showLabels: false,
          categories: ["类别 1", "类别 2", "类别 3"],
          series: [
            { seriesId: "series-1", name: "系列 1", values: [1, 2, 3] },
            { seriesId: "series-2", name: "系列 2", values: [2, 3, 1] },
          ],
        }
      : (JSON.parse(JSON.stringify(drawing!.chart)) as WordChart);
    replaceSameSideDock(parent, "right");
    controller = new AbortController();
    const signal = controller.signal;
    const chrome = createWordDockedPane(document, {
      ariaLabel: "图表编辑",
      title: insert ? "插入图表" : "设置图表格式",
      dock: "right",
      closeLabel: "关闭图表编辑",
      onClose: close,
      signal,
      width: "348px",
    });
    shell = chrome.panel;
    panel = chrome.body;
    parent.insertBefore(shell, before);
    const button = (
      name: string,
      run: () => void,
      container: HTMLElement = panel!,
    ): HTMLButtonElement => {
      const node = document.createElement("button");
      node.type = "button";
      node.textContent = name;
      node.addEventListener("click", () => attempt(run), { signal });
      container.append(node);
      return node;
    };
    const field = (
      name: string,
      value: string | number | undefined,
      type = "text",
      container: HTMLElement = panel!,
    ): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      input.type = type;
      input.value = value === undefined ? "" : String(value);
      if (type === "number") input.step = "any";
      label.append(input);
      container.append(label);
      return input;
    };
    const choose = (
      name: string,
      options: readonly (readonly [string, string])[],
      value: string,
      container: HTMLElement = panel!,
    ): HTMLSelectElement => {
      const label = document.createElement("label"),
        input = document.createElement("select");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      for (const [key, text] of options) {
        const node = document.createElement("option");
        node.value = key;
        node.textContent = text;
        input.append(node);
      }
      input.value = value;
      label.append(input);
      container.append(label);
      return input;
    };
    const number = (value: string): number | undefined => {
      if (!value.trim()) return undefined;
      const number = Number(value);
      if (!Number.isFinite(number))
        throw new DocumentError("InvalidOperation", "图表数据和轴范围必须是有限数值。");
      return number;
    };
    if (!insert) {
      selected = drawing!.nodeId;
      const target = choose(
        "当前图表",
        drawings.map((drawing, index) => [
          drawing.nodeId,
          drawing.chart.title ?? "图表 " + (index + 1),
        ]),
        drawing!.nodeId,
      );
      target.addEventListener(
        "change",
        () => {
          selected = target.value;
          open(false);
        },
        { signal },
      );
    }
    const kind = choose("图表类型", kinds, draft.kind),
      title = field("图表标题", draft.title ?? ""),
      legend = choose(
        "图例位置",
        [
          ["none", "隐藏"],
          ["left", "左"],
          ["right", "右"],
          ["top", "上"],
          ["bottom", "下"],
        ],
        draft.legend,
      ),
      labels = field("显示数据标签", "", "checkbox");
    labels.checked = draft.showLabels;
    title.addEventListener(
      "input",
      () => {
        draft.title = title.value;
      },
      { signal },
    );
    legend.addEventListener(
      "change",
      () => {
        draft.legend = legend.value as WordChart["legend"];
      },
      { signal },
    );
    labels.addEventListener(
      "change",
      () => {
        draft.showLabels = labels.checked;
      },
      { signal },
    );
    const labelKinds = choose(
      "标签内容",
      [
        ["value", "值"],
        ["category", "类别"],
        ["series", "系列"],
        ["percent", "百分比"],
      ],
      draft.labelContents?.[0] ?? "value",
    );
    labelKinds.addEventListener(
      "change",
      () => {
        draft.labelContents = [labelKinds.value as "value"];
      },
      { signal },
    );
    const grouping = choose(
      "系列排列",
      [
        ["clustered", "簇状／普通"],
        ["stacked", "堆积"],
        ["percentStacked", "百分比堆积"],
      ],
      draft.grouping === "stacked" || draft.grouping === "percentStacked"
        ? draft.grouping
        : "clustered",
    );
    grouping.addEventListener(
      "change",
      () => {
        draft.grouping = grouping.value as WordChart["grouping"];
      },
      { signal },
    );
    const blank = choose(
      "空值显示",
      [
        ["gap", "间隙"],
        ["zero", "零"],
        ["span", "连接"],
      ],
      draft.blankAs ?? "gap",
    );
    blank.addEventListener(
      "change",
      () => {
        draft.blankAs = blank.value as WordChart["blankAs"];
      },
      { signal },
    );
    const format = field("图表数字格式", draft.numberFormat ?? "General");
    format.addEventListener(
      "input",
      () => {
        draft.numberFormat = format.value;
      },
      { signal },
    );
    for (const [key, label] of [
      ["minimum", "纵轴最小值"],
      ["maximum", "纵轴最大值"],
      ["majorUnit", "纵轴主刻度"],
      ["xMinimum", "横轴最小值"],
      ["xMaximum", "横轴最大值"],
      ["xMajorUnit", "横轴主刻度"],
      ["secondaryXMinimum", "次横轴最小值"],
      ["secondaryXMaximum", "次横轴最大值"],
      ["secondaryXMajorUnit", "次横轴主刻度"],
      ["secondaryMinimum", "次轴最小值"],
      ["secondaryMaximum", "次轴最大值"],
      ["secondaryMajorUnit", "次轴主刻度"],
      ["gapWidth", "分类间距"],
      ["overlap", "系列重叠"],
      ["holeSize", "圆环孔径"],
      ["firstSliceAngle", "起始角度"],
    ] as const) {
      const input = field(label, draft[key], "number");
      input.addEventListener(
        "change",
        () =>
          attempt(() => {
            draft[key] = number(input.value);
          }),
        { signal },
      );
    }
    const scatterStyle = choose(
      "散点线型",
      [
        ["marker", "仅标记"],
        ["line", "直线"],
        ["lineMarker", "直线和标记"],
        ["smooth", "平滑线"],
        ["smoothMarker", "平滑线和标记"],
      ],
      draft.scatterStyle ?? "marker",
    );
    scatterStyle.addEventListener(
      "change",
      () => {
        draft.scatterStyle = scatterStyle.value as WordChart["scatterStyle"];
        for (const series of draft.series) {
          series.lineVisible = scatterStyle.value !== "marker";
          series.marker = ["line", "smooth"].includes(scatterStyle.value) ? "none" : "circle";
        }
      },
      { signal },
    );
    const axes = document.createElement("details"),
      summary = document.createElement("summary");
    summary.textContent = "坐标轴方向、交叉和标题";
    axes.append(summary);
    panel.append(axes);
    for (const [key, name] of [
      ["category", "类别／X 轴"],
      ["value", "数值／Y 轴"],
      ["secondaryCategory", "次类别／X 轴"],
      ["secondaryValue", "次数值／Y 轴"],
    ] as const) {
      const options = draft.axes?.[key],
        group = document.createElement("fieldset"),
        legend = document.createElement("legend");
      legend.textContent = name;
      group.append(legend);
      axes.append(group);
      const update = (change: Partial<NonNullable<WordChart["axes"]>["category"]>): void => {
        (draft.axes ??= {})[key] ??= {};
        Object.assign(draft.axes[key]!, change);
      };
      const visible = field(name + " 显示", "", "checkbox", group);
      visible.checked = options?.visible ?? key !== "secondaryCategory";
      visible.addEventListener("change", () => update({ visible: visible.checked }), { signal });
      const reverse = field(name + " 逆序", "", "checkbox", group);
      reverse.checked = options?.reverse ?? false;
      reverse.addEventListener("change", () => update({ reverse: reverse.checked }), { signal });
      const title = field(name + " 标题", options?.title ?? "", "text", group);
      title.addEventListener("input", () => update({ title: title.value }), { signal });
      const format = field(
        name + " 数字格式",
        options?.numberFormat ?? draft.numberFormat ?? "General",
        "text",
        group,
      );
      format.addEventListener("input", () => update({ numberFormat: format.value }), { signal });
      const label = choose(
        name + " 标签位置",
        [
          ["nextTo", "紧邻轴"],
          ["low", "低端"],
          ["high", "高端"],
          ["none", "隐藏"],
        ],
        options?.labelPosition ?? "nextTo",
        group,
      );
      label.addEventListener("change", () => update({ labelPosition: label.value as "nextTo" }), {
        signal,
      });
      const crossing = choose(
          name + " 交叉方式",
          [
            ["autoZero", "自动／零"],
            ["min", "最小值"],
            ["max", "最大值"],
            ["custom", "指定值"],
          ],
          options?.crossesAt !== undefined
            ? "custom"
            : (options?.crosses ?? (key.startsWith("secondary") ? "max" : "autoZero")),
          group,
        ),
        at = field(name + " 相交位置", options?.crossesAt ?? "", "number", group);
      at.disabled = crossing.value !== "custom";
      crossing.addEventListener(
        "change",
        () =>
          attempt(() => {
            at.disabled = crossing.value !== "custom";
            update(
              crossing.value === "custom"
                ? { crossesAt: number(at.value) ?? 0 }
                : { crosses: crossing.value as "autoZero", crossesAt: undefined },
            );
          }),
        { signal },
      );
      at.addEventListener("change", () => attempt(() => update({ crossesAt: number(at.value) })), {
        signal,
      });
      if (key === "value" || key === "secondaryValue") {
        const between = choose(
          name + " 类别刻度",
          [
            ["between", "刻度之间"],
            ["midCat", "刻度线上"],
          ],
          options?.crossBetween ?? "between",
          group,
        );
        between.addEventListener(
          "change",
          () => update({ crossBetween: between.value as "between" }),
          { signal },
        );
        const grid = field(name + " 主要网格线", "", "checkbox", group);
        grid.checked = options?.majorGridlines ?? key === "value";
        grid.addEventListener("change", () => update({ majorGridlines: grid.checked }), { signal });
      }
    }
    const textSettings = document.createElement("details"),
      textSummary = document.createElement("summary");
    textSummary.textContent = "图表文字和中西文字体";
    textSettings.append(textSummary);
    panel.append(textSettings);
    const scope = choose(
        "图表文字范围",
        [
          ["textStyle", "全图默认"],
          ["titleStyle", "标题"],
          ["legendStyle", "图例"],
          ["labelStyle", "数据标签"],
          ["category", "类别／X 轴"],
          ["value", "数值／Y 轴"],
          ["secondaryCategory", "次类别／X 轴"],
          ["secondaryValue", "次数值／Y 轴"],
        ],
        "textStyle",
        textSettings,
      ),
      latin = field("图表西文字体", "", "text", textSettings),
      east = field("图表东亚字体", "", "text", textSettings),
      complex = field("图表复杂文字字体", "", "text", textSettings),
      size = field("图表文字字号", "", "number", textSettings),
      color = field("图表文字颜色", "#333333", "color", textSettings),
      bold = field("图表文字加粗", "", "checkbox", textSettings),
      italic = field("图表文字斜体", "", "checkbox", textSettings);
    type TextStyle = NonNullable<WordChart["textStyle"]>;
    const style = (): TextStyle | undefined => {
        const key = scope.value;
        if (["category", "value", "secondaryCategory", "secondaryValue"].includes(key))
          return draft.axes?.[key as "category"]?.textStyle;
        return draft[key as "textStyle"];
      },
      setStyle = (value: TextStyle): void => {
        const key = scope.value;
        if (["category", "value", "secondaryCategory", "secondaryValue"].includes(key))
          ((draft.axes ??= {})[key as "category"] ??= {}).textStyle = value;
        else draft[key as "textStyle"] = value;
      };
    const loadStyle = (): void => {
      const current = style();
      latin.value = current?.fonts?.ascii ?? current?.fonts?.hAnsi ?? "";
      east.value = current?.fonts?.eastAsia ?? "";
      complex.value = current?.fonts?.cs ?? "";
      size.value = current?.fontSize === undefined ? "" : String(current.fontSize);
      color.value = current?.color && current.color !== "auto" ? current.color : "#333333";
      bold.checked = current?.bold ?? false;
      italic.checked = current?.italic ?? false;
    };
    scope.addEventListener("change", loadStyle, { signal });
    loadStyle();
    button(
      "应用图表文字设置",
      () => {
        setStyle({
          ...style(),
          fonts: {
            ascii: latin.value || undefined,
            hAnsi: latin.value || undefined,
            eastAsia: east.value || undefined,
            cs: complex.value || undefined,
          },
          fontSize: number(size.value),
          color: color.value,
          bold: bold.checked,
          italic: italic.checked,
        });
      },
      textSettings,
    );
    const data = document.createElement("div");
    data.style.cssText = "width:100%;overflow:auto";
    panel.append(data);
    const renderData = (): void => {
      data.replaceChildren();
      const table = document.createElement("table");
      table.style.cssText = "border-collapse:collapse;min-width:600px";
      data.append(table);
      const head = document.createElement("tr");
      table.append(head);
      const caption = document.createElement("th");
      caption.textContent = "系列与数据";
      head.append(caption);
      draft.categories.forEach((category, index) => {
        const cell = document.createElement("th");
        head.append(cell);
        const input = field("图表类别 " + (index + 1), category, "text", cell);
        input.addEventListener(
          "input",
          () => {
            draft.categories[index] = input.value;
          },
          { signal },
        );
        button(
          "删除类别 " + (index + 1),
          () => {
            draft.categories.splice(index, 1);
            for (const series of draft.series) {
              series.values.splice(index, 1);
              series.xValues?.splice(index, 1);
              series.pointColors?.splice(index, 1);
            }
            renderData();
          },
          cell,
        );
      });
      draft.series.forEach((series, ordinal) => {
        const row = document.createElement("tr"),
          settings = document.createElement("td");
        row.append(settings);
        table.append(row);
        const name = field("系列 " + (ordinal + 1) + " 名称", series.name, "text", settings);
        name.addEventListener(
          "input",
          () => {
            series.name = name.value;
          },
          { signal },
        );
        if (draft.kind === "combo") {
          const type = choose(
            "系列 " + (ordinal + 1) + " 类型",
            kinds.filter(([kind]) => ["column", "line", "area", "scatter", "bar"].includes(kind)),
            series.kind ?? "column",
            settings,
          );
          type.addEventListener(
            "change",
            () => {
              series.kind = type.value as NonNullable<typeof series.kind>;
              if (series.kind === "scatter") {
                series.xValues ??= draft.categories.map((_, index) => index + 1);
                series.axis = "secondary";
              }
              renderData();
            },
            { signal },
          );
        }
        const axis = choose(
          "系列 " + (ordinal + 1) + " 坐标轴",
          [
            ["primary", "主轴"],
            ["secondary", "次轴"],
          ],
          series.axis ?? "primary",
          settings,
        );
        axis.addEventListener(
          "change",
          () => {
            series.axis = axis.value as "primary" | "secondary";
          },
          { signal },
        );
        const color = field(
          "系列 " + (ordinal + 1) + " 颜色",
          series.color ?? ["#4472c4", "#ed7d31", "#a5a5a5"][ordinal % 3]!,
          "color",
          settings,
        );
        color.addEventListener(
          "input",
          () => {
            series.color = color.value;
          },
          { signal },
        );
        const marker = choose(
          "系列 " + (ordinal + 1) + " 标记",
          [
            ["none", "无"],
            ["circle", "圆"],
            ["square", "方"],
            ["diamond", "菱形"],
            ["triangle", "三角"],
            ["x", "叉"],
            ["plus", "加号"],
            ["star", "星"],
            ["dot", "点"],
            ["dash", "短横"],
            ["auto", "自动"],
          ],
          series.marker ?? (draft.kind === "scatter" ? "circle" : "none"),
          settings,
        );
        marker.addEventListener(
          "change",
          () => {
            series.marker = marker.value as typeof series.marker;
          },
          { signal },
        );
        const size = field(
          "系列 " + (ordinal + 1) + " 标记大小",
          series.markerSize ?? 5,
          "number",
          settings,
        );
        size.addEventListener(
          "change",
          () =>
            attempt(() => {
              series.markerSize = number(size.value);
            }),
          { signal },
        );
        const lineVisible = field("系列 " + (ordinal + 1) + " 显示线条", "", "checkbox", settings);
        lineVisible.checked =
          series.lineVisible ??
          (draft.kind !== "scatter" || (draft.scatterStyle ?? "marker") !== "marker");
        lineVisible.addEventListener(
          "change",
          () => {
            series.lineVisible = lineVisible.checked;
          },
          { signal },
        );
        const smooth = field("系列 " + (ordinal + 1) + " 平滑", "", "checkbox", settings);
        smooth.checked = series.smooth ?? false;
        smooth.addEventListener(
          "change",
          () => {
            series.smooth = smooth.checked;
          },
          { signal },
        );
        button(
          "删除系列 " + (ordinal + 1),
          () => {
            if (draft.series.length === 1)
              throw new DocumentError("InvalidOperation", "图表至少保留一个系列。");
            draft.series.splice(ordinal, 1);
            renderData();
          },
          settings,
        );
        series.values.forEach((value, index) => {
          const cell = document.createElement("td");
          row.append(cell);
          const input = field(
            "系列 " + (ordinal + 1) + " 值 " + (index + 1),
            value ?? "",
            "number",
            cell,
          );
          input.addEventListener(
            "change",
            () =>
              attempt(() => {
                series.values[index] = number(input.value) ?? null;
              }),
            { signal },
          );
          if (draft.kind === "scatter" || (draft.kind === "combo" && series.kind === "scatter")) {
            series.xValues ??= draft.categories.map((_, index) => index + 1);
            const x = field(
              "系列 " + (ordinal + 1) + " X " + (index + 1),
              series.xValues[index] ?? "",
              "number",
              cell,
            );
            x.addEventListener(
              "change",
              () =>
                attempt(() => {
                  series.xValues![index] = number(x.value) ?? null;
                }),
              { signal },
            );
          }
        });
      });
      button(
        "添加图表类别",
        () => {
          draft.categories.push("类别 " + (draft.categories.length + 1));
          for (const series of draft.series) {
            series.values.push(null);
            series.xValues?.push(draft.categories.length);
          }
          renderData();
        },
        data,
      );
      button(
        "添加图表系列",
        () => {
          let index = draft.series.length + 1;
          while (draft.series.some((series) => series.seriesId === "series-" + index)) index++;
          draft.series.push({
            seriesId: "series-" + index,
            name: "系列 " + index,
            values: draft.categories.map(() => null),
            ...(draft.kind === "combo" ? { kind: "line" as const } : {}),
            ...(draft.kind === "scatter"
              ? { xValues: draft.categories.map((_, index) => index + 1) }
              : {}),
          });
          renderData();
        },
        data,
      );
      if (draft.kind === "pie") {
        const note = document.createElement("p");
        note.textContent = "饼图显示第一个系列，其他系列数据继续保留。";
        data.append(note);
      }
    };
    kind.addEventListener(
      "change",
      () => {
        const previous = draft.kind;
        draft.kind = kind.value as WordChart["kind"];
        for (const series of draft.series) {
          if (draft.kind === "combo")
            series.kind ??= ["column", "line", "bar", "area", "scatter"].includes(previous)
              ? (previous as NonNullable<typeof series.kind>)
              : "column";
          if (draft.kind === "scatter")
            series.xValues ??= draft.categories.map((_, index) => index + 1);
        }
        renderData();
      },
      { signal },
    );
    renderData();
    button(insert ? "插入图表" : "应用图表修改", () => {
      if (session.snapshot().manifest.revision !== baseRevision)
        throw new DocumentError("RevisionConflict", "编辑图表期间文档已更改，请重新打开图表设置。");
      const { native: _native, ...chart } = draft;
      if (insert) {
        const ids = new Set(drawings.map((drawing) => drawing.nodeId));
        apply([
          {
            op: "insertWordDrawing",
            nodeId: location!.paragraphId,
            offset: location!.offset,
            geometry: {
              x: 0,
              y: 0,
              width: 360,
              height: 240,
              rotation: 0,
              flipH: false,
              flipV: false,
            },
            layout: { mode: "inline" },
            object: { kind: "chart", chart },
          },
        ]);
        selected = session
          .snapshot()
          .metadata.word!.drawings!.find(
            (drawing) => drawing.kind === "chart" && !ids.has(drawing.nodeId),
          )!.nodeId;
      } else apply([{ op: "setWordChart", nodeId: drawing!.nodeId, changes: chart }]);
      open(false);
    });
  }
  return {
    open,
    dispose: close,
    bind(root, signal) {
      root.addEventListener(
        "click",
        (event) => {
          const node = (event.target as Element).closest<HTMLElement>("[data-word-drawing=chart]");
          if (node) selected = node.dataset.nodeId;
        },
        { signal },
      );
      root.addEventListener(
        "dblclick",
        (event) => {
          const node = (event.target as Element).closest<HTMLElement>("[data-word-drawing=chart]");
          if (!node) return;
          event.preventDefault();
          selected = node.dataset.nodeId;
          attempt(() => open(false));
        },
        { signal },
      );
    },
  };
}
