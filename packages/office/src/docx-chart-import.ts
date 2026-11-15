import { readChartTextStyle, chartThemeFont } from "./docx-chart-text.js";
import { WordChartSchema, DocumentError } from "@doffi/core";
import type { WordChart } from "@doffi/core";
import type { OpcPackage } from "./ooxml/package.js";
import { A, C, R, attr, child, elements, text, number, fail, unsupported } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import { readChartGraph } from "./chart-native-graph.js";
const value = (node: XmlNode | undefined, name: string, fallback?: string): string | undefined =>
  attr(child(node, name, C), "val", "") ?? fallback;
const flag = (node: XmlNode | undefined, name: string): boolean =>
  ["1", "true", "on"].includes(value(node, name, "0")!);
const numeric = (node: XmlNode | undefined, name: string): number | undefined => {
  const raw = value(node, name);
  return raw === undefined ? undefined : number(raw, node!);
};
export function chartText(node: XmlNode | undefined): string {
  if (!node) return "";
  if ((node.uri === A && node.name === "t") || (node.uri === C && node.name === "v"))
    return text(node);
  return node.content.map((item) => (typeof item === "string" ? "" : chartText(item))).join("");
}
export function readWordChart(
  opc: OpcPackage,
  part: string,
  main: string,
  reference: XmlNode,
  asset: (part: string, type: string) => string,
): WordChart {
  const relation = opc.relationships(part).get(attr(reference, "id", R) ?? "");
  if (!relation || relation.external || !relation.type.endsWith("/chart"))
    fail(reference, "A native chart relationship is missing.");
  const source = opc.read(relation.target);
  if (source.uri !== C || source.name !== "chartSpace")
    fail(source, "A native chart part has an invalid root.");
  const chart = child(source, "chart", C),
    plot = child(chart, "plotArea", C);
  if (!chart || !plot) fail(source, "A native chart plot is missing.");
  const nativeKinds: Readonly<Record<string, WordChart["kind"]>> = {
      barChart: "column",
      lineChart: "line",
      pieChart: "pie",
      areaChart: "area",
      doughnutChart: "doughnut",
      scatterChart: "scatter",
    },
    groups = elements(plot, undefined, C).filter((node) => node.name.endsWith("Chart"));
  if (!groups.length || groups.some((group) => !Object.hasOwn(nativeKinds, group.name)))
    unsupported(reference, "This chart family is retained as native content.");
  const sourceSeries = new Map(
    groups.flatMap((group) => elements(group, "ser", C)).map((node, index) => [node, index]),
  );
  const groupKind = (node: XmlNode): WordChart["kind"] =>
    node.name === "barChart" && value(node, "barDir") === "bar" ? "bar" : nativeKinds[node.name]!;
  if (groups.length > 1 && groups.some((group) => ["pie", "doughnut"].includes(groupKind(group))))
    unsupported(reference, "This combination is retained as native content.");
  const result: WordChart = {
      kind: groups.length > 1 ? "combo" : groupKind(groups[0]!),
      legend: "none",
      showLabels: false,
      categories: [],
      series: [],
    },
    title = chartText(child(chart, "title", C));
  if (title) result.title = title;
  const legend = child(chart, "legend", C);
  if (legend)
    result.legend =
      ({ l: "left", r: "right", t: "top", b: "bottom", tr: "right" } as const)[
        value(legend, "legendPos", "r") as "l"
      ] ?? "right";
  const external = child(source, "externalData", C),
    bookRelation = external
      ? opc.relationships(source.part).get(attr(external, "id", R) ?? "")
      : undefined;
  if (
    external &&
    (!bookRelation || bookRelation.external || !bookRelation.type.endsWith("/package"))
  )
    unsupported(reference, "Externally linked chart data remains native.");
  if (
    bookRelation &&
    opc.contentType(bookRelation.target) !==
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  )
    unsupported(reference, "This embedded chart data format remains intact native content.");
  let cachedPoints = 0;
  const vector = (
    node: XmlNode | undefined,
    numericValues: boolean,
  ): {
    values: (string | number | null)[];
    formula?: string;
  } => {
    if (!node) return { values: [] };
    const ref = child(node, numericValues ? "numRef" : "strRef", C),
      literal = child(node, numericValues ? "numLit" : "strLit", C),
      cache = literal ?? child(ref, numericValues ? "numCache" : "strCache", C);
    if (!cache) {
      const direct = child(node, "v", C);
      if (direct) return { values: [text(direct)] };
      unsupported(
        reference,
        "A chart without usable cached vectors remains native until its data can be resolved.",
      );
    }
    const count =
      numeric(cache, "ptCount") ??
      elements(cache, "pt", C).reduce(
        (max, point) => Math.max(max, Number(attr(point, "idx", "")) + 1),
        0,
      );
    if (!Number.isSafeInteger(count) || count < 0) fail(cache, "A chart cache length is invalid.");
    cachedPoints += count;
    if (cachedPoints > opc.limits.maxNodes)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Chart vector caches exceed the document node budget.",
      );
    const values: (string | number | null)[] = Array(count).fill(numericValues ? null : "");
    const seen = new Set<number>();
    for (const point of elements(cache, "pt", C)) {
      const index = Number(attr(point, "idx", ""));
      if (!Number.isSafeInteger(index) || index < 0 || index >= count || seen.has(index))
        fail(point, "A chart point identity is invalid.");
      seen.add(index);
      const raw = text(child(point, "v", C) ?? point);
      if (numericValues) {
        if (!raw.trim()) {
          values[index] = null;
          continue;
        }
        const scalar = Number(raw);
        if (!Number.isFinite(scalar))
          unsupported(reference, "A native chart error value remains preserved.");
        values[index] = scalar;
      } else values[index] = raw;
    }
    const formula = child(ref, "f", C);
    return { values, ...(formula ? { formula: text(formula) } : {}) };
  };
  const themes = opc.related(main, "theme"),
    theme = themes ? opc.read(themes) : undefined,
    scheme = child(child(theme, "themeElements", A), "clrScheme", A);
  const color = (properties: XmlNode | undefined): string | undefined => {
    const fill = child(properties, "solidFill", A),
      direct = child(fill, "srgbClr", A),
      named = child(fill, "schemeClr", A);
    let raw = attr(direct, "val", "");
    if (!raw && named) {
      const slot = child(scheme, attr(named, "val", "") ?? "", A),
        native = slot ? elements(slot, undefined, A)[0] : undefined;
      raw = attr(native, "val", "") ?? attr(native, "lastClr", "");
    }
    return raw && /^[0-9a-f]{6}$/i.test(raw) ? "#" + raw : undefined;
  };
  const language = value(source, "lang", "zh-CN")!,
    globalStyle = readChartTextStyle(child(source, "txPr", C), theme, language),
    titleStyle = readChartTextStyle(child(chart, "title", C), theme, language),
    legendStyle = readChartTextStyle(child(child(chart, "legend", C), "txPr", C), theme, language);
  const themeFonts = Object.fromEntries(
    [
      ["ascii", "+mn-lt"],
      ["hAnsi", "+mn-lt"],
      ["eastAsia", "+mn-ea"],
      ["cs", "+mn-cs"],
    ].flatMap(([role, value]) => {
      const name = chartThemeFont(value, role!, theme, language);
      return name ? [[role, name]] : [];
    }),
  );
  if (globalStyle || Object.keys(themeFonts).length)
    result.textStyle = { ...globalStyle, fonts: { ...themeFonts, ...globalStyle?.fonts } };
  if (titleStyle) result.titleStyle = titleStyle;
  if (legendStyle) result.legendStyle = legendStyle;
  const bindings: NonNullable<WordChart["native"]>["bindings"] = [],
    byAxis = new Map(
      elements(plot, undefined, C)
        .filter((node) => ["valAx", "catAx", "dateAx"].includes(node.name))
        .map((node) => [value(node, "axId")!, node]),
    );
  const groupAxes = (group: XmlNode): string[] =>
    elements(group, "axId", C).map((node) => attr(node, "val", "")!);
  const primaryGroup =
      groups.find(
        (group) => value(byAxis.get(groupAxes(group).at(-1) ?? ""), "crosses") !== "max",
      ) ?? groups[0]!,
    primaryAxes = groupAxes(primaryGroup),
    secondaryAxes =
      groups.map(groupAxes).find((axes) => axes.some((id) => !primaryAxes.includes(id))) ?? [];
  if (
    elements(plot, undefined, C).some(
      (axis) => axis.name === "dateAx" || child(child(axis, "scaling", C), "logBase", C),
    )
  )
    unsupported(reference, "Date and logarithmic chart axes remain intact native objects.");
  const readAxis = (node: XmlNode | undefined): NonNullable<WordChart["axes"]>["category"] =>
    node
      ? {
          visible: !flag(node, "delete"),
          reverse: value(child(node, "scaling", C), "orientation", "minMax") === "maxMin",
          crosses: value(node, "crosses", "autoZero") as "autoZero",
          ...(numeric(node, "crossesAt") !== undefined
            ? { crossesAt: numeric(node, "crossesAt") }
            : {}),
          labelPosition: value(node, "tickLblPos", "nextTo") as "nextTo",
          ...(value(node, "crossBetween")
            ? { crossBetween: value(node, "crossBetween") as "between" }
            : {}),
          position: value(node, "axPos", "b") as "b",
          majorGridlines: !!child(node, "majorGridlines", C),
          ...(attr(child(node, "numFmt", C), "formatCode", "")
            ? { numberFormat: attr(child(node, "numFmt", C), "formatCode", "") }
            : {}),
          ...(chartText(child(node, "title", C))
            ? { title: chartText(child(node, "title", C)) }
            : {}),
          ...(readChartTextStyle(child(node, "txPr", C), theme, language)
            ? { textStyle: readChartTextStyle(child(node, "txPr", C), theme, language) }
            : {}),
        }
      : undefined;
  if (primaryAxes.length || secondaryAxes.length)
    result.axes = {
      category: readAxis(byAxis.get(primaryAxes[0] ?? "")),
      value: readAxis(byAxis.get(primaryAxes[1] ?? "")),
      secondaryCategory: readAxis(byAxis.get(secondaryAxes[0] ?? "")),
      secondaryValue: readAxis(byAxis.get(secondaryAxes[1] ?? "")),
    };
  for (const group of groups) {
    const kind = groupKind(group),
      axes = elements(group, "axId", C).map((node) => attr(node, "val", "")!),
      secondary = axes.some((id) => !primaryAxes.includes(id));
    const grouping = value(group, "grouping");
    if (grouping && ["standard", "clustered", "stacked", "percentStacked"].includes(grouping)) {
      if (
        result.grouping &&
        result.grouping !== grouping &&
        result.grouping !== "standard" &&
        grouping !== "standard"
      )
        unsupported(reference, "Mixed stacking groups remain native.");
      result.grouping = grouping as WordChart["grouping"];
    }
    for (const key of ["gapWidth", "overlap", "holeSize"] as const) {
      const number = numeric(group, key);
      if (number !== undefined) result[key] = number;
    }
    const angle = numeric(group, "firstSliceAng");
    if (angle !== undefined) result.firstSliceAngle = angle;
    if (kind === "scatter")
      result.scatterStyle = value(group, "scatterStyle", "lineMarker") as WordChart["scatterStyle"];
    const labels = child(group, "dLbls", C);
    const labelStyle = readChartTextStyle(child(labels, "txPr", C), theme, language);
    if (labelStyle) result.labelStyle = labelStyle;
    result.labelContents = (
      [
        ["showVal", "value"],
        ["showCatName", "category"],
        ["showSerName", "series"],
        ["showPercent", "percent"],
      ] as const
    )
      .filter(([name]) => flag(labels, name))
      .map(([, name]) => name);
    result.showLabels ||=
      flag(labels, "showVal") ||
      flag(labels, "showCatName") ||
      flag(labels, "showSerName") ||
      flag(labels, "showPercent");
    const series = elements(group, "ser", C)
      .map((node, index) => ({ node, index, order: numeric(node, "order") ?? index }))
      .sort((a, b) => a.order - b.order);
    for (const item of series) {
      const native = item.node,
        id = "series-" + result.series.length,
        titleNode = child(native, "tx", C),
        nameVector = titleNode
          ? child(titleNode, "strRef", C)
            ? vector(titleNode, false)
            : { values: [chartText(titleNode)] }
          : { values: [] },
        values = vector(child(native, kind === "scatter" ? "yVal" : "val", C), true),
        categoryNode = child(native, kind === "scatter" ? "xVal" : "cat", C),
        isNumeric =
          kind === "scatter" ||
          !!child(categoryNode, "numRef", C) ||
          !!child(categoryNode, "numLit", C),
        categories = categoryNode
          ? vector(categoryNode, isNumeric)
          : { values: values.values.map((_, index) => String(index + 1)) };
      const definition: WordChart["series"][number] = {
        seriesId: id,
        name: String(nameVector.values[0] ?? "Series " + (result.series.length + 1)),
        values: values.values as (number | null)[],
        ...(result.kind === "combo"
          ? { kind: kind as NonNullable<WordChart["series"][number]["kind"]> }
          : {}),
      };
      if (secondary) definition.axis = "secondary";
      if (kind === "scatter") definition.xValues = categories.values as (number | null)[];
      else {
        const labels = categories.values.map((value) => String(value ?? ""));
        if (
          result.categories.length &&
          labels.some(
            (label, index) =>
              result.categories[index] !== undefined && result.categories[index] !== label,
          )
        )
          unsupported(reference, "Different category vectors remain native.");
        if (labels.length > result.categories.length) result.categories = labels;
      }
      const seriesPaint = child(native, "spPr", C),
        line = child(seriesPaint, "ln", A);
      if (line) definition.lineVisible = !child(line, "noFill", A);
      const fill =
        (kind === "line" || kind === "scatter" ? color(line) : undefined) ?? color(seriesPaint);
      if (fill) definition.color = fill;
      const marker = child(native, "marker", C);
      if (marker) {
        definition.marker = value(
          marker,
          "symbol",
          "auto",
        ) as WordChart["series"][number]["marker"];
        const size = numeric(marker, "size");
        if (size !== undefined) definition.markerSize = size;
      }
      if (child(native, "smooth", C)) definition.smooth = flag(native, "smooth");
      const pointColors: (string | null)[] = [];
      for (const point of elements(native, "dPt", C)) {
        const index = numeric(point, "idx");
        if (
          index === undefined ||
          !Number.isSafeInteger(index) ||
          index < 0 ||
          index >= definition.values.length
        )
          fail(point, "A chart point style index is invalid.");
        while (pointColors.length <= index) pointColors.push(null);
        pointColors[index] = color(child(point, "spPr", C)) ?? null;
      }
      if (pointColors.some(Boolean)) definition.pointColors = pointColors;
      result.series.push(definition);
      bindings.push({
        seriesId: id,
        sourceIndex: sourceSeries.get(native)!,
        ...(fill ? { color: fill } : {}),
        ...(definition.pointColors ? { pointColors: definition.pointColors } : {}),
        ...("formula" in nameVector && nameVector.formula
          ? { nameFormula: nameVector.formula }
          : {}),
        ...(categories.formula
          ? { [kind === "scatter" ? "xFormula" : "categoryFormula"]: categories.formula }
          : {}),
        ...(values.formula ? { valueFormula: values.formula } : {}),
      });
    }
    const axis = byAxis.get(axes.at(-1) ?? "");
    if (axis) {
      const scale = child(axis, "scaling", C),
        min = numeric(scale, "min"),
        max = numeric(scale, "max"),
        unit = numeric(axis, "majorUnit");
      if (secondary) {
        if (min !== undefined) result.secondaryMinimum = min;
        if (max !== undefined) result.secondaryMaximum = max;
        if (unit !== undefined) result.secondaryMajorUnit = unit;
      } else {
        if (min !== undefined) result.minimum = min;
        if (max !== undefined) result.maximum = max;
        if (unit !== undefined) result.majorUnit = unit;
        const format = attr(child(axis, "numFmt", C), "formatCode", "");
        if (format) result.numberFormat = format;
      }
    }
    if (kind === "scatter") {
      const axis = byAxis.get(axes[0] ?? ""),
        scale = child(axis, "scaling", C);
      if (secondary) {
        result.secondaryXMinimum = numeric(scale, "min");
        result.secondaryXMaximum = numeric(scale, "max");
        result.secondaryXMajorUnit = numeric(axis, "majorUnit");
      } else {
        result.xMinimum = numeric(scale, "min");
        result.xMaximum = numeric(scale, "max");
        result.xMajorUnit = numeric(axis, "majorUnit");
      }
    }
  }
  if (
    !result.series.length ||
    (result.kind === "combo" &&
      result.series.some((series) => series.kind === "bar") &&
      !result.series.every((series) => series.kind === "bar"))
  )
    unsupported(reference, "This native series arrangement remains preserved.");
  const count = Math.max(
    result.categories.length,
    ...result.series.map((series) => series.values.length),
  );
  while (result.categories.length < count)
    result.categories.push(String(result.categories.length + 1));
  for (const series of result.series) {
    while (series.values.length < count) series.values.push(null);
    if (series.xValues) while (series.xValues.length < count) series.xValues.push(null);
  }
  if (result.kind === "scatter") {
    if (result.scatterStyle === "line") result.scatterStyle = "lineMarker";
    if (result.scatterStyle === "smooth") result.scatterStyle = "smoothMarker";
    if (result.scatterStyle === "marker") result.scatterStyle = "lineMarker";
    if (result.series.every((series) => series.lineVisible === false))
      result.scatterStyle = "marker";
    else if (result.series.every((series) => series.marker === "none"))
      result.scatterStyle = result.scatterStyle === "smoothMarker" ? "smooth" : "line";
  }
  result.blankAs = value(chart, "dispBlanksAs", "gap") as WordChart["blankAs"];
  const { categories: _categories, series: sourceData, ...properties } = result,
    baseline = {
      ...properties,
      series: sourceData.map((series) => {
        const { values: _values, xValues: _x, ...style } = series;
        return style;
      }),
    };
  const previewWarnings: string[] = [];
  if (child(source, "style", C)) previewWarnings.push("原生图表预设样式已保留；预览使用基础外观。");
  if (child(child(plot, "layout", C), "manualLayout", C))
    previewWarnings.push("原生绘图区位置已保留；预览使用自动布局。");
  result.native = {
    previewWarnings,
    baseline,
    primaryAxisIds: primaryAxes,
    secondaryAxisIds: secondaryAxes,
    chartPart: source.part,
    ...(bookRelation ? { workbookPart: bookRelation.target } : {}),
    ...readChartGraph(opc, source.part, asset),
    bindings,
  };
  const parsed = WordChartSchema.safeParse(result);
  if (!parsed.success) unsupported(reference, "This native chart configuration remains preserved.");
  return parsed.data;
}
