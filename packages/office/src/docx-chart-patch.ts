import type { DeepReadonly, WordChart } from "@doffi/core";
import { C, attr, child, elements, text } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
const equal = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b);
/** Ordinary edits patch the original native tree. In particular, changing a
 * vector must not repaint axes, flatten a rich title, discard point labels or
 * rewrite an unmodeled chart feature. Structural type/axis changes use the
 * compiler's new plot while retaining its carried native dependencies. */
export function patchWordChart(
  original: XmlNode,
  compiled: XmlNode,
  current: DeepReadonly<WordChart>,
): XmlNode {
  const baseline = current.native?.baseline;
  if (!baseline) return compiled;
  if (
    baseline.kind !== current.kind ||
    baseline.series.length !== current.series.length ||
    baseline.series.some(
      (series, index) =>
        series.seriesId !== current.series[index]?.seriesId ||
        series.kind !== current.series[index]?.kind ||
        (series.axis ?? "primary") !== (current.series[index]?.axis ?? "primary"),
    )
  )
    return compiled;
  const replace = (
    target: XmlNode | undefined,
    source: XmlNode | undefined,
    names: readonly string[],
    order: readonly string[],
  ): void => {
    if (!target || !source) return;
    for (const name of names) {
      target.content.splice(
        0,
        target.content.length,
        ...target.content.filter(
          (node) => typeof node === "string" || node.uri !== C || node.name !== name,
        ),
      );
      for (const node of elements(source, name, C)) {
        const at = target.content.findIndex(
          (other) =>
            typeof other !== "string" &&
            other.uri === C &&
            (order.indexOf(other.name) < 0 ? Infinity : order.indexOf(other.name)) >
              order.indexOf(name),
        );
        target.content.splice(at < 0 ? target.content.length : at, 0, node);
      }
    }
  };
  const oldChart = child(original, "chart", C),
    newChart = child(compiled, "chart", C),
    oldPlot = child(oldChart, "plotArea", C),
    newPlot = child(newChart, "plotArea", C),
    oldGroups = elements(oldPlot, undefined, C).filter((node) => node.name.endsWith("Chart")),
    newGroups = elements(newPlot, undefined, C).filter((node) => node.name.endsWith("Chart"));
  const flatOld = oldGroups.flatMap((group) => elements(group, "ser", C)),
    flatNew = newGroups.flatMap((group) => elements(group, "ser", C));
  const vector = (node: XmlNode | undefined): unknown => {
    if (!node) return undefined;
    const ref = elements(node, undefined, C)[0],
      cache = ref?.name.endsWith("Ref")
        ? elements(ref, undefined, C).find((node) => node.name.endsWith("Cache"))
        : ref;
    return {
      kind: ref?.name.replace("Ref", "").replace("Lit", ""),
      count: attr(child(cache, "ptCount", C), "val", ""),
      values: elements(cache, "pt", C).map((node) => [
        attr(node, "idx", ""),
        text(child(node, "v", C) ?? node),
      ]),
    };
  };
  const seriesOrder = [
    "idx",
    "order",
    "tx",
    "spPr",
    "invertIfNegative",
    "marker",
    "dPt",
    "dLbls",
    "trendline",
    "errBars",
    "cat",
    "val",
    "xVal",
    "yVal",
    "smooth",
    "extLst",
  ];
  const bindings = new Map(
    (current.native?.bindings ?? []).map((binding, index) => [
      binding.seriesId,
      binding.sourceIndex ?? index,
    ]),
  );
  current.series.forEach((series, index) => {
    const old = flatOld[bindings.get(series.seriesId) ?? index],
      fresh = flatNew[index],
      before = baseline.series[index]!;
    if (!old || !fresh) return;
    if (series.name !== before.name) replace(old, fresh, ["tx"], seriesOrder);
    for (const name of ["cat", "val", "xVal", "yVal"]) {
      const previous = child(old, name, C),
        next = child(fresh, name, C);
      if (!previous || !next) continue;
      const beforeValue = vector(previous),
        afterValue = vector(next);
      // Numeric categories and their source-linked formats remain native when
      // their string labels have not changed in the typed editing source.
      if (name === "cat") {
        const categoryCache = elements(previous, undefined, C)[0],
          cache = categoryCache?.name.endsWith("Ref")
            ? elements(categoryCache, undefined, C).find((node) => node.name.endsWith("Cache"))
            : categoryCache,
          count = Number(attr(child(cache, "ptCount", C), "val", "") ?? 0);
        if (count === current.categories.length) {
          const values = Array(current.categories.length).fill("") as string[];
          for (const point of elements(cache, "pt", C)) {
            const index = Number(attr(point, "idx", ""));
            if (Number.isSafeInteger(index) && index >= 0 && index < values.length)
              values[index] = text(child(point, "v", C) ?? point);
          }
          if (equal(values, current.categories)) continue;
        }
      }
      if (equal(beforeValue, afterValue)) continue;
      if (name === "val" || name === "yVal") {
        const previousRef = elements(previous, undefined, C)[0],
          previousCache = previousRef?.name.endsWith("Ref")
            ? elements(previousRef, undefined, C).find((node) => node.name === "numCache")
            : previousRef,
          format = child(previousCache, "formatCode", C),
          nextRef = elements(next, undefined, C)[0],
          nextCache = nextRef?.name.endsWith("Ref")
            ? elements(nextRef, undefined, C).find((node) => node.name === "numCache")
            : nextRef;
        if (format && nextCache) {
          nextCache.content.splice(
            0,
            nextCache.content.length,
            ...nextCache.content.filter(
              (node) => typeof node === "string" || node.name !== "formatCode",
            ),
            format,
          );
          const formats = elements(nextCache, "formatCode", C);
          for (const value of formats) {
            const index = nextCache.content.indexOf(value);
            nextCache.content.splice(index, 1);
            nextCache.content.unshift(value);
          }
        }
      }
      replace(old, fresh, [name], seriesOrder);
    }
    if (series.color !== before.color || series.lineVisible !== before.lineVisible)
      replace(old, fresh, ["spPr"], seriesOrder);
    if (series.marker !== before.marker || series.markerSize !== before.markerSize)
      replace(old, fresh, ["marker"], seriesOrder);
    if (!equal(series.pointColors, before.pointColors)) replace(old, fresh, ["dPt"], seriesOrder);
    if (series.smooth !== before.smooth) replace(old, fresh, ["smooth"], seriesOrder);
  });
  const chartOrder = [
    "title",
    "autoTitleDeleted",
    "pivotFmts",
    "view3D",
    "floor",
    "sideWall",
    "backWall",
    "plotArea",
    "legend",
    "plotVisOnly",
    "dispBlanksAs",
    "showDLblsOverMax",
    "extLst",
  ];
  if (current.title !== baseline.title || !equal(current.titleStyle, baseline.titleStyle))
    replace(oldChart, newChart, ["title", "autoTitleDeleted"], chartOrder);
  if (current.legend !== baseline.legend || !equal(current.legendStyle, baseline.legendStyle)) {
    const old = child(oldChart, "legend", C),
      fresh = child(newChart, "legend", C);
    if (old && fresh)
      replace(
        old,
        fresh,
        ["legendPos", ...(!equal(current.legendStyle, baseline.legendStyle) ? ["txPr"] : [])],
        ["legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr", "extLst"],
      );
    else replace(oldChart, newChart, ["legend"], chartOrder);
  }
  if (current.blankAs !== baseline.blankAs)
    replace(oldChart, newChart, ["dispBlanksAs"], chartOrder);
  oldGroups.forEach((group, index) => {
    const fresh = newGroups[index];
    if (!fresh) return;
    const circular = ["pieChart", "doughnutChart"].includes(group.name),
      groupOrder =
        group.name === "barChart"
          ? [
              "barDir",
              "grouping",
              "varyColors",
              "ser",
              "dLbls",
              "gapWidth",
              "overlap",
              "serLines",
              "axId",
              "extLst",
            ]
          : circular
            ? ["varyColors", "ser", "dLbls", "firstSliceAng", "holeSize", "extLst"]
            : group.name === "scatterChart"
              ? ["scatterStyle", "varyColors", "ser", "dLbls", "axId", "extLst"]
              : [
                  "grouping",
                  "varyColors",
                  "ser",
                  "dLbls",
                  "dropLines",
                  "hiLowLines",
                  "upDownBars",
                  "marker",
                  "smooth",
                  "axId",
                  "extLst",
                ];
    for (const [key, name] of [
      ["grouping", "grouping"],
      ["gapWidth", "gapWidth"],
      ["overlap", "overlap"],
      ["holeSize", "holeSize"],
      ["firstSliceAngle", "firstSliceAng"],
      ["scatterStyle", "scatterStyle"],
    ] as const)
      if (current[key] !== baseline[key]) replace(group, fresh, [name], groupOrder);
    if (
      current.showLabels !== baseline.showLabels ||
      !equal(current.labelContents, baseline.labelContents) ||
      !equal(current.labelStyle, baseline.labelStyle)
    )
      replace(group, fresh, ["dLbls"], groupOrder);
  });
  const axisIds = current.native?.primaryAxisIds
      ? [...current.native.primaryAxisIds, ...(current.native.secondaryAxisIds ?? [])]
      : [
          ...new Set(
            oldGroups.flatMap((group) =>
              elements(group, "axId", C).map((node) => attr(node, "val", "")),
            ),
          ),
        ],
    axes = elements(oldPlot, undefined, C)
      .filter((node) => ["catAx", "valAx", "dateAx"].includes(node.name))
      .sort(
        (a, b) =>
          axisIds.indexOf(attr(child(a, "axId", C), "val", "")) -
          axisIds.indexOf(attr(child(b, "axId", C), "val", "")),
      ),
    freshAxes = elements(newPlot, undefined, C).filter((node) =>
      ["catAx", "valAx"].includes(node.name),
    );
  axes.forEach((axis, ordinal) => {
    const id = attr(child(axis, "axId", C), "val", ""),
      primary = current.native?.primaryAxisIds,
      secondaryIds = current.native?.secondaryAxisIds,
      index =
        primary?.[1] === id
          ? 1
          : primary?.[0] === id
            ? 0
            : secondaryIds?.[1] === id
              ? 3
              : secondaryIds?.[0] === id
                ? 2
                : ordinal;
    const fresh = freshAxes[index];
    if (!fresh) return;
    const secondary = index >= 2,
      x =
        (index === 0 || index === 2) &&
        current.series.some(
          (series) =>
            (series.axis ?? "primary") === (secondary ? "secondary" : "primary") &&
            (current.kind === "scatter" || (current.kind === "combo" && series.kind === "scatter")),
        ),
      keys = x
        ? secondary
          ? (["secondaryXMinimum", "secondaryXMaximum", "secondaryXMajorUnit"] as const)
          : (["xMinimum", "xMaximum", "xMajorUnit"] as const)
        : secondary
          ? (["secondaryMinimum", "secondaryMaximum", "secondaryMajorUnit"] as const)
          : (["minimum", "maximum", "majorUnit"] as const),
      scale = child(axis, "scaling", C),
      freshScale = child(fresh, "scaling", C);
    if (current[keys[0]] !== baseline[keys[0]])
      replace(scale, freshScale, ["min"], ["logBase", "orientation", "max", "min", "extLst"]);
    if (current[keys[1]] !== baseline[keys[1]])
      replace(scale, freshScale, ["max"], ["logBase", "orientation", "max", "min", "extLst"]);
    if (current[keys[2]] !== baseline[keys[2]])
      replace(
        axis,
        fresh,
        ["majorUnit"],
        [
          "axId",
          "scaling",
          "delete",
          "axPos",
          "majorGridlines",
          "minorGridlines",
          "title",
          "numFmt",
          "majorTickMark",
          "minorTickMark",
          "tickLblPos",
          "spPr",
          "txPr",
          "crossAx",
          "crosses",
          "crossesAt",
          "crossBetween",
          "majorUnit",
          "minorUnit",
          "dispUnits",
          "extLst",
        ],
      );
    const axisKey =
        index === 0
          ? "category"
          : index === 1
            ? "value"
            : index === 2
              ? "secondaryCategory"
              : "secondaryValue",
      before = baseline.axes?.[axisKey],
      after = current.axes?.[axisKey],
      axisOrder = [
        "axId",
        "scaling",
        "delete",
        "axPos",
        "majorGridlines",
        "minorGridlines",
        "title",
        "numFmt",
        "majorTickMark",
        "minorTickMark",
        "tickLblPos",
        "spPr",
        "txPr",
        "crossAx",
        "crosses",
        "crossesAt",
        "auto",
        "lblAlgn",
        "lblOffset",
        "crossBetween",
        "majorUnit",
        "minorUnit",
        "dispUnits",
        "extLst",
      ];
    if (after?.reverse !== before?.reverse)
      replace(
        scale,
        freshScale,
        ["orientation"],
        ["logBase", "orientation", "max", "min", "extLst"],
      );
    for (const [key, name] of [
      ["visible", "delete"],
      ["position", "axPos"],
      ["majorGridlines", "majorGridlines"],
      ["labelPosition", "tickLblPos"],
      ["crossBetween", "crossBetween"],
      ["numberFormat", "numFmt"],
      ["title", "title"],
      ["textStyle", "txPr"],
    ] as const)
      if (after?.[key] !== before?.[key]) replace(axis, fresh, [name], axisOrder);
    if (after?.crosses !== before?.crosses || after?.crossesAt !== before?.crossesAt)
      replace(axis, fresh, ["crosses", "crossesAt"], axisOrder);
    if (current.numberFormat !== baseline.numberFormat)
      replace(
        axis,
        fresh,
        ["numFmt"],
        [
          "axId",
          "scaling",
          "delete",
          "axPos",
          "majorGridlines",
          "minorGridlines",
          "title",
          "numFmt",
          "majorTickMark",
          "minorTickMark",
          "tickLblPos",
          "spPr",
          "txPr",
          "crossAx",
          "crosses",
          "crossesAt",
          "crossBetween",
          "majorUnit",
          "minorUnit",
          "dispUnits",
          "extLst",
        ],
      );
  });
  if (!equal(current.textStyle, baseline.textStyle))
    replace(
      original,
      compiled,
      ["txPr"],
      [
        "date1904",
        "lang",
        "roundedCorners",
        "style",
        "clrMapOvr",
        "pivotSource",
        "protection",
        "chart",
        "spPr",
        "txPr",
        "externalData",
        "printSettings",
        "userShapes",
        "extLst",
      ],
    );
  if (!child(original, "externalData", C) && child(compiled, "externalData", C))
    replace(
      original,
      compiled,
      ["externalData"],
      [
        "date1904",
        "lang",
        "roundedCorners",
        "style",
        "clrMapOvr",
        "pivotSource",
        "protection",
        "chart",
        "spPr",
        "txPr",
        "externalData",
        "printSettings",
        "userShapes",
        "extLst",
      ],
    );
  return original;
}
