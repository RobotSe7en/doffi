import { checkWordDrawingCoverage } from "./docx-drawing-coverage.js";
import { readWordChart } from "./docx-chart-import.js";
import { WordDrawingSchema } from "@doffi/core";
import type { WordDrawing, WordDrawingLayout, WordMetadata } from "@doffi/core";
import type { OpcPackage } from "./ooxml/package.js";
import type { XmlNode } from "./ooxml/xml.js";
import {
  A,
  C,
  PIC,
  WP,
  W,
  R,
  attr,
  child,
  elements,
  number,
  text,
  fail,
  unsupported,
  unsupportedType,
  check,
} from "./ooxml/xml.js";
import { fill, outline, effects, paint, onlyColor } from "./drawingml/styles.js";
import { integer, points } from "./drawingml/values.js";
import { declarations } from "./ooxml/styles.js";
import type { Css } from "./ooxml/styles.js";
export const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
export interface WordDrawingHtml {
  tag: string;
  attrs: Record<string, string>;
  children: (WordDrawingHtml | string)[];
}
export interface WordDrawingReadContext {
  readonly word: WordMetadata;
  readonly opc: OpcPackage;
  readonly part: string;
  readonly main: string;
  make(tag: string, attrs: Record<string, string>): WordDrawingHtml;
  asset(part: string, type: string): string;
  image(part: string, source: XmlNode): string;
  story(source: XmlNode): string;
  native(source: XmlNode, node: WordDrawingHtml): void;
  identity(source: XmlNode, id: string, nativeId: string | undefined): void;
  connector(
    source: XmlNode,
    nodeId: string,
    startId: string | undefined,
    endId: string | undefined,
    local: ReadonlyMap<string, string>,
  ): void;
}
const valueBool = (node: XmlNode, name: string, fallback = false): boolean => {
  const value = attr(node, name, "");
  if (value === undefined) return fallback;
  if (["1", "true", "on"].includes(value)) return true;
  if (["0", "false", "off"].includes(value)) return false;
  return fail(node, "A drawing Boolean is invalid.");
};
function textBody(
  body: XmlNode | undefined,
  source: XmlNode,
): NonNullable<Extract<WordDrawing, { kind: "textBox" }>["text"]> {
  if (!body) return {};
  const vertical = attr(body, "anchor", "") ?? "t",
    direction = attr(body, "vert", "") ?? "horz",
    wrap = attr(body, "wrap", "") ?? "square";
  if (
    !["t", "ctr", "b"].includes(vertical) ||
    !["horz", "vert", "vert270"].includes(direction) ||
    !["square", "none"].includes(wrap) ||
    integer(body, "numCol", 1) !== 1 ||
    valueBool(body, "rtlCol") ||
    valueBool(body, "fromWordArt") ||
    child(body, "prstTxWarp", A)
  )
    unsupported(source, "This text body layout is preserved with its native drawing.");
  const fitting = child(body, "normAutofit", A),
    autoFit = child(body, "spAutoFit", A) ? "shape" : fitting ? "normal" : "none";
  return {
    insets: {
      top: points(body, "tIns", 3.6),
      right: points(body, "rIns", 7.2),
      bottom: points(body, "bIns", 3.6),
      left: points(body, "lIns", 7.2),
    },
    vertical: vertical === "ctr" ? "middle" : vertical === "b" ? "bottom" : "top",
    direction:
      direction === "vert" ? "vertical" : direction === "vert270" ? "vertical270" : "horizontal",
    rotation: integer(body, "rot", 0) / 60000,
    wrap: wrap !== "none",
    autoFit,
    ...(fitting
      ? {
          fontScale: integer(fitting, "fontScale", 100000) / 100000,
          lineSpacingReduction: integer(fitting, "lnSpcReduction", 0) / 100000,
        }
      : {}),
  };
}
function geometry(
  source: XmlNode | undefined,
  fallback: { width: number; height: number },
  root: boolean,
): WordDrawing["geometry"] {
  const off = child(source, "off", A),
    ext = child(source, "ext", A);
  return {
    x: root ? 0 : off ? points(off, "x") : 0,
    y: root ? 0 : off ? points(off, "y") : 0,
    width: root ? fallback.width : ext ? points(ext, "cx") : fallback.width,
    height: root ? fallback.height : ext ? points(ext, "cy") : fallback.height,
    rotation: source ? integer(source, "rot", 0) / 60000 : 0,
    flipH: source ? valueBool(source, "flipH") : false,
    flipV: source ? valueBool(source, "flipV") : false,
  };
}
function appearance(
  source: XmlNode | undefined,
  style: XmlNode | undefined,
  context: WordDrawingReadContext,
  original: XmlNode,
): Css {
  let result: Css = {};
  const themePath =
    context.opc.related(context.part, "theme") ?? context.opc.related(context.main, "theme");
  const theme = themePath ? context.opc.read(themePath) : undefined,
    scheme = child(child(theme, "themeElements", A), "fmtScheme", A);
  for (const [type, name, list] of [
    ["fill", "fillRef", "fillStyleLst"],
    ["line", "lnRef", "lnStyleLst"],
    ["effect", "effectRef", "effectStyleLst"],
  ] as const) {
    const ref = child(style, name, A);
    if (!ref) continue;
    const index = integer(ref, "idx");
    if (index === 0) continue;
    const actual = type === "fill" && index >= 1001 ? index - 1000 : index,
      entry = elements(
        child(scheme, type === "fill" && index >= 1001 ? "bgFillStyleLst" : list, A),
        undefined,
        A,
      )[actual - 1];
    if (!entry) unsupported(original, "The drawing theme reference cannot be represented.");
    try {
      checkWordDrawingCoverage(entry);
      let css =
        type === "fill"
          ? fill({ ...entry, content: [entry] })
          : type === "line"
            ? outline(entry)
            : effects(child(entry, "effectLst", A));
      const prefix = "--office-" + type + "-reference";
      paint(css, prefix, onlyColor(ref));
      for (const key of Object.keys(css)) {
        css[key] = css[key]!.replaceAll("--office-placeholder-color", prefix + "-color");
        if (key.endsWith("-source") && css[key] === "placeholder") {
          css[key] = "reference";
          css[key.slice(0, -7) + "-reference-prefix"] = prefix;
        }
      }
      result = { ...result, ...css };
    } catch (error) {
      if (error instanceof Error) unsupported(original, error.message);
      throw error;
    }
  }
  return {
    ...result,
    ...fill(source),
    ...outline(child(source, "ln", A)),
    ...effects(child(source, "effectLst", A)),
  };
}
export function importWordDrawing(
  source: XmlNode,
  context: WordDrawingReadContext,
): WordDrawingHtml {
  const inline = child(source, "inline", WP),
    anchor = child(source, "anchor", WP),
    frame = inline ?? anchor;
  if (!frame || (inline && anchor))
    fail(source, "A drawing requires one inline or floating frame.");
  const frameChildren = new Set([
    "simplePos",
    "positionH",
    "positionV",
    "extent",
    "effectExtent",
    "wrapNone",
    "wrapSquare",
    "wrapTight",
    "wrapThrough",
    "wrapTopAndBottom",
    "docPr",
    "cNvGraphicFramePr",
  ]);
  for (const node of frame.content) {
    if (typeof node === "string") continue;
    const known =
      (node.uri === WP && frameChildren.has(node.name)) ||
      (node.uri === A && node.name === "graphic");
    if (!known)
      unsupportedType(
        node,
        "This drawing frame extension is preserved with its complete native drawing.",
      );
  }
  for (const position of [child(frame, "positionH", WP), child(frame, "positionV", WP)])
    if (position)
      for (const node of position.content)
        if (
          typeof node !== "string" &&
          (node.uri !== WP || !["align", "posOffset"].includes(node.name))
        )
          unsupportedType(
            node,
            "This drawing position extension is preserved with its complete native drawing.",
          );
  checkWordDrawingCoverage(source);
  const extent = child(frame, "extent", WP);
  if (!extent) fail(frame, "A drawing extent is missing.");
  const size = { width: points(extent, "cx"), height: points(extent, "cy") };
  const layout: WordDrawingLayout = {
    mode: inline ? "inline" : "floating",
    distances: {
      top: points(frame, "distT", 0),
      right: points(frame, "distR", 0),
      bottom: points(frame, "distB", 0),
      left: points(frame, "distL", 0),
    },
  };
  let x = 0,
    y = 0;
  if (anchor) {
    const h = child(anchor, "positionH", WP),
      v = child(anchor, "positionV", WP);
    if (!h || !v) fail(anchor, "A floating drawing position is missing.");
    layout.horizontal = (attr(h, "relativeFrom", "") ?? "column") as NonNullable<
      WordDrawingLayout["horizontal"]
    >;
    layout.vertical = (attr(v, "relativeFrom", "") ?? "paragraph") as NonNullable<
      WordDrawingLayout["vertical"]
    >;
    const hx = child(h, "posOffset", WP),
      vy = child(v, "posOffset", WP),
      ha = child(h, "align", WP),
      va = child(v, "align", WP);
    if (hx) x = number(text(hx), hx) / 12700;
    if (vy) y = number(text(vy), vy) / 12700;
    if (ha) layout.alignH = text(ha) as NonNullable<WordDrawingLayout["alignH"]>;
    if (va) layout.alignV = text(va) as NonNullable<WordDrawingLayout["alignV"]>;
    const wrap = elements(anchor, undefined, WP).find((node) =>
      ["wrapNone", "wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom"].includes(
        node.name,
      ),
    );
    layout.wrap =
      (
        {
          wrapNone: "none",
          wrapSquare: "square",
          wrapTight: "tight",
          wrapThrough: "through",
          wrapTopAndBottom: "topAndBottom",
        } as const
      )[wrap?.name as "wrapNone"] ?? "square";
    if (wrap && attr(wrap, "wrapText", ""))
      layout.wrapSide = attr(wrap, "wrapText", "") as NonNullable<WordDrawingLayout["wrapSide"]>;
    const polygon = child(wrap, "wrapPolygon", WP);
    if (polygon) {
      const start = child(polygon, "start", WP);
      if (!start) fail(polygon, "A wrap polygon start is missing.");
      layout.wrapPolygon = [start, ...elements(polygon, "lineTo", WP)].map((node) => ({
        x: integer(node, "x"),
        y: integer(node, "y"),
      }));
    }
    layout.behindText = valueBool(anchor, "behindDoc");
    layout.allowOverlap = valueBool(anchor, "allowOverlap", true);
    layout.layoutInCell = valueBool(anchor, "layoutInCell", true);
    layout.locked = valueBool(anchor, "locked");
    layout.zOrder = integer(anchor, "relativeHeight", 0);
  }
  const data = child(child(frame, "graphic", A), "graphicData", A),
    objects = data?.content.filter((node): node is XmlNode => typeof node !== "string") ?? [];
  if (objects.length !== 1) unsupported(source, "A drawing graphic requires one supported object.");
  const localIds = new Map<string, string>();
  const parse = (native: XmlNode, root = false): WordDrawingHtml => {
    let definition: WordDrawing, node: WordDrawingHtml;
    const properties = child(
        native,
        native.uri === PIC ? "spPr" : native.uri === WPG ? "grpSpPr" : "spPr",
        native.uri,
      ),
      transform =
        native.uri === WPG && native.name === "graphicFrame"
          ? child(native, "xfrm", WPG)
          : child(properties, "xfrm", A),
      g = geometry(transform, size, root),
      mode = root ? layout : { mode: "inline" as const };
    if (root) {
      g.x = x;
      g.y = y;
    }
    const knownProperties =
      native.uri === WPG
        ? ["xfrm"]
        : native.uri === PIC
          ? ["xfrm", "prstGeom", "noFill", "ln", "effectLst"]
          : ["xfrm", "prstGeom", "solidFill", "noFill", "gradFill", "ln", "effectLst"];
    if (
      properties?.content.some(
        (node) =>
          typeof node !== "string" && (node.uri !== A || !knownProperties.includes(node.name)),
      )
    )
      unsupported(source, "This drawing appearance is preserved as a native object.");
    if (
      (native.uri === C && native.name === "chart") ||
      (native.uri === WPG && native.name === "graphicFrame")
    ) {
      const reference =
        native.uri === C
          ? native
          : child(child(child(native, "graphic", A), "graphicData", A), "chart", C);
      if (!reference) unsupported(source, "This grouped graphic remains native.");
      node = context.make("span", { "data-word-drawing": "chart" });
      node.children.push("\ufffc");
      let chart;
      try {
        chart = readWordChart(context.opc, context.part, context.main, reference, context.asset);
      } catch (error) {
        if (error instanceof Error && "code" in error && error.code === "UnsupportedFeature")
          unsupported(source, error.message);
        throw error;
      }
      definition = {
        nodeId: node.attrs["data-node-id"]!,
        kind: "chart",
        geometry: g,
        layout: mode,
        chart,
      };
      const properties = chart.native
          ? child(context.opc.read(chart.native.chartPart), "spPr", C)
          : undefined,
        css: Css = {};
      for (const read of [
        () => fill(properties),
        () => outline(child(properties, "ln", A)),
        () => effects(child(properties, "effectLst", A)),
      ])
        try {
          Object.assign(css, read());
        } catch (error) {
          if (!(error instanceof Error && "code" in error && error.code === "UnsupportedFeature"))
            throw error;
        }
      if (Object.keys(css).length) node.attrs.style = declarations(css);
    } else if (native.uri === PIC && native.name === "pic") {
      const fillSource = child(native, "blipFill", PIC),
        blip = child(fillSource, "blip", A),
        relation = context.opc.relationships(context.part).get(attr(blip, "embed", R) ?? "");
      if (!relation || relation.external || !relation.type.endsWith("/image"))
        unsupported(source, "Only authorized embedded image relationships can be displayed.");
      if (child(fillSource, "tile", A))
        unsupported(source, "Tiled image fills are preserved as native objects.");
      if (elements(blip, undefined, A).some((effect) => effect.name !== "extLst"))
        unsupported(source, "This picture color effect is preserved as native content.");
      node = context.make("img", {
        "data-word-drawing": "image",
        src: context.image(relation.target, source),
        alt:
          attr(child(child(native, "nvPicPr", PIC), "cNvPr", PIC), "descr", "") ??
          attr(child(frame, "docPr", WP), "descr", "") ??
          "",
      });
      const crop = child(fillSource, "srcRect", A);
      definition = {
        nodeId: node.attrs["data-node-id"]!,
        kind: "image",
        geometry: g,
        layout: mode,
        ...(crop
          ? {
              crop: {
                left: integer(crop, "l", 0) / 100000,
                top: integer(crop, "t", 0) / 100000,
                right: integer(crop, "r", 0) / 100000,
                bottom: integer(crop, "b", 0) / 100000,
              },
            }
          : {}),
      };
      const geometryNode = child(properties, "prstGeom", A);
      if (
        (geometryNode && attr(geometryNode, "prst", "") !== "rect") ||
        child(properties, "custGeom", A)
      )
        unsupported(source, "This image geometry is preserved as native content.");
      const line = child(properties, "ln", A),
        lineStyle = outline(line);
      if (line && !lineStyle["--office-line-kind"] && !lineStyle["--office-line-color"])
        Object.assign(lineStyle, {
          "--office-line-kind": "solid",
          "--office-line-color": "#000000",
        });
      node.attrs.style = declarations({
        ...lineStyle,
        ...effects(child(properties, "effectLst", A)),
      });
    } else if (native.uri === WPG && ["wgp", "grpSp"].includes(native.name)) {
      node = context.make("span", { "data-word-drawing": "group" });
      const off = child(transform, "chOff", A),
        ext = child(transform, "chExt", A);
      if (!off || !ext) fail(native, "A group child coordinate system is missing.");
      definition = {
        nodeId: node.attrs["data-node-id"]!,
        kind: "group",
        geometry: g,
        layout: mode,
        coordinateSpace: {
          x: points(off, "x"),
          y: points(off, "y"),
          width: points(ext, "cx"),
          height: points(ext, "cy"),
        },
      };
      for (const child of native.content)
        if (
          typeof child !== "string" &&
          !["cNvPr", "cNvGrpSpPr", "grpSpPr", "extLst"].includes(child.name)
        )
          node.children.push(parse(child));
    } else if (native.uri === WPS && native.name === "wsp") {
      check(
        native,
        ["cNvPr", "cNvSpPr", "cNvCnPr", "spPr", "style", "txbx", "bodyPr", "extLst"],
        WPS,
      );
      const connector = child(native, "cNvCnPr", WPS),
        preset = child(properties, "prstGeom", A),
        name = attr(preset, "prst", "") ?? "rect";
      if (child(properties, "custGeom", A))
        unsupported(source, "Custom shape geometry is preserved as a native object.");
      const kind = connector
        ? "connector"
        : valueBool(child(native, "cNvSpPr", WPS) ?? native, "txBox")
          ? "textBox"
          : "shape";
      node = context.make("span", { "data-word-drawing": kind });
      node.children.push("\ufffc");
      const base = { nodeId: node.attrs["data-node-id"]!, geometry: g, layout: mode };
      if (connector) {
        const start = child(connector, "stCxn", A),
          end = child(connector, "endCxn", A);
        definition = {
          ...base,
          kind: "connector",
          connector: {
            routing: name === "bentConnector3" ? "elbow" : "straight",
            ...(start ? { startSite: integer(start, "idx", 0) } : {}),
            ...(end ? { endSite: integer(end, "idx", 0) } : {}),
          },
        };
        context.connector(
          source,
          base.nodeId,
          attr(start, "id", ""),
          attr(end, "id", ""),
          localIds,
        );
      } else {
        const body = child(child(native, "txbx", WPS), "txbxContent", W),
          storyId = body ? context.story(body) : undefined;
        if (kind === "textBox" && !storyId) fail(native, "A text box is missing its Word content.");
        definition =
          kind === "textBox"
            ? { ...base, kind: "textBox", textStoryId: storyId! }
            : {
                ...base,
                kind: "shape",
                preset: name as Extract<WordDrawing, { kind: "shape" }>["preset"],
                ...(storyId ? { textStoryId: storyId } : {}),
              };
        if (definition.kind === "shape" && preset) {
          const adjustments = elements(child(preset, "avLst", A), "gd", A);
          if (adjustments.length)
            definition.adjustments = Object.fromEntries(
              adjustments.map((adjustment) => {
                const name = attr(adjustment, "name", ""),
                  formula = attr(adjustment, "fmla", ""),
                  match = /^val (-?\d+)$/.exec(formula ?? "");
                if (!name || !match)
                  unsupported(
                    source,
                    "Custom shape guide formulas are preserved as native content.",
                  );
                return [name, Number(match[1])];
              }),
            );
        }
      }
      node.attrs.style = declarations(
        appearance(properties, child(native, "style", WPS), context, source),
      );
      if (definition.kind === "shape" || definition.kind === "textBox")
        definition.text = textBody(child(native, "bodyPr", WPS), source);
    } else return unsupported(source, "This drawing object is preserved as native content.");
    const parsed = WordDrawingSchema.safeParse(definition);
    if (!parsed.success)
      unsupported(source, "This drawing geometry or layout is outside the editable model.");
    (context.word.drawings ??= []).push(parsed.data);
    const identity =
      native.uri === PIC
        ? child(child(native, "nvPicPr", PIC), "cNvPr", PIC)
        : child(native, "cNvPr", native.uri);
    if (!root && attr(identity, "name", "")) parsed.data.name = attr(identity, "name", "")!;
    const nativeId = attr(identity, "id", "");
    if (nativeId && nativeId !== "0") localIds.set(nativeId, node.attrs["data-node-id"]!);
    context.identity(native, node.attrs["data-node-id"]!, nativeId);
    context.native(native, node);
    return node;
  };
  const result = parse(objects[0]!, true),
    description = child(frame, "docPr", WP);
  if (description)
    context.identity(description, result.attrs["data-node-id"]!, attr(description, "id", ""));
  const definition = context.word.drawings!.find(
    (item) => item.nodeId === result.attrs["data-node-id"],
  )!;
  if (attr(description, "name", "")) {
    definition.name = attr(description, "name", "")!;
    if (/^(?:DoffiWatermark-|PowerPlusWaterMarkObject|WordPictureWatermark)/i.test(definition.name))
      definition.watermark = true;
  }
  context.native(source, result);
  return result;
}
