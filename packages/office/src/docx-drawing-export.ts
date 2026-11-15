import type { DeepReadonly, WordDrawing, WordMetadata } from "@doffi/core";
import type { SourceElement, SourceTree } from "./export/source.js";
import { children } from "./export/source.js";
import type { Css } from "./export/css.js";
import { attributes, empty, escapeXml } from "./export/xml.js";
import { emu, integer, unsupported } from "./export/drawing-values.js";
import { fillXml, lineXml, effectXml } from "./export/drawing-paint.js";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
export interface WordDrawingWriteContext {
  readonly word: DeepReadonly<WordMetadata>;
  readonly tree: SourceTree;
  css(node: SourceElement): Css;
  raw(node: SourceElement): Css;
  id(nodeId: string): number;
  image(node: SourceElement): string;
  chart(chart: DeepReadonly<Extract<WordDrawing, { kind: "chart" }>>): string;
  text(storyId: string, width: number): string;
  marker(node: SourceElement): string;
}
export function writeWordDrawing(
  context: WordDrawingWriteContext,
  drawing: DeepReadonly<WordDrawing>,
): string {
  const identity = (
    tag: string,
    attrs: Readonly<Record<string, string | number | undefined>>,
    node: SourceElement,
  ): string =>
    empty(tag, attrs).replace(
      "/>",
      (context.word.nativeContent
        ? ' xmlns:dn="urn:doffi:transient:native" dn:identity="' +
          escapeXml(node.attrs["data-node-id"]!) +
          '"'
        : "") + "/>",
    );
  const definitions = new Map((context.word.drawings ?? []).map((item) => [item.nodeId, item]));
  const payload = (item: DeepReadonly<WordDrawing>, nested = false): string => {
    const node = context.tree.byId.get(item.nodeId)!,
      g = item.geometry,
      id = context.id(item.nodeId),
      css = context.css(node),
      raw = context.raw(node),
      xfrm = (group = false): string =>
        "<a:xfrm" +
        attributes({
          rot: integer(((g.rotation % 360) + 360) % 360, 60000, 0, 21600000),
          flipH: g.flipH ? 1 : 0,
          flipV: g.flipV ? 1 : 0,
        }) +
        ">" +
        empty("a:off", { x: nested ? emu(g.x) : 0, y: nested ? emu(g.y) : 0 }) +
        empty("a:ext", { cx: emu(g.width, 0), cy: emu(g.height, 0) }) +
        (group && item.kind === "group"
          ? empty("a:chOff", { x: emu(item.coordinateSpace.x), y: emu(item.coordinateSpace.y) }) +
            empty("a:chExt", {
              cx: emu(item.coordinateSpace.width, 1),
              cy: emu(item.coordinateSpace.height, 1),
            })
          : "") +
        "</a:xfrm>";
    if (item.kind === "chart") {
      const chart = empty("c:chart", { "xmlns:c": C, "r:id": context.chart(item) });
      return nested
        ? "<wpg:graphicFrame>" +
            identity("wpg:cNvPr", { id, name: item.name ?? "Chart " + id }, node) +
            "<wpg:cNvFrPr/>" +
            xfrm().replace("<a:xfrm", "<wpg:xfrm").replace("</a:xfrm>", "</wpg:xfrm>") +
            '<a:graphic><a:graphicData uri="' +
            C +
            '">' +
            chart +
            "</a:graphicData></a:graphic></wpg:graphicFrame>"
        : chart;
    }
    if (item.kind === "image") {
      const crop = item.crop;
      return (
        "<pic:pic><pic:nvPicPr>" +
        identity(
          "pic:cNvPr",
          { id, name: item.name ?? "Picture " + id, descr: node.attrs.alt ?? "" },
          node,
        ) +
        "<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>" +
        empty("a:blip", { "r:embed": context.image(node) }) +
        (crop
          ? empty("a:srcRect", {
              l: integer(crop.left, 100000, 0, 99999),
              t: integer(crop.top, 100000, 0, 99999),
              r: integer(crop.right, 100000, 0, 99999),
              b: integer(crop.bottom, 100000, 0, 99999),
            })
          : "") +
        "<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>" +
        xfrm() +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
        lineXml(css, "--office-line", raw) +
        effectXml(css, raw) +
        "</pic:spPr></pic:pic>"
      );
    }
    if (item.kind === "group") {
      const tag = nested ? "wpg:grpSp" : "wpg:wgp";
      return (
        "<" +
        tag +
        ">" +
        (nested ? identity("wpg:cNvPr", { id, name: item.name ?? "Group " + id }, node) : "") +
        "<wpg:cNvGrpSpPr/><wpg:grpSpPr>" +
        xfrm(true) +
        "</wpg:grpSpPr>" +
        children(node)
          .map((child) => {
            const definition = definitions.get(child.attrs["data-node-id"]!);
            if (!definition) unsupported("A group source is missing its drawing definition.");
            return payload(definition!, true);
          })
          .join("") +
        "</" +
        tag +
        ">"
      );
    }
    const connector = item.kind === "connector",
      preset = connector
        ? item.connector.routing === "elbow"
          ? "bentConnector3"
          : "straightConnector1"
        : item.kind === "shape"
          ? item.preset
          : "rect",
      appearance = { ...css };
    if (
      !appearance["--office-fill-kind"] &&
      !appearance["--office-fill-color"] &&
      appearance["background-color"] &&
      appearance["background-color"] !== "transparent"
    ) {
      appearance["--office-fill-kind"] = "solid";
      appearance["--office-fill-color"] = appearance["background-color"]!;
    }
    const connections = connector
      ? (["start", "end"] as const)
          .map((side) => {
            const target = item.connector[(side + "NodeId") as "startNodeId" | "endNodeId"];
            return target
              ? empty("a:" + (side === "start" ? "stCxn" : "endCxn"), {
                  id: context.id(target),
                  idx: item.connector[(side + "Site") as "startSite" | "endSite"] ?? 0,
                })
              : "";
          })
          .join("")
      : "";
    const adjustment =
      item.kind === "shape"
        ? Object.entries(item.adjustments ?? {})
            .map(([name, value]) => empty("a:gd", { name, fmla: "val " + integer(value) }))
            .join("")
        : "";
    const body = "text" in item ? item.text : undefined,
      insets = body?.insets ?? { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
      text =
        "textStoryId" in item && item.textStoryId
          ? "<wps:txbx><w:txbxContent>" +
            context.text(item.textStoryId, Math.max(0.01, g.width - insets.left - insets.right)) +
            "</w:txbxContent></wps:txbx>"
          : "";
    const bodyXml =
      "<wps:bodyPr" +
      attributes({
        rot: integer(body?.rotation ?? 0, 60000),
        wrap: body?.wrap === false ? "none" : "square",
        lIns: emu(insets.left),
        rIns: emu(insets.right),
        tIns: emu(insets.top),
        bIns: emu(insets.bottom),
        anchor: body?.vertical === "middle" ? "ctr" : body?.vertical === "bottom" ? "b" : "t",
        vert:
          body?.direction === "vertical"
            ? "vert"
            : body?.direction === "vertical270"
              ? "vert270"
              : "horz",
      }) +
      ">" +
      (body?.autoFit === "shape"
        ? "<a:spAutoFit/>"
        : body?.autoFit === "normal"
          ? empty("a:normAutofit", {
              fontScale: integer(body.fontScale ?? 1, 100000, 1, 100000),
              lnSpcReduction: integer(body.lineSpacingReduction ?? 0, 100000, 0, 100000),
            })
          : "<a:noAutofit/>") +
      "</wps:bodyPr>";
    return (
      "<wps:wsp>" +
      (nested
        ? identity(
            "wps:cNvPr",
            { id, name: item.name ?? (connector ? "Connector " : "Shape ") + id },
            node,
          )
        : "") +
      (connector
        ? "<wps:cNvCnPr>" + connections + "</wps:cNvCnPr>"
        : empty("wps:cNvSpPr", { txBox: item.kind === "textBox" ? 1 : 0 })) +
      "<wps:spPr>" +
      xfrm() +
      "<a:prstGeom" +
      attributes({ prst: preset }) +
      "><a:avLst>" +
      adjustment +
      "</a:avLst></a:prstGeom>" +
      fillXml(
        appearance,
        "--office-fill",
        raw,
        false,
        connector || item.kind === "textBox" ? "none" : "#4472c4",
      ) +
      lineXml(appearance, "--office-line", raw, false, connector ? "#224466" : "none") +
      effectXml(appearance, raw) +
      "</wps:spPr>" +
      text +
      bodyXml +
      "</wps:wsp>"
    );
  };
  const g = drawing.geometry,
    layout = drawing.layout,
    inline = layout.mode === "inline",
    distances = layout.distances ?? { top: 0, right: 0, bottom: 0, left: 0 },
    distance = {
      distT: emu(distances.top, 0),
      distR: emu(distances.right, 0),
      distB: emu(distances.bottom, 0),
      distL: emu(distances.left, 0),
    };
  let opening = inline
    ? "<wp:inline" + attributes(distance) + ">"
    : "<wp:anchor" +
      attributes({
        ...distance,
        simplePos: 0,
        relativeHeight: integer(layout.zOrder ?? 0, 1, 0, 4294967295),
        behindDoc: layout.behindText ? 1 : 0,
        locked: layout.locked ? 1 : 0,
        layoutInCell: layout.layoutInCell === false ? 0 : 1,
        allowOverlap: layout.allowOverlap === false ? 0 : 1,
      }) +
      '><wp:simplePos x="0" y="0"/><wp:positionH' +
      attributes({ relativeFrom: layout.horizontal ?? "column" }) +
      ">" +
      (layout.alignH
        ? "<wp:align>" + escapeXml(layout.alignH) + "</wp:align>"
        : "<wp:posOffset>" + emu(g.x) + "</wp:posOffset>") +
      "</wp:positionH><wp:positionV" +
      attributes({ relativeFrom: layout.vertical ?? "paragraph" }) +
      ">" +
      (layout.alignV
        ? "<wp:align>" + escapeXml(layout.alignV) + "</wp:align>"
        : "<wp:posOffset>" + emu(g.y) + "</wp:posOffset>") +
      "</wp:positionV>";
  opening +=
    empty("wp:extent", { cx: emu(g.width, 0), cy: emu(g.height, 0) }) +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>';
  if (!inline) {
    const wrap = layout.wrap ?? "square";
    if (wrap === "none") opening += "<wp:wrapNone/>";
    else if (wrap === "topAndBottom") opening += "<wp:wrapTopAndBottom/>";
    else if (wrap === "square")
      opening += empty("wp:wrapSquare", { wrapText: layout.wrapSide ?? "bothSides" });
    else {
      const polygon = layout.wrapPolygon ?? [
          { x: 0, y: 0 },
          { x: 21600, y: 0 },
          { x: 21600, y: 21600 },
          { x: 0, y: 21600 },
        ],
        tag = wrap === "tight" ? "wrapTight" : "wrapThrough";
      opening +=
        "<wp:" +
        tag +
        attributes({ wrapText: layout.wrapSide ?? "bothSides" }) +
        '><wp:wrapPolygon edited="0">' +
        empty("wp:start", { x: integer(polygon[0]!.x), y: integer(polygon[0]!.y) }) +
        [...polygon.slice(1), polygon[0]!]
          .map((point) => empty("wp:lineTo", { x: integer(point.x), y: integer(point.y) }))
          .join("") +
        "</wp:wrapPolygon></wp:" +
        tag +
        ">";
    }
  }
  const carrier =
    drawing.kind === "chart" &&
    (drawing.geometry.rotation % 360 !== 0 || drawing.geometry.flipH || drawing.geometry.flipV);
  const uri = carrier
      ? WPG
      : drawing.kind === "chart"
        ? C
        : drawing.kind === "image"
          ? PIC
          : drawing.kind === "group"
            ? WPG
            : WPS,
    id = context.id(drawing.nodeId),
    node = context.tree.byId.get(drawing.nodeId)!;
  return (
    "<w:r" +
    context.marker(node) +
    '><w:drawing xmlns:wps="' +
    WPS +
    '" xmlns:wpg="' +
    WPG +
    '">' +
    opening +
    identity(
      "wp:docPr",
      { id, name: drawing.name ?? "Drawing " + id, descr: node.attrs.alt ?? "" },
      node,
    ) +
    "<wp:cNvGraphicFramePr/><a:graphic><a:graphicData" +
    attributes({ uri }) +
    ">" +
    (carrier
      ? '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>' +
        empty("a:ext", { cx: emu(g.width), cy: emu(g.height) }) +
        '<a:chOff x="0" y="0"/>' +
        empty("a:chExt", { cx: emu(g.width), cy: emu(g.height) }) +
        "</a:xfrm></wpg:grpSpPr>" +
        payload({ ...drawing, geometry: { ...g, x: 0, y: 0 } }, true) +
        "</wpg:wgp>"
      : payload(drawing)) +
    "</a:graphicData></a:graphic></wp:" +
    (inline ? "inline" : "anchor") +
    "></w:drawing></w:r>"
  );
}
