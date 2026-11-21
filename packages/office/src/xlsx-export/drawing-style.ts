import type { DeepReadonly, DrawingTextProperties, SourceStyleMap, SourceStyleNode, WorkbookSnapshot } from "@doffi/core";
import { effectXml, fillXml, lineXml } from "../export/drawing-paint.js";
import { cssString, points } from "../export/drawing-values.js";
import type { OutputOpc } from "../export/opc.js";
import { children, element as isElement } from "../export/source.js";
import type { SourceElement, SourceNode } from "../export/source.js";
import { empty } from "../export/xml.js";
import { paragraphProperties, runProperties } from "../pptx-export/text.js";
import { emu } from "./anchors.js";
import type { WorkbookStylesWriter } from "./styles.js";
import { bool, element, textElement } from "./xml.js";

interface DrawingStyle { readonly css: SourceStyleMap; readonly raw: SourceStyleMap }
export function shapeAppearance(style: DrawingStyle): string {
  const { css, raw } = style;
  const flatten = (values: SourceStyleMap): SourceStyleMap => Object.fromEntries(Object.entries(values).map(([key, value]) => {
    if (!key.endsWith("-color") || values[key.slice(0, -6) + "-source"] !== "reference") return [key, value];
    const prefix = key.startsWith("--office-fill") ? "--office-fill-reference" : key.startsWith("--office-line") ? "--office-line-reference" : "--office-effect-reference";
    return [key, values[prefix + "-color"] ?? value];
  }));
  const paint = flatten(css), authored = flatten(raw);
  const fill = paint["--office-fill-kind"] || paint["--office-fill-color"] ? fillXml(paint, "--office-fill", authored) : "";
  const line = paint["--office-line-kind"] || paint["--office-line-color"] ? lineXml(paint, "--office-line", authored) : "";
  return fill + line + effectXml(paint, authored);
}
export class WorkbookDrawingStyles {
  private readonly projections = new WeakMap<SourceElement, SourceStyleNode>();
  private readonly themeScope: SourceStyleNode;
  private readonly links: ReadonlyMap<string, NonNullable<WorkbookSnapshot["metadata"]["workbook"]["objectHyperlinks"]>[number]>;
  constructor(private readonly styles: WorkbookStylesWriter, private readonly opc: OutputOpc, snapshot: WorkbookSnapshot) {
    this.themeScope = { tag: "section", attrs: { class: snapshot.metadata.workbook.themeCssClass ?? "" }, parent: undefined };
    this.links = new Map(snapshot.metadata.workbook.objectHyperlinks?.map(link => [link.nodeId, link]));
  }
  private project(node: SourceElement): SourceStyleNode {
    const previous = this.projections.get(node);
    if (previous) return previous;
    const result = { tag: node.tag, attrs: node.attrs,
      parent: node.parent && node.parent.tag !== "body" ? this.project(node.parent) : this.themeScope };
    this.projections.set(node, result);
    return result;
  }
  style(node: SourceElement): DrawingStyle {
    const projection = this.project(node);
    return { css: this.styles.resolver.get(projection), raw: this.styles.resolver.specified(projection) };
  }
  classStyle(name: string | undefined): DrawingStyle {
    return name ? { css: this.styles.css(name), raw: this.styles.raw(name) } : { css: {}, raw: {} };
  }
  hyperlink(node: SourceElement, part: string): string {
    const link = this.links.get(node.attrs["data-node-id"] ?? "");
    if (!link) return "";
    const target = link.url ?? (link.location === undefined ? undefined : "#" + link.location);
    return target ? empty("a:hlinkClick", { "r:id": this.opc.relate(part, "hyperlink", target, true), tooltip: link.tooltip }) : "";
  }

  private bodyProperties(css: SourceStyleMap, text?: DeepReadonly<DrawingTextProperties>): string {
    const wrap = text?.wrap ?? (css["--office-text-wrap"] === undefined ? undefined : css["--office-text-wrap"] === "true");
    const fit = text?.autoFit ?? css["--office-text-auto-fit"];
    const scale = text?.fontScale ?? (Number(css["--office-text-font-scale"]) || 1);
    const reduction = text?.lineSpacingReduction ?? (Number(css["--office-text-line-spacing-reduction"]) || 0);
    const body = fit === "normal" ? empty("a:normAutofit", { fontScale: Math.round(scale * 100000), lnSpcReduction: Math.round(reduction * 100000) })
      : fit === "shape" ? "<a:spAutoFit/>" : fit === "none" ? "<a:noAutofit/>" : "";
    const margins = Object.fromEntries([["left", "lIns"], ["right", "rIns"], ["top", "tIns"], ["bottom", "bIns"]].map(([side, attribute]) => {
      const source = css["padding-" + side];
      return [attribute, source && source !== "initial" ? emu(points(source, css)) : undefined];
    }));
    return element("a:bodyPr", body, { ...margins, wrap: wrap === undefined ? undefined : wrap ? "square" : "none",
      anchor: css["--office-text-anchor"], anchorCtr: css["--office-text-anchor-center"] === undefined ? undefined : bool(css["--office-text-anchor-center"] === "true"),
      horzOverflow: css["--office-text-horizontal-overflow"], vertOverflow: css["--office-text-vertical-overflow"] });
  }

  private paragraph(node: SourceElement, part: string): string {
    const { css, raw } = this.style(node);
    const nativeLevel = node.attrs["data-office-level"];
    const inline = (child: SourceNode, inherited?: DrawingStyle, link = ""): string => {
      if (!isElement(child)) {
        if (!child.text) return "";
        const props = inherited ? "<a:rPr" + runProperties(inherited.css, inherited.raw, inherited.raw, { hyperlink: link || undefined }) + "</a:rPr>" : "";
        return element("a:r", props + textElement("a:t", child.text));
      }
      const style = this.style(child), combined = { css: style.css, raw: { ...inherited?.raw, ...style.raw } };
      const nextLink = this.hyperlink(child, part) || link;
      if (child.tag === "br") {
        const props = Object.keys(combined.raw).length ? "<a:rPr" + runProperties(combined.css, combined.raw) + "</a:rPr>" : "";
        return element("a:br", props);
      }
      return child.children.map(nested => inline(nested, combined, nextLink)).join("");
    };
    const end = this.classStyle(node.attrs["data-office-end-style"]);
    return element("a:p", "<a:pPr" + paragraphProperties(css, raw, raw, node, nativeLevel === undefined ? undefined : Number(nativeLevel)) +
      "</a:pPr>" + node.children.map(child => inline(child)).join("") +
      (Object.keys(end.raw).length ? "<a:endParaRPr" + runProperties(end.css, end.raw) + "</a:endParaRPr>" : ""));
  }
  richBody(tag: string, node: SourceElement, part: string, definition?: DeepReadonly<DrawingTextProperties>): string {
    const { css } = this.style(node), paragraphs = children(node, "p");
    const levels = new Set(paragraphs.flatMap(p => (p.attrs.class ?? "").split(/\s+/).map(name => /^(.*)-level-[0-8]$/.exec(name)?.[1]).filter(Boolean)));
    const style = [...levels].find(name => (node.attrs.class ?? "").split(/\s+/).includes(name!));
    const level = (name: string, className: string): string => {
      const value = this.classStyle(className);
      return Object.keys(value.raw).length ? "<a:" + name + paragraphProperties(value.css, value.raw, value.raw) + "</a:" + name + ">" : "";
    };
    const list = style ? level("defPPr", style) +
      Array.from({ length: 9 }, (_, index) => level("lvl" + (index + 1) + "pPr", style + "-level-" + index)).join("") : "";
    return element(tag, this.bodyProperties(css, definition) + element("a:lstStyle", list) +
      (paragraphs.length ? paragraphs.map(p => this.paragraph(p, part)).join("") : "<a:p/>"));
  }
  textStyle(name: string | undefined): string {
    const { raw, css } = this.classStyle(name);
    const textKeys = Object.keys(raw).filter(key => !/^--office-(fill|line|shadow|axis|label)-/.test(key));
    if (!textKeys.length) return "";
    const selected = Object.fromEntries(textKeys.map(key => [key, raw[key]!]));
    return element("c:txPr", this.bodyProperties(css) + "<a:lstStyle/><a:p><a:pPr" +
      paragraphProperties(css, selected, selected) + "</a:pPr></a:p>");
  }
}

export const drawingString = (value: string | undefined): string | undefined =>
  value === undefined ? undefined : cssString(value);
