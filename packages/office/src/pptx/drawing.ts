import { PresentationPresetSchema } from "@doffi/core";
import type { PresentationObject } from "@doffi/core";
import { A, attr, C, check, child, elements, fail, MC, P, R, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { bool, Context, integer, points, unsigned, value } from "./context.js";
import type { Html, Surface } from "./context.js";
import { colorMap, effects, fill, onlyColor, outline, paint } from "./styles.js";
import { importHyperlink, importText } from "./text.js";
import { importTable } from "./table.js";
import { importChart } from "./chart.js";
import { customGeometry } from "./custom-geometry.js";

/** Select markup compatibility by namespace support, never prefix spelling. */
export function compatible(source: XmlNode): XmlNode[] {
  const output: XmlNode[] = [];
  for (const node of source.content) {
    if (typeof node === "string") { if (node.trim()) fail(source, "Unexpected DrawingML text."); continue; }
    if (node.uri !== MC || node.name !== "AlternateContent") { output.push(node); continue; }
    const choices = elements(node, "Choice", MC), fallbacks = elements(node, "Fallback", MC);
    if (fallbacks.length > 1) fail(node, "Markup compatibility fallback is repeated.");
    const selected = choices.find(choice => {
      const requirements = value(choice, "Requires")?.trim().split(/\s+/);
      return !!requirements?.length && requirements.every(prefix => [A, P, R].includes(choice.namespaces[prefix] ?? ""));
    }) ?? fallbacks[0];
    if (!selected) unsupported(node, "No supported markup compatibility alternative is available.");
    output.push(...compatible(selected));
  }
  return output;
}

export function transform(source: XmlNode | undefined, target: Html, group = false): void {
  if (!source) return;
  check(source, group ? ["off", "ext", "chOff", "chExt"] : ["off", "ext"], A);
  const off = child(source, "off", A), ext = child(source, "ext", A);
  if (!off || !ext) fail(source, "An object transform is incomplete.");
  const geometry = { x: points(off, "x"), y: points(off, "y"), width: points(ext, "cx"), height: points(ext, "cy") };
  if (geometry.width < 0 || geometry.height < 0) fail(ext, "Object extents cannot be negative.");
  for (const [key, v] of Object.entries(geometry)) target.attrs["data-office-" + key] = String(v);
  if (value(source, "rot") !== undefined) target.attrs["data-office-rotation"] = String(integer(source, "rot") / 60000);
  for (const [native, sourceName] of [["flipH", "flip-h"], ["flipV", "flip-v"]]) if (value(source, native!) !== undefined) target.attrs["data-office-" + sourceName] = String(bool(source, native!));
  if (group) {
    const childOff = child(source, "chOff", A), childExt = child(source, "chExt", A);
    if (!childOff || !childExt) fail(source, "A group child coordinate transform is missing.");
    for (const [name, n, key] of [["x", childOff, "x"], ["y", childOff, "y"], ["width", childExt, "cx"], ["height", childExt, "cy"]] as const) {
      const v = points(n, key);
      if (["width", "height"].includes(name) && v <= 0) unsupported(source, "A group coordinate transform is not invertible.");
      target.attrs["data-office-child-" + name] = String(v);
    }
  }
}

function objectStyle(context: Context, source: XmlNode | undefined, properties: XmlNode | undefined, themeId: string, target: Html): void {
  const classes = [target.attrs.class];
  const css: Css = { ...fill(properties), ...outline(child(properties, "ln", A)), ...effects(child(properties, "effectLst", A)) };
  if (source) {
    check(source, ["lnRef", "fillRef", "effectRef", "fontRef"], A);
    for (const [name, kind] of [["lnRef", "line"], ["fillRef", "fill"], ["effectRef", "effect"]] as const) {
      const reference = child(source, name, A);
      if (!reference) continue;
      const index = integer(reference, "idx");
      if (index < 0) fail(reference, "A style reference cannot be negative.");
      if (index) {
        const style = themeId + "-" + kind + "-" + index;
        if (!context.css.some(rule => rule.startsWith("." + style + "{"))) fail(reference, "A theme style reference is missing.");
        classes.push(style);
      } else if (kind !== "effect") css["--office-" + kind + "-kind"] ??= "none";
      if (elements(reference, undefined, A).length) {
        const c = onlyColor(reference); paint(css, "--office-" + kind + "-reference", c);
      }
    }
    const font = child(source, "fontRef", A);
    if (font) {
      const role = value(font, "idx");
      if (!role || !["major", "minor", "none"].includes(role)) fail(font, "A theme font reference is invalid.");
      if (role !== "none") {
        for (const script of ["ascii", "h-ansi", "east-asia", "complex-script"]) css["--office-font-" + script] = "var(--office-" + role + "-" + script + ")";
        css["font-family"] = "var(--office-font-ascii),var(--office-font-east-asia)";
      }
      if (elements(font, undefined, A).length) { const c = onlyColor(font); css.color = c.css; paint(css, "--office-text", c); }
    }
  }
  target.attrs.class = classes.filter(Boolean).join(" "); context.style(target, css);
}

export class Drawings {
  private readonly connections: { object: Extract<PresentationObject, { kind: "connector" }>; start?: string; end?: string; surface: string }[] = [];
  private readonly objects: Map<string, PresentationObject> = new Map();
  private readonly placeholders: Map<string, Map<number, PresentationObject[]>> = new Map();
  constructor(readonly context: Context) {}
  private retainTags(owner: XmlNode | undefined, ownerId: string): void {
    const source = child(owner, "custDataLst", P);
    if (!source || !this.context.compatiblePreview) return;
    check(source, ["tags"], P);
    const tags = child(source, "tags", P);
    if (!tags) return;
    check(tags, [], P);
    const path = this.context.relation(tags.part, attr(tags, "id", R), "tags").target;
    const data = this.context.part(path, "presentationml.tags", "tagLst", P);
    check(data, ["tag"], P);
    for (const tag of elements(data, "tag", P)) {
      check(tag, [], P);
      if (value(tag, "name") === undefined || value(tag) === undefined) fail(tag, "A presentation tag requires a name and value.");
    }
    this.context.retainPreview("tags", ownerId, data, "Presentation tags require compatible preview.");
  }
  retainAnimation(surface: Surface, timing: XmlNode): void {
    const supported = "anim animEffect attrName attrNameLst bldLst bldP cBhvr cTn childTnLst cond nextCondLst pRg par prevCondLst seq set sldTgt spTgt stCondLst strVal tav tavLst tgtEl timing tn tnLst to txEl val".split(" ");
    const identities = this.context.nativeNodes.get(surface.id)!;
    const visit = (node: XmlNode): void => {
      if (node.uri !== P || !supported.includes(node.name) || [...node.attrs.keys()].some(key => key.startsWith(R + " ")))
        unsupported(node, "Only relationship-free slide object animations can be retained in static preview.");
      if (value(node, "spid") !== undefined && !identities.has(String(unsigned(node, "spid")))) fail(node, "An animation has a missing slide object target.");
      for (const item of node.content) {
        if (typeof item !== "string") visit(item);
        else if (item.trim() && node.name !== "attrName") fail(node, "An animation node contains unexpected text.");
      }
    };
    if (!this.context.compatiblePreview) unsupported(timing, "Animation is not supported in strict import.");
    visit(timing);
    this.context.retainPreview("animation", surface.id, timing, "Animation requires compatible preview.");
  }
  surface(surface: Surface): void {
    const { context } = this, source = surface.source;
    const content = child(source, "cSld", P), tree = child(content, "spTree", P);
    if (!content || !tree) fail(source, "A presentation drawing tree is missing.");
    this.retainTags(content, surface.root.attrs["data-node-id"]!);
    check(content, ["spTree", "bg", "custDataLst", "controls", "extLst"], P);
    if (child(content, "controls", P)) unsupported(content, "Embedded controls are not supported.");
    const map = child(source, "clrMap", P), override = child(source, "clrMapOvr", P);
    check(override, ["masterClrMapping", "overrideClrMapping"], A);
    if (override && elements(override, undefined, A).length !== 1) fail(override, "A color mapping override requires one mode.");
    const css: Css = colorMap(map ?? child(override, "overrideClrMapping", A));
    const headerFooter = child(source, "hf", P);
    if (headerFooter) for (const [key, name] of [["hdr", "header"], ["ftr", "footer"], ["dt", "date"], ["sldNum", "slide-number"]]) surface.root.attrs["data-office-show-" + name] = String(bool(headerFooter, key!, true));
    if (source.name === "sldLayout" && value(source, "type")) surface.root.attrs["data-office-layout-type"] = value(source, "type")!;
    if (source.name === "sldMaster" && value(content, "name") !== undefined) surface.root.attrs["data-office-name"] = value(content, "name")!;
    const background = child(content, "bg", P);
    if (background) {
      check(background, ["bgPr", "bgRef"], P);
      const properties = child(background, "bgPr", P), reference = child(background, "bgRef", P);
      if (!!properties === !!reference) fail(background, "A slide background requires exactly one representation.");
      if (properties) { check(properties, ["noFill", "solidFill", "gradFill", "effectLst"], A); Object.assign(css, fill(properties), effects(child(properties, "effectLst", A))); }
      if (reference) {
        const index = integer(reference, "idx") - 1000, name = surface.themeId + "-background-" + index;
        if (!context.css.some(rule => rule.startsWith("." + name + "{"))) fail(reference, "A theme background reference is missing.");
        surface.root.attrs.class += " " + name; paint(css, "--office-background-reference", onlyColor(reference));
      }
    }
    context.style(surface.root, css);
    const native = child(child(tree, "nvGrpSpPr", P), "cNvPr", P);
    if (!native) fail(tree, "A drawing tree identity is invalid.");
    const nativeId = String(unsigned(native, "id"));
    context.nativeNodes.set(surface.id, new Map([[nativeId, surface.root.attrs["data-node-id"]!]]));
    const rootTransform = child(child(tree, "grpSpPr", P), "xfrm", A);
    let parent = surface.root;
    if (rootTransform && ["rot", "flipH", "flipV"].some(name => value(rootTransform, name) !== undefined && value(rootTransform, name) !== "0") || rootTransform && [child(rootTransform, "off", A), child(rootTransform, "ext", A), child(rootTransform, "chOff", A), child(rootTransform, "chExt", A)].some(n => n && [...n.attrs.values()].some(v => v !== "0"))) {
      const group = context.make("div", { "data-office-kind": "group" }); transform(rootTransform, group, true);
      const object: PresentationObject = { kind: "group", nodeId: group.attrs["data-node-id"]!, nativeId };
      context.metadata.objects.push(object); this.objects.set(object.nodeId, object); parent.children.push(group); parent = group;
      context.nativeNodes.get(surface.id)!.set(nativeId, group.attrs["data-node-id"]!);
    }
    for (const node of compatible(tree)) {
      if (["nvGrpSpPr", "grpSpPr"].includes(node.name) && node.uri === P) continue;
      parent.children.push(this.object(node, surface));
    }
    const slots = new Map<number, PresentationObject[]>();
    for (const child of surface.root.children) if (typeof child !== "string") {
      const object = this.objects.get(child.attrs["data-node-id"]!);
      if (object?.placeholder) { const list = slots.get(object.placeholder.index) ?? []; list.push(object); slots.set(object.placeholder.index, list); }
    }
    this.placeholders.set(surface.id, slots);
  }
  object(source: XmlNode, surface: Surface, topLevel = true): Html {
    const { context } = this;
    if (source.uri !== P || !["sp", "pic", "grpSp", "cxnSp", "graphicFrame"].includes(source.name)) unsupported(source, "This slide object is not supported.");
    const n = source.name;
    const nvName = ({ sp: "nvSpPr", pic: "nvPicPr", grpSp: "nvGrpSpPr", cxnSp: "nvCxnSpPr", graphicFrame: "nvGraphicFramePr" } as Record<string, string>)[n]!;
    const content = compatible(source);
    check({ ...source, content: content.filter(node => !(n === "graphicFrame" && node.uri === A && node.name === "graphic")) },
      [nvName, n === "grpSp" ? "grpSpPr" : n === "graphicFrame" ? "xfrm" : "spPr", "style", ...(n === "sp" ? ["txBody"] : n === "pic" ? ["blipFill"] : n === "grpSp" ? ["sp", "pic", "grpSp", "cxnSp", "graphicFrame"] : []), "extLst"], P);
    // a:graphic is the sole cross-namespace child of a graphic frame.
    const nv = child(source, nvName, P), native = child(nv, "cNvPr", P);
    if (!native) fail(source, "An object identity is invalid.");
    const id = String(unsigned(native, "id"));
    const identities = context.nativeNodes.get(surface.id)!;
    if (identities.has(id)) fail(native, "Native object identities repeat inside a surface.");
    const target = context.make("div"), nodeId = target.attrs["data-node-id"]!;
    check(native, ["hlinkClick", "hlinkHover", "extLst"], A);
    if (child(native, "hlinkHover", A)) unsupported(native, "Hover actions are not supported.");
    importHyperlink(context, child(native, "hlinkClick", A), target, source.part);
    const applicationProperties = child(nv, "nvPr", P);
    check(applicationProperties, ["ph", "extLst", "custDataLst"], P);
    this.retainTags(applicationProperties, nodeId);
    identities.set(id, nodeId);
    const base: { nodeId: string; nativeId: string; name?: string; creationId?: string; placeholder?: NonNullable<PresentationObject["placeholder"]> } = { nodeId, nativeId: id };
    if (value(native, "name") !== undefined) base.name = value(native, "name")!;
    if (value(native, "descr") !== undefined) target.attrs["data-office-alt"] = value(native, "descr")!;
    if (bool(native, "hidden")) target.attrs["data-office-hidden"] = "true";
    const creation = elements(child(native, "extLst", A), "ext", A).flatMap(ext => ext.content.filter((c): c is XmlNode => typeof c !== "string")).find(c => c.name === "creationId" && c.uri === "http://schemas.microsoft.com/office/drawing/2014/main");
    if (creation && value(creation, "id")) base.creationId = value(creation, "id")!;
    const ph = child(child(nv, "nvPr", P), "ph", P);
    if (ph) {
      let type = value(ph, "type");
      const index = value(ph, "idx") === undefined ? 0 : unsigned(ph, "idx");
      if (topLevel && surface.layoutId && (type === undefined || type === "obj")) {
        const candidates = this.placeholders.get(surface.layoutId)?.get(index) ?? [];
        const matching = candidates.filter(o => o.placeholder!.type === (type ?? "obj"));
        if (matching.length === 1) type = matching[0]!.placeholder!.type;
        else if (candidates.length === 1) type = candidates[0]!.placeholder!.type;
        else if (candidates.length > 1) fail(ph, "An inherited native placeholder index is ambiguous.");
      }
      base.placeholder = { type: (type ?? "obj") as NonNullable<PresentationObject["placeholder"]>["type"], index };
    }
    const properties = child(source, n === "grpSp" ? "grpSpPr" : "spPr", P);
    if (properties) check(properties, ["xfrm", "prstGeom", ...(["sp", "pic"].includes(n) ? ["custGeom"] : []), "noFill", "solidFill", "gradFill", "ln", "effectLst", "extLst"], A);
    transform(n === "graphicFrame" ? child(source, "xfrm", P) : child(properties, "xfrm", A), target, n === "grpSp");
    objectStyle(context, child(source, "style", P), properties, surface.themeId, target);
    let object: PresentationObject;
    if (n === "grpSp") {
      object = { ...base, kind: "group" };
      for (const childNode of compatible(source)) if (![nvName, "grpSpPr", "extLst"].includes(childNode.name) || childNode.uri !== P) target.children.push(this.object(childNode, surface, false));
    } else if (n === "sp" || n === "cxnSp") {
      const geometry = child(properties, "prstGeom", A), custom = child(properties, "custGeom", A), textBox = bool(child(nv, "cNvSpPr", P) ?? native, "txBox");
      if (geometry && custom) fail(properties!, "A shape cannot use both preset and custom geometry.");
      const preset = geometry ? value(geometry, "prst") : custom || (n === "sp" && (ph || textBox)) ? "rect" : undefined;
      if (!preset) unsupported(source, "A shape geometry is missing.");
      if (n === "cxnSp") {
        if (!["line", "straightConnector1", "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5"].includes(preset)) unsupported(geometry ?? source, "Only straight and elbow connector presets are supported.");
        object = { ...base, kind: "connector", routing: preset === "line" || preset === "straightConnector1" ? "straight" : "elbow" };
        target.attrs["data-office-connector-preset"] = preset;
        const c = child(nv, "cNvCxnSpPr", P), start = child(c, "stCxn", A), end = child(c, "endCxn", A);
        if (start) object.startSite = unsigned(start, "idx");
        if (end) object.endSite = unsigned(end, "idx");
        this.connections.push({ object, surface: surface.id, ...(start ? { start: String(unsigned(start, "id")) } : {}), ...(end ? { end: String(unsigned(end, "id")) } : {}) });
      } else {
        const parsed = PresentationPresetSchema.safeParse(preset);
        if (!parsed.success) unsupported(geometry ?? source, "This preset geometry is not supported.");
        object = !custom && (textBox || !geometry) && preset === "rect" ? { ...base, kind: "text" } : { ...base, kind: "geometry", preset: parsed.data, ...(custom ? { customGeometry: customGeometry(custom) } : {}) };
        const text = child(source, "txBody", P);
        if (text) object.text = importText(context, text, target, source.part);
      }
      const adjustments = child(geometry, "avLst", A);
      check(geometry, ["avLst"], A);
      check(adjustments, ["gd"], A);
      for (const adjustment of elements(adjustments, "gd", A)) {
        const name = value(adjustment, "name"), formula = value(adjustment, "fmla");
        if (!name || !/^[a-z][a-z0-9]*$/i.test(name) || !formula || !/^val -?\d+$/.test(formula) || !Number.isSafeInteger(Number(formula.slice(4)))) unsupported(adjustment, "Only constant preset adjustments are supported.");
        target.attrs["data-office-adjust-" + name.toLowerCase()] = formula.slice(4);
      }
    } else if (n === "pic") {
      const geometry = child(properties, "prstGeom", A), custom = child(properties, "custGeom", A);
      if (geometry && custom) fail(properties!, "A picture cannot use both preset and custom geometry.");
      object = { ...base, kind: "image", ...(custom ? { customGeometry: customGeometry(custom) } : {}) };
      if (geometry && value(geometry, "prst") !== "rect") unsupported(geometry, "Only rectangular picture geometry is supported.");
      const blipFill = child(source, "blipFill", P), blip = child(blipFill, "blip", A);
      if (!blipFill || !blip) fail(source, "A picture source is missing.");
      check(blipFill, ["blip", "srcRect", "stretch"], A);
      if (attr(blip, "link", R)) unsupported(blip, "External picture references are not supported.");
      const src = context.image(source.part, attr(blip, "embed", R));
      check(blip, ["alphaModFix", "extLst"], A);
      const alpha = child(blip, "alphaModFix", A);
      if (alpha) {
        check(alpha, [], A);
        const amount = integer(alpha, "amt", 100000);
        if (amount < 0 || amount > 100000) unsupported(alpha, "Picture alpha modulation above 100% is not supported.");
        target.attrs["data-office-image-opacity"] = String(amount / 100000);
      }
      const crop = child(blipFill, "srcRect", A);
      if (crop) for (const [tag, name] of [["l", "left"], ["t", "top"], ["r", "right"], ["b", "bottom"]]) target.attrs["data-office-crop-" + name] = String(integer(crop, tag!, 0) / 100000);
      const stretch = child(blipFill, "stretch", A), rectangle = child(stretch, "fillRect", A);
      if (!stretch) unsupported(blipFill, "Tiled pictures are not supported.");
      if (rectangle && ["l", "t", "r", "b"].some(name => integer(rectangle, name, 0) !== 0)) unsupported(rectangle, "Offset picture stretch rectangles are not supported.");
      target.children.push(context.make("img", { src, alt: value(native, "descr") ?? "" }));
    } else {
      const graphic = child(source, "graphic", A), data = child(graphic, "graphicData", A);
      if (!data) fail(source, "A graphic frame has no graphic content.");
      check(graphic, ["graphicData"], A);
      if (["http://schemas.openxmlformats.org/drawingml/2006/table", "http://purl.oclc.org/ooxml/drawingml/table"].includes(value(data, "uri") ?? "")) {
        check(data, ["tbl"], A);
        const table = child(data, "tbl", A);
        if (!table) fail(data, "A table graphic is missing its native table.");
        object = { ...base, kind: "table", grid: importTable(context, table, target, source.part) };
      } else if ([C, "http://purl.oclc.org/ooxml/drawingml/chart"].includes(value(data, "uri") ?? "")) {
        check(data, ["chart"], C);
        const chart = child(data, "chart", C);
        if (!chart) fail(data, "A chart graphic is missing its chart reference.");
        object = { ...base, kind: "chart", chartId: importChart(context, chart, target) };
      } else unsupported(data, "This graphic frame type is not supported.");
    }
    target.attrs["data-office-kind"] = object.kind; context.metadata.objects.push(object); this.objects.set(object.nodeId, object);
    return target;
  }
  finish(): void {
    for (const connection of this.connections) {
      const identities = this.context.nativeNodes.get(connection.surface)!;
      if (connection.start !== undefined) connection.object.startNodeId = identities.get(connection.start) ?? fail({ part: connection.surface, name: "stCxn" }, "A connection start target is missing.");
      if (connection.end !== undefined) connection.object.endNodeId = identities.get(connection.end) ?? fail({ part: connection.surface, name: "endCxn" }, "A connection end target is missing.");
    }
  }
}
