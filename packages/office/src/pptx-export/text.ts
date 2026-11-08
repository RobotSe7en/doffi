import type { DeepReadonly, PresentationObject } from "@doffi/core";
import type { Css } from "../export/css.js";
import { fontName } from "../export/values.js";
import { children, element, sourceText } from "../export/source.js";
import type { SourceElement, SourceNode } from "../export/source.js";
import { attributes, empty, escapeXml } from "../export/xml.js";
import { colorXml } from "./appearance.js";
import type { NativeSurface, PresentationExportContext } from "./context.js";
import { relativePart } from "./context.js";
import { cssString, emu, integer, points, unsupported } from "./values.js";
const blocks = new Set(["p", "h1", "h2", "h3", "h4", "h5", "h6", "li"]);
const present = (own: Css, ...keys: string[]): boolean => keys.some(key => own[key] !== undefined);
function assertTextLayout(node: SourceElement, css: Css, body = false, cell = false): void {
    for (const key of ["min-width", "max-width", "min-height", "max-height"])
        if (css[key] !== undefined && !["none", "auto", "initial", ...(key.startsWith("min-") ? ["0", "0pt", "0px"] : [])].includes(css[key]!))
            unsupported("Minimum and maximum text box dimensions are outside the native PPTX text profile.");
    if (!body)
        for (const key of ["width", "height"])
            if (css[key] !== undefined && !["auto", "initial"].includes(css[key]!))
                unsupported("Nested text box dimensions are outside the native PPTX text profile.");
    for (const side of ["top", "right", "bottom", "left"]) {
        const border = css["border-" + side + "-style"], width = css["border-" + side + "-width"];
        if (!cell && border && !["none", "hidden"].includes(border) && !["0", "0pt", "0px"].includes(width ?? "medium"))
            unsupported("Text box borders are outside the native PPTX text profile.");
        if (!body && css["padding-" + side] !== undefined && points(css["padding-" + side], css) !== 0)
            unsupported("Nested text padding is outside the native PPTX text profile.");
        if (!body && !blocks.has(node.tag) && css["margin-" + side] !== undefined && points(css["margin-" + side], css) !== 0)
            unsupported("Inline text margins are outside the native PPTX text profile.");
    }
}
export interface NativeTextBody {
    readonly node: SourceElement;
    readonly paragraphs: readonly {
        node: SourceElement;
        start: number;
        text: string;
    }[];
    readonly text: string;
}
export function hyperlink(ctx: PresentationExportContext, node: SourceElement, part: string): string {
    const link = ctx.p.hyperlinks.find(link => link.nodeId === node.attrs["data-node-id"]), href = node.tag === "a" ? node.attrs.href : undefined;
    if (!link && !href)
        return "";
    if (link?.slideId) {
        const target = ctx.surfaces.find(surface => surface.kind === "slide" && surface.definitionId === link.slideId)!;
        return empty("a:hlinkClick", { "r:id": ctx.opc.relate(part, "slide", relativePart(part, target.part)), action: "ppaction://hlinksldjump" });
    }
    const target = link?.target ?? href!;
    if (!/^(https?:\/\/|mailto:)/i.test(target) || /[\u0000-\u0020\u007f]/.test(target))
        unsupported("This hyperlink target has no supported native PPTX representation.");
    return empty("a:hlinkClick", { "r:id": ctx.opc.relate(part, "hyperlink", target, true) });
}
function font(css: Css, raw: Css, role: string): string | undefined {
    const key = "--office-font-" + role, value = raw[key];
    if (value === undefined)
        return;
    const theme = /^var\(\s*--office-(major|minor)-(ascii|h-ansi|east-asia|complex-script)\s*\)$/.exec(value);
    if (theme)
        return "+" + (theme[1] === "major" ? "mj" : "mn") + "-" + ({ ascii: "lt", "h-ansi": "lt", "east-asia": "ea", "complex-script": "cs" }[theme[2]!]!);
    return fontName(css[key] ?? value);
}
/** Emit only authored overrides. Native placeholders continue to inherit from
 * their editable layout/master rather than receiving frozen computed styles. */
export function runProperties(css: Css, own: Css, raw: Css = own, extra: {
    lang?: string | undefined;
    hyperlink?: string | undefined;
} = {}): string {
    const attrs: Record<string, string | number | undefined> = { lang: extra.lang };
    let contents = "";
    if (css["word-spacing"] && !['normal', '0', '0pt', '0px'].includes(css["word-spacing"]!))
        unsupported("Independent word spacing has no supported native PPTX representation.");
    if (own["text-transform"] !== undefined && !["none", "uppercase"].includes(css["text-transform"]!))
        unsupported("This text transformation has no supported native PPTX representation.");
    if (own["font-style"] !== undefined && !["normal", "italic"].includes(css["font-style"]!))
        unsupported("This font slant has no supported native PPTX representation.");
    if (own["font-variant"] !== undefined && !["normal", "small-caps"].includes(css["font-variant"]!))
        unsupported("This font variant has no supported native PPTX representation.");
    if (present(own, "font-size"))
        attrs.sz = integer(points(css["font-size"], css, 18), 100, 100, 400000);
    if (present(own, "font-weight")) {
        if (!["normal", "bold", "400", "700"].includes(css["font-weight"]!))
            unsupported("PPTX supports regular and bold font weights in this profile.");
        attrs.b = ["bold", "700"].includes(css["font-weight"]!) ? 1 : 0;
    }
    if (present(own, "font-style"))
        attrs.i = css["font-style"] === "normal" ? 0 : 1;
    if (present(own, "letter-spacing"))
        attrs.spc = integer(css["letter-spacing"] === "normal" ? 0 : points(css["letter-spacing"], css), 100, -400000, 400000);
    if (present(own, "vertical-align")) {
        const v = css["vertical-align"]!;
        attrs.baseline = v === "super" ? 30000 : v === "sub" ? -25000 : v === "baseline" ? 0 : v.endsWith("%") ? integer(parseFloat(v), 1000, -100000, 100000) : unsupported("This text baseline has no native PPTX representation.");
    }
    if (present(own, "text-transform", "font-variant"))
        attrs.cap = css["text-transform"] === "uppercase" ? "all" : css["font-variant"] === "small-caps" ? "small" : "none";
    if (present(own, "text-decoration-line", "text-decoration-style", "--office-underline", "--office-strike", "--office-font-decoration-reset")) {
        const line = css["text-decoration-line"] ?? "none", style = css["text-decoration-style"] ?? "solid";
        if (line.includes("underline") && own["text-decoration-color"] !== undefined && css["text-decoration-color"]?.toLowerCase() !== "currentcolor")
            unsupported("Independent underline color is outside the native PPTX text profile.");
        if (line.includes("overline"))
            unsupported("Overline text is outside the native PPTX profile.");
        const native = css["--office-underline"], nativeStyle = native && ({ sng: "solid", heavy: "solid", dbl: "double", dotted: "dotted", dash: "dashed", wavy: "wavy" }[native]);
        attrs.u = line.includes("underline") ? nativeStyle === style ? native : ({ solid: "sng", double: "dbl", dotted: "dotted", dashed: "dash", wavy: "wavy" }[style] ?? unsupported("This underline style is outside the PPTX profile.")) : "none";
        if (line.includes("line-through") && !["solid", "double"].includes(style))
            unsupported("This strike style is outside the native PPTX text profile.");
        attrs.strike = line.includes("line-through") ? style === "double" ? "dblStrike" : "sngStrike" : "noStrike";
    }
    if (present(own, "color", "--office-text-color")) {
        const hasNative = css["--office-text-color"] !== undefined, paint = hasNative ? css : { ...css, "--office-text-color": css.color ?? "#000000" }, rawPaint = hasNative ? raw : { ...raw, "--office-text-color": raw.color ?? css.color ?? "#000000" };
        contents += "<a:solidFill>" + colorXml(paint, "--office-text", rawPaint) + "</a:solidFill>";
    }
    const roles = [["ascii", "latin"], ["east-asia", "ea"], ["complex-script", "cs"]] as const;
    for (const [role, tag] of roles) {
        let name = font(css, own, role);
        if (role === "ascii" && own["font-family"] !== undefined && own["--office-font-ascii"] === undefined)
            name = fontName(css["font-family"]!);
        if (name !== undefined)
            contents += empty("a:" + tag, { typeface: name });
    }
    if (own["--office-font-h-ansi"] !== undefined && font(css, own, "h-ansi") !== (font(css, own, "ascii") ?? fontName(css["font-family"]!)))
        unsupported("DrawingML has one Latin font slot; differing ASCII and HANSI faces cannot be exported losslessly.");
    contents += extra.hyperlink ?? "";
    return attributes(attrs) + ">" + contents;
}
function spacing(css: Css, name: string, tag: string): string {
    const value = css[name];
    if (value === undefined || value === "normal")
        return "";
    if (name !== "line-height" && value.includes("%"))
        unsupported("Percentage paragraph margins have no supported native PPTX representation.");
    const proportional = /^[-+.\d]+$/.test(value) || value.endsWith("em") || value.endsWith("%");
    return "<a:" + tag + ">" + empty(proportional ? "a:spcPct" : "a:spcPts", { val: proportional ? integer(parseFloat(value) * (value.endsWith("%") ? .01 : 1), 100000, 0, 2147483647) : integer(points(value, css), 100, 0, 2147483647) }) + "</a:" + tag + ">";
}
export function paragraphProperties(css: Css, own: Css, raw: Css = own, node?: SourceElement, level?: number): string {
    const attrs: Record<string, string | number | undefined> = { lvl: level };
    let contents = "";
    if (present(own, "text-align"))
        attrs.algn = ({ left: "l", start: "l", center: "ctr", right: "r", end: "r", justify: "just" }[css["text-align"]!] ?? unsupported("This paragraph alignment is outside the PPTX profile."));
    for (const [name, key] of [["marL", "margin-left"], ["marR", "margin-right"], ["indent", "text-indent"]])
        if (own[key!] !== undefined)
            attrs[name!] = emu(points(css[key!], css), -2147483648, 2147483647);
    if (own["tab-size"] !== undefined && css["tab-size"] && !/^\d+(?:\.\d+)?$/.test(css["tab-size"]!))
        attrs.defTabSz = emu(points(css["tab-size"], css), 1, 2147483647);
    for (const [key, tag] of [["line-height", "lnSpc"], ["margin-top", "spcBef"], ["margin-bottom", "spcAft"]])
        if (own[key!] !== undefined)
            contents += spacing(css, key!, tag!);
    if (present(own, "--office-bullet-color"))
        contents += "<a:buClr>" + colorXml(css, "--office-bullet", raw) + "</a:buClr>";
    if (css["--office-bullet-size"] && own["--office-bullet-size"] !== undefined) {
        const size = css["--office-bullet-size"]!;
        if (size.endsWith("em") && Math.abs(parseFloat(size) * 100 - Math.round(parseFloat(size) * 100)) > 1e-8)
            unsupported("A fractional bullet percentage has no exact native representation in this profile.");
        contents += empty(size.endsWith("em") ? "a:buSzPct" : "a:buSzPts", { val: size.endsWith("em") ? String(integer(parseFloat(size), 100, 25, 400)) + "%" : integer(points(size, css), 100, 100, 400000) });
    }
    if (own["--office-bullet-font"] !== undefined)
        contents += empty("a:buFont", { typeface: fontName(css["--office-bullet-font"]!) });
    const kind = node?.attrs["data-office-list-none"] === "true" ? "none" : node?.attrs["data-office-bullet"] !== undefined ? "bullet" : node?.attrs["data-office-numbering"] !== undefined ? "number" : own["--office-list-kind"] !== undefined ? css["--office-list-kind"] : node?.tag === "li" ? (node.parent?.tag === "ol" ? "number" : "bullet") : undefined;
    if (kind === "none")
        contents += "<a:buNone/>";
    else if (kind === "bullet")
        contents += empty("a:buChar", { char: node?.attrs["data-office-bullet"] ?? cssString(css["--office-list-character"] ?? '"•"') });
    else if (kind === "number")
        contents += empty("a:buAutoNum", { type: node?.attrs["data-office-numbering"] ?? css["--office-list-numbering"] ?? ({ "lower-alpha": "alphaLcPeriod", "upper-alpha": "alphaUcPeriod", "lower-roman": "romanLcPeriod", "upper-roman": "romanUcPeriod" }[css["list-style-type"] ?? ""] ?? "arabicPeriod"), startAt: integer(Number(node?.attrs["data-office-numbering-start"] ?? css["--office-list-start"] ?? node?.parent?.attrs.start ?? 1), 1, 1, 32767) });
    if (own["--office-tab-origin"] !== undefined && css["--office-tab-origin"] === "left")
        contents += '<a:tabLst><a:tab pos="0" algn="l"/></a:tabLst>';
    const run = runProperties(css, own, raw);
    contents += "<a:defRPr" + run + "</a:defRPr>";
    return attributes(attrs) + ">" + contents;
}
export class PresentationTextWriter {
    readonly bodies: Map<string, NativeTextBody> = new Map<string, NativeTextBody>();
    constructor(private readonly ctx: PresentationExportContext) { }
    body(node: SourceElement, surface: NativeSurface, tag: string = "p:txBody", part: string = surface.part, object?: DeepReadonly<Extract<PresentationObject, {
        kind: "text" | "geometry";
    }>>, cellRoot: boolean = false): string {
        const css = this.ctx.styles.get(node, surface), own = this.ctx.styles.own(node, surface), raw = this.ctx.styles.raw(node, surface), attrs: Record<string, string | number | undefined> = { wrap: object?.text?.wrap === false || css["white-space"] === "nowrap" ? "none" : "square" };
        assertTextLayout(node, css, true, cellRoot);
        for (const [side, key] of [["left", "lIns"], ["right", "rIns"], ["top", "tIns"], ["bottom", "bIns"]])
            if (cellRoot)
                attrs[key!] = 0;
            else if (css["padding-" + side] !== undefined)
                attrs[key!] = emu(points(css["padding-" + side], css), 0, 2147483647);
        if (own["--office-text-anchor"] !== undefined)
            attrs.anchor = css["--office-text-anchor"];
        if (own["--office-text-anchor-center"] !== undefined)
            attrs.anchorCtr = css["--office-text-anchor-center"] === "true" ? 1 : 0;
        let result = "<" + tag + "><a:bodyPr" + attributes(attrs) + ">", fit = object?.text?.autoFit;
        result += fit === "normal" ? empty("a:normAutofit", { fontScale: integer(object?.text?.fontScale ?? 1, 100000, 1, 100000), lnSpcReduction: integer(object?.text?.lineSpacingReduction ?? 0, 100000, 0, 100000) }) : fit === "shape" ? "<a:spAutoFit/>" : "<a:noAutofit/>";
        result += "</a:bodyPr><a:lstStyle><a:defPPr" + paragraphProperties(css, ["td", "th"].includes(node.tag) ? Object.fromEntries(Object.entries(own).filter(([key]) => key !== "vertical-align")) : own, raw) + "</a:defPPr></a:lstStyle>";
        const paragraphs: {
            node: SourceElement;
            start: number;
            text: string;
        }[] = [], strings: string[] = [];
        let offset = 0;
        const write = (p: SourceElement, nodes: readonly SourceNode[] = p.children, level = Number(p.attrs["data-office-level"] ?? 0)): void => {
            const computed = this.ctx.styles.get(p, surface), authored = { ...this.ctx.styles.own(p, surface) }, refs = this.ctx.styles.raw(p, surface);
            assertTextLayout(p, computed, p === node, cellRoot && p === node);
            if (cellRoot && p === node)
                delete authored["vertical-align"]; // Cell alignment belongs to tcPr, never a run baseline.
            if (/^h[1-6]$/.test(p.tag) && authored["font-size"] === undefined)
                authored["font-size"] = computed["font-size"]!;
            if (blocks.has(p.tag))
                for (const side of ["top", "right", "bottom", "left"])
                    if (authored["padding-" + side] !== undefined && points(computed["padding-" + side], computed) !== 0)
                        unsupported("Paragraph padding is outside the native PPTX text profile.");
            if (blocks.has(p.tag) && authored["background-color"] !== undefined && computed["background-color"] !== "transparent")
                unsupported("Paragraph shading has no supported native PPTX representation.");
            result += "<a:p><a:pPr" + paragraphProperties(computed, authored, refs, p, level) + "</a:pPr>";
            let plain = "";
            const visit = (child: SourceNode, overrides: Css = {}, link = ""): void => {
                if (!element(child)) {
                    if (!child.text)
                        return;
                    const parent = child.parent, style = this.ctx.styles.get(parent, surface), rawStyle = this.ctx.styles.raw(parent, surface);
                    result += "<a:r><a:rPr" + runProperties(style, overrides, rawStyle, { lang: parent.attrs.lang, hyperlink: link }) + '</a:rPr><a:t>' + escapeXml(child.text) + "</a:t></a:r>";
                    plain += child.text;
                    return;
                }
                if (child.tag === "br") {
                    result += "<a:br/>";
                    plain += "\v";
                    return;
                }
                if (child.tag === "img" || blocks.has(child.tag) || ["table", "ul", "ol"].includes(child.tag))
                    unsupported("Nested block content cannot be placed in a native DrawingML run.");
                const style = this.ctx.styles.get(child, surface), local = { ...overrides, ...this.ctx.styles.own(child, surface) }, semantic: Record<string, string> = {};
                assertTextLayout(child, style);
                if (["strong", "b"].includes(child.tag))
                    semantic["font-weight"] = style["font-weight"]!;
                if (["em", "i"].includes(child.tag))
                    semantic["font-style"] = style["font-style"]!;
                if (["u", "s", "del", "ins"].includes(child.tag))
                    semantic["text-decoration-line"] = style["text-decoration-line"]!;
                if (["sub", "sup"].includes(child.tag))
                    semantic["vertical-align"] = style["vertical-align"]!;
                if (local["background-color"] !== undefined && style["background-color"] !== "transparent")
                    unsupported("Text shading has no supported native PPTX representation.");
                Object.assign(local, semantic);
                const nextLink = hyperlink(this.ctx, child, part) || link;
                const field = child.attrs["data-office-field"];
                if (field) {
                    const text = sourceText(child);
                    result += "<a:fld" + attributes({ id: this.ctx.guid(child.attrs["data-office-field-id"]), type: field }) + "><a:rPr" + runProperties(style, local, this.ctx.styles.raw(child, surface), { lang: child.attrs.lang, hyperlink: nextLink }) + '</a:rPr><a:t>' + escapeXml(text) + "</a:t></a:fld>";
                    plain += text;
                    return;
                }
                for (const nested of child.children)
                    visit(nested, local, nextLink);
            };
            nodes.forEach(child => visit(child));
            const end = p.attrs["data-office-end-style"], endOwn = end ? this.ctx.styles.css.specifiedClass(end) : {}, endCss = end ? this.ctx.styles.css.classStyle(end, computed) : computed;
            result += "<a:endParaRPr" + runProperties(endCss, endOwn, endOwn) + "</a:endParaRPr></a:p>";
            paragraphs.push({ node: p, start: offset, text: plain });
            strings.push(plain);
            offset += plain.length + 1;
        };
        const walk = (parent: SourceElement, level = 0): void => {
            if (parent !== node)
                assertTextLayout(parent, this.ctx.styles.get(parent, surface));
            let loose: SourceNode[] = [];
            const flush = (): void => {
                if (loose.some(n => element(n) || n.text.trim().length)) {
                    write(parent, loose, level);
                    loose = [];
                }
            };
            for (const child of parent.children) {
                if (element(child) && blocks.has(child.tag)) {
                    flush();
                    if (child.tag === "li") {
                        const inline = child.children.filter(n => !element(n) || !["ul", "ol"].includes(n.tag));
                        write(child, inline, level);
                        for (const list of children(child).filter(n => ["ul", "ol"].includes(n.tag)))
                            walk(list, Math.min(8, level + 1));
                    }
                    else
                        write(child);
                }
                else if (element(child) && ["div", "section", "article", "ul", "ol"].includes(child.tag)) {
                    flush();
                    walk(child, level);
                }
                else
                    loose.push(child);
            }
            flush();
        };
        walk(node);
        if (!paragraphs.length)
            write(node, []);
        this.bodies.set(node.attrs["data-node-id"]!, { node, paragraphs, text: strings.join("\r") });
        return result + "</" + tag + ">";
    }
    plain(value: string, tag = "p188:txBody"): string { return "<" + tag + "><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t >" + escapeXml(value) + "</a:t></a:r></a:p></" + tag + ">"; }
}
