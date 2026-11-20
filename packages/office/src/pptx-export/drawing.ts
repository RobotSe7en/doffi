import type { DeepReadonly, PresentationObject } from "@doffi/core";
import type { Css } from "../export/css.js";
import { children } from "../export/source.js";
import type { SourceElement } from "../export/source.js";
import { attributes, empty } from "../export/xml.js";
import { effectXml, fillXml, lineXml, objectAppearance } from "./appearance.js";
import type { NativeSurface, PresentationExportContext } from "./context.js";
import { PresentationTextWriter, hyperlink } from "./text.js";
import { emu, integer, points, unsupported } from "./values.js";
export class PresentationDrawingWriter {
    constructor(readonly ctx: PresentationExportContext, readonly text: PresentationTextWriter, private readonly chart: (object: DeepReadonly<Extract<PresentationObject, {
        kind: "chart";
    }>>, node: SourceElement, surface: NativeSurface) => string) { }
    private transform(node: SourceElement, group = false, tag = "a:xfrm"): string {
        if (node.attrs["data-office-x"] === undefined)
            return "";
        const value = (name: string, fallback = 0): number => Number(node.attrs["data-office-" + name] ?? fallback), attrs: Record<string, number> = {};
        if (value("rotation"))
            attrs.rot = integer(value("rotation"), 60000, -2147483648, 2147483647);
        if (node.attrs["data-office-flip-h"] !== undefined)
            attrs.flipH = node.attrs["data-office-flip-h"] === "true" ? 1 : 0;
        if (node.attrs["data-office-flip-v"] !== undefined)
            attrs.flipV = node.attrs["data-office-flip-v"] === "true" ? 1 : 0;
        return "<" + tag + attributes(attrs) + ">" + empty("a:off", { x: emu(value("x")), y: emu(value("y")) }) + empty("a:ext", { cx: emu(value("width"), 0), cy: emu(value("height"), 0) }) + (group ? empty("a:chOff", { x: emu(value("child-x")), y: emu(value("child-y")) }) + empty("a:chExt", { cx: emu(value("child-width", value("width")), 1), cy: emu(value("child-height", value("height")), 1) }) : "") + "</" + tag + ">";
    }
    private identity(object: DeepReadonly<PresentationObject>, node: SourceElement, surface: NativeSurface, tag: string, properties = ""): string {
        const id = object.nodeId, ph = object.placeholder;
        return "<p:" + tag + "><p:cNvPr" + attributes({ id: this.ctx.nativeIds.get(id)!, name: object.name ?? "Object " + this.ctx.nativeIds.get(id), descr: node.attrs["data-office-alt"], hidden: node.attrs["data-office-hidden"] === "true" ? 1 : undefined }) + ">" + hyperlink(this.ctx, node, surface.part) + '<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"' + attributes({ id: this.ctx.creationIds.get(id)! }) + '/></a:ext></a:extLst></p:cNvPr>' + properties + "<p:nvPr>" + (ph ? empty("p:ph", { type: ph.type, idx: integer(ph.index, 1, 0, 4294967295) }) : "") + "</p:nvPr></p:" + tag + ">";
    }
    private geometry(node: SourceElement, preset: string): string {
        let result = "<a:prstGeom" + attributes({ prst: preset }) + "><a:avLst>";
        for (const [key, value] of Object.entries(node.attrs))
            if (key.startsWith("data-office-adjust-"))
                result += empty("a:gd", { name: key.slice(19), fmla: "val " + integer(Number(value), 1, -2147483648, 2147483647) });
        return result + "</a:avLst></a:prstGeom>";
    }
    private shapeGeometry(node: SourceElement, object: DeepReadonly<Extract<PresentationObject, { kind: "geometry" | "image" }>>, preset: string): string {
        const custom = object.customGeometry;
        if (!custom) return this.geometry(node, preset);
        const guides = custom.guides.map(({ name, operator, args }) => empty("a:gd", { name, fmla: [operator, ...args].join(" ") })).join(""),
            sites = custom.connections.map(({ angle, x, y }) => "<a:cxn" + attributes({ ang: angle }) + ">" + empty("a:pos", { x, y }) + "</a:cxn>").join(""),
            rect = custom.textRect ? empty("a:rect", { l: custom.textRect.left, t: custom.textRect.top, r: custom.textRect.right, b: custom.textRect.bottom }) : "",
            paths = custom.paths.map(path => "<a:path" + attributes({ w: path.width, h: path.height, fill: path.fill ? "norm" : "none", stroke: path.stroke ? 1 : 0 }) + ">" +
                path.commands.map(command => {
                    if (command.kind === "close") return "<a:close/>";
                    const tag = command.kind === "move" ? "a:moveTo" : "a:lnTo";
                    return "<" + tag + ">" + empty("a:pt", { x: command.x, y: command.y }) + "</" + tag + ">";
                }).join("") + "</a:path>").join("");
        return "<a:custGeom><a:gdLst>" + guides + "</a:gdLst><a:cxnLst>" + sites + "</a:cxnLst>" + rect + "<a:pathLst>" + paths + "</a:pathLst></a:custGeom>";
    }
    object(node: SourceElement, surface: NativeSurface): string {
        const result = this.compileObject(node, surface);
        if (this.ctx.objects.get(node.attrs["data-node-id"]!)?.kind !== "group")
            this.ctx.checkGenerated(result);
        return result;
    }
    private compileObject(node: SourceElement, surface: NativeSurface): string {
        const object = this.ctx.objects.get(node.attrs["data-node-id"]!);
        if (!object)
            unsupported("A presentation surface has an unregistered drawing.");
        const transform = this.transform(node, object.kind === "group");
        if (object.kind === "group" && this.ctx.styles.get(node, surface)["--office-shadow-kind"] === "outer" && this.ctx.styles.get(node, surface)["--office-shadow-source"] === "reference")
            unsupported("A group theme effect reference cannot be represented by this native group profile.");
        if (object.kind === "group") {
            const head = "<p:grpSp>" + this.identity(object, node, surface, "nvGrpSpPr", "<p:cNvGrpSpPr/>") + "<p:grpSpPr>" + transform + effectXml(this.ctx.styles.get(node, surface), this.ctx.styles.raw(node, surface)) + "</p:grpSpPr>";
            this.ctx.checkGenerated(head + "</p:grpSp>");
            return head + children(node).map(child => this.object(child, surface)).join("") + "</p:grpSp>";
        }
        if (object.kind === "text" || object.kind === "geometry") {
            const preset = object.kind === "text" ? "rect" : object.preset, appearance = objectAppearance(this.ctx, node, surface, object.kind, preset);
            return "<p:sp>" + this.identity(object, node, surface, "nvSpPr", empty("p:cNvSpPr", { txBox: object.kind === "text" ? 1 : undefined })) + "<p:spPr>" + transform + (object.kind === "geometry" ? this.shapeGeometry(node, object, preset) : this.geometry(node, preset)) + appearance.properties + "</p:spPr>" + appearance.style + this.text.body(node, surface, "p:txBody", surface.part, object) + "</p:sp>";
        }
        if (object.kind === "connector") {
            const appearance = objectAppearance(this.ctx, node, surface, "connector"), connection = (side: "start" | "end"): string => { const target = object[side + "NodeId" as "startNodeId" | "endNodeId"]; return target ? empty(side === "start" ? "a:stCxn" : "a:endCxn", { id: this.ctx.nativeIds.get(target)!, idx: object[side + "Site" as "startSite" | "endSite"] ?? 0 }) : ""; };
            return "<p:cxnSp>" + this.identity(object, node, surface, "nvCxnSpPr", "<p:cNvCxnSpPr>" + connection("start") + connection("end") + "</p:cNvCxnSpPr>") + "<p:spPr>" + transform + this.geometry(node, node.attrs["data-office-connector-preset"] ?? (object.routing === "straight" ? "straightConnector1" : "bentConnector3")) + appearance.properties + "</p:spPr>" + appearance.style + "</p:cxnSp>";
        }
        if (object.kind === "image") {
            const image = children(node, "img")[0];
            if (!image)
                unsupported("An exported picture has no source image.");
            const relation = this.ctx.image(image, surface.part), crop: Record<string, number> = {};
            for (const [side, native] of [["left", "l"], ["top", "t"], ["right", "r"], ["bottom", "b"]])
                if (node.attrs["data-office-crop-" + side] !== undefined)
                    crop[native!] = integer(Number(node.attrs["data-office-crop-" + side]), 100000, -2147483648, 2147483647);
            const appearance = objectAppearance(this.ctx, node, surface, "text");
            const opacity = node.attrs["data-office-image-opacity"], blip = opacity === undefined ? empty("a:blip", { "r:embed": relation }) :
                "<a:blip" + attributes({ "r:embed": relation }) + ">" + empty("a:alphaModFix", { amt: integer(Number(opacity), 100000, 0, 100000) }) + "</a:blip>";
            return "<p:pic>" + this.identity(object, node, surface, "nvPicPr", "<p:cNvPicPr/>") + "<p:blipFill>" + blip + (Object.keys(crop).length ? empty("a:srcRect", crop) : "") + "<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>" + transform + this.shapeGeometry(node, object, "rect") + appearance.properties + "</p:spPr>" + appearance.style + "</p:pic>";
        }
        const graphic = object.kind === "table" ? this.table(object, node, surface) : this.chart(object, node, surface), uri = object.kind === "table" ? "http://schemas.openxmlformats.org/drawingml/2006/table" : "http://schemas.openxmlformats.org/drawingml/2006/chart";
        return "<p:graphicFrame>" + this.identity(object, node, surface, "nvGraphicFramePr", "<p:cNvGraphicFramePr/>") + this.transform(node, false, "p:xfrm") + "<a:graphic><a:graphicData" + attributes({ uri }) + ">" + graphic + "</a:graphicData></a:graphic></p:graphicFrame>";
    }
    tree(surface: NativeSurface): string {
        let result = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
        this.ctx.checkGenerated(result + "</p:spTree>");
        const objects = children(surface.node).filter(node => this.ctx.objects.has(node.attrs["data-node-id"]!));
        result += objects.map(node => this.object(node, surface)).join("");
        if (surface.kind === "notes" && !objects.length) {
            const id = this.ctx.extraId(surface);
            const notes = '<p:sp><p:nvSpPr><p:cNvPr' + attributes({ id, name: "Notes" }) + '/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext' + attributes({ cx: emu((this.ctx.p.notesSize ?? this.ctx.p.size).width, 1), cy: emu((this.ctx.p.notesSize ?? this.ctx.p.size).height, 1) }) + '/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' + this.text.body(surface.node, surface) + "</p:sp>";
            this.ctx.checkGenerated(notes);
            result += notes;
        }
        return result + "</p:spTree>";
    }
    private table(object: DeepReadonly<Extract<PresentationObject, {
        kind: "table";
    }>>, node: SourceElement, surface: NativeSurface): string {
        const table = children(node, "table")[0]!, rows = children(table).flatMap(child => child.tag === "tr" ? [child] : children(child, "tr")), grid = object.grid;
        if (rows.length * grid.length > this.ctx.limits.maxNodes)
            unsupported("A native table exceeds its output cell budget.");
        const occupied: Array<Array<{
            cell: SourceElement;
            row: number;
            column: number;
            height: number;
        } | undefined>> = Array.from({ length: rows.length }, () => Array(grid.length));
        for (const [rowIndex, row] of rows.entries()) {
            let column = 0;
            for (const cell of children(row)) {
                while (occupied[rowIndex]![column])
                    column++;
                const width = Number(cell.attrs.colspan ?? 1), height = Number(cell.attrs.rowspan ?? 1) || rows.length - rowIndex, origin = { cell, row: rowIndex, column, height };
                for (let y = rowIndex; y < rowIndex + height; y++)
                    for (let x = column; x < column + width; x++) {
                        if (!occupied[y] || x >= grid.length || occupied[y]![x])
                            unsupported("A native table merge is invalid.");
                        occupied[y]![x] = origin;
                    }
                column += width;
            }
        }
        const css = this.ctx.styles.get(table, surface), raw = this.ctx.styles.raw(table, surface);
        let result = "<a:tbl><a:tblPr>" + fillXml(css, "--office-table-fill", raw) + "</a:tblPr><a:tblGrid>" + grid.map(width => empty("a:gridCol", { w: emu(width, 1) })).join("") + "</a:tblGrid>";
        for (const [rowIndex, row] of rows.entries()) {
            const rowCss = this.ctx.styles.get(row, surface);
            result += "<a:tr" + attributes({ h: emu(points(rowCss.height, rowCss, Number(node.attrs["data-office-height"]) / rows.length), 1) }) + ">";
            for (let column = 0; column < grid.length; column++) {
                const origin = occupied[rowIndex]![column];
                if (!origin)
                    unsupported("A native table grid contains an uncovered cell.");
                const cell = origin.cell, continuation = origin.row !== rowIndex || origin.column !== column, span = { rowSpan: !continuation && origin.height > 1 ? origin.height : undefined, gridSpan: !continuation && Number(cell.attrs.colspan ?? 1) > 1 ? Number(cell.attrs.colspan) : undefined, hMerge: origin.column < column ? 1 : undefined, vMerge: origin.row < rowIndex ? 1 : undefined };
                result += "<a:tc" + attributes(span) + ">";
                if (continuation)
                    result += "<a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/>";
                else {
                    const c = this.ctx.styles.get(cell, surface), own = this.ctx.styles.own(cell, surface), r = this.ctx.styles.raw(cell, surface), body = children(cell).length === 1 && children(cell)[0]!.tag === "div" ? children(cell)[0]! : cell;
                    result += this.text.body(body, surface, "a:txBody", surface.part, undefined, body === cell);
                    const attrs: Record<string, string | number> = { anchor: ({ top: "t", middle: "ctr", bottom: "b" }[c["vertical-align"] ?? "top"] ?? "t") };
                    for (const [side, key] of [["left", "marL"], ["right", "marR"], ["top", "marT"], ["bottom", "marB"]])
                        attrs[key!] = emu(points(c["padding-" + side], c, 0), 0, 2147483647);
                    result += "<a:tcPr" + attributes(attrs) + ">";
                    for (const [side, tag] of [["left", "lnL"], ["right", "lnR"], ["top", "lnT"], ["bottom", "lnB"]]) {
                        const prefix = "--office-cell-border-" + side, style = c["border-" + side + "-style"];
                        let paint: Css = c;
                        if (own["border-" + side + "-style"] !== undefined && !Object.keys(own).some(key => key.startsWith(prefix)))
                            paint = { ...c, [prefix + "-kind"]: style === "none" || style === "hidden" ? "none" : "solid", [prefix + "-width"]: c["border-" + side + "-width"] ?? "0pt", [prefix + "-color"]: c["border-" + side + "-color"] ?? c.color!, [prefix + "-dash"]: style === "dotted" ? "dot" : style === "dashed" ? "dash" : "solid" };
                        result += lineXml(paint, prefix, r, false, "none", "a:" + tag);
                    }
                    const paint = c["--office-cell-fill-kind"] ? c : c["background-color"] !== "transparent" ? { ...c, "--office-cell-fill-kind": "solid", "--office-cell-fill-color": c["background-color"]! } : c;
                    result += fillXml(paint, "--office-cell-fill", r) + "</a:tcPr>";
                }
                result += "</a:tc>";
            }
            result += "</a:tr>";
        }
        return result + "</a:tbl>";
    }
}
