import type { DeepReadonly, DocumentReadLimits, PresentationMetadata, PresentationObject } from "@doffi/core";
import { ExportCss } from "../export/css.js";
import type { Css } from "../export/css.js";
import type { SourceElement, SourceTree } from "../export/source.js";
import { children } from "../export/source.js";
type Master = DeepReadonly<PresentationMetadata["masters"][number]>;
type Layout = DeepReadonly<PresentationMetadata["layouts"][number]>;
type NotesMaster = DeepReadonly<PresentationMetadata["notesMasters"][number]>;
export interface PresentationStyleSurface {
    readonly rootNodeId: string;
    readonly themeId: string;
    readonly master?: Master;
    readonly layout?: Layout;
    readonly notesMaster?: NotesMaster;
}
const inheritedNames = new Set(["font-family", "font-size", "font-weight", "font-style", "color", "line-height", "letter-spacing", "text-align", "text-indent", "margin-left", "margin-right", "margin-top", "margin-bottom", "padding-top", "padding-right", "padding-bottom", "padding-left", "tab-size", "text-transform", "font-variant"]);
const objectDefaults: Css = { "--office-fill-kind": "initial", "--office-fill-color": "initial", "--office-line-kind": "initial", "--office-line-color": "initial", "--office-shadow-kind": "initial", "--office-text-anchor": "initial" };
const defaults = 'p,h1,h2,h3,h4,h5,h6{margin:0;line-height:1.2;font-weight:inherit}table{border-spacing:0}td,th{font-weight:inherit;text-align:inherit;vertical-align:top}';
const compatible = (left: NonNullable<PresentationObject["placeholder"]>, right: NonNullable<PresentationObject["placeholder"]>): boolean => left.index === right.index && (left.type === right.type || [left.type, right.type].every(type => type === "title" || type === "ctrTitle"));
function paragraphs(root: SourceElement): SourceElement[] {
    const result: SourceElement[] = [];
    const visit = (node: SourceElement): void => {
        if (node.tag === "p")
            result.push(node);
        for (const child of children(node))
            visit(child);
    };
    visit(root);
    return result;
}
/** Portable counterpart of the preview's derived native inheritance. Only
 * private source projections are themed; the document package is never changed. */
export class PresentationExportStyles {
    readonly css: ExportCss;
    private readonly objects: ReadonlyMap<string, DeepReadonly<PresentationObject>>;
    private readonly scopes = new Map<string, WeakMap<SourceElement, SourceElement>>();
    private readonly rawCache = new WeakMap<SourceElement, Css>();
    private readonly rawFallbacks = new WeakMap<SourceElement, Css>();
    private readonly projected = new Map<string, Set<string>>();
    private readonly themeClasses: ReadonlySet<string>;
    constructor(private readonly p: DeepReadonly<PresentationMetadata>, private readonly tree: SourceTree, sourceCss: readonly string[], limits: DocumentReadLimits) {
        this.css = new ExportCss([defaults, ...sourceCss], limits);
        this.objects = new Map(p.objects.map(object => [object.nodeId, object]));
        this.themeClasses = new Set(p.themes.map(theme => theme.cssClass));
    }
    private virtual(node: SourceElement, surface: PresentationStyleSurface): SourceElement {
        let scope = this.scopes.get(surface.rootNodeId);
        if (!scope) {
            scope = new WeakMap();
            this.scopes.set(surface.rootNodeId, scope);
        }
        const known = scope.get(node);
        if (known)
            return known;
        const attrs = { ...node.attrs };
        if (attrs["data-node-id"] === surface.rootNodeId) {
            const theme = this.p.themes.find(theme => theme.themeId === surface.themeId)!;
            attrs.class = [...(attrs.class ?? "").split(/\s+/).filter(name => name && !this.themeClasses.has(name)), theme.cssClass].join(" ");
        }
        const result: SourceElement = { ...node, attrs, parent: node.parent ? this.virtual(node.parent, surface) : undefined, children: [] };
        scope.set(node, result);
        return result;
    }
    private values(css: Css): Css { return Object.fromEntries(Object.entries(css).filter(([name]) => inheritedNames.has(name) || name.startsWith("--office-"))); }
    private probe(tag: string, attrs: Readonly<Record<string, string>>, parent: SourceElement): SourceElement {
        return { tag, attrs, parent, children: [], from: 0, to: 0 };
    }
    private prototype(node: SourceElement, parent: SourceElement, values: Css, raw: Css): Css {
        const attrs: Record<string, string> = { class: node.attrs.class ?? "", style: node.attrs.style ?? "" };
        if (node.attrs["data-office-kind"])
            attrs["data-office-kind"] = node.attrs["data-office-kind"];
        const probe = this.probe(node.tag, attrs, parent);
        const result = this.values(this.css.withFallback(probe, values, attrs["data-office-kind"] ? objectDefaults : {}));
        Object.assign(raw, this.css.specified(probe));
        return result;
    }
    private placeholder(root: string, object: DeepReadonly<PresentationObject>): SourceElement | undefined {
        if (!object.placeholder)
            return;
        const objects = children(this.tree.byId.get(root)!).map(node => ({ node, definition: this.objects.get(node.attrs["data-node-id"]!) })).filter(entry => entry.definition?.placeholder);
        return objects.find(entry => entry.definition!.placeholder!.index === object.placeholder!.index && entry.definition!.placeholder!.type === object.placeholder!.type)?.node
            ?? objects.find(entry => compatible(entry.definition!.placeholder!, object.placeholder!))?.node;
    }
    private prototypes(object: DeepReadonly<PresentationObject>, surface: PresentationStyleSurface): SourceElement[] {
        const roots = [surface.master?.rootNodeId, surface.layout?.rootNodeId, surface.notesMaster?.rootNodeId];
        return roots.flatMap(root => root && root !== surface.rootNodeId ? this.placeholder(root, object) ?? [] : []);
    }
    private textClass(object: DeepReadonly<PresentationObject>, surface: PresentationStyleSurface): string | undefined {
        if (surface.notesMaster)
            return surface.notesMaster.textStyle;
        const role = object.placeholder?.type, master = surface.master;
        return role === "title" || role === "ctrTitle" ? master?.titleStyle : role === "body" ? master?.bodyStyle : master?.otherStyle;
    }
    private prepare(object: DeepReadonly<PresentationObject>, surface: PresentationStyleSurface): void {
        let projected = this.projected.get(surface.rootNodeId);
        if (!projected) {
            projected = new Set();
            this.projected.set(surface.rootNodeId, projected);
        }
        if (projected.has(object.nodeId))
            return;
        projected.add(object.nodeId);
        const source = this.tree.byId.get(object.nodeId)!, node = this.virtual(source, surface), parentObject = source.parent && this.objects.get(source.parent.attrs["data-node-id"]!);
        if (parentObject)
            this.prepare(parentObject, surface);
        let values: Css = {}, raw: Css = {};
        const prototypes = this.prototypes(object, surface);
        for (const prototype of prototypes)
            values = this.prototype(prototype, node.parent!, values, raw);
        this.css.withFallback(node, values, objectDefaults, object.kind === "chart" ? { "font-size": "10pt" } : {});
        let computed = this.css.get(node);
        if (object.kind === "geometry" && object.preset !== "line" && !["linear", "none"].includes(computed["--office-fill-kind"] ?? "") && !computed["--office-text-color"]) {
            values = { ...values, color: "var(--office-theme-lt1,#ffffff)" };
            raw = { ...raw, color: "var(--office-theme-lt1,#ffffff)" };
            this.css.withFallback(node, values, objectDefaults);
            computed = this.css.get(node);
        }
        this.rawFallbacks.set(node, { ...raw });
        if (object.kind === "chart") {
            const titleId = this.p.charts.find(chart => chart.chartId === object.chartId)?.titleNodeId, title = titleId && this.tree.byId.get(titleId);
            if (title && title.tag === "div")
                this.css.withFallback(this.virtual(title, surface), {}, {}, { "font-weight": "bold" });
        }
        if (object.kind !== "text" && object.kind !== "geometry")
            return;
        const textClass = this.textClass(object, surface);
        for (const paragraph of paragraphs(source)) {
            const rendered = this.virtual(paragraph, surface), level = paragraph.attrs["data-office-level"] ?? "0", base = this.probe("p", { class: textClass ? textClass + " " + textClass + "-level-" + level : "" }, rendered.parent!);
            let inherited = this.values(this.css.get(base)), rawText = { ...this.css.specified(base) };
            for (const prototype of prototypes) {
                const match = paragraphs(prototype).find(p => (p.attrs["data-office-level"] ?? "0") === level);
                if (match)
                    inherited = this.prototype(match, rendered.parent!, inherited, rawText);
            }
            this.css.withFallback(rendered, inherited);
            this.rawFallbacks.set(rendered, rawText);
        }
    }
    private ensure(node: SourceElement, surface: PresentationStyleSurface): SourceElement {
        for (let current: SourceElement | undefined = node; current; current = current.parent) {
            const object = this.objects.get(current.attrs["data-node-id"]!);
            if (object) {
                this.prepare(object, surface);
                break;
            }
        }
        return this.virtual(node, surface);
    }
    get(node: SourceElement, surface: PresentationStyleSurface): Css { return this.css.get(this.ensure(node, surface)); }
    own(node: SourceElement, surface: PresentationStyleSurface): Css { return this.css.specified(this.ensure(node, surface)); }
    raw(node: SourceElement, surface: PresentationStyleSurface): Css {
        const projected = this.ensure(node, surface);
        const visit = (current: SourceElement | undefined): Css => {
            if (!current)
                return {};
            const cached = this.rawCache.get(current);
            if (cached)
                return cached;
            const inherited = Object.fromEntries(Object.entries(visit(current.parent)).filter(([name]) => name.startsWith("--office-") || ["font-family", "font-size", "font-weight", "font-style", "color", "line-height", "letter-spacing", "text-align", "tab-size", "text-transform", "font-variant"].includes(name)));
            const result = { ...inherited, ...this.rawFallbacks.get(current), ...this.css.specified(current) };
            this.rawCache.set(current, result);
            return result;
        };
        return visit(projected);
    }
    probeClass(names: string | readonly string[] | undefined, node: SourceElement, surface: PresentationStyleSurface): {
        css: Css;
        raw: Css;
        own: Css;
    } {
        const layers = typeof names === "string" ? [names] : names?.length ? names : [""];
        let probe = this.ensure(node, surface), raw = this.raw(node, surface), own: Css = {};
        for (const name of layers) { probe = this.probe("span", { class: name }, probe); own = this.css.specified(probe); raw = { ...raw, ...own }; }
        return { css: this.css.get(probe), own, raw };
    }
    theme(themeId: string): Css { return this.css.classStyle(this.p.themes.find(theme => theme.themeId === themeId)!.cssClass, this.css.get(this.tree.root)); }
    themeOwn(themeId: string): Css { return this.css.specifiedClass(this.p.themes.find(theme => theme.themeId === themeId)!.cssClass); }
    textDefaults(className: string | undefined, level: number, surface: PresentationStyleSurface): {
        css: Css;
        own: Css;
    } {
        const parent = this.virtual(this.tree.byId.get(surface.rootNodeId)!, surface), node = this.probe("p", { class: className ? className + " " + className + "-level-" + level : "" }, parent);
        return { css: this.css.get(node), own: this.css.specified(node) };
    }
}
