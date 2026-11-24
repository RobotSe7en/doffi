import { PresentationOperationSchema } from "../contracts/presentation-operations.js";
import type { PresentationOperation } from "../contracts/presentation-operations.js";
import type { DocumentOperation } from "../contracts/operations.js";
import type { PresentationMetadata, PresentationObject } from "../contracts/presentation.js";
import type { DocumentState } from "../document/state.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import type { DocumentReadLimits } from "../runtime.js";
import { invalid, checkLimit } from "../document/validation.js";
import { resolveReference } from "../document/package.js";
import { DocumentError } from "../errors.js";
import { parse, walk, ident, generate } from "css-tree";
import { evaluatePresentationPreset } from "../presentation-geometry.js";
import { applyPresentationGeometry, geometryAttributes, inverse, local, multiply, point, presentationTransforms } from "./presentation-geometry.js";
import type { PresentationGeometry } from "../contracts/presentation.js";
import type { Matrix } from "./presentation-geometry.js";
import { editPresentationTable } from "./presentation-table.js";
import type { PresentationStyleEdit } from "./presentation-styles.js";
import { definitionIds, definitionOperation, editPresentationDefinition, instantiateLayout } from "./presentation-definitions.js";
import { chartOperation, editPresentationChart } from "./presentation-charts.js";
import { commentOperation, editPresentationComment, validatePresentationComment } from "./presentation-comments.js";
import type { Annotation } from "../contracts/annotations.js";
const names: ReadonlySet<string> = new Set(PresentationOperationSchema.options.map(schema => schema.shape.op.value));
export function isPresentationOperation(operation: DocumentOperation): operation is PresentationOperation { return names.has(operation.op); }
export interface PresentationEditContext {
    readonly base: DocumentState;
    readonly p: PresentationMetadata;
    readonly edit: DocumentEdit;
    readonly parse: () => ParsedHtml;
    readonly allocate: () => string;
    readonly limits: DocumentReadLimits;
    readonly styles: PresentationStyleEdit;
    readonly annotations: Annotation[];
    readonly explicit: Set<string>;
}
const copy = <T>(input: T): T => JSON.parse(JSON.stringify(input)) as T;
const escape = (value: string): string => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\r/g, "&#13;");
function attributes(values: Readonly<Record<string, string>>): string { return Object.entries(values).map(([key, value]) => " " + key + '="' + escape(value) + '"').join(""); }
export function validatePresentationTargets(base: DocumentState, operation: PresentationOperation): void {
    const p = base.snapshot.metadata.presentation;
    if (!p)
        invalid("Presentation operations cannot target a Word document.");
    const nodes = ["nodeId" in operation ? operation.nodeId : undefined, "parentNodeId" in operation ? operation.parentNodeId : undefined,
        "beforeNodeId" in operation ? operation.beforeNodeId : undefined, ...("nodeIds" in operation ? operation.nodeIds : []),
        ...(operation.op === "insertPresentationObject" && operation.object.kind === "connector" ? [operation.object.startNodeId, operation.object.endNodeId] : [])].filter((id): id is string => id !== undefined);
    if (operation.op === "setPresentationConnector")
        for (const endpoint of [operation.changes.start, operation.changes.end])
            if (endpoint)
                nodes.push(endpoint.nodeId);
    if (nodes.some(id => !base.html.hasNode(id)))
        invalid("Presentation targets must belong to the base snapshot.");
    if ("nodeIds" in operation && new Set(operation.nodeIds).size !== operation.nodeIds.length)
        invalid("Presentation selections cannot repeat identities.");
    for (const id of ["slideId" in operation ? operation.slideId : undefined, "afterSlideId" in operation ? operation.afterSlideId : undefined, "beforeSlideId" in operation ? operation.beforeSlideId : undefined])
        if (id !== undefined && !p.slides.some(s => s.slideId === id))
            invalid("Slide targets must belong to the base snapshot.");
    if (operation.op === "insertSlide" && !p.layouts.some(l => l.layoutId === operation.layoutId))
        invalid("The new slide layout must belong to the base snapshot.");
    const reference = (kind: "theme" | "master" | "layout" | "notes-master", id: string | undefined): void => { if (id !== undefined && !definitionIds(p, kind).includes(id))
        invalid("Definition references must belong to the base snapshot."); };
    if (operation.op === "setPresentationTheme")
        reference("theme", operation.themeId);
    if (operation.op === "setPresentationSlideLayout")
        reference("layout", operation.layoutId);
    if (operation.op === "insertPresentationDefinition") {
        reference(operation.kind, operation.copyFromId);
        reference("theme", operation.themeId);
        reference("master", operation.masterId);
    }
    if (operation.op === "updatePresentationDefinition") {
        reference(operation.kind, operation.definitionId);
        reference("theme", operation.changes.themeId);
        reference("master", operation.changes.masterId);
    }
    if (operation.op === "removePresentationDefinition") {
        reference(operation.kind, operation.definitionId);
        reference(operation.kind, operation.replacementId);
    }
    if ((operation.op === "insertPresentationNotes" || operation.op === "setPresentationNotesMaster") && operation.notesMasterId)
        reference("notes-master", operation.notesMasterId);
    if (operation.op === "setPresentationChartStyle" && operation.seriesId !== undefined) {
        const chart = p.charts.find(c => c.nodeId === operation.nodeId), data = p.dataSets.find(d => d.dataSetId === chart?.dataSetId);
        if (!data?.series.some(s => s.seriesId === operation.seriesId))
            invalid("Chart series targets must belong to the base data set.");
    }
    if (commentOperation(operation))
        validatePresentationComment(base, operation);
}
function root(ctx: PresentationEditContext): string { const html = ctx.parse(); return html.nodeIds().find(id => html.peekNode(id)?.attributes["data-office-kind"] === "presentation")!; }
function slide(ctx: PresentationEditContext, id: string): PresentationMetadata["slides"][number] { return ctx.p.slides.find(s => s.slideId === id) ?? invalid("An earlier operation removed the slide."); }
function following(ctx: PresentationEditContext, nodeId: string): string | undefined {
    const html = ctx.parse(), node = html.peekNode(nodeId)!;
    const children = html.peekNode(node.parentNodeId!)!.childNodeIds, index = children.indexOf(nodeId);
    return children[index + 1];
}
export function registerCopies(ctx: PresentationEditContext, identities: ReadonlyMap<string, string>, detachPlaceholders: boolean, slides?: ReadonlyMap<string, string>): void {
    const html = ctx.parse();
    for (const item of [...ctx.p.nativePreview ?? []]) {
        const ownerId = identities.get(item.ownerId);
        if (["tags", "transition", "table-style", "emf-image"].includes(item.feature) && ownerId) ctx.p.nativePreview!.push({ ...item, ownerId });
    }
    const copiedTransforms = presentationTransforms(ctx.p, html);
    for (const original of [...ctx.p.objects]) {
        const id = identities.get(original.nodeId);
        if (!id)
            continue;
        const object = copy(original);
        object.nodeId = id;
        delete object.nativeId;
        delete object.creationId;
        if (detachPlaceholders && object.placeholder) {
            ctx.edit.attributes(id, geometryAttributes(copiedTransforms.geometry(original.nodeId)));
            delete object.placeholder;
        }
        if (object.kind === "connector")
            for (const end of ["start", "end"] as const) {
                const key = end === "start" ? "startNodeId" : "endNodeId", target = object[key];
                if (target && identities.has(target))
                    object[key] = identities.get(target)!;
                else if (target && copiedTransforms.owner(target) !== copiedTransforms.owner(id)) {
                    delete object[key];
                    delete object[end === "start" ? "startSite" : "endSite"];
                }
            }
        if (object.kind === "chart") {
            const source = ctx.p.charts.find(c => c.chartId === object.chartId)!;
            const chart = copy(source), data = copy(ctx.p.dataSets.find(d => d.dataSetId === source.dataSetId)!);
            chart.chartId = ctx.allocate();
            chart.nodeId = id;
            data.dataSetId = ctx.allocate();
            chart.dataSetId = data.dataSetId;
            const series = new Map(data.series.map(s => [s.seriesId, ctx.allocate()]));
            for (const item of data.series)
                item.seriesId = series.get(item.seriesId)!;
            for (const item of [...(chart.seriesStyles ?? []), ...(chart.pointStyles ?? [])])
                item.seriesId = series.get(item.seriesId)!;
            if (chart.titleNodeId)
                chart.titleNodeId = identities.get(chart.titleNodeId)!;
            object.chartId = chart.chartId;
            ctx.p.charts.push(chart);
            ctx.p.dataSets.push(data);
        }
        ctx.p.objects.push(object);
    }
    for (const original of [...ctx.p.hyperlinks])
        if (identities.has(original.nodeId))
            ctx.p.hyperlinks.push({ ...copy(original), nodeId: identities.get(original.nodeId)!, ...(original.slideId && slides?.has(original.slideId) ? { slideId: slides.get(original.slideId)! } : {}) });
}
function insertObject(ctx: PresentationEditContext, operation: Extract<PresentationOperation, {
    op: "insertPresentationObject";
}>): void {
    const html = ctx.parse(), parent = html.peekNode(operation.parentNodeId), group = ctx.p.objects.find(o => o.nodeId === operation.parentNodeId);
    if (!parent || !["slide", "master", "layout", "notes-master", "notes"].includes(parent.attributes["data-office-kind"] ?? "") && group?.kind !== "group")
        invalid("Objects require a presentation surface or group parent.");
    const input = operation.object;
    let body = "", object: PresentationObject;
    if (input.kind === "text" || input.kind === "geometry") {
        body = "<p>" + escape(input.text ?? "").replace(/\n/g, "<br>") + "</p>";
        object = input.kind === "text" ? { kind: "text", nodeId: "pending" } : { kind: "geometry", nodeId: "pending", preset: input.preset };
    }
    else if (input.kind === "image") {
        const manifest = ctx.base.snapshot.manifest, directory = manifest.entrypoint.split("/").slice(0, -1), parts = input.assetPath.split("/");
        while (directory.length && directory[0] === parts[0]) {
            directory.shift();
            parts.shift();
        }
        const relative = [...directory.map(() => ".."), ...parts.map(encodeURIComponent)].join("/");
        if (resolveReference(relative, manifest.entrypoint) !== input.assetPath)
            invalid("The image path cannot be resolved.");
        body = '<img src="' + escape(relative) + '" alt="' + escape(input.alt) + '">';
        object = { kind: "image", nodeId: "pending" };
    }
    else if (input.kind === "table") {
        checkLimit(input.rows * (input.columns * 2 + 1) + 4, ctx.limits.maxNodes);
        body = '<table style="width:100%;border-collapse:collapse"><tbody>' + Array.from({ length: input.rows }, () => '<tr style="height:' + operation.geometry.height / input.rows + 'pt">' + "<td><p></p></td>".repeat(input.columns) + "</tr>").join("") + "</tbody></table>";
        object = { kind: "table", nodeId: "pending", grid: Array.from({ length: input.columns }, () => operation.geometry.width / input.columns) };
    }
    else {
        const { style: _style, ...definition } = input;
        object = { ...definition, nodeId: "pending" };
    }
    const inserted = ctx.edit.insertSource(operation.parentNodeId, "<div" + attributes({ ...geometryAttributes(operation.geometry), "data-office-kind": input.kind, ...(input.style ? { style: input.style } : {}) }) + ">" + body + "</div>", operation.beforeNodeId);
    object.nodeId = inserted.rootNodeIds[0]!;
    ctx.p.objects.push(object);
}
function selection(ctx: PresentationEditContext, ids: readonly string[]) {
    const html = ctx.parse(), objects = new Map(ctx.p.objects.map(o => [o.nodeId, o])), selected = new Set(ids);
    if (ids.some(id => !objects.has(id) || !html.hasNode(id)))
        invalid("Selections require current presentation objects.");
    const parent = html.peekNode(ids[0]!)!.parentNodeId!;
    if (ids.some(id => html.peekNode(id)!.parentNodeId !== parent))
        invalid("This object operation requires one shared parent coordinate system.");
    const siblings = html.peekNode(parent)!.childNodeIds, ordered = siblings.filter(id => selected.has(id));
    const transforms = presentationTransforms(ctx.p, html);
    const entries = ordered.map(id => {
        const geometry = transforms.geometry(id), matrix = local(geometry);
        const corners = [[0, 0], [geometry.width, 0], [0, geometry.height], [geometry.width, geometry.height]].map(([x, y]) => point(matrix, { x: x!, y: y! }));
        const left = Math.min(...corners.map(p => p.x)), top = Math.min(...corners.map(p => p.y)), right = Math.max(...corners.map(p => p.x)), bottom = Math.max(...corners.map(p => p.y));
        return { id, geometry, bounds: { x: left, y: top, width: right - left, height: bottom - top } };
    });
    return { html, parent, siblings, ordered, selected, entries };
}
function duplicateObjects(ctx: PresentationEditContext, operation: Extract<PresentationOperation, {
    op: "duplicatePresentationObjects";
}>): void {
    const selected = selection(ctx, operation.nodeIds), before = following(ctx, selected.ordered.at(-1)!);
    const identities = new Map<string, string>();
    for (const id of selected.ordered) {
        const cloned = ctx.edit.cloneSource(id, selected.parent, before);
        for (const [original, copied] of cloned.identities)
            identities.set(original, copied);
    }
    registerCopies(ctx, identities, true);
    for (const entry of selected.entries)
        ctx.edit.attributes(identities.get(entry.id)!, geometryAttributes({ ...entry.geometry, x: entry.geometry.x + operation.offset.x, y: entry.geometry.y + operation.offset.y }));
}
function arrangeObjects(ctx: PresentationEditContext, operation: Extract<PresentationOperation, {
    op: "alignPresentationObjects" | "distributePresentationObjects" | "reorderPresentationObjects";
}>): void {
    const selected = selection(ctx, operation.nodeIds), { entries } = selected;
    if (operation.op === "reorderPresentationObjects") {
        const order = [...selected.siblings];
        if (operation.position === "front" || operation.position === "back") {
            const other = order.filter(id => !selected.selected.has(id));
            order.splice(0, order.length, ...(operation.position === "front" ? [...other, ...selected.ordered] : [...selected.ordered, ...other]));
        }
        else if (operation.position === "forward") {
            for (let i = order.length - 2; i >= 0; i--)
                if (selected.selected.has(order[i]!) && !selected.selected.has(order[i + 1]!))
                    [order[i], order[i + 1]] = [order[i + 1]!, order[i]!];
        }
        else
            for (let i = 1; i < order.length; i++)
                if (selected.selected.has(order[i]!) && !selected.selected.has(order[i - 1]!))
                    [order[i], order[i - 1]] = [order[i - 1]!, order[i]!];
        if (order.some((id, i) => id !== selected.siblings[i]))
            ctx.edit.moveSource(order, selected.parent);
        return;
    }
    if (operation.op === "alignPresentationObjects") {
        if (operation.relativeTo === "slide" && selected.html.peekNode(selected.parent)!.attributes["data-office-kind"] !== "slide")
            invalid("Slide alignment requires top-level slide objects.");
        const x = Math.min(...entries.map(e => e.bounds.x)), y = Math.min(...entries.map(e => e.bounds.y)), right = Math.max(...entries.map(e => e.bounds.x + e.bounds.width)), bottom = Math.max(...entries.map(e => e.bounds.y + e.bounds.height));
        const area = operation.relativeTo === "slide" ? { x: 0, y: 0, width: ctx.p.size.width, height: ctx.p.size.height } : { x, y, width: right - x, height: bottom - y };
        for (const entry of entries) {
            const { bounds: b, geometry: g } = entry;
            const delta = operation.alignment === "left" ? area.x - b.x : operation.alignment === "center" ? area.x + area.width / 2 - b.x - b.width / 2 : operation.alignment === "right" ? area.x + area.width - b.x - b.width
                : operation.alignment === "top" ? area.y - b.y : operation.alignment === "middle" ? area.y + area.height / 2 - b.y - b.height / 2 : area.y + area.height - b.y - b.height;
            const horizontal = ["left", "center", "right"].includes(operation.alignment);
            if (Math.abs(delta) > 1e-10)
                ctx.edit.attributes(entry.id, geometryAttributes({ ...g, x: g.x + (horizontal ? delta : 0), y: g.y + (horizontal ? 0 : delta) }));
        }
        return;
    }
    if (entries.length < 3)
        invalid("Distribution requires at least three objects.");
    const axis = operation.axis === "horizontal" ? "x" : "y", size = operation.axis === "horizontal" ? "width" : "height";
    const ordered = [...entries].sort((a, b) => a.bounds[axis] - b.bounds[axis]), first = ordered[0]!, last = ordered.at(-1)!;
    const gap = (last.bounds[axis] + last.bounds[size] - first.bounds[axis] - ordered.reduce((sum, e) => sum + e.bounds[size], 0)) / (ordered.length - 1);
    let cursor = first.bounds[axis];
    for (const entry of ordered) {
        const delta = cursor - entry.bounds[axis];
        if (Math.abs(delta) > 1e-10)
            ctx.edit.attributes(entry.id, geometryAttributes({ ...entry.geometry, [axis]: entry.geometry[axis] + delta }));
        cursor += entry.bounds[size] + gap;
    }
}
/** Changing ancestry is safe only when source styles do not depend on that group. */
function guardGroupingStyles(ctx: PresentationEditContext, groupId?: string): void {
    if (groupId) {
        const attrs = ctx.parse().peekNode(groupId)!.attributes;
        if (attrs.class?.trim() || attrs.style?.trim() || attrs.id || attrs.lang || attrs.dir || attrs["data-office-hidden"] === "true")
            throw new DocumentError("UnsupportedFeature", "Ungrouping a painted or styled group requires explicit appearance materialization.");
    }
    for (const style of ctx.base.snapshot.styles)
        walk(parse(style.css), node => {
            if (node.type === "TypeSelector" && ident.decode(node.name).toLowerCase() === "div"
                || node.type === "AttributeSelector" && (/^data-(?:node-id$|office-(?:kind$|x$|y$|width$|height$|child-|rotation$|flip-))/.test(ident.decode(node.name.name)) || groupId !== undefined && ctx.parse().peekNode(groupId)!.attributes[ident.decode(node.name.name)] !== undefined))
                throw new DocumentError("UnsupportedFeature", "Grouping source with ancestry-sensitive object selectors requires explicit appearance materialization.");
        });
}
function groupObjects(ctx: PresentationEditContext, ids: readonly string[]): void {
    const selected = selection(ctx, ids);
    if (ids.length < 2)
        invalid("Grouping requires at least two sibling objects.");
    guardGroupingStyles(ctx);
    if (ctx.p.objects.some(o => selected.selected.has(o.nodeId) && o.placeholder))
        throw new DocumentError("UnsupportedFeature", "Native placeholders cannot be grouped directly.");
    const x = Math.min(...selected.entries.map(e => e.bounds.x)), y = Math.min(...selected.entries.map(e => e.bounds.y));
    const width = Math.max(1, Math.max(...selected.entries.map(e => e.bounds.x + e.bounds.width)) - x), height = Math.max(1, Math.max(...selected.entries.map(e => e.bounds.y + e.bounds.height)) - y);
    const geometry = { x, y, width, height, rotation: 0, flipH: false, flipV: false };
    const result = ctx.edit.insertSource(selected.parent, "<div" + attributes({ "data-office-kind": "group", ...geometryAttributes(geometry), "data-office-child-x": "0", "data-office-child-y": "0", "data-office-child-width": String(width), "data-office-child-height": String(height) }) + "></div>", following(ctx, selected.ordered.at(-1)!));
    const groupId = result.rootNodeIds[0]!;
    ctx.p.objects.push({ kind: "group", nodeId: groupId });
    ctx.edit.moveSource(selected.ordered, groupId);
    for (const entry of selected.entries)
        ctx.edit.attributes(entry.id, geometryAttributes({ ...entry.geometry, x: entry.geometry.x - x, y: entry.geometry.y - y }));
}
function decompose(matrix: Matrix, width: number, height: number, flipH: boolean): PresentationGeometry {
    const sx = Math.hypot(matrix[0], matrix[1]), sy = Math.hypot(matrix[2], matrix[3]), determinant = matrix[0] * matrix[3] - matrix[1] * matrix[2];
    if (!sx || !sy || !Number.isFinite(sx + sy) || Math.abs(matrix[0] * matrix[2] + matrix[1] * matrix[3]) > sx * sy * 1e-9)
        throw new DocumentError("UnsupportedFeature", "Ungrouping would require shear or a singular native transform.");
    const rotation = Math.atan2(matrix[1] * (flipH ? -1 : 1), matrix[0] * (flipH ? -1 : 1)) * 180 / Math.PI, w = width * sx, h = height * sy, center = point(matrix, { x: width / 2, y: height / 2 });
    return { x: center.x - w / 2, y: center.y - h / 2, width: w, height: h, rotation, flipH, flipV: determinant < 0 !== flipH };
}
/** Conservative paint proof: unrelated theme format classes do not paint this object. */
function scaledPaintRequiresMaterialization(ctx: PresentationEditContext, id: string): boolean {
    const html = ctx.parse(), ancestors: {
        tagName: string;
        attributes: Readonly<Record<string, string>>;
    }[] = [];
    for (let current: string | null = id; current !== null; current = html.peekNode(current)!.parentNodeId)
        ancestors.push(html.peekNode(current)!);
    const inheritedNodes = [...ancestors, ...ctx.edit.shellAncestors()];
    const classes = new Set(inheritedNodes.flatMap(node => (node.attributes.class ?? "").split(/\s+/))), tags = new Set(inheritedNodes.map(node => node.tagName));
    let noLine = false, line = false, color = false, shadow = false, border = false;
    const declaration = (name: string, value: string, certain: boolean): void => {
        if (name === "--office-line-kind") {
            if (value === "none") {
                if (certain)
                    noLine = true;
            }
            else
                line = true;
        }
        if (name === "--office-line-color")
            color = true;
        if (name === "--office-shadow-kind" && value !== "none")
            shadow = true;
        if (/^--office-line-(start|end)-arrow$/.test(name) && value !== "none")
            line = true;
        if (/^border(?:-|$)/.test(name) && !/^(?:none|0(?:pt|px)?)$/.test(value))
            border = true;
    };
    for (const source of ctx.base.snapshot.styles) {
        const ast = parse(source.css);
        walk(ast, node => {
            if (node.type !== "Rule" || node.prelude?.type !== "SelectorList")
                return;
            const possible = node.prelude.children.toArray().some(selector => {
                let matches = true;
                walk(selector, part => {
                    if (part.type === "ClassSelector" && !classes.has(ident.decode(part.name)))
                        matches = false;
                    if (part.type === "TypeSelector" && !tags.has(ident.decode(part.name).toLowerCase()))
                        matches = false;
                    if (part.type === "IdSelector" && !inheritedNodes.some(node => node.attributes.id === ident.decode(part.name)))
                        matches = false;
                    if (part.type === "AttributeSelector" && !inheritedNodes.some(node => {
                        const name = ident.decode(part.name.name), value = part.value?.type === "String" ? part.value.value : part.value?.type === "Identifier" ? ident.decode(part.value.name) : undefined;
                        return node.attributes[name] !== undefined && (value === undefined || node.attributes[name] === value);
                    }))
                        matches = false;
                });
                return matches;
            });
            const certain = node.prelude.children.toArray().some(selector => selector.type === "Selector" && inheritedNodes.some(target => selector.children.toArray().every(part => part.type === "ClassSelector" ? (target.attributes.class ?? "").split(/\s+/).includes(ident.decode(part.name))
                : part.type === "TypeSelector" ? target.tagName === ident.decode(part.name).toLowerCase()
                    : part.type === "IdSelector" ? target.attributes.id === ident.decode(part.name)
                        : part.type === "PseudoClassSelector" ? target.tagName === "html" && part.name === "root" : false)));
            if (possible)
                for (const decl of node.block.children)
                    if (decl.type === "Declaration")
                        declaration(ident.decode(decl.property), generate(decl.value), certain);
        });
    }
    for (const node of inheritedNodes)
        if (node.attributes.style)
            walk(parse(node.attributes.style, { context: "declarationList" }), decl => { if (decl.type === "Declaration")
                declaration(ident.decode(decl.property), generate(decl.value), true); });
    return line || color && !noLine || shadow || border;
}
function ungroupObject(ctx: PresentationEditContext, id: string): void {
    const html = ctx.parse(), group = ctx.p.objects.find(o => o.nodeId === id), node = html.peekNode(id);
    if (group?.kind !== "group" || !node?.parentNodeId)
        invalid("Ungrouping requires a current group.");
    guardGroupingStyles(ctx, id);
    const transforms = presentationTransforms(ctx.p, html), parentInverse = inverse(transforms.parent(id)), groupGeometry = transforms.geometry(id), a = node.attributes;
    const groupScaleX = groupGeometry.width / Number(a["data-office-child-width"]), groupScaleY = groupGeometry.height / Number(a["data-office-child-height"]);
    const scaled = Math.abs(groupScaleX - 1) > 1e-9 || Math.abs(groupScaleY - 1) > 1e-9;
    const children = node.childNodeIds.map(childId => {
        const object = ctx.p.objects.find(o => o.nodeId === childId)!, before = transforms.geometry(childId), matrix = multiply(parentInverse, transforms.world(childId));
        if (object.placeholder)
            throw new DocumentError("UnsupportedFeature", "Ungrouping native placeholders requires explicit inheritance materialization.");
        const geometry = decompose(matrix, before.width, before.height, groupGeometry.flipH !== before.flipH);
        if (scaled && object.kind !== "group") {
            if (!["text", "geometry"].includes(object.kind))
                throw new DocumentError("UnsupportedFeature", "Scaled object appearance cannot be flattened by this geometry operation.");
            // A painted outline/shadow is anisotropic under a nonuniform group, and
            // requires an explicit style edit before it can share native dimensions.
            if (scaledPaintRequiresMaterialization(ctx, childId) || object.kind === "geometry" && (object.preset === "line" || object.customGeometry))
                throw new DocumentError("UnsupportedFeature", "Scaled strokes and shadows require explicit appearance materialization.");
            const preset = object.kind === "geometry" ? object.preset : "rect", adjustments = Object.fromEntries(Object.entries(html.peekNode(childId)!.attributes).filter(([key]) => key.startsWith("data-office-adjust-")).map(([key, value]) => [key.slice(19), Number(value)]));
            const old = evaluatePresentationPreset(preset, before.width, before.height, adjustments), next = evaluatePresentationPreset(preset, geometry.width, geometry.height, adjustments);
            const scaleX = geometry.width / before.width, scaleY = geometry.height / before.height;
            const scaledPaths = old.paths.map(path => ({ ...path, commands: path.commands.map(command => {
                    if (command.kind === "move" || command.kind === "line")
                        return { ...command, x: command.x * scaleX, y: command.y * scaleY };
                    if (command.kind === "arc")
                        return { ...command, radiusX: command.radiusX * scaleX, radiusY: command.radiusY * scaleY };
                    return command;
                }) }));
            const equal = (left: unknown, right: unknown): boolean => typeof left === "number" && typeof right === "number" ? Math.abs(left - right) <= 1e-7 * Math.max(1, Math.abs(left), Math.abs(right))
                : Array.isArray(left) && Array.isArray(right) ? left.length === right.length && left.every((value, index) => equal(value, right[index]))
                    : left !== null && right !== null && typeof left === "object" && typeof right === "object" ? Object.keys(left).length === Object.keys(right).length && Object.entries(left).every(([key, value]) => equal(value, (right as Record<string, unknown>)[key])) : left === right;
            const textRect = old.textRect ? { x: old.textRect.x * scaleX, y: old.textRect.y * scaleY, width: old.textRect.width * scaleX, height: old.textRect.height * scaleY } : null;
            if (!equal(scaledPaths, next.paths) || !equal(textRect, next.textRect))
                throw new DocumentError("UnsupportedFeature", "The preset cannot preserve its paths or text rectangle after nonuniform ungrouping.");
        }
        return { id: childId, geometry };
    });
    const before = following(ctx, id);
    ctx.edit.moveSource(node.childNodeIds, node.parentNodeId, before);
    for (const child of children)
        ctx.edit.attributes(child.id, geometryAttributes(child.geometry));
    ctx.edit.apply({ op: "removeNode", nodeId: id });
}
/** Apply a typed format command to the same isolated source tree and metadata. */
export function applyPresentationOperation(ctx: PresentationEditContext, operation: PresentationOperation): void {
    if (definitionOperation(operation)) {
        editPresentationDefinition(ctx, operation);
        return;
    }
    if (chartOperation(operation)) {
        editPresentationChart(ctx, operation);
        return;
    }
    if (commentOperation(operation)) {
        editPresentationComment(ctx, operation);
        return;
    }
    const { p, edit } = ctx;
    switch (operation.op) {
        case "setPresentationSize":
            p.size = operation.size;
            return;
        case "setPresentationGeometry":
            applyPresentationGeometry(operation, p, edit, ctx.base.html);
            return;
        case "insertPresentationObject":
            insertObject(ctx, operation);
            return;
        case "duplicatePresentationObjects":
            duplicateObjects(ctx, operation);
            return;
        case "alignPresentationObjects":
        case "distributePresentationObjects":
        case "reorderPresentationObjects":
            arrangeObjects(ctx, operation);
            return;
        case "groupPresentationObjects":
            groupObjects(ctx, operation.nodeIds);
            return;
        case "ungroupPresentationObject":
            ungroupObject(ctx, operation.nodeId);
            return;
        case "insertSlide": {
            const layout = p.layouts.find(l => l.layoutId === operation.layoutId) ?? invalid("The new slide layout is missing.");
            const after = operation.afterSlideId ? slide(ctx, operation.afterSlideId) : undefined, before = after ? following(ctx, after.rootNodeId) : undefined;
            const inserted = edit.insertSource(root(ctx), '<section data-office-kind="slide" data-office-story="slide"></section>', before);
            const entry = { slideId: ctx.allocate(), rootNodeId: inserted.rootNodeIds[0]!, layoutId: layout.layoutId };
            p.slides.push(entry);
            instantiateLayout(ctx, entry.slideId);
            return;
        }
        case "duplicateSlide": {
            const original = slide(ctx, operation.slideId), after = operation.afterSlideId ? slide(ctx, operation.afterSlideId) : original;
            if (p.nativePreview?.some(item => item.feature === "animation" && item.ownerId === original.slideId))
                throw new DocumentError("UnsupportedFeature", "Duplicating a slide with preserved animation would lose its native object bindings.");
            const cloned = edit.cloneSource(original.rootNodeId, root(ctx), following(ctx, after.rootNodeId));
            const entry = { ...copy(original), slideId: ctx.allocate(), rootNodeId: cloned.rootNodeIds[0]! };
            p.slides.push(entry);
            const slideMapping = new Map([[original.slideId, entry.slideId]]);
            registerCopies(ctx, cloned.identities, false, slideMapping);
            for (const story of [...p.stories])
                if (story.kind === "notes" && story.slideId === original.slideId) {
                    const notes = edit.cloneSource(story.rootNodeId, root(ctx));
                    p.stories.push({ ...copy(story), storyId: ctx.allocate(), rootNodeId: notes.rootNodeIds[0]!, slideId: entry.slideId, ...(story.bodyNodeId ? { bodyNodeId: notes.identities.get(story.bodyNodeId)! } : {}) });
                    registerCopies(ctx, notes.identities, false, slideMapping);
                }
            return;
        }
        case "removeSlide": {
            const entry = slide(ctx, operation.slideId);
            if (p.slides.length === 1)
                invalid("A presentation must retain at least one slide.");
            edit.apply({ op: "removeNode", nodeId: entry.rootNodeId });
            for (const story of p.stories)
                if (story.slideId === entry.slideId) {
                    if (story.kind === "notes")
                        edit.apply({ op: "removeNode", nodeId: story.rootNodeId });
                    else
                        delete story.slideId;
                }
            p.stories = p.stories.filter(s => s.kind !== "notes" || s.slideId !== entry.slideId);
            if (p.nativePreview) p.nativePreview = p.nativePreview.filter(item => item.feature !== "animation" || item.ownerId !== entry.slideId);
            p.slides = p.slides.filter(s => s !== entry);
            p.hyperlinks = p.hyperlinks.filter(link => link.slideId !== entry.slideId);
            return;
        }
        case "moveSlide": {
            const entry = slide(ctx, operation.slideId), before = operation.beforeSlideId ? slide(ctx, operation.beforeSlideId) : undefined;
            const roots = new Set(p.slides.map(slide => slide.rootNodeId)), order = ctx.parse().peekNode(root(ctx))!.childNodeIds.filter(id => roots.has(id));
            if (entry === before || before?.rootNodeId === order[order.indexOf(entry.rootNodeId) + 1] || !before && order.at(-1) === entry.rootNodeId)
                return;
            edit.moveSource([entry.rootNodeId], root(ctx), before?.rootNodeId);
            return;
        }
        case "setSlideProperties": {
            const entry = slide(ctx, operation.slideId), original = ctx.base.snapshot.metadata.presentation!.slides.find(slide => slide.slideId === entry.slideId)!;
            Object.assign(entry, operation.changes);
            if (operation.changes.hidden !== undefined && operation.changes.hidden === (original.hidden ?? false)) {
                if (original.hidden === undefined)
                    delete entry.hidden;
                else
                    entry.hidden = original.hidden;
            }
            return;
        }
        case "setPresentationCrop": {
            if (p.objects.find(o => o.nodeId === operation.nodeId)?.kind !== "image")
                invalid("Crop operations require an image object.");
            if (operation.crop.left + operation.crop.right >= 1 || operation.crop.top + operation.crop.bottom >= 1)
                invalid("An image crop must retain a visible area.");
            const original = ctx.base.html.peekNode(operation.nodeId)!.attributes;
            edit.attributes(operation.nodeId, Object.fromEntries(Object.entries(operation.crop).map(([name, value]) => { const key = "data-office-crop-" + name; return [key, value === Number(original[key] ?? 0) ? original[key] : String(value)]; })));
            return;
        }
        case "setPresentationPreset": {
            const object = p.objects.find(o => o.nodeId === operation.nodeId);
            if (object?.kind !== "geometry")
                invalid("Preset operations require a geometry object.");
            if (object.preset === operation.preset && !object.customGeometry)
                return;
            delete object.customGeometry;
            object.preset = operation.preset;
            const original = ctx.base.snapshot.metadata.presentation!.objects.find(o => o.nodeId === operation.nodeId);
            const restored = original?.kind === "geometry" && original.preset === operation.preset ? Object.fromEntries(Object.entries(ctx.base.html.peekNode(operation.nodeId)!.attributes).filter(([key]) => key.startsWith("data-office-adjust-"))) : {};
            edit.attributes(operation.nodeId, { ...Object.fromEntries(Object.keys(ctx.parse().peekNode(operation.nodeId)!.attributes).filter(key => key.startsWith("data-office-adjust-")).map(key => [key, undefined])), ...restored });
            if (original?.kind === "geometry" && original.preset === operation.preset)
                edit.orderAttributes(operation.nodeId, Object.keys(ctx.base.html.peekNode(operation.nodeId)!.attributes));
            return;
        }
        case "setPresentationConnector": {
            const object = p.objects.find(o => o.nodeId === operation.nodeId);
            if (object?.kind !== "connector")
                invalid("Connector operations require a current connector.");
            const original = ctx.base.snapshot.metadata.presentation!.objects.find(o => o.nodeId === operation.nodeId);
            for (const end of ["start", "end"] as const) {
                const endpoint = operation.changes[end], nodeKey = end === "start" ? "startNodeId" : "endNodeId", siteKey = end === "start" ? "startSite" : "endSite";
                if (endpoint === null) {
                    delete object[nodeKey];
                    delete object[siteKey];
                }
                else if (endpoint) {
                    object[nodeKey] = endpoint.nodeId;
                    if (original?.kind === "connector" && original[nodeKey] === endpoint.nodeId && (original[siteKey] ?? 0) === endpoint.site) {
                        if (original[siteKey] === undefined)
                            delete object[siteKey];
                        else
                            object[siteKey] = original[siteKey];
                    }
                    else
                        object[siteKey] = endpoint.site;
                }
            }
            if (operation.changes.routing && operation.changes.routing !== object.routing) {
                object.routing = operation.changes.routing;
                const old = ctx.base.html.peekNode(operation.nodeId)!.attributes, restored = original?.kind === "connector" && original.routing === object.routing ? Object.fromEntries(Object.entries(old).filter(([key]) => key.startsWith("data-office-adjust-") || key === "data-office-connector-preset")) : { "data-office-connector-preset": object.routing === "straight" ? "straightConnector1" : "bentConnector3" };
                edit.attributes(operation.nodeId, { ...Object.fromEntries(Object.keys(ctx.parse().peekNode(operation.nodeId)!.attributes).filter(key => key.startsWith("data-office-adjust-") || key === "data-office-connector-preset").map(key => [key, undefined])), ...restored });
                if (original?.kind === "connector" && original.routing === object.routing)
                    edit.orderAttributes(operation.nodeId, Object.keys(old));
            }
            return;
        }
        case "insertPresentationNotes": {
            const entry = slide(ctx, operation.slideId);
            if (p.stories.some(story => story.kind === "notes" && story.slideId === entry.slideId))
                invalid("This slide already has a notes story.");
            const notesMaster = operation.notesMasterId ? p.notesMasters.find(n => n.notesMasterId === operation.notesMasterId) : p.notesMasters[0], layout = p.layouts.find(layout => layout.layoutId === entry.layoutId)!, master = p.masters.find(master => master.masterId === layout.masterId)!, theme = p.themes.find(theme => theme.themeId === (notesMaster?.themeId ?? master.themeId))!;
            if (operation.notesMasterId && !notesMaster)
                invalid("An earlier operation removed the requested notes master.");
            const inserted = edit.insertSource(root(ctx), '<div data-office-kind="notes" data-office-story="notes" class="' + escape(theme.cssClass) + '">' + operation.html + '</div>');
            p.stories.push({ kind: "notes", storyId: ctx.allocate(), rootNodeId: inserted.rootNodeIds[0]!, slideId: entry.slideId, ...(notesMaster ? { notesMasterId: notesMaster.notesMasterId } : {}) });
            return;
        }
        case "editPresentationTable":
            editPresentationTable(ctx, operation);
            return;
        case "setPresentationTextProperties": {
            const object = p.objects.find(object => object.nodeId === operation.nodeId);
            if (object?.kind !== "text" && object?.kind !== "geometry")
                invalid("Text properties require a text-capable object.");
            const text = { ...object.text };
            for (const key of ["autoFit", "wrap", "fontScale", "lineSpacingReduction"] as const) {
                const value = operation.changes[key];
                if (value === null)
                    delete text[key];
                else if (value !== undefined)
                    Object.assign(text, { [key]: value });
            }
            if (operation.changes.autoFit && operation.changes.autoFit !== "normal") {
                text.fontScale = 1;
                text.lineSpacingReduction = 0;
            }
            if (Object.keys(text).length)
                object.text = text;
            else
                delete object.text;
            return;
        }
    }
}
