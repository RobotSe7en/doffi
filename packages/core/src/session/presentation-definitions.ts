import type { PresentationOperation } from "../contracts/presentation-operations.js";
import type { PresentationMetadata } from "../contracts/presentation.js";
import type { DeepReadonly } from "../runtime.js";
import { invalid } from "../document/validation.js";
import { geometryAttributes, presentationTransforms } from "./presentation-geometry.js";
import { registerCopies } from "./presentation-edit.js";
import type { PresentationEditContext } from "./presentation-edit.js";
import { allocatePresentationClass } from "./presentation-styles.js";
import { placeholderReferenceKeys } from "../document/presentation.js";
import { setInlineStyle } from "../document/css.js";
type Kind = "theme" | "master" | "layout" | "notes-master";
export function definitionIds(p: DeepReadonly<PresentationMetadata>, kind: Kind): readonly string[] { return kind === "theme" ? p.themes.map(d => d.themeId) : kind === "master" ? p.masters.map(d => d.masterId) : kind === "layout" ? p.layouts.map(d => d.layoutId) : p.notesMasters.map(d => d.notesMasterId); }
export function definitionOperation(operation: PresentationOperation): boolean { return ["setPresentationTheme", "insertPresentationDefinition", "updatePresentationDefinition", "removePresentationDefinition", "setPresentationSlideLayout", "setPresentationNotesMaster", "setPresentationPlaceholder", "resetPresentationOverrides"].includes(operation.op); }
export function instantiateLayout(ctx: PresentationEditContext, slideId: string): void {
    const { p, edit } = ctx, slide = p.slides.find(s => s.slideId === slideId)!, layout = p.layouts.find(l => l.layoutId === slide.layoutId)!, html = ctx.parse();
    const existing = p.objects.filter(o => html.peekNode(o.nodeId)?.parentNodeId === slide.rootNodeId && o.placeholder).map(o => o.placeholder!);
    for (const original of [...p.objects])
        if (original.placeholder && html.peekNode(original.nodeId)?.parentNodeId === layout.rootNodeId && !existing.some(ph => placeholderReferenceKeys(ph).includes(original.placeholder!.type + ":" + original.placeholder!.index))) {
            if (original.kind === "text" || original.kind === "geometry") {
                const nodeId = edit.insertSource(slide.rootNodeId, '<div data-office-kind="' + original.kind + '"><p></p></div>').rootNodeIds[0]!;
                p.objects.push(original.kind === "text" ? { kind: "text", nodeId, placeholder: { ...original.placeholder } } : { kind: "geometry", nodeId, preset: original.preset, placeholder: { ...original.placeholder } });
            }
            else {
                const copied = edit.cloneSource(original.nodeId, slide.rootNodeId);
                registerCopies(ctx, copied.identities, false);
            }
        }
}
function switchLayout(ctx: PresentationEditContext, slideId: string, layoutId: string): void {
    const { p, edit } = ctx, slide = p.slides.find(s => s.slideId === slideId) ?? invalid("The slide is missing."), layout = p.layouts.find(l => l.layoutId === layoutId) ?? invalid("The layout is missing.");
    if (slide.layoutId === layoutId)
        return;
    const html = ctx.parse(), transforms = presentationTransforms(p, html), geometries = new Map(p.objects.filter(o => html.peekNode(o.nodeId)?.parentNodeId === slide.rootNodeId && o.placeholder).map(o => [o.nodeId, transforms.geometry(o.nodeId)]));
    slide.layoutId = layout.layoutId;
    // Unmatched placeholders retain their text and former resolved geometry.
    const inherited = presentationTransforms(p, html);
    for (const [id, g] of geometries)
        try {
            inherited.geometry(id);
        }
        catch {
            edit.attributes(id, geometryAttributes(g));
        }
    instantiateLayout(ctx, slideId);
}
export function editPresentationDefinition(ctx: PresentationEditContext, operation: PresentationOperation): void {
    const { p, edit } = ctx;
    if (operation.op === "setPresentationTheme") {
        const theme = p.themes.find(t => t.themeId === operation.themeId) ?? invalid("The theme is missing.");
        for (const [property, value] of Object.entries(operation.properties)) {
            if (!/^--office-(?:theme-(?:dk[12]|lt[12]|accent[1-6]|hlink|folHlink)|(?:major|minor)-(?:ascii|h-ansi|east-asia|complex-script))$/.test(property))
                invalid("Theme editing accepts the declared color and font roles.");
            setInlineStyle("", property.startsWith("--office-theme-") ? "color" : "font-family", value, ctx.limits);
        }
        const properties: Record<string, string> = {};
        for (const [property, value] of Object.entries(operation.properties)) {
            if (property.startsWith("--office-theme-")) {
                const palette = property.replace("--office-theme-", "--office-palette-");
                properties[palette] = value;
                properties[property] = "var(" + palette + ")";
            }
            else
                properties[property] = value;
        }
        for (const group of ["major", "minor"]) {
            const ascii = "--office-" + group + "-ascii", ansi = "--office-" + group + "-h-ansi", a = properties[ascii], b = properties[ansi];
            if (a !== undefined && b !== undefined && a !== b)
                invalid("A native presentation theme has one Latin font slot; use text runs for local font distinctions.");
            if (a !== undefined || b !== undefined)
                properties[ascii] = properties[ansi] = a ?? b!;
        }
        ctx.styles.setClass(theme.cssClass, properties);
        return;
    }
    if (operation.op === "setPresentationSlideLayout") {
        switchLayout(ctx, operation.slideId, operation.layoutId);
        return;
    }
    if (operation.op === "setPresentationNotesMaster") {
        const note = p.stories.find(s => s.kind === "notes" && s.slideId === operation.slideId);
        if (!note || note.kind !== "notes")
            invalid("Assigning a notes master requires an existing notes story.");
        const html = ctx.parse(), before = presentationTransforms(p, html), geometry = new Map(p.objects.filter(o => o.placeholder && before.owner(o.nodeId) === note.rootNodeId).map(o => [o.nodeId, before.geometry(o.nodeId)]));
        if (operation.notesMasterId === null)
            delete note.notesMasterId;
        else
            note.notesMasterId = operation.notesMasterId;
        const after = presentationTransforms(p, html);
        for (const [id, g] of geometry)
            try {
                after.geometry(id);
            }
            catch {
                edit.attributes(id, geometryAttributes(g));
            }
        return;
    }
    if (operation.op === "setPresentationPlaceholder") {
        const object = p.objects.find(o => o.nodeId === operation.nodeId) ?? invalid("The placeholder object is missing."), html = ctx.parse(), surface = html.peekNode(operation.nodeId)!.parentNodeId;
        if (![...p.slides, ...p.masters, ...p.layouts, ...p.notesMasters, ...p.stories.filter(s => s.kind === "notes")].some(s => s.rootNodeId === surface))
            invalid("Only surface objects can be placeholders.");
        if (operation.placeholder)
            object.placeholder = { ...operation.placeholder };
        else {
            const g = presentationTransforms(p, html).geometry(object.nodeId);
            delete object.placeholder;
            edit.attributes(object.nodeId, geometryAttributes(g));
        }
        return;
    }
    if (operation.op === "resetPresentationOverrides") {
        const object = p.objects.find(o => o.nodeId === operation.nodeId) ?? invalid("The object is missing.");
        if (!object.placeholder)
            invalid("Resetting local overrides requires a placeholder with inherited defaults.");
        if (operation.geometry)
            edit.attributes(object.nodeId, Object.fromEntries(Object.keys(geometryAttributes(presentationTransforms(p, ctx.parse()).geometry(object.nodeId))).map(name => [name, undefined])));
        if (operation.appearance)
            edit.attributes(object.nodeId, { style: undefined, class: undefined });
        if (operation.text && (object.kind === "text" || object.kind === "geometry"))
            delete object.text;
        presentationTransforms(p, ctx.parse()).geometry(object.nodeId);
        return;
    }
    if (operation.op === "insertPresentationDefinition") {
        const kind = operation.kind, id = ctx.allocate(), html = ctx.parse(), root = html.nodeIds().find(id => html.peekNode(id)?.attributes["data-office-kind"] === "presentation")!;
        if (kind === "theme" && (operation.themeId !== undefined || operation.masterId !== undefined || operation.name !== undefined) || kind !== "layout" && (operation.masterId !== undefined || operation.name !== undefined) || kind === "layout" && operation.themeId !== undefined)
            invalid("Definition creation includes properties outside its kind.");
        if (kind === "theme") {
            const original = p.themes.find(t => t.themeId === (operation.copyFromId ?? p.themes[0]!.themeId)) ?? invalid("The source theme is missing."), cssClass = allocatePresentationClass(ctx, "theme-");
            ctx.styles.cloneClass(original.cssClass, cssClass);
            p.themes.push({ themeId: id, cssClass });
            const retained = p.nativePreview?.filter(item => ["theme-defaults", "theme-format"].includes(item.feature) && item.ownerId === original.themeId);
            if (retained?.length) p.nativePreview!.push(...retained.map(item => ({ ...item, ownerId: id })));
            return;
        }
        const originals = kind === "master" ? p.masters : kind === "layout" ? p.layouts : p.notesMasters;
        const original = operation.copyFromId ? originals.find(d => "masterId" in d && kind === "master" ? d.masterId === operation.copyFromId : "layoutId" in d && kind === "layout" ? d.layoutId === operation.copyFromId : "notesMasterId" in d && d.notesMasterId === operation.copyFromId) : undefined;
        if (operation.copyFromId && !original)
            invalid("The source definition is missing.");
        const copied = original ? edit.cloneSource(original.rootNodeId, root) : edit.insertSource(root, '<section data-office-kind="' + kind + '" data-office-story="' + kind + '"></section>'), rootNodeId = copied.rootNodeIds[0]!;
        if (kind === "master")
            p.masters.push({ ...original, masterId: id, rootNodeId, themeId: operation.themeId ?? (original && "themeId" in original ? original.themeId : p.themes[0]!.themeId) });
        else if (kind === "layout")
            p.layouts.push({ ...original, layoutId: id, rootNodeId, masterId: operation.masterId ?? (original && "masterId" in original ? original.masterId : p.masters[0]!.masterId), ...(operation.name !== undefined ? { name: operation.name } : {}) });
        else
            p.notesMasters.push({ ...original, notesMasterId: id, rootNodeId, themeId: operation.themeId ?? (original && "themeId" in original ? original.themeId : p.themes[0]!.themeId) });
        if (original)
            registerCopies(ctx, copied.identities, false);
        return;
    }
    if (operation.op === "updatePresentationDefinition") {
        const { kind, definitionId, changes } = operation;
        if (kind === "master") {
            if (changes.masterId !== undefined || changes.name !== undefined || changes.showMasterShapes !== undefined)
                invalid("A master accepts a theme reference.");
            const d = p.masters.find(d => d.masterId === definitionId) ?? invalid("The master is missing.");
            if (changes.themeId !== undefined)
                d.themeId = changes.themeId;
        }
        else if (kind === "layout") {
            if (changes.themeId !== undefined)
                invalid("A layout obtains its theme from its master.");
            const d = p.layouts.find(d => d.layoutId === definitionId) ?? invalid("The layout is missing.");
            Object.assign(d, changes);
        }
        else {
            if (changes.masterId !== undefined || changes.name !== undefined || changes.showMasterShapes !== undefined)
                invalid("A notes master accepts a theme reference.");
            const d = p.notesMasters.find(d => d.notesMasterId === definitionId) ?? invalid("The notes master is missing.");
            if (changes.themeId !== undefined)
                d.themeId = changes.themeId;
        }
        return;
    }
    if (operation.op === "removePresentationDefinition") {
        const { kind, definitionId: id, replacementId: replacement } = operation;
        if (replacement === id)
            invalid("A removed definition cannot replace itself.");
        if (kind === "theme") {
            const theme = p.themes.find(t => t.themeId === id) ?? invalid("The theme is missing.");
            for (const d of [...p.masters, ...p.notesMasters])
                if (d.themeId === id) {
                    if (!replacement)
                        invalid("A referenced theme requires a replacement.");
                    d.themeId = replacement;
                }
            const target = p.themes.find(t => t.themeId === replacement)?.cssClass;
            const html = ctx.parse();
            for (const nodeId of html.nodeIds()) {
                const value = html.peekNode(nodeId)!.attributes.class, classes = value?.split(/\s+/);
                if (classes?.includes(theme.cssClass))
                    edit.attributes(nodeId, { class: classes.flatMap(name => name === theme.cssClass ? target ? [target] : [] : [name]).join(" ") || undefined });
            }
            p.themes = p.themes.filter(t => t !== theme);
            if (p.nativePreview) p.nativePreview = p.nativePreview.filter(item => !["theme-defaults", "theme-format"].includes(item.feature) || item.ownerId !== id);
            return;
        }
        let removedRoot: string;
        if (kind === "master") {
            const d = p.masters.find(d => d.masterId === id) ?? invalid("The master is missing.");
            for (const l of p.layouts)
                if (l.masterId === id) {
                    if (!replacement)
                        invalid("A referenced master requires a replacement.");
                    l.masterId = replacement;
                }
            removedRoot = d.rootNodeId;
            p.masters = p.masters.filter(m => m !== d);
        }
        else if (kind === "layout") {
            const d = p.layouts.find(d => d.layoutId === id) ?? invalid("The layout is missing.");
            for (const s of p.slides)
                if (s.layoutId === id) {
                    if (!replacement)
                        invalid("A referenced layout requires a replacement.");
                    switchLayout(ctx, s.slideId, replacement);
                }
            removedRoot = d.rootNodeId;
            p.layouts = p.layouts.filter(l => l !== d);
        }
        else {
            const d = p.notesMasters.find(d => d.notesMasterId === id) ?? invalid("The notes master is missing.");
            for (const s of p.stories)
                if (s.kind === "notes" && s.notesMasterId === id) {
                    if (!replacement)
                        invalid("A referenced notes master requires a replacement.");
                    s.notesMasterId = replacement;
                }
            removedRoot = d.rootNodeId;
            p.notesMasters = p.notesMasters.filter(n => n !== d);
        }
        edit.removeSource(removedRoot);
        return;
    }
    invalid("Unsupported presentation definition operation.");
}
