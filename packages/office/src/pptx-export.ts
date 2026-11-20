import { DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, openDocument } from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";
import { writeOpcZip } from "./document-zip.js";
import { attributes, declaration, empty, escapeXml } from "./export/xml.js";
import { xml } from "./ooxml/xml.js";
import { PresentationChartWriter } from "./pptx-export/chart.js";
import { PresentationCommentWriter } from "./pptx-export/comments.js";
import { mainPart, nativeNamespaces, pmlType, PresentationExportContext, relativePart } from "./pptx-export/context.js";
import { PresentationDrawingWriter } from "./pptx-export/drawing.js";
import { PresentationTextWriter } from "./pptx-export/text.js";
import { background, colorMap, textStyle, writeThemes } from "./pptx-export/themes.js";
import type { PptxExportOptions, PptxExportResult } from "./pptx-export/types.js";
import { emu } from "./pptx-export/values.js";
export type { PptxExportOptions, PptxExportReport, PptxExportResult, PptxAnnotationLoss, UnexportedPptxAnnotation } from "./pptx-export/types.js";
/** Compile the current validated HTML/CSS package into native editable OOXML.
 * Original Office bytes, browser DOM, network and file system are never read. */
export function exportPptx(document: ContentDocumentModel, options: PptxExportOptions = {}): PptxExportResult {
    if (options.commentFormat !== undefined && !["preserve", "modern", "classic"].includes(options.commentFormat))
        throw new DocumentError("InvalidOperation", "Unknown native PPTX comment format.");
    const model = openDocument(document.serialize(), { ...options, kind: "presentation" }), limits = { ...DEFAULT_DOCUMENT_READ_LIMITS, ...options.limits };
    if (model.snapshot().manifest.kind !== "presentation")
        throw new DocumentError("UnsupportedFormat", "PPTX export requires a presentation package.");
    const retained = model.snapshot().metadata.presentation?.nativePreview;
    if (retained?.length)
        throw new DocumentError("UnsupportedFeature", "Native PPTX export would lose preserved preview-only definitions.", retained.map(({ part, feature }) => ({ code: "UnsupportedFeature", path: [part, feature], message: "This native fragment is retained in the document package." })));
    const ctx = new PresentationExportContext(model, limits, options), text = new PresentationTextWriter(ctx), charts = new PresentationChartWriter(ctx, text), drawings = new PresentationDrawingWriter(ctx, text, (object, node, surface) => charts.write(object, node, surface)), comments = new PresentationCommentWriter(ctx, text);
    const content = new Map(ctx.surfaces.map(surface => { const fill = background(ctx, surface); ctx.checkGenerated(fill); return [surface.part, fill + drawings.tree(surface)] as const; }));
    comments.write();
    ctx.checkGenerated("");
    const slides = ctx.surfaces.filter(surface => surface.kind === "slide"), masters = ctx.surfaces.filter(surface => surface.kind === "master"), notesMasters = ctx.surfaces.filter(surface => surface.kind === "notes-master");
    for (const surface of ctx.surfaces) {
        const slide = ctx.p.slides.find(slide => slide.slideId === surface.definitionId), layout = surface.kind === "layout" ? surface.layout : undefined, story = ctx.p.stories.find(story => story.storyId === surface.definitionId), tag = ({ slide: "sld", master: "sldMaster", layout: "sldLayout", "notes-master": "notesMaster", notes: "notes" }[surface.kind]), attrs: Record<string, string | number | undefined> = {};
        if (slide) {
            attrs.show = slide.hidden ? 0 : 1;
            attrs.showMasterSp = slide.showMasterShapes === false ? 0 : 1;
        }
        if (layout) {
            attrs.showMasterSp = layout.showMasterShapes === false ? 0 : 1;
            attrs.type = surface.node.attrs["data-office-layout-type"] ?? "cust";
            attrs.preserve = 1;
        }
        if (surface.kind === "notes")
            attrs.showMasterSp = surface.node.attrs["data-office-show-master-shapes"] === "false" ? 0 : 1;
        const part = ctx.opc.part(surface.part, pmlType(({ slide: "slide", master: "slideMaster", layout: "slideLayout", "notes-master": "notesMaster", notes: "notesSlide" }[surface.kind])));
        part.add(declaration + "<p:" + tag + nativeNamespaces + attributes(attrs) + "><p:cSld" + attributes({ name: slide?.name ?? layout?.name ?? surface.node.attrs["data-office-name"] }) + ">" + content.get(surface.part)! + "</p:cSld>" + colorMap(ctx, surface, surface.kind === "master" || surface.kind === "notes-master"));
        if (surface.kind === "master") {
            ctx.opc.relate(surface.part, "theme", relativePart(surface.part, ctx.themePart(surface.themeId)));
            part.add("<p:sldLayoutIdLst>");
            for (const child of ctx.surfaces.filter(child => child.kind === "layout" && child.master?.masterId === surface.definitionId))
                part.add(empty("p:sldLayoutId", { id: 2147483649 + ctx.p.layouts.findIndex(layout => layout.layoutId === child.definitionId), "r:id": ctx.opc.relate(surface.part, "slideLayout", relativePart(surface.part, child.part)) }));
            part.add("</p:sldLayoutIdLst>");
        }
        else if (surface.kind === "layout")
            ctx.opc.relate(surface.part, "slideMaster", relativePart(surface.part, masters.find(master => master.definitionId === surface.master!.masterId)!.part));
        else if (surface.kind === "slide") {
            ctx.opc.relate(surface.part, "slideLayout", relativePart(surface.part, ctx.surfaces.find(layout => layout.kind === "layout" && layout.definitionId === slide!.layoutId)!.part));
            const notes = ctx.surfaces.find(notes => notes.kind === "notes" && ctx.p.stories.some(story => story.storyId === notes.definitionId && story.kind === "notes" && story.slideId === slide!.slideId));
            if (notes)
                ctx.opc.relate(surface.part, "notesSlide", relativePart(surface.part, notes.part));
        }
        else if (surface.kind === "notes-master")
            ctx.opc.relate(surface.part, "theme", relativePart(surface.part, ctx.themePart(surface.themeId)));
        else if (story?.kind === "notes") {
            const owner = slides.find(slide => slide.definitionId === story.slideId)!;
            ctx.opc.relate(surface.part, "slide", relativePart(surface.part, owner.part));
            if (story.notesMasterId) {
                const master = notesMasters.find(master => master.definitionId === story.notesMasterId)!;
                ctx.opc.relate(surface.part, "notesMaster", relativePart(surface.part, master.part));
            }
        }
        if (["master", "layout", "notes-master"].includes(surface.kind)) {
            const hf: Record<string, number> = {};
            for (const [key, name] of [["hdr", "header"], ["ftr", "footer"], ["dt", "date"], ["sldNum", "slide-number"]])
                if (surface.node.attrs["data-office-show-" + name] !== undefined)
                    hf[key!] = surface.node.attrs["data-office-show-" + name] === "true" ? 1 : 0;
            if (Object.keys(hf).length)
                part.add(empty("p:hf", hf));
        }
        if (surface.kind === "master")
            part.add("<p:txStyles>" + textStyle(ctx, surface, surface.master!.titleStyle, "p:titleStyle") + textStyle(ctx, surface, surface.master!.bodyStyle, "p:bodyStyle") + textStyle(ctx, surface, surface.master!.otherStyle, "p:otherStyle") + "</p:txStyles>");
        if (surface.kind === "notes-master")
            part.add(textStyle(ctx, surface, surface.notesMaster!.textStyle, "p:notesStyle"));
        part.add((comments.extensions.get(surface.part) ?? "") + "</p:" + tag + ">");
        part.finish();
    }
    const main = ctx.opc.part(mainPart, pmlType("presentation.main"));
    main.add(declaration + "<p:presentation" + nativeNamespaces + "><p:sldMasterIdLst>");
    for (const [index, master] of masters.entries())
        main.add(empty("p:sldMasterId", { id: 2147483648 + index, "r:id": ctx.opc.relate(mainPart, "slideMaster", relativePart(mainPart, master.part)) }));
    main.add("</p:sldMasterIdLst>");
    if (notesMasters.length) {
        main.add("<p:notesMasterIdLst>");
        for (const master of notesMasters)
            main.add(empty("p:notesMasterId", { "r:id": ctx.opc.relate(mainPart, "notesMaster", relativePart(mainPart, master.part)) }));
        main.add("</p:notesMasterIdLst>");
    }
    main.add("<p:sldIdLst>");
    for (const slide of slides)
        main.add(empty("p:sldId", { id: slide.nativeSlideId!, "r:id": ctx.opc.relate(mainPart, "slide", relativePart(mainPart, slide.part)) }));
    main.add("</p:sldIdLst>" + empty("p:sldSz", { cx: emu(ctx.p.size.width, 914400, 51206400), cy: emu(ctx.p.size.height, 914400, 51206400) }) + empty("p:notesSz", { cx: emu((ctx.p.notesSize ?? ctx.p.size).width, 1), cy: emu((ctx.p.notesSize ?? ctx.p.size).height, 1) }) + "</p:presentation>");
    main.finish();
    writeThemes(ctx);
    const core = ctx.opc.part("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
    core.add(declaration + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>' + escapeXml(ctx.snapshot.metadata.title ?? "") + "</dc:title><dc:language>" + escapeXml(ctx.snapshot.metadata.locale ?? "") + "</dc:language></cp:coreProperties>");
    core.finish();
    ctx.opc.relate("", "officeDocument", mainPart);
    ctx.opc.relate("", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
    ctx.opc.finish();
    for (const [path, bytes] of ctx.opc.parts.files)
        if (path.endsWith(".xml") || path.endsWith(".rels"))
            xml(bytes, path, limits, ctx.xmlBudget);
    return Object.freeze({ bytes: writeOpcZip(ctx.opc.parts.files, options), report: Object.freeze({ documentId: ctx.snapshot.manifest.documentId, revision: ctx.snapshot.manifest.revision, exportedAnnotationIds: Object.freeze(ctx.exportedAnnotationIds), unexportedAnnotations: Object.freeze(ctx.unexportedAnnotations.map(item => Object.freeze(item))), annotationLosses: Object.freeze(ctx.annotationLosses.map(item => Object.freeze(item))) }) });
}
