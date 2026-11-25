import { DocumentError, openDocument } from "@doffi/core";
import type { ContentDocumentManifest, DocumentPackage, PresentationDocumentMetadata } from "@doffi/core";
import type { DocumentZipOptions } from "./document-zip.js";
import { openOpc } from "./ooxml/package.js";
import type { OpcPackage } from "./ooxml/package.js";
import { A, attr, check, checkAttributes, child, elements, fail, P, R, text, unsupported } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import { escapeHtml } from "./ooxml/styles.js";
import { bool, Context, encode, integer, points, serialize, unsigned, value } from "./pptx/context.js";
import type { Html, Surface } from "./pptx/context.js";
import { compatible, Drawings } from "./pptx/drawing.js";
import { importTheme } from "./pptx/styles.js";
import { importTextStyle } from "./pptx/text.js";
import { Comments } from "./pptx/comments.js";

export interface PptxImportOptions extends DocumentZipOptions { readonly documentId: string; readonly mode?: "strict" | "compatible" }

function activeTiming(source: XmlNode): XmlNode | undefined {
  const timing = child(source, "timing", P);
  if (!timing) return undefined;
  const content = (node: XmlNode) => node.content.filter(item => typeof item !== "string" || item.trim());
  // Office can write a time-root container with no playable timeline.
  let node = timing;
  for (const name of ["tnLst", "par", "cTn"]) {
    const items = content(node), next = items[0];
    if (items.length !== 1 || !next || typeof next === "string" || next.uri !== P || next.name !== name ||
        [...node.attrs.keys()].some(key => !key.startsWith("http://www.w3.org/2000/xmlns/ "))) return timing;
    node = next;
  }
  return !content(node).length && value(node, "nodeType") === "tmRoot" && value(node, "dur") === "indefinite" &&
    value(node, "restart") === "never" && /^\d+$/.test(value(node, "id") ?? "") &&
    [...node.attrs.keys()].every(key => [" id", " dur", " restart", " nodeType"].includes(key) ||
      key.startsWith("http://www.w3.org/2000/xmlns/ ")) ? undefined : timing;
}

class Importer {
  readonly context: Context;
  readonly drawings: Drawings;
  readonly roots: Html[] = [];
  constructor(readonly main: string, readonly options: PptxImportOptions, opc: OpcPackage) {
    this.context = new Context(opc, options.mode === "compatible"); this.context.bytes(options.documentId.length * 6); this.drawings = new Drawings(this.context);
  }
  createSurface(path: string, kind: "master" | "layout" | "slide" | "notes-master" | "notes", id: string, themeId: string, source: XmlNode): Surface {
    const root = this.context.make("section", { "data-office-kind": kind, "data-office-story": kind, class: themeId });
    const transition = child(source, "transition", P);
    // LibreOffice emits a container with no effect on some static surfaces.
    if (transition && (transition.content.some(item => typeof item !== "string" || item.trim()) ||
        [...transition.attrs].some(([name, v]) => name === " advClick" ? !["1", "true"].includes(v) : name !== " spd" && !name.startsWith("http://www.w3.org/2000/xmlns/ ")))) {
      const effect = child(transition, "wipe", P);
      if (!this.context.compatiblePreview || !effect) unsupported(transition, "Slide transitions are outside the static PPT scope.");
      check(transition, ["wipe"], P);
      check(effect, [], P);
      checkAttributes(transition, ["spd", "advClick"]);
      checkAttributes(effect, ["dir"]);
      if (!["l", "r", "u", "d"].includes(value(effect, "dir") ?? "l")) fail(effect, "A wipe direction is invalid.");
      bool(transition, "advClick", true);
      this.context.retainPreview("transition", root.attrs["data-node-id"]!, transition, "Slide transitions require compatible preview.");
    }
    const surface = { id, root, source, themeId }; this.context.surfaces.set(path, surface); this.roots.push(root); return surface;
  }
  readSurface(path: string, type: string, name: string, allowed: string[]): XmlNode {
    const source = this.context.part(path, "presentationml." + type, name, P);
    const selected = { ...source, content: compatible(source) };
    check(selected, ["cSld", "clrMap", "clrMapOvr", "extLst", ...allowed], P);
    return selected;
  }
  handoutParts(list: XmlNode | undefined): { master: XmlNode; theme?: XmlNode } | undefined {
    if (!list) return undefined;
    if (!this.context.compatiblePreview) unsupported(list, "Handout masters require compatible preview.");
    check(list, ["handoutMasterId"], P);
    const entry = child(list, "handoutMasterId", P);
    if (!entry) return undefined;
    check(entry, [], P);
    const { context } = this, path = context.relation(this.main, attr(entry, "id", R), "handoutMaster").target;
    const master = context.part(path, "presentationml.handoutMaster", "handoutMaster", P);
    if ([...context.opc.relationships(path).values()].some(relation => relation.type !== R + "/theme"))
      unsupported(master, "A preview-only handout master may depend only on its internal theme.");
    const themePath = context.related(path, "theme", false);
    const theme = themePath ? context.part(themePath, "theme", "theme", A) : undefined;
    if (theme && context.opc.relationships(theme.part).size) unsupported(theme, "A preview-only handout theme cannot have package relationships.");
    return { master, ...(theme ? { theme } : {}) };
  }
  import(): DocumentPackage {
    const { context } = this, { opc, metadata } = context;
    const source = context.part(this.main, "presentationml.presentation.main", "presentation", P);
    if (value(source, "firstSlideNum") !== undefined) {
      const first = integer(source, "firstSlideNum");
      if (first < 0 || first > 9999) fail(source, "The first slide number is outside the Office range.");
      metadata.firstSlideNumber = first;
    }
    check(source, ["sldMasterIdLst", "notesMasterIdLst", "handoutMasterIdLst", "sldIdLst", "sldSz", "notesSz", "defaultTextStyle", "extLst", "embeddedFontLst", "custDataLst", "photoAlbum", "custShowLst", "kinsoku"], P);
    for (const name of ["embeddedFontLst", "photoAlbum", "custShowLst"]) {
      const feature = child(source, name, P);
      if (feature) unsupported(feature, "This presentation-level feature is outside the reading scope.");
    }
    const handout = this.handoutParts(child(source, "handoutMasterIdLst", P));
    const size = child(source, "sldSz", P);
    if (!size) fail(source, "The presentation slide size is missing.");
    metadata.size = { unit: "pt", width: points(size, "cx"), height: points(size, "cy") };
    if (metadata.size.width <= 0 || metadata.size.height <= 0) fail(size, "Slide dimensions must be positive.");
    const notesSize = child(source, "notesSz", P);
    if (notesSize) metadata.notesSize = { unit: "pt", width: points(notesSize, "cx"), height: points(notesSize, "cy") };
    const defaultStyle = child(source, "defaultTextStyle", P);
    importTextStyle(context, defaultStyle, "presentation-default");
    const seenSlides = new Set<string>(), nativeSlides = new Set<number>();
    for (const entry of elements(child(source, "sldIdLst", P), "sldId", P)) {
      const native = integer(entry, "id"), path = context.relation(this.main, attr(entry, "id", R), "slide").target;
      if (native < 256 || native > 2147483647 || nativeSlides.has(native) || seenSlides.has(path)) fail(entry, "Slide identities or references are invalid or duplicate.");
      seenSlides.add(path); nativeSlides.add(native); context.slidePaths.set(path, "slide-" + native);
    }
    if (!seenSlides.size) fail(source, "A presentation must contain at least one slide.");
    const seenMasters = new Set<string>(), nativeMasters = new Set<number>(), nativeLayouts = new Set<number>();
    for (const entry of elements(child(source, "sldMasterIdLst", P), "sldMasterId", P)) {
      const path = context.relation(this.main, attr(entry, "id", R), "slideMaster").target;
      const nativeId = unsigned(entry, "id", 2147483648);
      if (seenMasters.has(path) || nativeMasters.has(nativeId)) fail(entry, "Slide master references or numeric identities are duplicated.");
      nativeMasters.add(nativeId);
      seenMasters.add(path);
      const masterSource = this.readSurface(path, "slideMaster", "sldMaster", ["sldLayoutIdLst", "txStyles", "hf", "timing", "transition"]);
      const masterTiming = activeTiming(masterSource);
      if (masterTiming) unsupported(masterTiming, "Master animation is not supported.");
      const themeId = importTheme(context, context.related(path, "theme")!), id = context.id("master");
      const master = this.createSurface(path, "master", id, themeId, masterSource);
      const definition: typeof metadata.masters[number] = { masterId: id, rootNodeId: master.root.attrs["data-node-id"]!, themeId };
      const styles = child(masterSource, "txStyles", P);
      for (const [name, key] of [["titleStyle", "titleStyle"], ["bodyStyle", "bodyStyle"], ["otherStyle", "otherStyle"]] as const) {
        const nativeStyle = child(styles, name, P);
        if (nativeStyle) { const cssClass = id + "-" + name; importTextStyle(context, nativeStyle, cssClass); definition[key] = cssClass; }
      }
      metadata.masters.push(definition);
      this.drawings.surface(master);
      const layoutPaths = new Set<string>();
      for (const layoutEntry of elements(child(masterSource, "sldLayoutIdLst", P), "sldLayoutId", P)) {
        const layoutPath = context.relation(path, attr(layoutEntry, "id", R), "slideLayout").target;
        const nativeId = unsigned(layoutEntry, "id", 2147483648);
        if (layoutPaths.has(layoutPath) || context.surfaces.has(layoutPath) || nativeLayouts.has(nativeId)) fail(layoutEntry, "A layout has multiple native owners or duplicate identities.");
        nativeLayouts.add(nativeId);
        layoutPaths.add(layoutPath);
        if (context.related(layoutPath, "slideMaster") !== path) fail(layoutEntry, "A layout's master relation disagrees with its owner.");
        const layoutSource = this.readSurface(layoutPath, "slideLayout", "sldLayout", ["hf", "timing", "transition"]);
        const layoutTiming = activeTiming(layoutSource);
        if (layoutTiming) unsupported(layoutTiming, "Layout animation is not supported.");
        const layoutId = context.id("layout"), layout = this.createSurface(layoutPath, "layout", layoutId, themeId, layoutSource);
        const name = value(child(layoutSource, "cSld", P), "name");
        metadata.layouts.push({ layoutId, rootNodeId: layout.root.attrs["data-node-id"]!, masterId: id, ...(name !== undefined ? { name } : {}), showMasterShapes: bool(layoutSource, "showMasterSp", true) });
        this.drawings.surface(layout);
      }
    }
    if (!metadata.masters.length || !metadata.layouts.length) fail(source, "A presentation requires master and layout definitions.");
    for (const entry of elements(child(source, "notesMasterIdLst", P), "notesMasterId", P)) {
      const path = context.relation(this.main, attr(entry, "id", R), "notesMaster").target;
      if (context.surfaces.has(path)) fail(entry, "Notes master references are duplicated or have the wrong kind.");
      const native = this.readSurface(path, "notesMaster", "notesMaster", ["hf", "notesStyle"]);
      const themeId = importTheme(context, context.related(path, "theme")!), id = context.id("notes-master");
      const surface = this.createSurface(path, "notes-master", id, themeId, native), textStyle = id + "-text-style";
      importTextStyle(context, child(native, "notesStyle", P), textStyle);
      metadata.notesMasters.push({ notesMasterId: id, rootNodeId: surface.root.attrs["data-node-id"]!, themeId, ...(child(native, "notesStyle", P) ? { textStyle } : {}) });
      this.drawings.surface(surface);
    }
    for (const [path, id] of context.slidePaths) {
      const layout = context.surfaces.get(context.related(path, "slideLayout")!);
      if (!layout || !metadata.layouts.some(item => item.layoutId === layout.id)) fail(source, "A slide refers to a layout outside its declared masters.");
      const slideSource = this.readSurface(path, "slide", "sld", ["transition", "timing"]);
      const slide = this.createSurface(path, "slide", id, layout.themeId, slideSource), name = value(child(slideSource, "cSld", P), "name");
      slide.layoutId = layout.id;
      metadata.slides.push({ slideId: id, rootNodeId: slide.root.attrs["data-node-id"]!, layoutId: layout.id, ...(name !== undefined ? { name } : {}), hidden: !bool(slideSource, "show", true), showMasterShapes: bool(slideSource, "showMasterSp", true) });
      this.drawings.surface(slide);
      const timing = activeTiming(slideSource);
      if (timing) this.drawings.retainAnimation(slide, timing);
      const notesPath = context.related(path, "notesSlide", false);
      if (notesPath) {
        if (context.surfaces.has(notesPath) || context.related(notesPath, "slide") !== path) fail(slideSource, "A notes part has duplicate or conflicting slide ownership.");
        const native = this.readSurface(notesPath, "notesSlide", "notes", []);
        const masterPath = context.related(notesPath, "notesMaster", false), master = masterPath ? context.surfaces.get(masterPath) : undefined;
        if (masterPath && (!master || !metadata.notesMasters.some(m => m.notesMasterId === master.id))) fail(native, "A notes master is outside the declared definitions.");
        const notesId = context.id("notes"), notes = this.createSurface(notesPath, "notes", notesId, master?.themeId ?? slide.themeId, native);
        notes.root.attrs["data-office-show-master-shapes"] = String(bool(native, "showMasterSp", true));
        this.drawings.surface(notes);
        const body = metadata.objects.filter(o => o.placeholder?.type === "body" && notes.root.children.some(n => typeof n !== "string" && n.attrs["data-node-id"] === o.nodeId));
        if (body.length > 1) unsupported(native, "Multiple notes text bodies are not supported.");
        metadata.stories.push({ storyId: notesId, rootNodeId: notes.root.attrs["data-node-id"]!, kind: "notes", slideId: id, ...(master ? { notesMasterId: master.id } : {}), ...(body[0] ? { bodyNodeId: body[0].nodeId } : {}) });
      }
    }
    this.drawings.finish();
    const comments = new Comments(context, this.main, this.options.documentId, this.roots);
    for (const [path, id] of context.slidePaths) {
      const slide = context.surfaces.get(path)!;
      if (slide.id !== id) fail(slide.source, "A comment slide identity is inconsistent.");
      comments.read(slide, path);
    }
    const corePath = opc.related("", "metadata/core-properties");
    const core = corePath ? opc.read(corePath) : undefined;
    const title = text(child(core, "title", "http://purl.org/dc/elements/1.1/")) || "未命名演示文稿";
    const locale = text(child(core, "language", "http://purl.org/dc/elements/1.1/")) || "zh-CN";
    const documentMetadata: PresentationDocumentMetadata = { title, locale, presentation: metadata };
    const root = context.make("article", { "data-office-kind": "presentation", class: "presentation-default" }, this.roots);
    if (handout) {
      const owner = root.attrs["data-node-id"]!;
      context.retainPreview("handout-master", owner, handout.master, "Handout masters require compatible preview.");
      if (handout.theme) context.retainPreview("handout-theme", owner, handout.theme, "Handout themes require compatible preview.");
    }
    const html = '<!doctype html><html lang="' + escapeHtml(locale) + '"><head><meta charset="utf-8"><title>' + escapeHtml(title) + '</title><link rel="stylesheet" href="styles.css"></head><body>' + serialize(root) + '</body></html>';
    const manifest: ContentDocumentManifest = { format: "doffi", formatVersion: 1, documentId: this.options.documentId, kind: "presentation", revision: 0,
      entrypoint: "content.html", styles: ["styles.css"], metadata: "metadata.json", annotations: "annotations.json", assets: context.assets };
    context.files.set("manifest.json", encode(JSON.stringify(manifest))); context.files.set("content.html", encode(html));
    context.files.set("metadata.json", encode(JSON.stringify(documentMetadata))); context.files.set("annotations.json", encode(JSON.stringify(context.annotations)));
    context.files.set("styles.css", encode("p{margin:0}" + context.css.join("\n")));
    return openDocument(context.files, this.options).serialize();
  }
}

/** Map authorized PPTX bytes into the one persistent HTML/CSS source. */
export function importPptx(bytes: Uint8Array, options: PptxImportOptions): DocumentPackage {
  if (!options || typeof options.documentId !== "string" || !options.documentId) throw new DocumentError("InvalidDocument", "PPTX import requires a document identity supplied by its host.");
  if (options.mode !== undefined && options.mode !== "strict" && options.mode !== "compatible") throw new DocumentError("InvalidOperation", "Unknown PPTX import mode.");
  const opc = openOpc(bytes, options), main = new Context(opc).related("", "officeDocument", false);
  if (!main || opc.contentType(main) !== "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml") throw new DocumentError("UnsupportedFormat", "The input is not a supported PPTX presentation.");
  try { return new Importer(main, options, opc).import(); }
  catch (error) { if (error instanceof DocumentError) throw error; throw new DocumentError("InvalidDocument", "PPTX could not be mapped to the supported presentation model."); }
}
