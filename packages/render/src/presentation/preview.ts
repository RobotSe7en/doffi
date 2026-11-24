import { DocumentError } from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";
import type { DocumentPreview, DocumentPreviewOptions } from "../preview.js";
import { renderLimits } from "../contracts.js";
import type { PresentationRenderReport, FontResolution } from "../contracts.js";
import { collectDocumentFontNames, hasUnloadedWordWebFonts, loadWordWebFonts, resolvePreviewFonts, resolvePreviewFontsInScope } from "../fonts.js";
import type { HostedWordFonts } from "../fonts.js";
import { FontProbeCache } from "../font-probe-cache.js";
import { preparePreviewPackage, releasePreviewUrls } from "../preview-package.js";
import { SourceProjection } from "../projection.js";
import { copyPreviewSelection } from "../clipboard.js";
import { createPrintDocument } from "../print.js";
import type { PreviewSource } from "../source-index.js";
import { composePresentation } from "./canvas.js";
import type { PresentationCanvas } from "./canvas.js";

export function mountPresentationPreview(container: HTMLElement, input: ContentDocumentModel, options: DocumentPreviewOptions): DocumentPreview & { source(): PreviewSource | undefined; showCanvasSlide(slideId: string, active?: boolean): boolean } {
  if (options.layout === "continuous") throw new DocumentError("InvalidOperation", "Presentations use fixed pages layout.");
  const limits = renderLimits(options.limits), host = container.ownerDocument, frame = host.createElement("iframe");
  const fontProbes = new FontProbeCache();
  const prepare = (model: ContentDocumentModel, previous?: ReturnType<typeof preparePreviewPackage>): ReturnType<typeof preparePreviewPackage> => {
    const snapshot = model.snapshot(); if (snapshot.manifest.kind !== "presentation") throw new DocumentError("UnsupportedFormat", "Use a new preview when changing the document kind.");
    const result = preparePreviewPackage(model, previous);
    if (result.model.snapshot().manifest.kind !== "presentation") { releasePreviewUrls(result, previous); throw new DocumentError("UnsupportedFormat", "A presentation package is required."); }
    return result;
  };
  let prepared = prepare(input), canvas: PresentationCanvas | undefined, projection: SourceProjection | undefined, fonts: readonly FontResolution[] = [], hostedFonts: HostedWordFonts | undefined;
  let disposed = false, loaded = false,composing=false, generation = 0, observer: ResizeObserver | undefined, scheduled: number | undefined;
  let editingSlide = options.presentationEditingSlideId;
  const fontSelector = "[data-node-id],[data-office-script],.sdk-ppt-marker";
  function addScopeFonts(document: Document, scopes: readonly HTMLElement[]): void {
    const results = new Map(fonts.map(font => [font.requestedFamily + "\0" + font.selectedFamily, font]));
    for (const scope of scopes) for (const font of resolvePreviewFontsInScope(document, scope, fontSelector, fontProbes, hostedFonts)) {
      const key = font.requestedFamily + "\0" + font.selectedFamily, previous = results.get(key);
      results.set(key, previous ? Object.freeze({
        ...font,
        glyphAuditComplete: previous.glyphAuditComplete && font.glyphAuditComplete,
        missingCharacters: Object.freeze([...new Set([...previous.missingCharacters, ...font.missingCharacters])]),
      }) : font);
    }
    fonts = Object.freeze([...results.values()]);
  }
  let report: PresentationRenderReport = { kind:"presentation",slides:[], documentId: prepared.model.snapshot().manifest.documentId, revision: prepared.model.snapshot().manifest.revision, layout: "pages", pages: [], nodePages: {}, fields: [], fonts: [], diagnostics: [] };
  let resolve: () => void = () => {}, reject: (reason: unknown) => void = () => {};
  const ready = new Promise<void>((r, j) => { resolve = r; reject = j; });
  frame.title = options.title ?? "幻灯片只读预览"; frame.setAttribute("sandbox", "allow-same-origin allow-scripts"); frame.style.cssText = "width:100%;height:480px;display:block;border:0";
  function measure(syncHeight = true): void {
    if (disposed || !canvas || composing) return;
    const document = frame.contentDocument!;
    const visibility = document.head.querySelector<HTMLStyleElement>("[data-sdk-slide-visibility]");
    const restoreVisibility = !!visibility && !visibility.disabled;
    if (restoreVisibility) visibility.disabled = true;
    const snapshot = prepared.model.snapshot();
    let measured: ReturnType<PresentationCanvas["measure"]>;
    try {
      measured = canvas.measure();
    } finally {
      if (restoreVisibility) visibility.disabled = false;
    }
    const p = snapshot.metadata.presentation!;
    const byId = new Map(p.slides.map(slide => [slide.slideId, slide]));
    const slides=measured.pages.flatMap(page=>{const slide=byId.get(page.pageId);return slide?[Object.freeze({slideId:slide.slideId,rootNodeId:slide.rootNodeId,index:page.index,hidden:slide.hidden??false,rect:page.rect})]:[];});
    report = Object.freeze({ kind:"presentation",slides:Object.freeze(slides), documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, layout: "pages", ...measured, fields: canvas.fields, fonts });
    if (!syncHeight) return;
    const scrollbar = Math.max(0, frame.clientHeight - document.documentElement.clientHeight);
    const height = Math.ceil(Math.max(200, document.body.scrollHeight) + scrollbar) + "px";
    if (frame.style.height !== height) frame.style.height = height;
  }
  function render(): void {
    observer?.disconnect();
    const document = frame.contentDocument!; canvas = composePresentation(document, prepared.model, limits, options.presentationSurfaceId, editingSlide);
    if (editingSlide) {
      fonts = [];
      addScopeFonts(document, canvas.fontScopes());
    } else fonts = resolvePreviewFonts(document, fontSelector, hostedFonts, fontProbes);
    canvas.finalizeCharts(); canvas.source.reindex();
    projection = new SourceProjection(); const raw = new DOMParser().parseFromString(prepared.model.snapshot().html, "text/html").querySelector<HTMLElement>('[data-office-kind="presentation"]')!;
    projection.capture(canvas.source.root, raw); measure();
    const substitutions = fonts.filter(f => f.substituted), missing = [...new Set(fonts.flatMap(f => [...f.missingCharacters]))];
    if (substitutions.length || missing.length || fonts.some(f => !f.glyphAuditComplete)) {
      const note = document.createElement("p"); note.setAttribute("role", "note"); note.style.cssText = "font:12px sans-serif;color:#694800";
      note.textContent = (substitutions.length ? "字体替代：" + substitutions.map(f => f.requestedFamily + " → " + f.selectedFamily).join("；") + "。" : "") + (missing.length ? "缺少字形：" + missing.map(c => "U+" + c.codePointAt(0)!.toString(16).toUpperCase()).join("、") + "。" : "") + (fonts.some(f => !f.glyphAuditComplete) ? "部分字形未核验。" : ""); document.body.append(note); measure();
    }
    observer = new ResizeObserver(() => {
      if (disposed || scheduled !== undefined) return;
      scheduled = host.defaultView!.requestAnimationFrame(() => { scheduled = undefined; measure(); });
    }); observer.observe(document.body);
  }
  function showCanvasSlide(slideId: string, active = true): boolean {
    if (!editingSlide || options.presentationSurfaceId || !canvas || disposed) return false;
    if (active) editingSlide = slideId;
    const document = frame.contentDocument!;
    const visibility = document.head.querySelector<HTMLStyleElement>("[data-sdk-slide-visibility]");
    const restore = !!visibility && !visibility.disabled;
    if (restore) visibility.disabled = true;
    try {
      const count = canvas.fontScopes().length;
      if (!canvas.ensureSlide(slideId)) return false;
      addScopeFonts(document, canvas.fontScopes().slice(count));
      canvas.finalizeCharts();
      canvas.source.reindex();
      measure(false);
      return true;
    } finally {
      if (restore) visibility.disabled = false;
    }
  }
  const load = (): void => {
    if (disposed) return;
    void frame.contentDocument!.fonts.ready.then(async () => {
      if (options.wordWebFonts?.length && !disposed)
        hostedFonts = await loadWordWebFonts(frame.contentDocument!, options.wordWebFonts, collectDocumentFontNames(prepared.model.snapshot()));
      if (disposed) return;
      render(); loaded = true;
      frame.contentDocument!.addEventListener("compositionstart",()=>{composing=true;},true);
      frame.contentDocument!.addEventListener("compositionend",()=>{composing=false;},true);
      frame.contentDocument!.addEventListener("copy", event => {
        const selection = frame.contentDocument!.getSelection();
        if (selection?.anchorNode && projection?.ignored(selection.anchorNode) || selection?.focusNode && projection?.ignored(selection.focusNode)) return;
        if (canvas && projection) copyPreviewSelection(event, canvas.source.root, projection);
      }); resolve();
    }).catch(error => { disposed = true; generation++; observer?.disconnect(); if (scheduled !== undefined) host.defaultView!.cancelAnimationFrame(scheduled); fontProbes.clear(); canvas = undefined; projection = undefined; frame.remove(); releasePreviewUrls(prepared); reject(error); });
  };
  frame.addEventListener("load", load, { once: true }); frame.srcdoc = "<!doctype html>" + prepared.document.documentElement.outerHTML; container.append(frame);
  return { element: frame, ready, source: () => canvas?.source, fonts: () => fonts, layout: () => report, showCanvasSlide,
    update(model) {
      if (disposed || !loaded) throw new DocumentError("InvalidOperation", "Wait for an active preview before updating.");
      const previous = prepared, incoming = prepare(model, previous); generation++;
      const document = frame.contentDocument!;
      const replace = (value: typeof prepared): void => { prepared = value; document.head.replaceWith(document.importNode(value.document.head, true)); document.body.replaceWith(document.importNode(value.document.body, true)); render(); };
      try { replace(incoming); } catch (error) { releasePreviewUrls(incoming, previous); replace(previous); throw error; }
      releasePreviewUrls(previous, incoming);
      if (options.wordWebFonts?.length) {
        const requested = collectDocumentFontNames(incoming.model.snapshot());
        if (hasUnloadedWordWebFonts(document, options.wordWebFonts, requested)) {
          const expected = generation;
          void loadWordWebFonts(document, options.wordWebFonts, requested).then(next => {
            if (disposed || expected !== generation) return;
            hostedFonts = next;
            replace(prepared);
          }).catch(() => {});
        }
      }
    },
    async settled() {
      const expected = generation; const current = (): void => { if (disposed || expected !== generation) throw new DocumentError("InvalidOperation", "The pending render was disposed or replaced."); };
      await ready; current(); await frame.contentDocument!.fonts.ready; current();
      await Promise.all([...frame.contentDocument!.images].map(image => image.complete && image.naturalWidth ? Promise.resolve() : image.decode())); current(); measure(); return report;
    },
    printDocument() {
      if (disposed || !loaded || !canvas) throw new DocumentError("InvalidOperation", "The presentation preview is not ready for printing.");
      const document = frame.contentDocument!;
      const visibility = document.head.querySelector<HTMLStyleElement>("[data-sdk-slide-visibility]");
      const restoreVisibility = !!visibility && !visibility.disabled;
      if (restoreVisibility) visibility.disabled = true;
      try {
        if (editingSlide) {
          const count = canvas.fontScopes().length;
          for (const slide of prepared.model.snapshot().metadata.presentation!.slides)
            canvas.ensureSlide(slide.slideId);
          addScopeFonts(document, canvas.fontScopes().slice(count));
          canvas.finalizeCharts();
        }
        measure(false);
        return createPrintDocument(document, canvas.source.root, report.pages, limits);
      } finally {
        if (restoreVisibility) visibility.disabled = false;
      }
    },
    dispose() { if (disposed) return; disposed = true; generation++; observer?.disconnect(); if (scheduled !== undefined) host.defaultView!.cancelAnimationFrame(scheduled); fontProbes.clear(); canvas = undefined; projection = undefined; frame.removeEventListener("load", load); frame.remove(); frame.srcdoc = ""; releasePreviewUrls(prepared); resolve(); },
  };
}
