import { DocumentError } from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";
import type { DocumentPreview, DocumentPreviewOptions } from "../preview.js";
import { renderLimits } from "../contracts.js";
import type { PresentationRenderReport, FontResolution } from "../contracts.js";
import { resolvePreviewFonts } from "../fonts.js";
import { preparePreviewPackage } from "../preview-package.js";
import { SourceProjection } from "../projection.js";
import { copyPreviewSelection } from "../clipboard.js";
import { createPrintDocument } from "../print.js";
import type { PreviewSource } from "../source-index.js";
import { composePresentation } from "./canvas.js";
import type { PresentationCanvas } from "./canvas.js";

export function mountPresentationPreview(container: HTMLElement, input: ContentDocumentModel, options: DocumentPreviewOptions): DocumentPreview & { source(): PreviewSource | undefined } {
  if (options.layout === "continuous") throw new DocumentError("InvalidOperation", "Presentations use fixed pages layout.");
  const limits = renderLimits(options.limits), host = container.ownerDocument, frame = host.createElement("iframe");
  const prepare = (model: ContentDocumentModel): ReturnType<typeof preparePreviewPackage> => {
    const snapshot = model.snapshot(); if (snapshot.manifest.kind !== "presentation") throw new DocumentError("UnsupportedFormat", "Use a new preview when changing the document kind.");
    const result = preparePreviewPackage(model);
    if (result.model.snapshot().manifest.kind !== "presentation") { for (const url of result.urls) URL.revokeObjectURL(url); throw new DocumentError("UnsupportedFormat", "A presentation package is required."); }
    return result;
  };
  let prepared = prepare(input), canvas: PresentationCanvas | undefined, projection: SourceProjection | undefined, fonts: readonly FontResolution[] = [];
  let disposed = false, loaded = false,composing=false, generation = 0, observer: ResizeObserver | undefined, scheduled: number | undefined;
  let report: PresentationRenderReport = { kind:"presentation",slides:[], documentId: prepared.model.snapshot().manifest.documentId, revision: prepared.model.snapshot().manifest.revision, layout: "pages", pages: [], nodePages: {}, fields: [], fonts: [], diagnostics: [] };
  let resolve: () => void = () => {}, reject: (reason: unknown) => void = () => {};
  const ready = new Promise<void>((r, j) => { resolve = r; reject = j; });
  frame.title = options.title ?? "幻灯片只读预览"; frame.setAttribute("sandbox", "allow-same-origin allow-scripts"); frame.style.cssText = "width:100%;height:480px;display:block;border:0";
  function measure(): void {
    if (disposed || !canvas || composing) return;
    const snapshot = prepared.model.snapshot();
    const measured=canvas.measure(),p=snapshot.metadata.presentation!;
    const slides=measured.pages.flatMap(page=>{const slide=p.slides.find(slide=>slide.slideId===page.pageId);return slide?[Object.freeze({slideId:slide.slideId,rootNodeId:slide.rootNodeId,index:page.index,hidden:slide.hidden??false,rect:page.rect})]:[];});
    report = Object.freeze({ kind:"presentation",slides:Object.freeze(slides), documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, layout: "pages", ...measured, fields: canvas.fields, fonts });
    const document = frame.contentDocument!, scrollbar = Math.max(0, frame.clientHeight - document.documentElement.clientHeight);
    const height = Math.ceil(Math.max(200, document.body.scrollHeight) + scrollbar) + "px";
    if (frame.style.height !== height) frame.style.height = height;
  }
  function render(): void {
    observer?.disconnect();
    const document = frame.contentDocument!; canvas = composePresentation(document, prepared.model, limits,options.presentationSurfaceId);
    fonts = resolvePreviewFonts(document, "[data-node-id],[data-office-script],.sdk-ppt-marker"); canvas.finalizeCharts(); canvas.source.reindex();
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
  const load = (): void => {
    if (disposed) return;
    void frame.contentDocument!.fonts.ready.then(() => {
      if (disposed) return;
      render(); loaded = true;
      frame.contentDocument!.addEventListener("compositionstart",()=>{composing=true;},true);
      frame.contentDocument!.addEventListener("compositionend",()=>{composing=false;},true);
      frame.contentDocument!.addEventListener("copy", event => {
        const selection = frame.contentDocument!.getSelection();
        if (selection?.anchorNode && projection?.ignored(selection.anchorNode) || selection?.focusNode && projection?.ignored(selection.focusNode)) return;
        if (canvas && projection) copyPreviewSelection(event, canvas.source.root, projection);
      }); resolve();
    }).catch(error => { disposed = true; generation++; observer?.disconnect(); if (scheduled !== undefined) host.defaultView!.cancelAnimationFrame(scheduled); canvas = undefined; projection = undefined; frame.remove(); for (const url of prepared.urls) URL.revokeObjectURL(url); reject(error); });
  };
  frame.addEventListener("load", load, { once: true }); frame.srcdoc = "<!doctype html>" + prepared.document.documentElement.outerHTML; container.append(frame);
  return { element: frame, ready, source: () => canvas?.source, fonts: () => fonts, layout: () => report,
    update(model) {
      if (disposed || !loaded) throw new DocumentError("InvalidOperation", "Wait for an active preview before updating.");
      const incoming = prepare(model), previous = prepared; generation++;
      const document = frame.contentDocument!;
      const replace = (value: typeof prepared): void => { prepared = value; document.head.replaceWith(document.importNode(value.document.head, true)); document.body.replaceWith(document.importNode(value.document.body, true)); render(); };
      try { replace(incoming); } catch (error) { for (const url of incoming.urls) URL.revokeObjectURL(url); replace(previous); throw error; }
      for (const url of previous.urls) URL.revokeObjectURL(url);
    },
    async settled() {
      const expected = generation; const current = (): void => { if (disposed || expected !== generation) throw new DocumentError("InvalidOperation", "The pending render was disposed or replaced."); };
      await ready; current(); await frame.contentDocument!.fonts.ready; current();
      await Promise.all([...frame.contentDocument!.images].map(image => image.complete && image.naturalWidth ? Promise.resolve() : image.decode())); current(); measure(); return report;
    },
    printDocument() { if (disposed || !loaded || !canvas) throw new DocumentError("InvalidOperation", "The presentation preview is not ready for printing."); measure(); return createPrintDocument(frame.contentDocument!, canvas.source.root, report.pages, limits); },
    dispose() { if (disposed) return; disposed = true; generation++; observer?.disconnect(); if (scheduled !== undefined) host.defaultView!.cancelAnimationFrame(scheduled); canvas = undefined; projection = undefined; frame.removeEventListener("load", load); frame.remove(); frame.srcdoc = ""; for (const url of prepared.urls) URL.revokeObjectURL(url); resolve(); },
  };
}
