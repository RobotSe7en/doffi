import { DocumentError, openDocument } from "@doffi/core";
import type { BridgeWebFont, BrowserRenderBridge, BrowserRenderResult } from "./contracts.js";
import { collectDocumentFontNames, loadWordWebFonts } from "./fonts.js";
import type { WordWebFont } from "./fonts.js";
import { mountPreview } from "./preview.js";
import { renderWorkbookPrint } from "./workbook/print.js";
import type { WorkbookPrintDocument } from "./workbook/print.js";

declare global { interface Window { __doffiRender: BrowserRenderBridge } }
let workbook: WorkbookPrintDocument | undefined;
let renderer: ReturnType<typeof mountPreview> | undefined;
let activeWebFonts: readonly WordWebFont[] | undefined;
let activeRequestedFonts: ReadonlySet<string> | undefined;

function toWordWebFonts(entries: readonly BridgeWebFont[] | undefined): readonly WordWebFont[] | undefined {
  if (!entries?.length) return undefined;
  return entries.map((entry) => {
    let pending: Promise<ArrayBuffer> | undefined;
    return {
      family: entry.family,
      aliases: entry.aliases,
      ...(entry.weight !== undefined ? { weight: entry.weight } : {}),
      ...(entry.style !== undefined ? { style: entry.style } : {}),
      ...(entry.equivalentAliases !== undefined ? { equivalentAliases: entry.equivalentAliases } : {}),
      ...(entry.fallbackOnly !== undefined ? { fallbackOnly: entry.fallbackOnly } : {}),
      load() {
        return pending ??= fetch(entry.url).then((response) => {
          if (!response.ok) throw new Error("Failed to load host font.");
          return response.arrayBuffer();
        }).catch((error) => {
          pending = undefined;
          throw error;
        });
      },
    };
  });
}

function failure(reason: unknown): BrowserRenderResult {
  const error = reason instanceof DocumentError ? reason : new DocumentError("ExportFailed", "Browser rendering failed.");
  return { ok: false, error: error.toJSON() };
}
window.__doffiRender = {
  async mount(files, options) {
    try {
      window.__doffiRender.dispose();
      const input = new Map(files.map(([path, base64]) => [path, Uint8Array.from(atob(base64), (value) => value.charCodeAt(0))]));
      document.body.style.margin = "0";
      const host = document.createElement("div"); document.body.replaceChildren(host);
      const model = openDocument(input);
      activeWebFonts = toWordWebFonts(options?.webFonts);
      if (model.kind === "workbook") {
        workbook = await renderWorkbookPrint(host, model, { ...options?.workbook,
          ...(options?.limits ? { limits: options.limits } : {}), ...(activeWebFonts ? { webFonts: activeWebFonts } : {}) });
        return { ok: true, report: workbook.report };
      }
      activeRequestedFonts = activeWebFonts?.length ? collectDocumentFontNames(model.snapshot()) : undefined;
      renderer = mountPreview(host, model, {
        ...options,
        ...(activeWebFonts ? { wordWebFonts: activeWebFonts } : {}),
      });
      const report = await renderer.settled();
      // Renderer-owned labels are preview UI, never part of a PNG or PDF.
      const source = renderer.source();
      if (!source) throw new DocumentError("InvalidOperation", "The rendered source is unavailable.");
      const canvasDocument = source.root.ownerDocument;
      const clean = canvasDocument.createElement("style"); clean.textContent = ".sdk-ppt-hidden-label{display:none!important}"; canvasDocument.head.append(clean);
      return { ok: true, report };
    } catch (reason) { return failure(reason); }
  },
  async preparePrint() {
    try {
      if (workbook) return { ok: true, report: workbook.report };
      if (!renderer) throw new DocumentError("InvalidOperation", "A document must be rendered before printing.");
      const report = await renderer.settled(), print = renderer.printDocument();
      if (activeWebFonts?.length) await loadWordWebFonts(document, activeWebFonts, activeRequestedFonts);
      document.head.replaceWith(document.importNode(print.head, true)); document.body.replaceWith(document.importNode(print.body, true));
      await document.fonts.ready;
      await Promise.all([...document.images].map(async (image) => { if (!image.complete || !image.naturalWidth) await image.decode(); }));
      return { ok: true, report };
    } catch (reason) { return failure(reason); }
  },
  pageClip(index) {
    if (workbook) {
      const selected = workbook.pages[index];
      if (!selected) throw new DocumentError("InvalidOperation", "The requested page does not exist.");
      for (const page of workbook.pages) page.style.display = page === selected ? "block" : "none";
      const { x, y, width, height } = selected.getBoundingClientRect();
      return { x, y, width, height };
    }
    const page = renderer?.layout().pages[index], source = renderer?.source();
    if (!page || !source || !renderer) throw new DocumentError("InvalidOperation", "The requested page does not exist.");
    const outer = renderer.element.getBoundingClientRect(), inner = source.root.getBoundingClientRect();
    return { x: outer.left + inner.left + page.rect.x * 96 / 72, y: outer.top + inner.top + page.rect.y * 96 / 72,
      width: page.rect.width * 96 / 72, height: page.rect.height * 96 / 72 };
  },
  dispose() { workbook?.dispose(); workbook = undefined; renderer?.dispose(); renderer = undefined; activeWebFonts = undefined; activeRequestedFonts = undefined; },
};
