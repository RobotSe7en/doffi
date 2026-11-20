import { readFile } from "node:fs/promises";
import { chromium } from "playwright-core";
import type { Browser, BrowserContext } from "playwright-core";
import { DocumentError, openDocument } from "@doffi/core";
import type { DocumentErrorCode, ContentDocumentModel } from "@doffi/core";
import { DEFAULT_RENDER_LIMITS, renderLimits } from "@doffi/render/contracts";
import type { BridgeWebFont, BrowserRenderBridge, BrowserRenderResult, DocumentRenderReport, RenderLimits } from "@doffi/render/contracts";
import { conversionEnvironment } from "./environment.js";

/** Host-authorized font asset for Node PDF/PNG rendering. Never supplied by untrusted conversion requests. */
export interface NodeWebFont {
  readonly family: string;
  readonly aliases?: readonly string[];
  readonly weight?: string;
  readonly style?: string;
  readonly equivalentAliases?: boolean;
  readonly fallbackOnly?: boolean;
  readonly path?: string;
  readonly load?: () => Promise<Uint8Array | ArrayBuffer> | Uint8Array | ArrayBuffer;
}

export type ConversionErrorCode = DocumentErrorCode | "Cancelled" | "Timeout" | "Busy" | "InvalidRequest";
export class ConversionError extends Error {
  readonly code: ConversionErrorCode;
  constructor(code: ConversionErrorCode, message: string) { super(message.slice(0, 256)); this.name = "ConversionError"; this.code = code; }
  toJSON(): { code: ConversionErrorCode; message: string } { return { code: this.code, message: this.message }; }
}
export interface ConverterOptions {
  /** Host configuration, never a path supplied by an untrusted conversion request. */
  readonly browserExecutable?: string;
  readonly browserChannel?: string;
  readonly webFonts?: readonly NodeWebFont[];
  readonly maxConcurrency?: number;
  readonly timeoutMs?: number;
  readonly maxInputBytes?: number;
  readonly maxOutputBytes?: number;
  readonly maxPixels?: number;
  readonly limits?: Partial<RenderLimits>;
}
export interface RenderOptions {
  readonly signal?: AbortSignal;
  readonly timeoutMs?: number;
  readonly limits?: Partial<RenderLimits>;
}
export interface PageRenderOptions extends RenderOptions { readonly pageIndex?: number; readonly scale?: number }
export interface RenderedOutput {
  readonly bytes: Uint8Array;
  readonly mimeType: "application/pdf" | "image/png";
  readonly report: DocumentRenderReport;
  readonly engine: { readonly name: "chromium"; readonly version: string };
}
export interface DocumentConverter {
  renderPdf(model: ContentDocumentModel, options?: RenderOptions): Promise<RenderedOutput>;
  renderPage(model: ContentDocumentModel, options?: PageRenderOptions): Promise<RenderedOutput>;
  /** Cancels outstanding jobs and closes every context and the owned browser. */
  close(): Promise<void>;
}
export const DEFAULT_CONVERTER_LIMITS: Readonly<Required<Omit<ConverterOptions, "browserExecutable" | "browserChannel" | "limits" | "webFonts">>> = Object.freeze({
  maxConcurrency: 2, timeoutMs: 30_000, maxInputBytes: 64 * 1024 * 1024, maxOutputBytes: 64 * 1024 * 1024, maxPixels: 16 * 1024 * 1024,
});
const FONT_ROUTE_PREFIX = "https://fonts.doffi.invalid/";
function validateWebFonts(entries: readonly NodeWebFont[] | undefined): readonly NodeWebFont[] {
  if (!entries?.length) return [];
  if (entries.length > 256) throw new ConversionError("InvalidRequest", "Configured web fonts exceed the host limit.");
  for (const entry of entries) {
    if (!entry.family || (!entry.path && !entry.load))
      throw new ConversionError("InvalidRequest", "Each configured web font requires a family and a path or byte loader.");
  }
  return entries;
}
function fontMimeType(path: string | undefined): string {
  if (path?.endsWith(".woff2")) return "font/woff2";
  if (path?.endsWith(".woff")) return "font/woff";
  if (path?.endsWith(".otf")) return "font/otf";
  return "font/ttf";
}
function positive(value: number, name: string): number {
  if (!Number.isSafeInteger(value) || value <= 0) throw new ConversionError("InvalidRequest", name + " must be a positive safe integer."); return value;
}
function abortError(signal: AbortSignal): ConversionError {
  return new ConversionError(signal.reason instanceof Error && signal.reason.name === "TimeoutError" ? "Timeout" : "Cancelled",
    signal.reason instanceof Error && signal.reason.name === "TimeoutError" ? "Document rendering timed out." : "Document rendering was cancelled.");
}
function check(result: BrowserRenderResult): DocumentRenderReport {
  if (!result.ok) throw new DocumentError(result.error.code, result.error.message, result.error.details); return result.report;
}

/** Every job owns a fresh browser/context and only authorized package bytes. */
export function createDocumentConverter(options: ConverterOptions = {}): DocumentConverter {
  const quota = { ...DEFAULT_CONVERTER_LIMITS, ...options }, globalLimits = renderLimits(options.limits ?? DEFAULT_RENDER_LIMITS);
  const hostFonts = validateWebFonts(options.webFonts);
  const bridgeFonts: readonly BridgeWebFont[] | undefined = hostFonts.length
    ? hostFonts.map((font, index) => ({
      family: font.family,
      aliases: font.aliases ?? [],
      url: FONT_ROUTE_PREFIX + index,
      ...(font.weight !== undefined ? { weight: font.weight } : {}),
      ...(font.style !== undefined ? { style: font.style } : {}),
      ...(font.equivalentAliases !== undefined ? { equivalentAliases: font.equivalentAliases } : {}),
      ...(font.fallbackOnly !== undefined ? { fallbackOnly: font.fallbackOnly } : {}),
    }))
    : undefined;
  for (const key of ["maxConcurrency", "timeoutMs", "maxInputBytes", "maxOutputBytes", "maxPixels"] as const) positive(quota[key], key);
  let script: Promise<string> | undefined, closing: Promise<void> | undefined, closed = false, cleanupFailed = false;
  const jobs = new Set<{ controller: AbortController; completion: Promise<unknown> }>();
  async function render(model: ContentDocumentModel, format: "pdf" | "png", settings: PageRenderOptions = {}): Promise<RenderedOutput> {
    if (closed) throw new ConversionError("InvalidOperation", "The converter has been closed.");
    if (settings.signal?.aborted) throw abortError(settings.signal);
    if (jobs.size >= quota.maxConcurrency) throw new ConversionError("Busy", "The converter is at its configured concurrency limit.");
    const timeout = positive(settings.timeoutMs ?? quota.timeoutMs, "timeoutMs");
    if (timeout > quota.timeoutMs) throw new ConversionError("InvalidRequest", "Requested timeout exceeds the host quota.");
    const limits = renderLimits({ ...globalLimits, ...settings.limits });
    for (const key of Object.keys(globalLimits) as (keyof RenderLimits)[]) if (limits[key] > globalLimits[key]) throw new ConversionError("InvalidRequest", "Requested render limits exceed the host quota.");
    const files = model.serialize(); let size = 0;
    for (const [path, value] of files) {
      size += new TextEncoder().encode(path).length + value.byteLength;
      if (size > quota.maxInputBytes) throw new ConversionError("ResourceLimitExceeded", "Document exceeds the conversion input byte budget.");
    }
    const checked = openDocument(files, { kind: model.kind }), metadata = checked.snapshot().metadata;
    const width = metadata.presentation ? metadata.presentation.size.width : Math.max(metadata.pageSetup.width, ...metadata.sections.map((section) => (section.pageSetup ?? metadata.pageSetup).width));
    if (width > 20_000) throw new ConversionError("UnsupportedFeature", "Page width exceeds the renderer size limit.");
    const scale = settings.scale ?? 1, pageIndex = settings.pageIndex ?? 0;
    if (!Number.isFinite(scale) || scale < 1 || scale > 4 || !Number.isSafeInteger(pageIndex) || pageIndex < 0) throw new ConversionError("InvalidRequest", "Page index or image scale is invalid.");
    const controller = new AbortController(), signal = AbortSignal.any([controller.signal, AbortSignal.timeout(timeout), ...(settings.signal ? [settings.signal] : [])]);
    let context: BrowserContext | undefined, ownedBrowser: Browser | undefined, backendCleanup: Promise<void> | undefined;
    const stopBackend = (): Promise<void> => {
      if (!ownedBrowser) return Promise.resolve();
      return backendCleanup ??= ownedBrowser.close();
    };
    const job = { controller, completion: Promise.resolve() }; jobs.add(job);
    let rejectAbort: (reason: unknown) => void = () => {};
    const aborted = new Promise<never>((_resolve, reject) => { rejectAbort = reject; });
    const onAbort = (): void => { rejectAbort(abortError(signal)); void stopBackend().catch(() => {}); };
    signal.addEventListener("abort", onAbort, { once: true });
    const operation = (async (): Promise<RenderedOutput> => {
      try {
      const current = await chromium.launch({ headless: true, timeout, env: conversionEnvironment(),
        ...(options.browserExecutable ? { executablePath: options.browserExecutable } : {}), ...(options.browserChannel ? { channel: options.browserChannel } : {}),
      });
      ownedBrowser = current; if (signal.aborted) throw abortError(signal);
      const viewportWidth = Math.max(1280, Math.ceil(width * 96 / 72) + 48);
      const created = await current.newContext({ viewport: { width: viewportWidth, height: Math.min(900, Math.max(200, Math.floor(quota.maxPixels / viewportWidth))) },
        deviceScaleFactor: format === "png" ? scale : 1, locale: "zh-CN", colorScheme: "light", reducedMotion: "reduce", serviceWorkers: "block", acceptDownloads: false });
      if (signal.aborted) { await created.close(); throw abortError(signal); } context = created;
      await context.route("**/*", async (route) => {
        const url = route.request().url();
        if (url.startsWith("blob:") || url === "about:blank") return route.continue();
        if (hostFonts.length && url.startsWith(FONT_ROUTE_PREFIX)) {
          const index = Number(url.slice(FONT_ROUTE_PREFIX.length));
          const font = Number.isSafeInteger(index) && index >= 0 && index < hostFonts.length ? hostFonts[index] : undefined;
          if (!font) return route.abort("blockedbyclient");
          try {
            const raw = font.load ? await font.load() : await readFile(font.path!);
            const body = raw instanceof Uint8Array ? Buffer.from(raw.buffer, raw.byteOffset, raw.byteLength) : Buffer.from(raw);
            return route.fulfill({
              status: 200,
              contentType: fontMimeType(font.path),
              headers: { "Access-Control-Allow-Origin": "*" },
              body,
            });
          } catch {
            return route.abort("failed");
          }
        }
        return route.abort("blockedbyclient");
      });
      const page = await context.newPage(); page.setDefaultTimeout(timeout); await page.emulateMedia({ media: "screen" });
      script ??= readFile(new URL(import.meta.resolve("@doffi/render/browser-script")), "utf8");
      await page.addScriptTag({ content: await script });
      const result = await page.evaluate(async ({ entries, bounds, webFonts }) => {
        const bridge = (globalThis as unknown as { __doffiRender: BrowserRenderBridge }).__doffiRender;
        return bridge.mount(entries, { layout: "pages", limits: bounds, ...(webFonts ? { webFonts } : {}) });
      }, { entries: [...files].map(([path, bytes]) => [path, Buffer.from(bytes).toString("base64")] as const), bounds: limits, webFonts: bridgeFonts });
      let report = check(result), bytes: Uint8Array;
      if (format === "pdf") {
        report = check(await page.evaluate(() => (globalThis as unknown as { __doffiRender: BrowserRenderBridge }).__doffiRender.preparePrint()));
        bytes = await page.pdf({ preferCSSPageSize: true, printBackground: true, tagged: true });
      } else {
        if (pageIndex >= report.pages.length) throw new ConversionError("InvalidRequest", "Requested image page does not exist.");
        const clip = await page.evaluate((index) => (globalThis as unknown as { __doffiRender: BrowserRenderBridge }).__doffiRender.pageClip(index), pageIndex);
        if (clip.width * clip.height * scale * scale > quota.maxPixels) throw new ConversionError("ResourceLimitExceeded", "Page image exceeds the pixel budget.");
        bytes = await page.screenshot({ type: "png", clip, fullPage: true });
      }
      if (bytes.byteLength > quota.maxOutputBytes) throw new ConversionError("ResourceLimitExceeded", "Rendered output exceeds its byte budget.");
      return { bytes: new Uint8Array(bytes), mimeType: format === "pdf" ? "application/pdf" : "image/png", report, engine: { name: "chromium", version: current.version() } };
      } finally {
        try { await stopBackend(); }
        catch { cleanupFailed = true; throw new ConversionError("ExportFailed", "Browser cleanup failed."); }
        finally { jobs.delete(job); }
      }
    })();
    // Caller cancellation is prompt; the quota remains occupied until the actual
    // backend operation and its owned browser cleanup have both completed.
    job.completion = operation.then(() => {}, () => {});
    try { return await Promise.race([operation, aborted]); }
    catch (reason) {
      if (reason instanceof ConversionError || reason instanceof DocumentError) throw reason;
      if (signal.aborted) throw abortError(signal);
      if (reason instanceof Error && reason.name === "TimeoutError") throw new ConversionError("Timeout", "Document rendering timed out.");
      throw new ConversionError("ExportFailed", "Chromium rendering failed.");
    } finally {
      signal.removeEventListener("abort", onAbort);
    }
  }
  return {
    renderPdf: (model, settings) => render(model, "pdf", settings), renderPage: (model, settings) => render(model, "png", settings),
    close() {
      closing ??= (async () => {
        closed = true; for (const job of jobs) job.controller.abort();
        await Promise.allSettled([...jobs].map((job) => job.completion));
        if (cleanupFailed) throw new ConversionError("ExportFailed", "Browser cleanup failed.");
      })(); return closing;
    },
  };
}
export async function renderDocumentPdf(model: ContentDocumentModel, options: ConverterOptions & RenderOptions = {}): Promise<RenderedOutput> {
  const converter = createDocumentConverter(options);
  try { return await converter.renderPdf(model, options); } finally { await converter.close(); }
}
