import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { chromium } from "playwright-core";
import type { Browser, BrowserContext } from "playwright-core";
import { expect, test, vi } from "vitest";
import {} from "@doffi/core";
import { createDocumentConverter } from "@doffi/node";
function deferred<T>(): { promise: Promise<T>; resolve(value: T): void } {
  let resolve: (value: T) => void = () => {};
  const promise = new Promise<T>((done) => { resolve = done; }); return { promise, resolve };
}

test("a timed-out pending context keeps its quota slot until backend work and cleanup finish", async () => {
  const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path,
    new Uint8Array(readFileSync(new URL("../../../examples/report/" + path, import.meta.url))),
  ]));
  const context = deferred<BrowserContext>(), cleanup = deferred<void>();
  const contextClose = vi.fn(async () => {}), browserClose = vi.fn(() => cleanup.promise);
  const browser = { newContext: vi.fn(() => context.promise), close: browserClose } as unknown as Browser;
  const launch = vi.spyOn(chromium, "launch").mockResolvedValue(browser);
  const converter = createDocumentConverter({ maxConcurrency: 1, timeoutMs: 1000 }), model = openDocument(files);
  try {
    const pending = converter.renderPdf(model, { timeoutMs: 30 });
    await expect(pending).rejects.toMatchObject({ code: "Timeout" });
    expect(browserClose).toHaveBeenCalledTimes(1);
    await expect(converter.renderPdf(model)).rejects.toMatchObject({ code: "Busy" });
    expect(launch).toHaveBeenCalledTimes(1);
    let closed = false; const closing = converter.close().then(() => { closed = true; });
    await Promise.resolve(); expect(closed).toBe(false);
    context.resolve({ close: contextClose } as unknown as BrowserContext);
    await Promise.resolve(); expect(closed).toBe(false);
    cleanup.resolve(); await closing;
    expect(contextClose).toHaveBeenCalledTimes(1); expect(browserClose).toHaveBeenCalledTimes(1);
  } finally {
    context.resolve({ close: contextClose } as unknown as BrowserContext); cleanup.resolve(); await converter.close(); launch.mockRestore();
  }
});

test("host webFonts are validated and fulfilled offline through the isolated font route", async () => {
  expect(() => createDocumentConverter({ webFonts: [{ family: "", path: "/tmp/a.woff" }] })).toThrowError(expect.objectContaining({ code: "InvalidRequest" }));
  expect(() => createDocumentConverter({ webFonts: [{ family: "HostFont" }] })).toThrowError(expect.objectContaining({ code: "InvalidRequest" }));
  expect(() => createDocumentConverter({ webFonts: Array.from({ length: 257 }, (_, i) => ({ family: "F" + i, path: "/tmp/f.woff" })) })).toThrowError(expect.objectContaining({ code: "InvalidRequest" }));

  const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path,
    new Uint8Array(readFileSync(new URL("../../../examples/report/" + path, import.meta.url))),
  ]));
  const fontBytes = new Uint8Array([0x77, 0x4f, 0x46, 0x46]);
  let routeHandler: ((route: { request(): { url(): string }; continue(): Promise<void>; abort(reason?: string): Promise<void>; fulfill(opts: { status: number; contentType: string; headers: Record<string, string>; body: Buffer }): Promise<void> }) => Promise<void>) | undefined;
  let capturedWebFonts: unknown;
  const fakeReport = { documentId: "doc", revision: 0, layout: "pages", pages: [{ pageId: "p1", index: 1, sectionId: "s1", number: 1, label: "1", rect: { x: 0, y: 0, width: 612, height: 792 }, setup: { unit: "pt", width: 612, height: 792, margins: { top: 0, right: 0, bottom: 0, left: 0 } }, blank: false }], nodePages: {}, fields: [], fonts: [], diagnostics: [] };
  const fakePage = {
    setDefaultTimeout: vi.fn(),
    emulateMedia: vi.fn(async () => {}),
    addScriptTag: vi.fn(async () => {}),
    evaluate: vi.fn(async (_fn: unknown, arg?: { webFonts?: unknown }) => {
      if (arg && "webFonts" in arg) capturedWebFonts = arg.webFonts;
      return { ok: true, report: fakeReport };
    }),
    pdf: vi.fn(async () => Buffer.from("%PDF-1.4\n")),
  };
  const fakeContext = {
    route: vi.fn(async (_pattern: string, handler: typeof routeHandler) => { routeHandler = handler; }),
    newPage: vi.fn(async () => fakePage),
    close: vi.fn(async () => {}),
  } as unknown as BrowserContext;
  const fakeBrowser = {
    version: () => "130.0",
    newContext: vi.fn(async () => fakeContext),
    close: vi.fn(async () => {}),
  } as unknown as Browser;
  const launch = vi.spyOn(chromium, "launch").mockResolvedValue(fakeBrowser);
  const converter = createDocumentConverter({
    webFonts: [{ family: "Host Sans", aliases: ["微软雅黑"], weight: "400", equivalentAliases: true, load: () => fontBytes }],
  });
  try {
    const out = await converter.renderPdf(openDocument(files));
    expect(out.mimeType).toBe("application/pdf");
    expect(capturedWebFonts).toEqual([{ family: "Host Sans", aliases: ["微软雅黑"], url: "https://fonts.doffi.invalid/0", weight: "400", equivalentAliases: true }]);
    expect(routeHandler).toBeDefined();

    const fulfilled: unknown[] = [], aborted: string[] = [];
    const callRoute = (url: string) => routeHandler!({
      request: () => ({ url: () => url }),
      continue: async () => {},
      abort: async (reason) => { aborted.push(url + ":" + (reason ?? "")); },
      fulfill: async (opts) => { fulfilled.push(opts); },
    });
    await callRoute("https://fonts.doffi.invalid/0");
    await callRoute("https://fonts.doffi.invalid/99");
    await callRoute("https://example.com/malicious.woff2");
    expect(fulfilled).toHaveLength(1);
    expect(fulfilled[0]).toMatchObject({ status: 200, headers: { "Access-Control-Allow-Origin": "*" } });
    expect(new Uint8Array((fulfilled[0] as { body: Buffer }).body)).toEqual(fontBytes);
    expect(aborted).toEqual(["https://fonts.doffi.invalid/99:blockedbyclient", "https://example.com/malicious.woff2:blockedbyclient"]);
  } finally {
    await converter.close(); launch.mockRestore();
  }
});
