import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createPresentation, openDocument as openAnyDocument } from "@doffi/core";
import type { ContentDocumentModel as DocumentModel } from "@doffi/core";
import { writeDocumentZip, exportPptx, importXlsx } from "@doffi/office";
import { createDocumentConverter, executeConversionRequest, parseConversionRequest, serveConversionProtocol } from "@doffi/node";
import type { WorkbookModel } from "@doffi/core";
import type { ConversionRequestV1, DocumentConverter, RenderedOutput, WorkbookDocumentConverter } from "@doffi/node";
import type { WorkbookPrintReport } from "@doffi/render/contracts";

const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path,
  new Uint8Array(readFileSync(new URL("../../../examples/report/" + path, import.meta.url))),
]));
function request(taskId = "任务😀"): ConversionRequestV1 {
  return { protocolVersion: 1, taskId, format: "pdf", input: { format: "document-package-zip", documentId: "report-demo", bytesBase64: Buffer.from(writeDocumentZip(files)).toString("base64") } };
}
function fake(bytes = 8): DocumentConverter & { calls: string[] } {
  const calls: string[] = [];
  const render = async (model: DocumentModel): Promise<RenderedOutput> => {
    const snapshot = model.snapshot(); calls.push(snapshot.manifest.documentId);
    return { bytes: new Uint8Array(bytes), mimeType: "application/pdf", engine: { name: "chromium", version: "fixture" },
      report: { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, layout: "pages", pages: [], nodePages: {}, fields: [], diagnostics: [], fonts: [] } };
  };
  return { calls, renderPdf: render, renderPage: render, async close() {} };
}

test("conversion accepts only versioned authorized bytes and preserves asserted package identity", async () => {
  const converter = fake(), input = request(), result = await executeConversionRequest(converter, input);
  expect(result).toMatchObject({ protocolVersion: 1, taskId: "任务😀", ok: true, output: { mimeType: "application/pdf", report: { documentId: "report-demo", revision: 0 } } });
  expect(converter.calls).toEqual(["report-demo"]);
  expect(await executeConversionRequest(converter, { ...input, input: { ...input.input, documentId: "another-private-document" } })).toMatchObject({ ok: false, error: { code: "InvalidDocument" } });
  expect(converter.calls).toHaveLength(1);
});
test.each([
  [{ ...request(), protocolVersion: 2 }, "UnsupportedVersion"],
  [{ ...request(), format: "unknown" }, "UnsupportedFormat"],
  [{ ...request(), input: { format: "docx", bytesBase64: "AA==" } }, "InvalidRequest"],
  [{ ...request(), input: { ...request().input, path: "/private/file.docx" } }, "InvalidRequest"],
  [{ ...request(), input: { ...request().input, url: "https://example.com/private.docx" } }, "InvalidRequest"],
  [{ ...request(), options: { browserExecutable: "/tmp/executable" } }, "InvalidRequest"],
  [{ ...request(), options: { pageIndex: 0 } }, "InvalidRequest"],
  [{ ...request(), options: { limits: { maxPages: 0 } } }, "InvalidRequest"],
  [{ ...request(), extra: true }, "InvalidRequest"],
])("rejects an unsupported conversion request without invoking the renderer", async (input, code) => {
  const converter = fake();
  expect(await executeConversionRequest(converter, input)).toMatchObject({ taskId: "任务😀", ok: false, error: { code } });
  expect(converter.calls).toEqual([]);
});
test.each(["Z===", "Zh==", "AA A", "AAAA=", "AAéA"])("rejects noncanonical base64 %s", async (bytesBase64) => {
  expect(await executeConversionRequest(fake(), { ...request(), input: { ...request().input, bytesBase64 } })).toMatchObject({ ok: false, error: { code: "InvalidRequest" } });
});
test("conversion byte budgets are checked before decoding or rendering", async () => {
  const converter = fake();
  expect(await executeConversionRequest(converter, request(), { maxInputBytes: 1 })).toMatchObject({ ok: false, error: { code: "ResourceLimitExceeded" } });
  expect(converter.calls).toEqual([]);
});
test("NDJSON handles split UTF-8 and serializes requests behind output backpressure", async () => {
  const converter = fake(), data = new TextEncoder().encode(JSON.stringify(request("任务😀")) + "\n" + JSON.stringify(request("next")) + "\n");
  const output: string[] = [];
  async function* chunks(): AsyncGenerator<Uint8Array> { for (const byte of data) yield new Uint8Array([byte]); }
  await serveConversionProtocol(chunks(), async (line) => {
    expect(converter.calls).toHaveLength(output.length + 1); await Promise.resolve(); output.push(line);
  }, converter);
  expect(output.map((line) => JSON.parse(line).taskId)).toEqual(["任务😀", "next"]);
});
test("invalid JSON is a recoverable frame and oversized frames stop before the renderer", async () => {
  const converter = fake(), output: string[] = [], value = JSON.stringify(request());
  async function* input(): AsyncGenerator<Uint8Array> { yield new TextEncoder().encode("{bad}\n" + value + "\n"); }
  await serveConversionProtocol(input(), async (line) => { output.push(line); }, converter);
  expect(JSON.parse(output[0]!)).toMatchObject({ ok: false, taskId: null, error: { code: "InvalidRequest" } });
  expect(JSON.parse(output[1]!)).toMatchObject({ ok: true, taskId: "任务😀" });
  const rejected = fake();
  await expect(serveConversionProtocol(input(), async () => {}, rejected, { maxMessageBytes: 10 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  expect(rejected.calls).toEqual([]);
});
test("output frames are bounded and malformed UTF-8 is rejected", async () => {
  const output: string[] = [];
  async function* input(): AsyncGenerator<Uint8Array> { yield new TextEncoder().encode(JSON.stringify(request()) + "\n"); }
  await serveConversionProtocol(input(), async (line) => { output.push(line); }, fake(1024), { maxOutputMessageBytes: 512 });
  expect(JSON.parse(output[0]!)).toMatchObject({ ok: false, error: { code: "ResourceLimitExceeded" } });
  async function* invalid(): AsyncGenerator<Uint8Array> { yield new Uint8Array([0xff, 10]); }
  await expect(serveConversionProtocol(invalid(), async () => {}, fake())).rejects.toThrow();
  async function* malformed(): AsyncGenerator<Uint8Array> { yield new TextEncoder().encode("{bad}\n"); }
  const rejected: string[] = [];
  await expect(serveConversionProtocol(malformed(), async (line) => { rejected.push(line); }, fake(), { maxOutputMessageBytes: 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  expect(rejected).toEqual([]);
});
test("converter configuration, cancellation and quota failures do not require browser startup", async () => {
  expect(() => createDocumentConverter({ maxConcurrency: 0 })).toThrow();
  const model = openDocument(files), converter = createDocumentConverter({ maxInputBytes: 1 });
  await expect(converter.renderPdf(model)).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  const signal = new AbortController(); signal.abort();
  await expect(converter.renderPdf(model, { signal: signal.signal })).rejects.toMatchObject({ code: "Cancelled" });
  await expect(converter.renderPdf(model, { limits: { maxPages: 257 } })).rejects.toMatchObject({ code: "InvalidRequest" });
  await converter.close(); await converter.close();
  await expect(converter.renderPdf(model)).rejects.toMatchObject({ code: "InvalidOperation" });
  expect(() => parseConversionRequest({})).toThrow();
});

test("native PPTX bytes and PPT packages use the same bounded conversion boundary",async()=>{
  const files=createPresentation({documentId:"private-ppt"}),native=exportPptx(openDocument(files)).bytes;
  for(const[format,bytes]of[["pptx",native],["document-package-zip",writeDocumentZip(files)]]as const){
    const converter=fake(),result=await executeConversionRequest(converter,{protocolVersion:1,taskId:"ppt-conversion",format:"pdf",input:{format,documentId:"private-ppt",bytesBase64:Buffer.from(bytes).toString("base64")}});
    expect(result).toMatchObject({ok:true,output:{report:{documentId:"private-ppt"}}});expect(converter.calls).toEqual(["private-ppt"]);
  }
  expect(()=>parseConversionRequest({protocolVersion:1,taskId:"missing-id",format:"pdf",input:{format:"pptx",bytesBase64:"AA=="}})).toThrow();
});

const nativeWorkbook = readFileSync(new URL("../../../examples/xlsx/11-print-layout.xlsx", import.meta.url));
const workbookRequest: ConversionRequestV1 = {
  protocolVersion: 1, taskId: "print-book", format: "png",
  input: { format: "xlsx", documentId: "print-source", bytesBase64: nativeWorkbook.toString("base64") },
  options: { pageIndex: 1, workbook: { formulaValues: "cached", context: { locale: "zh-CN" } } },
};
test("workbook protocol routes native and package bytes with workbook-only settings and reports", async () => {
  const files = await importXlsx(nativeWorkbook, { documentId: "print-source" });
  const sheetId = openAnyDocument(files, { kind: "workbook" }).snapshot().metadata.workbook.sheets[0]!.sheetId;
  const workbook = { ...workbookRequest.options!.workbook, sheets: [{ sheetId, areas: [{ from: { row: 0, column: 0 }, to: { row: 2, column: 2 } }] }] };
  const methods: string[] = [];
  function rendered(model: WorkbookModel, mimeType: RenderedOutput["mimeType"]): RenderedOutput<WorkbookPrintReport> {
    const { documentId, revision } = model.snapshot().manifest;
    return { bytes: new Uint8Array([1]), mimeType, engine: { name: "chromium", version: "fixture" },
      report: { kind: "workbook", documentId, revision, layout: "pages", pages: [], nodePages: {}, fields: [], diagnostics: [], fonts: [],
        formulaValues: "cached", formulaCells: { valid: 0, imported: 0, stale: 0, unavailable: 0 } } };
  }
  const converter: WorkbookDocumentConverter = {
    ...fake(),
    async renderWorkbookPdf(model, options) {
      expect(options).toEqual(workbook);
      methods.push("pdf");
      return rendered(model, "application/pdf");
    },
    async renderWorkbookPage(model, options) {
      expect(options).toEqual({ pageIndex: 1, ...workbook });
      methods.push("png");
      return rendered(model, "image/png");
    },
  };
  const native = await executeConversionRequest(converter, { ...workbookRequest, options: { pageIndex: 1, workbook } });
  const input = { format: "document-package-zip", documentId: "print-source", bytesBase64: Buffer.from(writeDocumentZip(files)).toString("base64") };
  const packed = await executeConversionRequest(converter, { ...workbookRequest, input, format: "pdf", options: { workbook } });
  expect(native).toMatchObject({ ok: true, taskId: "print-book", output: { mimeType: "image/png", report: { kind: "workbook", documentId: "print-source", formulaValues: "cached" } } });
  expect(packed).toMatchObject({ ok: true, output: { mimeType: "application/pdf", report: { kind: "workbook", documentId: "print-source" } } });
  expect(await executeConversionRequest(converter, { ...workbookRequest, input: { ...input, documentId: "other" } })).toMatchObject({ ok: false, error: { code: "InvalidDocument" } });
  expect(await executeConversionRequest(converter, { ...request(), options: { workbook } })).toMatchObject({ ok: false, error: { code: "InvalidRequest" } });
  expect(await executeConversionRequest(fake(), workbookRequest)).toMatchObject({ ok: false, error: { code: "UnsupportedFeature" } });
  expect(methods).toEqual(["png", "pdf"]);
});

test.each([
  { url: "https://example.com/data" },
  { formulaValues: "auto" },
  { context: { now: { unixMilliseconds: 0, utcOffsetMinutes: 1441 } } },
  { fileName: "a".repeat(4097) },
  { maxPaperPixels: 0 },
  { sheets: [] },
  { sheets: [{ sheetId: "x" }, { sheetId: "x" }] },
  { sheets: [{ sheetId: "x", areas: [] }] },
  { sheets: [{ sheetId: "x", areas: [{ from: { row: 3, column: 0 }, to: { row: 2, column: 0 } }] }] },
  { sheets: [{ sheetId: "x", areas: Array.from({ length: 257 }, () => ({ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } })) }] },
])("validates nested workbook print settings before opening input", async workbook => {
  const result = await executeConversionRequest(fake(), { ...workbookRequest, options: { workbook } }, { maxInputBytes: 1 });
  expect(result).toMatchObject({ ok: false, taskId: "print-book", error: { code: "InvalidRequest" } });
});

test("native workbook conversion requires identity and can cancel before import finishes", async () => {
  expect(() => parseConversionRequest({ ...workbookRequest, input: { format: "xlsx", bytesBase64: "AA==" } })).toThrow();
  const controller = new AbortController();
  const pending = executeConversionRequest(fake(), workbookRequest, {}, controller.signal);
  controller.abort();
  expect(await pending).toMatchObject({ ok: false, taskId: "print-book", error: { code: "Cancelled" } });
});
