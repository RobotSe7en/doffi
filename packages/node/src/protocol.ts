import { copyWorkbookScalarContext, DocumentError, openDocument, WorkbookRangeSchema } from "@doffi/core";
import type { WorkbookScalarContext } from "@doffi/core";
import { importDocx, importPptx, importXlsx, readDocumentZip } from "@doffi/office";
import { DEFAULT_RENDER_LIMITS } from "@doffi/render/contracts";
import type { ConversionRequestV1, ConversionResponseV1, PortableRenderReport, RenderLimits } from "@doffi/render/contracts";
import { ConversionError, DEFAULT_CONVERTER_LIMITS } from "./render.js";
import type { DocumentConverter, RenderedOutput, WorkbookDocumentConverter } from "./render.js";

type ProtocolConverter = DocumentConverter & Partial<Pick<WorkbookDocumentConverter, "renderWorkbookPdf" | "renderWorkbookPage">>;
export type { ConversionRequestV1, ConversionResponseV1 } from "@doffi/render/contracts";

export interface ProtocolLimits { readonly maxInputBytes?: number; readonly maxMessageBytes?: number; readonly maxOutputMessageBytes?: number }
function object(value: unknown, keys: readonly string[]): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value) || Object.keys(value).some((key) => !keys.includes(key))) throw new ConversionError("InvalidRequest", "Conversion object has an invalid shape or unknown fields.");
  return value as Record<string, unknown>;
}
function identifier(value: unknown): value is string { return typeof value === "string" && value.length > 0 && value.length <= 256; }
function integer(value: unknown, minimum: number): value is number { return typeof value === "number" && Number.isSafeInteger(value) && value >= minimum; }
function validateWorkbookOptions(value: unknown, maximum: number): void {
  const options = object(value, ["sheets", "formulaValues", "context", "fileName", "filePath", "maxPaperPixels"]);
  if (options.formulaValues !== undefined && options.formulaValues !== "current" && options.formulaValues !== "cached"
    || options.maxPaperPixels !== undefined && !integer(options.maxPaperPixels, 1)
    || [options.fileName, options.filePath].some(text => text !== undefined && (typeof text !== "string" || text.length > 4096)))
    throw new ConversionError("InvalidRequest", "Workbook print options are invalid.");
  if (options.context !== undefined) {
    try { copyWorkbookScalarContext(options.context as WorkbookScalarContext); }
    catch { throw new ConversionError("InvalidRequest", "Workbook print time or locale is invalid."); }
  }
  if (options.sheets === undefined) return;
  // A selected worksheet or explicit area consumes at least one print page.
  const seen = new Set<string>();
  if (!Array.isArray(options.sheets) || !options.sheets.length || options.sheets.length > maximum)
    throw new ConversionError("InvalidRequest", "Workbook print sheet selection is invalid.");
  let areas = 0;
  for (const value of options.sheets) {
    const selection = object(value, ["sheetId", "areas"]);
    if (!identifier(selection.sheetId) || seen.has(selection.sheetId))
      throw new ConversionError("InvalidRequest", "Workbook print sheet identifiers must be distinct.");
    seen.add(selection.sheetId);
    if (selection.areas === undefined) { areas++; continue; }
    if (!Array.isArray(selection.areas) || !selection.areas.length || selection.areas.length > maximum - areas)
      throw new ConversionError("InvalidRequest", "Workbook print area selection is invalid.");
    areas += selection.areas.length;
    for (const value of selection.areas) {
      const range = WorkbookRangeSchema.safeParse(value);
      if (!range.success || range.data.from.row > range.data.to.row || range.data.from.column > range.data.to.column)
        throw new ConversionError("InvalidRequest", "Workbook print ranges must use ordered zero-based coordinates.");
    }
  }
  if (areas > maximum) throw new ConversionError("InvalidRequest", "Too many workbook print areas were selected.");
}
export function parseConversionRequest(input: unknown): ConversionRequestV1 {
  const request = object(input, ["protocolVersion", "taskId", "format", "input", "options"]);
  if (request.protocolVersion !== 1) throw new ConversionError("UnsupportedVersion", "Unknown conversion protocol version.");
  if (!identifier(request.taskId)) throw new ConversionError("InvalidRequest", "A bounded task identifier is required.");
  if (!["pdf", "png"].includes(request.format as string)) throw new ConversionError("UnsupportedFormat", "Unknown conversion output format.");
  const source = object(request.input, ["format", "bytesBase64", "documentId"]);
  if (!["document-package-zip", "docx", "pptx", "xlsx"].includes(source.format as string)) throw new ConversionError("UnsupportedFormat", "Unknown conversion input format.");
  if (typeof source.bytesBase64 !== "string" || !source.bytesBase64.length) throw new ConversionError("InvalidRequest", "Authorized input bytes are required.");
  if ((source.documentId !== undefined || source.format !== "document-package-zip") && !identifier(source.documentId)) throw new ConversionError("InvalidRequest", "Native Office conversion requires a bounded document identifier.");
  if (request.options !== undefined) {
    const options = object(request.options, ["pageIndex", "scale", "timeoutMs", "limits", "workbook"]);
    if (options.pageIndex !== undefined && !integer(options.pageIndex, 0) || options.timeoutMs !== undefined && !integer(options.timeoutMs, 1)
      || options.scale !== undefined && (typeof options.scale !== "number" || !Number.isFinite(options.scale) || options.scale < 1 || options.scale > 4)) throw new ConversionError("InvalidRequest", "Conversion output options are invalid.");
    if (request.format === "pdf" && (options.pageIndex !== undefined || options.scale !== undefined)) throw new ConversionError("InvalidRequest", "PDF output uses all measured pages at document scale.");
    let maxPages = DEFAULT_RENDER_LIMITS.maxPages;
    if (options.limits !== undefined) {
      const limits = object(options.limits, ["maxNodes", "maxPages", "maxLayoutPasses", "maxPrintNodes"] satisfies (keyof RenderLimits)[]);
      if (Object.values(limits).some((value) => !integer(value, 1))) throw new ConversionError("InvalidRequest", "Conversion render limits are invalid.");
      if (integer(limits.maxPages, 1)) maxPages = limits.maxPages;
    }
    if (options.workbook !== undefined) {
      if (source.format === "docx" || source.format === "pptx")
        throw new ConversionError("InvalidRequest", "Workbook print options require workbook input.");
      validateWorkbookOptions(options.workbook, maxPages);
    }
  }
  return request as unknown as ConversionRequestV1;
}
function failure(reason: unknown, taskId: string | null): ConversionResponseV1 {
  const error = reason instanceof ConversionError || reason instanceof DocumentError ? reason.toJSON() : { code: "ExportFailed" as const, message: "Conversion failed." };
  return { protocolVersion: 1, taskId, ok: false, error };
}
export async function executeConversionRequest(converter: ProtocolConverter, input: unknown, limits: ProtocolLimits = {}, signal?: AbortSignal): Promise<ConversionResponseV1> {
  let taskId: string | null = input && typeof input === "object" && Object.hasOwn(input, "taskId") && identifier((input as Record<string,unknown>).taskId) ? (input as { taskId: string }).taskId : null;
  try {
    const request = parseConversionRequest(input); taskId = request.taskId;
    const checkCancellation = (): void => {
      if (signal?.aborted) throw new ConversionError("Cancelled", "Conversion was cancelled.");
    };
    checkCancellation();
    const maximum = limits.maxInputBytes ?? DEFAULT_CONVERTER_LIMITS.maxInputBytes, encoded = request.input.bytesBase64;
    if (!integer(maximum, 1)) throw new ConversionError("InvalidRequest", "Protocol input limit must be a positive safe integer.");
    if (encoded.length > Math.ceil(maximum / 3) * 4) throw new ConversionError("ResourceLimitExceeded", "Conversion input exceeds its byte budget.");
    if (encoded.length % 4 || /[^A-Za-z0-9+/=]/.test(encoded)) throw new ConversionError("InvalidRequest", "Input bytes must use canonical base64.");
    const bytes = Buffer.from(encoded, "base64");
    if (bytes.length > maximum) throw new ConversionError("ResourceLimitExceeded", "Conversion input exceeds its byte budget.");
    if (bytes.toString("base64") !== encoded) throw new ConversionError("InvalidRequest", "Input bytes must use canonical base64.");
    const { format, documentId } = request.input;
    const nativeOptions = { documentId: documentId! };
    const files = format === "xlsx" ? await importXlsx(bytes, { ...nativeOptions, ...(signal ? { signal } : {}) })
      : format === "docx" ? importDocx(bytes, nativeOptions)
      : format === "pptx" ? importPptx(bytes, nativeOptions) : readDocumentZip(bytes);
    checkCancellation();
    const model = openDocument(files);
    if (request.input.documentId !== undefined && model.snapshot().manifest.documentId !== request.input.documentId) throw new ConversionError("InvalidDocument", "Package identity differs from the asserted document identifier.");
    const { workbook, ...options } = request.options ?? {};
    const settings = { ...options, ...(signal ? { signal } : {}) };
    let output: RenderedOutput<PortableRenderReport>;
    if (model.kind === "workbook") {
      if (!converter.renderWorkbookPdf || !converter.renderWorkbookPage)
        throw new ConversionError("UnsupportedFeature", "This converter does not provide workbook printing.");
      const print = { ...settings, ...workbook };
      output = request.format === "pdf" ? await converter.renderWorkbookPdf(model, print) : await converter.renderWorkbookPage(model, print);
    } else {
      if (workbook) throw new ConversionError("InvalidRequest", "Workbook print options require workbook input.");
      const content = signal ? settings : request.options;
      output = request.format === "pdf" ? await converter.renderPdf(model, content) : await converter.renderPage(model, content);
    }
    checkCancellation();
    return { protocolVersion: 1, taskId, ok: true, output: { bytesBase64: Buffer.from(output.bytes).toString("base64"), mimeType: output.mimeType, report: output.report, engine: output.engine } };
  } catch (reason) {
    return failure(signal?.aborted ? new ConversionError("Cancelled", "Conversion was cancelled.") : reason, taskId);
  }
}

/** NDJSON byte transport with backpressure and bounded frames; no file paths or URLs. */
export async function serveConversionProtocol(input: AsyncIterable<Uint8Array>, write: (line: string) => Promise<void>, converter: ProtocolConverter, limits: ProtocolLimits = {}, signal?: AbortSignal): Promise<void> {
  const maximum = limits.maxMessageBytes ?? 96 * 1024 * 1024, outputMaximum = limits.maxOutputMessageBytes ?? 96 * 1024 * 1024;
  if (!integer(maximum, 1) || !integer(outputMaximum, 1)) throw new ConversionError("InvalidRequest", "Protocol frame limits must be positive safe integers.");
  const decoder = new TextDecoder("utf-8", { fatal: true }); let pending = "", pendingBytes = 0;
  async function emit(response: ConversionResponseV1): Promise<void> {
    let serialized = JSON.stringify(response);
    if (Buffer.byteLength(serialized) > outputMaximum) {
      response = failure(new ConversionError("ResourceLimitExceeded", "Conversion response exceeds its frame budget."), response.taskId); serialized = JSON.stringify(response);
      if (Buffer.byteLength(serialized) > outputMaximum) throw new ConversionError("ResourceLimitExceeded", "Response budget is too small for a bounded error frame.");
    }
    await write(serialized + "\n");
  }
  async function respond(line: string): Promise<void> {
    let value: unknown;
    try { value = JSON.parse(line); } catch { await emit(failure(new ConversionError("InvalidRequest", "Invalid conversion JSON."), null)); return; }
    const response = await executeConversionRequest(converter, value, limits, signal);
    if (!signal?.aborted) await emit(response);
  }
  for await (const chunk of input) {
    if (signal?.aborted) return;
    let offset = 0;
    for (let index = 0; index < chunk.length; index++) {
      if (chunk[index] !== 10) continue;
      pendingBytes += index - offset;
      if (pendingBytes > maximum) throw new ConversionError("ResourceLimitExceeded", "Conversion request exceeds its frame budget.");
      pending += decoder.decode(chunk.subarray(offset, index), { stream: true });
      pending += decoder.decode(); await respond(pending); pending = ""; pendingBytes = 0; offset = index + 1;
      if (signal?.aborted) return;
    }
    pendingBytes += chunk.length - offset;
    if (pendingBytes > maximum) throw new ConversionError("ResourceLimitExceeded", "Conversion request exceeds its frame budget.");
    pending += decoder.decode(chunk.subarray(offset), { stream: true });
  }
  pending += decoder.decode(); if (pending.length && !signal?.aborted) await respond(pending);
}
