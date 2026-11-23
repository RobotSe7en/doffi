import { DocumentError } from "@doffi/core";
import type { DocumentErrorData, WorkbookScalarContext, WorkbookSession } from "@doffi/core";
import { exportXlsx, writeDocumentZipAsync } from "@doffi/office";
import type { XlsxExportProgress, XlsxExportReport } from "@doffi/office";
import type { ConversionRequestV1, ConversionResponseV1, WorkbookPrintReport, WorkbookPrintSettings } from "@doffi/render/contracts";

export type WorkbookOutputOptions = { readonly format: "zip" }
  | { readonly format: "xlsx"; readonly context: WorkbookScalarContext }
  | { readonly format: "pdf" | "png"; readonly print: WorkbookPrintSettings; readonly pageIndex?: number };
export interface WorkbookOutputRequest {
  readonly kind: "export";
  readonly taskId: string;
  readonly documentId: string;
  readonly revision: number;
  readonly options: WorkbookOutputOptions;
}
export interface WorkbookOutputResult {
  readonly bytes: Uint8Array;
  readonly mimeType: string;
  readonly report?: XlsxExportReport | WorkbookPrintReport;
}
export type WorkbookOutputResponse =
  | { readonly taskId: string; readonly kind: "output"; readonly result: WorkbookOutputResult }
  | { readonly taskId: string; readonly kind: "output-progress"; readonly progress: XlsxExportProgress }
  | { readonly taskId: string; readonly kind: "output-error"; readonly error: DocumentErrorData };
export type CancelWorkbookOutput = { readonly kind: "cancel-export"; readonly taskId: string };

function base64(bytes: Uint8Array): string {
  let value = "";
  for (let offset = 0; offset < bytes.length; offset += 32768)
    value += String.fromCharCode(...bytes.subarray(offset, offset + 32768));
  return btoa(value);
}

export async function runWorkbookOutput(
  session: WorkbookSession,
  request: WorkbookOutputRequest,
  signal: AbortSignal,
  onProgress: (progress: XlsxExportProgress) => void,
): Promise<WorkbookOutputResult> {
  const { taskId, documentId, revision, options } = request;
  const guard = (): void => {
    if (signal.aborted) throw new DocumentError("OperationCancelled", "已取消工作簿导出。");
    const manifest = session.snapshot().manifest;
    if (manifest.documentId !== documentId || manifest.revision !== revision)
      throw new DocumentError("RevisionConflict", "导出期间工作簿已变化，请重新导出。");
  };
  guard();
  if (options.format === "xlsx") {
    const result = await exportXlsx(session, { signal, onProgress, calculationContext: options.context });
    return { ...result, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
  }
  const bytes = await writeDocumentZipAsync(session.serialize(), { kind: "workbook", signal });
  guard();
  if (options.format === "zip") return { bytes, mimeType: "application/zip" };

  const { format, print, pageIndex } = options;
  const conversion: ConversionRequestV1 = {
    protocolVersion: 1, taskId, format,
    input: { format: "document-package-zip", documentId, bytesBase64: base64(bytes) },
    options: { workbook: print, ...(pageIndex === undefined ? {} : { pageIndex }) },
  };
  const response = await fetch(new URL("/__doffi/convert", globalThis.location.origin), {
    method: "POST", signal,
    headers: { "Content-Type": "application/json", "X-Doffi-Dev": "1" },
    body: JSON.stringify(conversion),
  });
  let result: ConversionResponseV1;
  try { result = await response.json() as ConversionResponseV1; }
  catch { throw new DocumentError("ExportFailed", "工作簿转换服务不可用。"); }
  guard();
  if (!response.ok || !result?.ok)
    throw new DocumentError("ExportFailed", result && !result.ok && typeof result.error?.message === "string"
      ? result.error.message.slice(0, 256) : "工作簿转换失败。");
  const { output } = result;
  if (result.protocolVersion !== 1 || result.taskId !== taskId
    || output?.mimeType !== (format === "pdf" ? "application/pdf" : "image/png")
    || typeof output.bytesBase64 !== "string" || !output.bytesBase64
    || output.report?.kind !== "workbook" || !Array.isArray(output.report.pages)
    || output.report.documentId !== documentId || output.report.revision !== revision)
    throw new DocumentError("ExportFailed", "工作簿打印结果与当前请求不匹配。");
  const binary = atob(output.bytesBase64), rendered = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) rendered[index] = binary.charCodeAt(index);
  return { bytes: rendered, mimeType: output.mimeType, report: output.report };
}
