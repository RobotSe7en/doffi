import { createDocumentSession, DocumentError, openDocument } from "@doffi/core";
import type { WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip } from "@doffi/office";
import { serveWorkbookSession } from "@doffi/editor";
import type { WorkbookOpenRequest, WorkbookOpenResponse } from "./open-workbook.js";
import { runWorkbookOutput } from "./workbook-output.js";
import type { CancelWorkbookOutput, WorkbookOutputRequest, WorkbookOutputResponse } from "./workbook-output.js";

const scope = globalThis as unknown as {
  onmessage: ((event: MessageEvent<WorkbookOpenRequest | WorkbookOutputRequest | CancelWorkbookOutput>) => void) | null;
  postMessage(message: WorkbookOpenResponse | WorkbookOutputResponse, transfer?: Transferable[]): void;
};
let opened = false, session: WorkbookSession | undefined;
let output: { taskId: string; controller: AbortController; completion: Promise<void> } | undefined;

function exportWorkbook(current: WorkbookSession, request: WorkbookOutputRequest): void {
  const previous = output;
  previous?.controller.abort();
  const controller = new AbortController(), { taskId } = request;
  const completion = (async () => {
    await previous?.completion;
    if (controller.signal.aborted) return;
    try {
      const result = await runWorkbookOutput(current, request, controller.signal, progress => {
        scope.postMessage({ taskId, kind: "output-progress", progress });
      });
      if (!controller.signal.aborted) {
        const bytes = result.bytes.buffer instanceof ArrayBuffer ? result.bytes : new Uint8Array(result.bytes);
        scope.postMessage({ taskId, kind: "output", result: { ...result, bytes } }, [bytes.buffer]);
      }
    } catch (reason) {
      if (controller.signal.aborted) return;
      const error = reason instanceof DocumentError ? reason : new DocumentError("ExportFailed", "无法导出工作簿，请重试。");
      scope.postMessage({ taskId, kind: "output-error", error: error.toJSON() });
    }
  })().finally(() => { if (output?.controller === controller) output = undefined; });
  output = { taskId, controller, completion };
}

scope.onmessage = event => {
  const request = event.data;
  if (request.kind === "cancel-export") {
    if (output?.taskId === request.taskId) output.controller.abort();
    return;
  }
  if (request.kind === "export") {
    if (session) exportWorkbook(session, request);
    return;
  }
  if (opened) return; opened = true;
  void (async () => {
    try {
      const files = request.kind === "xlsx" ? await importXlsx(request.bytes, { documentId: request.documentId,
        onProgress(progress) { scope.postMessage({ taskId: request.taskId, kind: "progress", progress }); } }) : readDocumentZip(request.bytes);
      const model = openDocument(files);
      if (model.kind === "workbook") {
        session = createDocumentSession(model, { sessionLimits: { maxCandidateBytes: 512 * 1024 * 1024, maxTransactionBytes: 512 * 1024 * 1024, maxHistoryBytes: 512 * 1024 * 1024 } });
        serveWorkbookSession(request.port, session);
      }
      else {
        const entries = [...files].map(([path, value]) => [path, new Uint8Array(value)] as [string, Uint8Array<ArrayBuffer>]);
        scope.postMessage({ taskId: request.taskId, kind: "content", files: entries }, entries.map(([, value]) => value.buffer));
        request.port.close();
      }
    } catch (reason) {
      const error = reason instanceof DocumentError ? reason : new DocumentError("InvalidDocument", "无法读取工作簿。");
      scope.postMessage({ taskId: request.taskId, kind: "error", error: error.toJSON() }); request.port.close();
    }
  })();
};
