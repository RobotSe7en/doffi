import { createDocumentSession, DocumentError, openDocument } from "@doffi/core";
import { importXlsx, readDocumentZip } from "@doffi/office";
import { serveWorkbookSession } from "@doffi/editor";
import type { WorkbookOpenRequest, WorkbookOpenResponse } from "./open-workbook.js";

const scope = globalThis as unknown as {
  onmessage: ((event: MessageEvent<WorkbookOpenRequest>) => void) | null;
  postMessage(message: WorkbookOpenResponse, transfer?: Transferable[]): void;
};
let opened = false;
scope.onmessage = event => {
  if (opened) return; opened = true;
  const request = event.data;
  void (async () => {
    try {
      const files = request.kind === "xlsx" ? await importXlsx(request.bytes, { documentId: request.documentId,
        onProgress(progress) { scope.postMessage({ taskId: request.taskId, kind: "progress", progress }); } }) : readDocumentZip(request.bytes);
      const model = openDocument(files);
      if (model.kind === "workbook") {
        const session = createDocumentSession(model, { sessionLimits: { maxCandidateBytes: 512 * 1024 * 1024, maxTransactionBytes: 512 * 1024 * 1024, maxHistoryBytes: 512 * 1024 * 1024 } });
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
