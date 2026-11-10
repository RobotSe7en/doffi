import { DocumentError, openDocument, parseContract } from "@doffi/core";
import type { DocumentErrorData, ContentDocumentModel, WorkbookSnapshot } from "@doffi/core";
import { connectWorkbookSession } from "@doffi/editor";
import type { WorkbookSessionConnection } from "@doffi/editor";
import type { OfficeImportProgress } from "@doffi/office";

export interface WorkbookOpenRequest { readonly kind: "xlsx" | "package"; readonly taskId: string; readonly documentId: string; readonly bytes: Uint8Array; readonly port: MessagePort }
export type WorkbookOpenResponse = { readonly taskId: string; readonly kind: "progress"; readonly progress: OfficeImportProgress }
  | { readonly taskId: string; readonly kind: "content"; readonly files: [string, Uint8Array][] }
  | { readonly taskId: string; readonly kind: "error"; readonly error: DocumentErrorData };
export interface HostedWorkbook {
  readonly kind: "workbook";
  readonly session: WorkbookSessionConnection;
  snapshot(): WorkbookSnapshot;
  acceptView(snapshot: WorkbookSnapshot): void;
  readonly pendingInput: boolean;
  setPendingInput(pending: boolean): void;
  readonly disposed: boolean;
  dispose(): void;
}
export type OpenedWorkspaceDocument = ContentDocumentModel | HostedWorkbook;

/** This application owns the Worker, authorized input bytes and lifetime. The
 * SDK owns the dedicated session connection and bounded read leases. Package classification
 * also runs off the UI thread, so a workbook ZIP never opens on the UI thread. */
export function openWorkbookFile(bytes: Uint8Array, kind: "xlsx" | "package", taskId: string, documentId: string,
  options: { readonly signal: AbortSignal; readonly onProgress?: (value: OfficeImportProgress) => void; readonly onFailure?: (error: Error) => void }
): Promise<OpenedWorkspaceDocument> {
  if (options.signal.aborted) return Promise.reject(new DocumentError("OperationCancelled", "已取消读取。"));
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("./workbook-open-worker.ts", import.meta.url), { type: "module" });
    const channel = new MessageChannel(), controller = new AbortController();
    let settled = false, disposed = false, session: WorkbookSessionConnection | undefined;
    const detachOpen = (): void => { clearTimeout(timer); options.signal.removeEventListener("abort", cancel); };
    const dispose = (): void => {
      if (disposed) return; disposed = true; detachOpen(); controller.abort(); session?.dispose(); session = undefined;
      channel.port1.close(); channel.port2.close(); worker.onmessage = null; worker.onerror = null; worker.onmessageerror = null; worker.terminate();
    };
    const fail = (error: Error): void => {
      if (disposed) return;
      const wasSettled = settled; settled = true; dispose();
      if (wasSettled) options.onFailure?.(error); else reject(error);
    };
    const cancel = (): void => fail(new DocumentError("OperationCancelled", "已取消读取。"));
    const timer = setTimeout(() => fail(new Error("工作簿读取超时，请重试。")), 60_000);
    worker.onerror = () => fail(new Error("工作簿读取进程已停止，请重新打开文件。"));
    worker.onmessageerror = () => fail(new Error("工作簿读取结果无法接收。"));
    worker.onmessage = (event: MessageEvent<WorkbookOpenResponse>) => {
      if (disposed) return;
      try {
        const message = event.data;
        if (!message || message.taskId !== taskId) throw new Error("工作簿读取结果与请求不匹配。");
        if (message.kind === "progress") { if (!settled) options.onProgress?.(message.progress); }
        else if (message.kind === "error") { const error = parseContract("document-error", message.error); fail(new DocumentError(error.code, error.message, error.details)); }
        else if (message.kind === "content" && !settled) {
          const model = openDocument(new Map(message.files));
          if (model.kind === "workbook") throw new Error("工作簿必须保留在所属读取进程内。");
          settled = true; dispose(); resolve(model);
        } else throw new Error("工作簿读取返回了未知响应。");
      } catch (error) { fail(error instanceof Error ? error : new Error("工作簿读取失败。")); }
    };
    options.signal.addEventListener("abort", cancel, { once: true });
    if (options.signal.aborted) { cancel(); return; }
    void connectWorkbookSession(channel.port2, {}, controller.signal).then(async connection => {
      if (disposed || settled) { connection.dispose(); return; }
      session = connection; const initial = await connection.view(undefined, { calculate: false, signal: controller.signal });
      let snapshot = initial.cells.snapshot, pendingInput = false; initial.dispose();
      if (disposed || settled) return; settled = true; detachOpen();
      resolve(Object.freeze({ kind: "workbook", session: connection, snapshot: () => snapshot,
        acceptView(value: WorkbookSnapshot) { if (!disposed && value.manifest.documentId === connection.state().documentId && value.manifest.revision <= connection.state().revision) snapshot = value; },
        get pendingInput() { return pendingInput; }, setPendingInput(value: boolean) { pendingInput = value; }, get disposed() { return disposed; }, dispose }));
    }).catch(error => { if (!disposed && !settled) fail(error instanceof Error ? error : new Error("工作簿读取失败。")); });
    try {
      const input = new Uint8Array(bytes);
      worker.postMessage({ kind, taskId, documentId, bytes: input, port: channel.port1 } satisfies WorkbookOpenRequest, [input.buffer, channel.port1]);
    } catch (error) { fail(error instanceof Error ? error : new Error("工作簿无法开始读取。")); }
  });
}
