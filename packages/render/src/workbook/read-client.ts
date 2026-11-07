import { DocumentError, DocumentErrorDataSchema } from "@doffi/core";
import type { WorkbookRuleReadRequest, WorkbookRuleReadResponse } from "@doffi/core";
import type { WorkbookPreviewSource } from "./preview-contracts.js";
import type { WorkbookWindowRequest, WorkbookWindowResponse } from "./window-source.js";
import { checkWindowIdentity } from "./window-source.js";
import type { WorkbookAssetRequest, WorkbookAssetResponse } from "./assets.js";
import type { WorkbookChartReadRequest, WorkbookChartReadResponse } from "./chart-source.js";
import { isRecord, keys, requestId, transportConfig, transportReadRequest, transportSnapshot, WORKBOOK_READ_PROTOCOL } from "./read-transport.js";
import type { WorkbookReadOperation, WorkbookReadTransportOptions } from "./read-transport.js";

export interface WorkbookReadConnection extends WorkbookPreviewSource { dispose(): void }
interface PendingRead { readonly operation: string; resolve(value: unknown): void; reject(error: unknown): void; clean(): void }
/** Connect to the host's dedicated authorized port. Only its metadata snapshot
 * and requested windows cross to this side; no package chunks are requested.
 * Disposal closes this port, never the host-owned Worker. */
export async function connectWorkbookReads(port: MessagePort, options: WorkbookReadTransportOptions = {}, signal?: AbortSignal): Promise<WorkbookReadConnection> {
  const config = transportConfig(options), pending = new Map<number, PendingRead>(); let disposed = false, nextId = 0, currentValues = false;
  function dispose(error = new DocumentError("OperationCancelled", "The workbook read connection was disposed.")): void {
    if (disposed) return; disposed = true;
    try { port.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, event: "close" }); } catch { /* A failed port is still released. */ }
    for (const [id, read] of pending) { try { port.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, id, operation: "cancel" }); } catch { /* Port may already be unavailable. */ } read.clean(); read.reject(error); }
    pending.clear(); port.removeEventListener("message", receive); port.removeEventListener("messageerror", invalidMessage); port.close();
  }
  function invalidMessage(): void { dispose(new DocumentError("InvalidDocument", "The workbook transport returned an invalid response.")); }
  function receive(event: MessageEvent<unknown>): void {
    const message = event.data;
    if (!isRecord(message) || message.protocol !== WORKBOOK_READ_PROTOCOL) return;
    if (message.event === "close" && keys(message, ["protocol", "event"])) { dispose(); return; }
    if (!requestId(message.id) || typeof message.ok !== "boolean" || !keys(message, message.ok ? ["protocol", "id", "ok", "value", "currentValues"] : ["protocol", "id", "ok", "error"])) { invalidMessage(); return; }
    const read = pending.get(message.id); if (!read) return; // Cancelled/timed-out responses cannot become a newer window.
    if (message.currentValues !== undefined && (read.operation !== "snapshot" || typeof message.currentValues !== "boolean")) { invalidMessage(); return; }
    if (read.operation === "snapshot") currentValues = message.currentValues === true;
    pending.delete(message.id); read.clean();
    if (message.ok) read.resolve(message.value);
    else {
      const parsed = DocumentErrorDataSchema.safeParse(message.error);
      if (!parsed.success) { read.reject(new DocumentError("InvalidDocument", "The workbook transport returned an invalid error.")); invalidMessage(); }
      else read.reject(new DocumentError(parsed.data.code, parsed.data.message, parsed.data.details));
    }
  }
  function request(operation: WorkbookReadOperation | "snapshot", payload: unknown, signal?: AbortSignal): Promise<unknown> {
    if (disposed) return Promise.reject(new DocumentError("InvalidOperation", "The workbook read connection has been disposed."));
    if (signal?.aborted) return Promise.reject(new DocumentError("OperationCancelled", "The workbook transport read was cancelled."));
    if (pending.size >= config.maxPendingRequests || nextId >= Number.MAX_SAFE_INTEGER) return Promise.reject(new DocumentError("ResourceLimitExceeded", "Too many pending workbook transport reads."));
    if (operation !== "snapshot") { try { transportReadRequest(operation, payload, config); } catch (error) { return Promise.reject(error); } }
    const id = ++nextId;
    return new Promise((resolve, reject) => {
      const cancel = (): void => {
        if (!pending.delete(id)) return; clean();
        try { port.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, id, operation: "cancel" }); } catch { /* Local cancellation still settles. */ }
        reject(new DocumentError("OperationCancelled", "The workbook transport read was cancelled or exceeded its deadline."));
      };
      const timer = setTimeout(cancel, config.requestTimeoutMs);
      const clean = (): void => { clearTimeout(timer); signal?.removeEventListener("abort", cancel); };
      pending.set(id, { operation, resolve, reject, clean }); signal?.addEventListener("abort", cancel, { once: true });
      try { port.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, id, operation, ...(operation === "snapshot" ? {} : { payload }) }); }
      catch { pending.delete(id); clean(); reject(new DocumentError("InvalidOperation", "The workbook transport could not send this read.")); }
    });
  }
  port.addEventListener("message", receive); port.addEventListener("messageerror", invalidMessage); port.start();
  try {
    const value = await request("snapshot", undefined, signal);
    if (signal?.aborted) throw new DocumentError("OperationCancelled", "The workbook transport connection was cancelled.");
    const snapshot = transportSnapshot(value, config, true);
    const read = async <T>(operation: WorkbookReadOperation, input: WorkbookWindowRequest | WorkbookAssetRequest | WorkbookChartReadRequest | WorkbookRuleReadRequest, signal: AbortSignal): Promise<T> => {
      checkWindowIdentity(input, snapshot);
      const result = await request(operation, input, signal);
      if (!isRecord(result) || typeof result.documentId !== "string" || !Number.isSafeInteger(result.revision)) throw new DocumentError("InvalidDocument", "Workbook transport returned an invalid read identity.");
      checkWindowIdentity(result as { documentId: string; revision: number }, snapshot);
      return result as T;
    };
    return Object.freeze({ kind: "workbook", cells: Object.freeze({ snapshot, allowReadAhead: true,
      read: (input: WorkbookWindowRequest, signal: AbortSignal) => read<WorkbookWindowResponse>("cells", input, signal) }),
    assets: Object.freeze({ snapshot, readAsset: (input: WorkbookAssetRequest, signal: AbortSignal) => read<WorkbookAssetResponse>("asset", input, signal) }),
    charts: Object.freeze({ snapshot, read: (input: WorkbookChartReadRequest, signal: AbortSignal) => read<WorkbookChartReadResponse>("charts", input, signal) }),
    rules: Object.freeze({ snapshot, ...(currentValues ? { currentValues: true } : {}), read: (input: WorkbookRuleReadRequest, signal: AbortSignal) => read<WorkbookRuleReadResponse>("rules", input, signal) }),
    dispose: () => dispose() });
  } catch (error) { dispose(); throw error; }
}
