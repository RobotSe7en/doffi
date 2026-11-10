import { DocumentError } from "@doffi/core";
import type { WorkbookCalculationResult, WorkbookModel, WorkbookRuleReadRequest, WorkbookSnapshot } from "@doffi/core";
import { modelRuleSource } from "./rule-source.js";
import { modelAssetSource } from "./assets.js";
import type { WorkbookAssetRequest } from "./assets.js";
import { modelChartSource } from "./chart-source.js";
import type { WorkbookChartReadRequest } from "./chart-source.js";
import { modelWindowSource } from "./window-source.js";
import type { WorkbookWindowRequest } from "./window-source.js";
import { isRecord, keys, requestId, transportConfig, transportError, transportReadRequest, transportSnapshot, WORKBOOK_READ_PROTOCOL } from "./read-transport.js";
import type { WorkbookReadTransportOptions } from "./read-transport.js";

export interface WorkbookReadServer { readonly disposed: boolean; dispose(): void }
/** Bind one exclusive host-authorized port to one immutable model. Importing,
 * switching models, owning/terminating Workers and permissions stay with the
 * host. Document IDs in requests can never select another server-side model. */
export function serveWorkbookReads(port: MessagePort, model: WorkbookModel, options: WorkbookReadTransportOptions = {}, calculation?: WorkbookCalculationResult): WorkbookReadServer {
  if (model.kind !== "workbook") throw new DocumentError("UnsupportedFormat", "Workbook read servers require a workbook model.");
  const config = transportConfig(options);
  let snapshot: WorkbookSnapshot | undefined = transportSnapshot(model.snapshot(), config);
  let providers: { cells: ReturnType<typeof modelWindowSource>; assets: ReturnType<typeof modelAssetSource>; charts: ReturnType<typeof modelChartSource>; rules: ReturnType<typeof modelRuleSource> } | undefined
    = { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model, calculation), rules: modelRuleSource(model, calculation) };
  let disposed = false, highestId = 0, snapshotSent = false;
  const pending = new Map<number, AbortController>();
  const post = (id: number, ok: boolean, value: unknown, transfer: Transferable[] = [], currentValues?: boolean): void => {
    if (disposed) return;
    try { port.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, id, ok, ...(ok ? { value } : { error: transportError(value).toJSON() }), ...(currentValues === undefined ? {} : { currentValues }) }, transfer); }
    catch { dispose(); }
  };
  const receive = (event: MessageEvent<unknown>): void => {
    const message = event.data;
    if (isRecord(message) && message.protocol === WORKBOOK_READ_PROTOCOL && message.event === "close" && keys(message, ["protocol", "event"])) { dispose(); return; }
    if (!isRecord(message) || message.protocol !== WORKBOOK_READ_PROTOCOL || !requestId(message.id)) return;
    const { id, operation } = message;
    if (!keys(message, ["protocol", "id", "operation", "payload"])) { post(id, false, new DocumentError("InvalidOperation", "Unknown workbook transport request fields.")); return; }
    if (operation === "cancel") { if (!Object.hasOwn(message, "payload")) pending.get(id)?.abort(); return; }
    if (id <= highestId) { post(id, false, new DocumentError("InvalidOperation", "Workbook transport request IDs must increase.")); return; }
    highestId = id;
    if (pending.size >= config.maxPendingRequests) { post(id, false, new DocumentError("ResourceLimitExceeded", "Too many pending workbook transport reads.")); return; }
    if (operation === "snapshot") {
      if (Object.hasOwn(message, "payload")) post(id, false, new DocumentError("InvalidOperation", "Workbook snapshot requests have no payload."));
      else if (snapshotSent) post(id, false, new DocumentError("InvalidOperation", "This connection already received its immutable snapshot."));
      else { snapshotSent = true; post(id, true, snapshot, [], providers?.rules.currentValues); }
      return;
    }
    const owned = providers; if (!owned) return;
    try {
      if (operation !== "cells" && operation !== "asset" && operation !== "charts" && operation !== "rules") throw new DocumentError("InvalidOperation", "Unknown workbook transport read operation.");
      transportReadRequest(operation, message.payload, config);
      const abort = new AbortController(); pending.set(id, abort);
      const timer = setTimeout(() => { if (!abort.signal.aborted) { post(id, false, new DocumentError("OperationCancelled", "Workbook transport read exceeded its deadline.")); abort.abort(); } }, config.requestTimeoutMs);
      void (async () => {
        try {
          if (operation === "asset") {
            const result = await owned.assets.readAsset(message.payload as WorkbookAssetRequest, abort.signal);
            if (!abort.signal.aborted) post(id, true, result, [result.bytes.buffer as ArrayBuffer]);
          } else {
            const result = operation === "cells" ? await owned.cells.read(message.payload as WorkbookWindowRequest, abort.signal)
              : operation === "charts" ? await owned.charts.read(message.payload as WorkbookChartReadRequest, abort.signal)
              : await owned.rules.read(message.payload as WorkbookRuleReadRequest, abort.signal);
            if (!abort.signal.aborted) post(id, true, result);
          }
        } catch (error) { if (!abort.signal.aborted) post(id, false, error); }
        finally { clearTimeout(timer); pending.delete(id); }
      })();
    } catch (error) { post(id, false, error); }
  };
  const invalidMessage = (): void => dispose();
  function dispose(): void {
    if (disposed) return; disposed = true;
    try { port.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, event: "close" }); } catch { /* A failed port is still released. */ }
    for (const abort of pending.values()) abort.abort(); pending.clear(); providers = undefined; snapshot = undefined;
    port.removeEventListener("message", receive); port.removeEventListener("messageerror", invalidMessage); port.close();
  }
  port.addEventListener("message", receive); port.addEventListener("messageerror", invalidMessage); port.start();
  return Object.freeze({ get disposed() { return disposed; }, dispose });
}
