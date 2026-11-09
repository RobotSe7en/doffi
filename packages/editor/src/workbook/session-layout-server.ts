import { DocumentError, DocumentErrorDataSchema, readWorkbookChangeImpact } from "@doffi/core";
import type { WorkbookCalculationResult, WorkbookLayoutMeasurement, WorkbookLayoutRequest } from "@doffi/core";
import { serveWorkbookReads } from "@doffi/render";
import { checkSessionMessage, invalidSessionTransport, keys, record, WORKBOOK_SESSION_PROTOCOL } from "./session-transport.js";
import type { sessionTransportConfig } from "./session-transport.js";

/** Each measurement receives two capabilities bound to this exact private
 * model: bounded source reads and one reply. Neither port can commit edits. */
export function requestWorkbookLayout(request: WorkbookLayoutRequest, calculation: WorkbookCalculationResult | undefined, signal: AbortSignal,
  config: ReturnType<typeof sessionTransportConfig>, send: (value: unknown, reads: MessagePort, reply: MessagePort) => void,
  previous?: { readonly request: WorkbookLayoutRequest; readonly calculation?: WorkbookCalculationResult }): Promise<WorkbookLayoutMeasurement> {
  if (signal.aborted) return Promise.reject(new DocumentError("OperationCancelled", "Workbook layout measurement was cancelled."));
  const value: Record<string, unknown> = { sourceDigest: request.sourceDigest, sheetIds: request.sheetIds, ...(request.calculationContext ? { calculationContext: request.calculationContext } : {}) };
  if (previous) {
    const hint = { previousSourceDigest: previous.request.sourceDigest, impact: readWorkbookChangeImpact(previous.request.model, request.model, previous.calculation, calculation) };
    try { checkSessionMessage({ ...value, ...hint }, config); Object.assign(value, hint); }
    catch (error) { if (!(error instanceof DocumentError) || error.code !== "ResourceLimitExceeded") throw error; }
  }
  const reads = new MessageChannel(), reply = new MessageChannel();
  let reader: ReturnType<typeof serveWorkbookReads>;
  try { reader = serveWorkbookReads(reads.port1, request.model, config.reads, calculation); }
  catch (error) { for (const port of [reads.port1, reads.port2, reply.port1, reply.port2]) port.close(); return Promise.reject(error); }
  return new Promise((resolve, reject) => {
    let finished = false;
    const clean = (): void => { signal.removeEventListener("abort", stop); reply.port1.removeEventListener("message", receive); reply.port1.removeEventListener("messageerror", malformed);
      reader.dispose(); reads.port2.close(); reply.port1.close(); reply.port2.close(); };
    const fail = (error: unknown): void => { if (finished) return; finished = true; clean(); reject(error); };
    const stop = (): void => fail(new DocumentError("OperationCancelled", "Workbook layout measurement was cancelled."));
    const malformed = (): void => fail(new DocumentError("InvalidDocument", "Invalid workbook layout measurement reply."));
    function receive(event: MessageEvent<unknown>): void {
      if (finished) return;
      try {
        const message = event.data;
        if (!record(message) || message.protocol !== WORKBOOK_SESSION_PROTOCOL || typeof message.ok !== "boolean"
          || !keys(message, message.ok ? ["protocol", "ok", "value"] : ["protocol", "ok", "error"])) invalidSessionTransport("Invalid workbook layout reply envelope.");
        checkSessionMessage(message.ok ? message.value : message.error, config);
        if (!message.ok) {
          const error = DocumentErrorDataSchema.safeParse(message.error); if (!error.success) invalidSessionTransport("Invalid workbook layout error.");
          fail(new DocumentError(error.data.code, error.data.message, error.data.details)); return;
        }
        finished = true; clean(); resolve(message.value as WorkbookLayoutMeasurement);
      } catch (error) { fail(error); }
    }
    signal.addEventListener("abort", stop, { once: true }); reply.port1.addEventListener("message", receive); reply.port1.addEventListener("messageerror", malformed); reply.port1.start();
    try { checkSessionMessage(value, config); if (signal.aborted) stop(); else send(value, reads.port2, reply.port2); }
    catch (error) { fail(error); }
  });
}
