import { copyWorkbookScalarContext, DocumentError } from "@doffi/core";
import type { WorkbookChangeImpact, WorkbookScalarContext } from "@doffi/core";
import { connectWorkbookReads, linkWorkbookPreviewSource } from "@doffi/render";
import type { WorkbookPreviewSource, WorkbookReadConnection } from "@doffi/render";
import { checkSessionMessage, invalidSessionTransport, keys, record, WORKBOOK_SESSION_PROTOCOL } from "./session-transport.js";
import type { sessionTransportConfig, WorkbookSessionIdentity, WorkbookSessionPreparationExecution } from "./session-transport.js";

/** Only authorized window/rule queries cross these temporary read ports.
 * Closing an exchange releases its view, never the owning Worker session. */
export interface WorkbookLayoutHistory { source?: WorkbookPreviewSource; sourceDigest?: string }
export async function answerWorkbookLayout(value: unknown, reads: MessagePort, reply: MessagePort,
  measure: NonNullable<WorkbookSessionPreparationExecution["measureLayout"]>, signal: AbortSignal,
  config: ReturnType<typeof sessionTransportConfig>, identity: WorkbookSessionIdentity, history: WorkbookLayoutHistory): Promise<void> {
  let source: WorkbookReadConnection | undefined;
  const check = (): void => { if (signal.aborted) throw new DocumentError("OperationCancelled", "Workbook layout measurement was cancelled."); };
  try {
    check(); checkSessionMessage(value, config);
    if (!record(value) || !keys(value, ["sourceDigest", "sheetIds", "calculationContext", "previousSourceDigest", "impact"]) || typeof value.sourceDigest !== "string" || !value.sourceDigest.length
      || (value.previousSourceDigest === undefined) !== (value.impact === undefined) || value.previousSourceDigest !== undefined && typeof value.previousSourceDigest !== "string"
      || !Array.isArray(value.sheetIds) || !value.sheetIds.length || value.sheetIds.some(id => typeof id !== "string" || !id) || new Set(value.sheetIds).size !== value.sheetIds.length)
      invalidSessionTransport("Invalid workbook layout source request.");
    const calculationContext = value.calculationContext === undefined ? undefined : copyWorkbookScalarContext(value.calculationContext as WorkbookScalarContext);
    source = await connectWorkbookReads(reads, config.reads, signal); check();
    const snapshot = source.cells.snapshot;
    if (snapshot.manifest.documentId !== identity.documentId || snapshot.manifest.revision !== identity.revision)
      throw new DocumentError("RevisionConflict", "Workbook layout source belongs to another preparation revision.");
    const measured = history.source && value.previousSourceDigest === history.sourceDigest
      ? linkWorkbookPreviewSource({ ...source, impact: value.impact as WorkbookChangeImpact }, history.source, value.impact as WorkbookChangeImpact) : source;
    const result = await measure(Object.freeze({ source: measured, sourceDigest: value.sourceDigest, sheetIds: Object.freeze([...value.sheetIds]) as readonly string[],
      ...(calculationContext ? { calculationContext } : {}), signal }));
    check(); checkSessionMessage(result, config); reply.postMessage({ protocol: WORKBOOK_SESSION_PROTOCOL, ok: true, value: result });
    history.source = measured; history.sourceDigest = value.sourceDigest;
  } catch (reason) {
    const error = reason instanceof DocumentError ? reason : new DocumentError("InvalidOperation", "The authorized workbook layout measurer could not complete.");
    try { const data = error.toJSON(); checkSessionMessage(data, config); reply.postMessage({ protocol: WORKBOOK_SESSION_PROTOCOL, ok: false, error: data }); }
    catch { try { reply.postMessage({ protocol: WORKBOOK_SESSION_PROTOCOL, ok: false, error: new DocumentError("ResourceLimitExceeded", "Workbook layout error exceeds the reply budget.").toJSON() }); } catch { /* Its owning preparation may already be closed. */ } }
  } finally { source?.dispose(); reads.close(); reply.close(); }
}
