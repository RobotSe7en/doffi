import { createDocumentSession, DocumentError, openDocument } from "@doffi/core";
import type { WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip } from "@doffi/office";
import { serveWorkbookSession } from "@doffi/editor";
import type { WorkbookSessionServer } from "@doffi/editor";

const scope = globalThis as unknown as { onmessage: ((event: MessageEvent<{ kind: string; format?: "package"; documentId: string; bytes: Uint8Array; port: MessagePort }>) => void) | null; postMessage(value: unknown): void };
let session: WorkbookSession | undefined, server: WorkbookSessionServer | undefined, opened = false, sessionsCreated = 0;
scope.onmessage = event => {
  const request = event.data;
  void (async () => {
    try {
      if (request.kind === "open" && !opened) {
        opened = true;
        const files = request.format === "package" ? readDocumentZip(request.bytes) : await importXlsx(request.bytes, { documentId: request.documentId, onProgress: progress => scope.postMessage({ kind: "progress", progress }) });
        session = createDocumentSession(openDocument(files, { kind: "workbook" }), { sessionLimits: {
          maxCandidateBytes: 512 * 1024 * 1024, maxTransactionBytes: 512 * 1024 * 1024, maxHistoryBytes: 512 * 1024 * 1024 } }); sessionsCreated++;
        server = serveWorkbookSession(request.port, session); return;
      }
      if (request.kind !== "closed-audit" || !session || !server) return;
      const deadline = performance.now() + 5000;
      while (!server.disposed && performance.now() < deadline) await new Promise(resolve => setTimeout(resolve, 0));
      let cells = 0, formulas = 0;
      for (const chunk of session.snapshot().manifest.chunks) for (let row = chunk.row; row < Math.min(chunk.row + 256, 1_048_576); row += 64) {
        const values = session.readWindow({ sheetId: chunk.sheetId, row, column: chunk.column, rows: Math.min(64, 1_048_576 - row), columns: Math.min(32, 16_384 - chunk.column) });
        cells += values.length; formulas += values.filter(cell => cell.formula).length;
      }
      const snapshot = session.snapshot(), reopened = openDocument(session.serialize(), { kind: "workbook" });
      const audit = { kind: "audit", cells, formulas, chunks: snapshot.manifest.chunks, history: session.history(), revision: snapshot.manifest.revision,
        sessionsCreated, endpointDisposed: server.disposed, reopenedEqual: JSON.stringify(reopened.snapshot()) === JSON.stringify(snapshot) };
      // The endpoint owns its port, while this host retains its original core
      // session and can continue editing after the main client disconnects.
      session.commit(session.prepare({ protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision,
        operations: [{ op: "setWorkbookCells", sheetId: snapshot.metadata.workbook.activeSheetId, cells: [{ row: 13000, column: 35, source: { kind: "value", value: { type: "number", value: 42 } } }] }] }));
      scope.postMessage({ ...audit, hostUsable: session.getCell(snapshot.metadata.workbook.activeSheetId, 13000, 35)?.value });
    } catch (reason) {
      const error = reason instanceof DocumentError ? reason : new DocumentError("InvalidDocument", "The workbook editing Worker could not complete its host action.");
      scope.postMessage({ kind: "error", error: error.toJSON() });
    }
  })();
};
