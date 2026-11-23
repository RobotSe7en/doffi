import { createDocumentSession, openDocument } from "@doffi/core";
import { exportXlsx, importXlsx, writeDocumentZip } from "@doffi/office";
import { mountDocumentPreview } from "@doffi/render";

export async function startWorkbookHost(mount) {
  const response = await fetch("./source.xlsx");
  if (!response.ok) throw Error(`Could not read the source XLSX (${response.status}).`);
  const files = await importXlsx(new Uint8Array(await response.arrayBuffer()), { documentId: "packed-workbook" });
  const model = openDocument(files, { kind: "workbook" });
  const session = createDocumentSession(model);
  const sheetId = model.snapshot().metadata.workbook.sheets[0].sheetId;
  const preview = mountDocumentPreview(document.querySelector("#preview"), model, {
    title: "独立工作簿只读预览", widthPt: 420, heightPt: 220,
  });
  const errors = [];
  const { view, unmount } = await mount(document.querySelector("#workbench"), session, {
    widthPt: 420,
    heightPt: 240,
    allowGlobalScope: true,
    onError(error) { errors.push(error.message); },
    onSubmitAnnotations(request) {
      return {
        protocolVersion: 1,
        requestId: request.requestId,
        documentId: request.documentId,
        baseRevision: request.baseRevision,
        annotationIds: request.annotations.map(annotation => annotation.annotationId),
        proposalId: "packed-formula",
        summary: "引用人工输入并计算三倍。",
        operations: [{ op: "setWorkbookCells", sheetId, cells: [
          { row: 4, column: 1, source: { kind: "formula", expression: "=A5*3" } },
        ] }],
      };
    },
  });
  await Promise.all([preview.ready, view.ready]);
  await view.settled();

  window.host = {
    async selectPreviewCell() {
      const point = { row: 0, column: 0 };
      await preview.selectRange({ from: point, to: point });
      return preview.layout();
    },
    inspect() {
      return {
        revision: session.snapshot().manifest.revision,
        input: session.getCell(sheetId, 4, 0)?.value?.value ?? null,
        formula: session.getCell(sheetId, 4, 1)?.formula?.expression ?? null,
        undoSteps: session.history().undoSteps,
        errors,
      };
    },
    async save() {
      const xlsx = await exportXlsx(session);
      return { xlsx: [...xlsx.bytes], report: xlsx.report, zip: [...writeDocumentZip(session.serialize())] };
    },
    dispose() {
      unmount();
      preview.dispose();
      session.dispose();
    },
  };
}
