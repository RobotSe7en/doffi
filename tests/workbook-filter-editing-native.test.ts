import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument, parseContract } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

function apply(session: WorkbookSession, operations: DocumentOperation[]) {
  return session.commit(session.prepare({ protocolVersion: 1, documentId: "filter-native", baseRevision: session.snapshot().manifest.revision, operations }));
}
test("native filter removal/restoration preserves source cells, manual dimensions, definitions and package reopening", async () => {
  const session = createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/12-multi-sheet-filter.xlsx"), { documentId: "filter-native" }), { kind: "workbook" }));
  try {
    const original = session.snapshot(), source = original.metadata.workbook.sheets.find(sheet => sheet.autoFilter); expect(source?.autoFilter).toBeDefined();
    const sheetId = source!.sheetId, restore = parseContract("document-operation", { op: "setWorkbookAutoFilter", sheetId, filter: source!.autoFilter }), before = session.serialize();
    expect(apply(session, [{ op: "setWorkbookAutoFilter", sheetId, filter: null }, restore]).changed).toBe(false); expect(session.serialize()).toEqual(before);
    apply(session, [{ op: "setWorkbookAutoFilter", sheetId, filter: null }]);
    expect(session.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)?.autoFilter).toBeUndefined();
    expect(session.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
    expect(session.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)?.rowDimensions).toEqual(source!.rowDimensions);
    const reopened = openDocument(readDocumentZip(writeDocumentZip(session.serialize())), { kind: "workbook" }); expect(reopened.snapshot()).toEqual(session.snapshot());
    apply(session, [restore]); expect(session.snapshot().metadata).toEqual(original.metadata); expect(session.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
    expect(session.history().undoSteps).toBe(2); session.undo(); session.undo(); expect(session.snapshot().metadata).toEqual(original.metadata);
  } finally { session.dispose(); }
});
