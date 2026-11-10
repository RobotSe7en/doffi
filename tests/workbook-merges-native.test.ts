import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

function apply(session: WorkbookSession, operations: DocumentOperation[]) {
  return session.commit(session.prepare({ protocolVersion: 1, documentId: "merge-native", baseRevision: session.snapshot().manifest.revision, operations }));
}
test("native merges unmerge and restore while exact cell bytes and package reopening remain intact", async () => {
  const session = createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/03-structure.xlsx"), { documentId: "merge-native" }), { kind: "workbook" }));
  try {
    const original = session.snapshot(), sheet = original.metadata.workbook.sheets[0]!; expect(sheet.merges.length).toBeGreaterThan(0);
    apply(session, sheet.merges.map(range => ({ op: "unmergeWorkbookCells", sheetId: sheet.sheetId, range })));
    expect(session.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([]); expect(session.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
    apply(session, sheet.merges.map(range => ({ op: "mergeWorkbookCells", sheetId: sheet.sheetId, range })));
    expect(session.snapshot().metadata).toEqual(original.metadata); expect(session.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
    expect(openDocument(readDocumentZip(writeDocumentZip(session.serialize())), { kind: "workbook" }).snapshot()).toEqual(session.snapshot());
    session.undo(); session.undo(); expect(session.snapshot().metadata).toEqual(original.metadata); expect(session.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
  } finally { session.dispose(); }
});

test.each(["09-classic-notes", "10-modern-threads"])("native %s keeps comment identities, anchors and rich stories through an explicit merge and undo", async name => {
  const session = createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/" + name + ".xlsx"), { documentId: "merge-native" }), { kind: "workbook" }));
  try {
    const original = session.snapshot(), sheetId = original.metadata.workbook.activeSheetId, bodies = original.annotations.map(annotation => annotation.bodyNodeId).filter((id): id is string => !!id).map(id => ({ id, node: session.getNode(id) }));
    apply(session, [{ op: "mergeWorkbookCells", sheetId, range: { from: { row: 0, column: 0 }, to: { row: 0, column: 1 } }, coveredContent: "clear" }]);
    const reopened = openDocument(readDocumentZip(writeDocumentZip(session.serialize())), { kind: "workbook" });
    expect(reopened.snapshot().metadata.workbook.comments).toEqual(original.metadata.workbook.comments);
    expect(reopened.snapshot().annotations.map(({ baseRevision: _revision, ...annotation }) => annotation)).toEqual(original.annotations.map(({ baseRevision: _revision, ...annotation }) => annotation));
    for (const body of bodies) { expect(body.node).toBeDefined(); expect(reopened.getNode(body.id)).toEqual(body.node); }
    expect(session.history().undoSteps).toBe(1); session.undo();
    expect(session.snapshot().manifest.chunks).toEqual(original.manifest.chunks); expect(session.snapshot().metadata).toEqual(original.metadata);
  } finally { session.dispose(); }
});
