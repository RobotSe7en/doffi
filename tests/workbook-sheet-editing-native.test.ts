import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
function commit(session: WorkbookSession, operations: DocumentOperation[]) {
  return session.commit(session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: session.snapshot().manifest.revision, operations }));
}
test.each(["04-formulas-names", "07-three-charts", "08-images-shapes", "09-classic-notes", "10-modern-threads"])("native %s supports rename/deletion and exact history recovery", async name => {
  const input = readFileSync(new URL("../examples/xlsx/" + name + ".xlsx", import.meta.url));
  const session = createDocumentSession(openDocument(await importXlsx(input, { documentId: "sheet-native" }), { kind: "workbook" }));
  const original = session.snapshot(), sheet = original.metadata.workbook.sheets[0]!;
  commit(session, [{ op: "insertWorkbookSheet", name: "Retained" }]);
  commit(session, [{ op: "updateWorkbookSheet", sheetId: sheet.sheetId, changes: { name: "改名 O'Brien" } }]);
  expect(session.snapshot().metadata.workbook.sheets[0]!.sheetId).toBe(sheet.sheetId);
  const renamed = session.snapshot();
  commit(session, [{ op: "removeWorkbookSheet", sheetId: sheet.sheetId }]);
  expect(session.snapshot().manifest.chunks.every(chunk => chunk.sheetId !== sheet.sheetId)).toBe(true);
  expect(session.snapshot().metadata.workbook.objects?.every(object => object.sheetId !== sheet.sheetId) ?? true).toBe(true);
  expect(session.getAnnotations().filter(annotation => annotation.anchor.kind === "cells" && annotation.anchor.sheetId === sheet.sheetId).every(annotation => annotation.status === "needs_reanchor")).toBe(true);
  expect(openDocument(session.serialize(), { kind: "workbook" }).snapshot()).toEqual(session.snapshot());
  session.undo();
  expect(session.snapshot().metadata).toEqual(renamed.metadata); expect(session.snapshot().manifest.chunks).toEqual(renamed.manifest.chunks);
  session.undo(); session.undo();
  expect(session.snapshot().metadata).toEqual(original.metadata); expect(session.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
  expect(session.snapshot().annotations.map(({ baseRevision: _revision, ...annotation }) => annotation)).toEqual(original.annotations.map(({ baseRevision: _revision, ...annotation }) => annotation));
  session.dispose();
});
test("renaming an independently authored formula target retains its historical cache through package ZIP reopen", async () => {
  const input = readFileSync(new URL("../examples/xlsx/04-formulas-names.xlsx", import.meta.url));
  const model = openDocument(await importXlsx(input, { documentId: "native-formula-cache" }), { kind: "workbook" });
  const session = createDocumentSession(model), before = model.snapshot();
  try {
    const owner = before.metadata.workbook.sheets.find(sheet => sheet.name === "Data")!, target = before.metadata.workbook.sheets.find(sheet => sheet.name === "基础 2")!;
    const original = session.getCell(owner.sheetId, 1, 3)!;
    expect(original).toMatchObject({ value: { type: "number", value: 7 }, formula: { expression: "='基础 2'!$B$2+A2", cacheState: "imported" } });
    commit(session, [{ op: "updateWorkbookSheet", sheetId: target.sheetId, changes: { name: "历史 O'Brien" } }]);
    const current = session.getCell(owner.sheetId, 1, 3)!;
    expect(current).toMatchObject({ value: original.value, text: original.text, formula: { expression: "=('历史 O''Brien'!$B$2+A2)", cacheState: "stale" } });
    session.calculateSync(); expect(session.getCalculatedCell(owner.sheetId, 1, 3)).toMatchObject({ state: "valid", value: { type: "number", value: 7 } });
    const reopened = openDocument(await readDocumentZip(await writeDocumentZip(session.serialize())), { kind: "workbook" });
    expect(reopened.getCell(owner.sheetId, 1, 3)).toEqual(current);
    expect(session.history().undoSteps).toBe(1); session.undo();
    expect(session.getCell(owner.sheetId, 1, 3)?.html).toBe(original.html);
    expect(session.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { session.dispose(); }
});
