import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookConditionalRule, WorkbookSession, WorkbookTable } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

function apply(session: WorkbookSession, operations: DocumentOperation[]) {
  return session.commit(session.prepare({ protocolVersion: 1, documentId: "definitions-native", baseRevision: session.snapshot().manifest.revision, operations }));
}
async function session(name: string) {
  return createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/" + name + ".xlsx"), { documentId: "definitions-native" }), { kind: "workbook" }));
}
test("independent native table no-op/rename keeps cached content, calculated exceptions and exact undo/package reopening", async () => {
  const s = await session("05-native-table");
  try {
    const original = s.snapshot(), sheet = original.metadata.workbook.sheets[0]!, source = sheet.tables![0]!, before = s.serialize();
    const table = JSON.parse(JSON.stringify(source)) as WorkbookTable;
    expect(apply(s, [{ op: "setWorkbookTable", sheetId: sheet.sheetId, table }]).changed).toBe(false); expect(s.serialize()).toEqual(before);
    // A deliberate exception is actual source, not a separate calculated table.
    apply(s, [{ op: "setWorkbookCells", sheetId: sheet.sheetId, cells: [{ row: 2, column: 3, source: { kind: "value", value: { type: "number", value: 123 } } }] }]); s.clearHistory();
    const saved = s.snapshot(), total = s.getCell(sheet.sheetId, 5, 3)!;
    apply(s, [{ op: "setWorkbookTable", sheetId: sheet.sheetId, table: { ...table, name: "CurrentSales" } }]);
    expect(s.getCell(sheet.sheetId, 2, 3)?.value).toEqual({ type: "number", value: 123 });
    expect(s.getCell(sheet.sheetId, 5, 3)).toMatchObject({ value: total.value, formula: { cacheState: "stale" } });
    expect(s.getCell(sheet.sheetId, 1, 3)?.formula?.expression).toContain("CurrentSales");
    const reopened = openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" });
    expect(reopened.snapshot()).toEqual(s.snapshot()); expect(s.history().undoSteps).toBe(1);
    s.undo(); expect(s.snapshot().metadata).toEqual(saved.metadata); expect(s.snapshot().manifest.chunks).toEqual(saved.manifest.chunks);
    s.redo(); expect(s.snapshot().metadata).toEqual(reopened.snapshot().metadata);
  } finally { s.dispose(); }
});
test("independent native conditional rules can change priority/threshold and undo without rewriting cell caches or imported validation violations", async () => {
  const s = await session("06-rules-validation");
  try {
    const original = s.snapshot(), sheet = original.metadata.workbook.sheets[0]!, source = sheet.conditionalRules![0]!;
    const rule = JSON.parse(JSON.stringify(source)) as WorkbookConditionalRule;
    const operation: DocumentOperation = { op: "setWorkbookConditionalRule", sheetId: sheet.sheetId, rule: { ...rule, priority: 20, ...(rule.kind === "cellIs" ? { formulas: ["=150"] } : {}) } };
    apply(s, [operation]); expect(s.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
    expect(s.snapshot().metadata.workbook.sheets[0]!.dataValidations).toEqual(sheet.dataValidations);
    const reopened = openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }); expect(reopened.snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(original.metadata); expect(s.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
  } finally { s.dispose(); }
});
