import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { connectWorkbookSession, createWorkbookRangeCopyTransaction, serveWorkbookSession } from "@doffi/editor";

for (const producer of [false, true]) test(`native ${producer ? "producer" : "independent"} hyperlinks copy with exact text, frozen candidates and package reopen`, async () => {
  const model = openDocument(await importXlsx(readFileSync(`examples/xlsx/${producer ? "producer/" : ""}01-values-text.xlsx`), { documentId: "copy-native-hyperlinks" }), { kind: "workbook" }), s = createDocumentSession(model);
  try {
    const original = s.snapshot().metadata.workbook.sheets[0]!.hyperlinks![0]!, text = s.getCell("sheet-1", original.range.from.row, original.range.from.column)!.text;
    const tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", original.range, { row: 256, column: 32 }), before = s.serialize();
    const candidate = s.prepare(tx); expect(s.serialize()).toEqual(before); expect(candidate.preview.metadata.workbook.sheets[0]!.hyperlinks).toHaveLength(2); s.discard(candidate); expect(s.serialize()).toEqual(before);
    s.commit(s.prepare(tx)); const current = s.snapshot(), copied = current.metadata.workbook.sheets[0]!.hyperlinks!.find(link => link.range.from.row === 256)!;
    expect(copied).toMatchObject({ ...original, range: { from: { row: 256, column: 32 }, to: { row: 256, column: 32 } } }); expect(s.getCell("sheet-1", 256, 32)!.text).toBe(text);
    const reopened = openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }); expect(reopened.snapshot().metadata).toEqual(current.metadata); expect(reopened.getCell("sheet-1", 256, 32)!.text).toBe(text);
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata.workbook.sheets[0]!.hyperlinks).toEqual([original]); expect(s.getCell("sheet-1", 256, 32)).toBeUndefined(); s.redo(); expect(s.snapshot().metadata.workbook.sheets[0]!.hyperlinks).toHaveLength(2);
  } finally { s.dispose(); }
});

test("the sole Worker-style session exposes copied links only in the private candidate until atomic publication", async () => {
  const model = openDocument(await importXlsx(readFileSync("examples/xlsx/01-values-text.xlsx"), { documentId: "copy-link-transport" }), { kind: "workbook" }), s = createDocumentSession(model, { sessionLimits: { maxCandidates: 1 } });
  const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const original = s.snapshot().metadata.workbook.sheets[0]!.hyperlinks![0]!, identity = client.state();
    const candidate = await client.prepare({ kind: "copy", documentId: identity.documentId, revision: identity.revision, sheetId: "sheet-1", source: original.range, target: { row: 256, column: 32 } });
    const before = await client.view(), preview = await client.view(candidate);
    try { expect(before.cells.snapshot.metadata.workbook.sheets[0]!.hyperlinks).toHaveLength(1); expect(preview.cells.snapshot.metadata.workbook.sheets[0]!.hyperlinks).toHaveLength(2); expect(s.history().undoSteps).toBe(0); } finally { before.dispose(); preview.dispose(); }
    await client.commit(candidate); expect(s.history().undoSteps).toBe(1); expect(s.snapshot().metadata.workbook.sheets[0]!.hyperlinks).toHaveLength(2); await client.undo(); expect(s.snapshot().metadata.workbook.sheets[0]!.hyperlinks).toEqual([original]);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});

for (const name of ["copy", "overlap", "other"]) test(`actual pinned Calc ${name} links remain editable without fabricating lost producer fields`, async () => {
  const model = openDocument(await importXlsx(readFileSync(`packages/office/tests/fixtures/workbook-copy-hyperlinks-${name}-calc.xlsx`), { documentId: "copy-produced-hyperlinks" }), { kind: "workbook" }), s = createDocumentSession(model);
  try {
    const before = s.snapshot(), source = before.metadata.workbook.sheets[0]!.hyperlinks![0]!, targetSheetId = name === "other" ? before.metadata.workbook.sheets[1]!.sheetId : "sheet-1";
    expect(source.tooltip).toBeUndefined(); expect(source.display).toBe("编号 001");
    const tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", source.range, { row: 257, column: 32 }, { targetSheetId }); s.commit(s.prepare(tx));
    const current = s.snapshot(), copied = current.metadata.workbook.sheets.find(sheet => sheet.sheetId === targetSheetId)!.hyperlinks!.find(link => link.range.from.row === 257)!;
    expect(copied.url).toBe(source.url); expect(copied.display).toBe(source.display); expect(copied.tooltip).toBeUndefined();
    const reopened = openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }); expect(reopened.snapshot().metadata).toEqual(current.metadata); expect(s.history().undoSteps).toBe(1);
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); s.redo(); expect(s.snapshot().metadata).toEqual(current.metadata);
  } finally { s.dispose(); }
});

test("Calc-created sheet default font header controls remain an explicit unsupported input without byte changes", async () => {
  const bytes = readFileSync("packages/office/tests/fixtures/workbook-copy-hyperlinks-new-sheet-calc.xlsx"), original = Buffer.from(bytes);
  await expect(importXlsx(bytes, { documentId: "copy-produced-unsupported-print" })).rejects.toMatchObject({ code: "UnsupportedFeature" }); expect(bytes).toEqual(original);
});
