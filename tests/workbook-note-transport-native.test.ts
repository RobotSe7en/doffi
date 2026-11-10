import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookSession } from "@doffi/core";
import { connectWorkbookSession, serveWorkbookSession } from "@doffi/editor";
import type { WorkbookLayoutSourceRequest } from "@doffi/render";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

interface Producer { name: string; outputFile: string; outputSha256: string; observed: { nativeBoxes: { commentCells: string[]; commentPr: (string | null)[] } } }
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-note-transport.json", "utf8")) as { sources: { name: string; source: string; sourceSha256: string; cases: Producer[] }[] };
function metrics(request: WorkbookLayoutRequest): WorkbookLayoutMeasurement {
  return { sourceDigest: request.sourceDigest, fontEnvironment: "native-note-explicit-point-metrics", sheets: request.sheetIds.map(sheetId => ({ sheetId,
    rows: { defaultSizePt: request.model.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!.defaultRowHeightPt, overrides: [] },
    columns: { defaultSizePt: 45, overrides: [] } })) };
}
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const cases = evidence.sources.flatMap(source => source.cases.map(producer => ({ source, producer, label: source.name + "/" + producer.name })));
test("eight native note edits retain immutable sources, both original boxes and default-header negatives", () => {
  for (const source of evidence.sources) expect(createHash("sha256").update(readFileSync(source.source)).digest("hex")).toBe(source.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-note-transport.py"], { encoding: "utf8" })).toContain("eight independent");
});
test.each(cases)("native source $label maps current note cells, preserves body/native identity and reopens one-step history", async ({ source, producer }) => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync(source.source), { documentId: "native-note-transport" }), { kind: "workbook" }));
  try {
    const sheetId = s.snapshot().metadata.workbook.sheets[0]!.sheetId;
    if (producer.name.startsWith("cut-other")) s.commit(s.prepare(transaction(s, [{ op: "insertWorkbookSheet", name: "Other" }]), { measureWorkbookLayout: metrics }));
    if (producer.name === "sort") s.commit(s.prepare(transaction(s, [{ op: "setWorkbookCells", sheetId, cells: [3, 1, 2].map((value, index) => ({ row: index + 1, column: 1, source: { kind: "value" as const, value: { type: "number" as const, value } } })) }]), { measureWorkbookLayout: metrics }));
    s.clearHistory(); const before = s.snapshot(), original = before.annotations[0]!, oldBinding = before.metadata.workbook.comments![0]!, body = s.getNode(original.bodyNodeId!)!.html;
    const targetSheetId = producer.name.startsWith("cut-other") ? before.metadata.workbook.sheets.at(-1)!.sheetId : sheetId;
    const operation: DocumentOperation = producer.name === "sort" ? { op: "reorderWorkbookRows", sheetId, range: range(1, 1, 3, 2), rows: [3, 1, 2] }
      : { op: "moveWorkbookRange", sheetId, range: range(1, 1), targetSheetId, target: { row: 19, column: 7 } };
    const candidate = s.prepare(transaction(s, [operation]), { measureWorkbookLayout: metrics }); expect(s.snapshot()).toEqual(before); s.commit(candidate);
    expect(s.snapshot().annotations[0]).toMatchObject({ annotationId: original.annotationId, bodyNodeId: original.bodyNodeId, anchor: { sheetId: targetSheetId,
      ranges: [producer.name === "sort" ? range(3, 1) : range(19, 7)] } });
    const next = s.snapshot().metadata.workbook.comments![0]!; expect(next.format).toBe("classic");
    if (next.format !== "classic" || oldBinding.format !== "classic") throw Error("Native binding changed kind");
    const { box: _oldBox, commentPrBox: _oldAlternate, ...oldFields } = oldBinding, { box: _box, commentPrBox: _alternate, ...fields } = next;
    expect(fields).toEqual(oldFields); expect(Boolean(next.commentPrBox)).toBe(Boolean(oldBinding.commentPrBox));
    expect(s.getNode(original.bodyNodeId!)!.html).toBe(body); expect(s.history().undoSteps).toBe(1);
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
test.each(cases)("native producer $label reopens its actual moved note and alternate-box loss without fabricating a second source", async ({ producer }) => {
  const bytes = readFileSync(producer.outputFile); expect(createHash("sha256").update(bytes).digest("hex")).toBe(producer.outputSha256);
  if (producer.name === "cut-other-default-header") {
    await expect(importXlsx(bytes, { documentId: "note-default-header-negative" })).rejects.toMatchObject({ code: "UnsupportedFeature", message: "This print header/footer control is not implemented." }); return;
  }
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "note-transport-producer" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), annotation = before.annotations[0]!, expected = producer.name === "sort" ? range(3, 1) : range(19, 7);
    expect(annotation.anchor).toMatchObject({ ranges: [expected] }); expect(producer.observed.nativeBoxes.commentCells).toEqual([producer.name === "sort" ? "B4" : "H20"]);
    expect(before.metadata.workbook.comments![0]).not.toHaveProperty("commentPrBox"); expect(producer.observed.nativeBoxes.commentPr).toEqual([null]);
    s.commit(s.prepare(transaction(s, [{ op: "setWorkbookCommentBody", annotationId: annotation.annotationId, html: "<p>原生再编辑禰󠄀😀</p>" }]), { measureWorkbookLayout: metrics }));
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
  } finally { s.dispose(); }
});
test("sole Worker note movement measures both private sheets, discards/cancels candidates and commits one undo step", async () => {
  const source = evidence.sources[1]!, s = createDocumentSession(openDocument(await importXlsx(readFileSync(source.source), { documentId: "worker-note-transport" }), { kind: "workbook" }));
  s.commit(s.prepare(transaction(s, [{ op: "insertWorkbookSheet", name: "Other" }]), { measureWorkbookLayout: metrics })); s.clearHistory();
  const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), targetSheetId = before.metadata.workbook.sheets.at(-1)!.sheetId, command = { kind: "transaction" as const, transaction: transaction(s, [{ op: "moveWorkbookRange", sheetId: before.metadata.workbook.sheets[0]!.sheetId, range: range(1, 1), targetSheetId, target: { row: 19, column: 7 } }]) };
    const measureLayout = (request: WorkbookLayoutSourceRequest): WorkbookLayoutMeasurement => ({ sourceDigest: request.sourceDigest, fontEnvironment: "native-note-private-point-metrics", sheets: request.sheetIds.map(sheetId => ({ sheetId,
      rows: { defaultSizePt: request.source.cells.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!.defaultRowHeightPt, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
    await expect(client.prepare(command)).rejects.toMatchObject({ code: "InvalidOperation" }); expect(s.snapshot()).toEqual(before);
    const candidate = await client.prepare(command, undefined, { measureLayout }); expect(s.snapshot()).toEqual(before); await client.discard(candidate); expect(s.snapshot()).toEqual(before);
    const abort = new AbortController(); await expect(client.prepare(command, undefined, { signal: abort.signal, measureLayout: request => { abort.abort(); return measureLayout(request); } })).rejects.toMatchObject({ code: "OperationCancelled" });
    expect(s.snapshot()).toEqual(before); await client.commit(await client.prepare(command, undefined, { measureLayout })); expect(s.history().undoSteps).toBe(1);
    const view = await client.view(); expect(view.cells.snapshot.annotations[0]!.anchor).toMatchObject({ sheetId: targetSheetId, ranges: [range(19, 7)] }); view.dispose();
    await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    await expect(client.prepare(command, undefined, { measureLayout })).rejects.toMatchObject({ code: "RevisionConflict" });
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
