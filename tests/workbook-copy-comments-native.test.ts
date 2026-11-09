import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { connectWorkbookSession, serveWorkbookSession } from "@doffi/editor";
import { writeOpcZip } from "../packages/office/src/document-zip.js";

const paths = ["examples/xlsx/09-classic-notes.xlsx", "examples/xlsx/producer/09-classic-notes.xlsx", "examples/xlsx/10-modern-threads.xlsx",
  "packages/office/tests/fixtures/xlsx-comment-mentions.xlsx", "packages/office/tests/fixtures/xlsx-comment-uids.xlsx", "packages/office/tests/fixtures/xlsx-dual-note-boxes-conflict.xlsx"];
interface Producer { name: string; source: string; sourceSha256: string; outputFile: string; outputSha256: string;
  nativeCall: { sourceText: string; targetText: string; hadClassicNote: boolean; copiedClassicNote: boolean };
  observed: { classic: { cell: string; text: string }[]; modern: unknown[]; duplicateVmlIds: unknown[] } }
const producerEvidence = JSON.parse(readFileSync("docs/evidence/workbook-copy-comments.json", "utf8")) as { cases: Producer[] };
test("four actual comment-copy producer observations retain independently authored inputs and exact immutable byte hashes", () => {
  expect(execFileSync("python3", ["scripts/check-workbook-copy-comments.py"], { encoding: "utf8" })).toContain("four independent");
  for (const item of producerEvidence.cases) {
    expect(createHash("sha256").update(readFileSync(item.source)).digest("hex")).toBe(item.sourceSha256);
    expect(createHash("sha256").update(readFileSync(item.outputFile)).digest("hex")).toBe(item.outputSha256);
    expect(item.nativeCall.targetText).toBe(item.nativeCall.sourceText); expect(item.nativeCall.sourceText).not.toBe("");
  }
});
test.each(producerEvidence.cases)("the unmodified actual $name comment copy reopens or explicitly rejects malformed native identities", async item => {
  if (item.observed.duplicateVmlIds.length) {
    expect(item.nativeCall.copiedClassicNote).toBe(true);
    await expect(importXlsx(readFileSync(item.outputFile), { documentId: "malformed-copied-note" })).rejects.toMatchObject({ code: "InvalidDocument" }); return;
  }
  const model = openDocument(await importXlsx(readFileSync(item.outputFile), { documentId: "copied-note-producer" }), { kind: "workbook" });
  const snapshot = model.snapshot(); expect(snapshot.metadata.workbook.comments?.filter(binding => binding.format === "classic").length ?? 0).toBe(item.observed.classic.length);
  expect(snapshot.metadata.workbook.comments?.filter(binding => binding.format === "modern").length ?? 0).toBe(item.observed.modern.length);
  for (const note of item.observed.classic) {
    const coordinate = /^([A-Z]+)(\d+)$/.exec(note.cell)!; let column = 0; for (const letter of coordinate[1]!) column = column * 26 + letter.charCodeAt(0) - 64;
    const annotation = snapshot.annotations.find(annotation => annotation.anchor.kind === "cells" && annotation.anchor.ranges.some(range => range.from.row === Number(coordinate[2]) - 1 && range.from.column === column - 1))!;
    expect(model.getNode(annotation.bodyNodeId!)!.text).toBe(note.text);
  }
  expect(openDocument(readDocumentZip(writeDocumentZip(model.serialize())), { kind: "workbook" }).snapshot()).toEqual(snapshot);
});
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
/** Explicit point metrics test the core's native box transport, not font or
 * Excel geometry acceptance. Public browser tests use the real render service. */
function metrics(request: WorkbookLayoutRequest): WorkbookLayoutMeasurement {
  return { sourceDigest: request.sourceDigest, fontEnvironment: "independent-source-test-metrics", sheets: request.sheetIds.map(sheetId => ({ sheetId,
    rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) };
}
test.each(paths)("independent %s clones native comment bodies, reply graphs and retained recovery into a different worksheet", async path => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync(path), { documentId: "native-copy-comments" }), { kind: "workbook" }));
  try {
    s.commit(s.prepare(transaction(s, [{ op: "insertWorkbookSheet", name: "Destination" }]), { measureWorkbookLayout: metrics })); s.clearHistory();
    const before = s.snapshot(), targetSheetId = before.metadata.workbook.sheets.at(-1)!.sheetId;
    const original = before.annotations.find(annotation => annotation.anchor.kind === "cells" && annotation.anchor.ranges.length)!;
    if (original.anchor.kind !== "cells") throw Error("Expected native cell anchor");
    const source = original.anchor.ranges[0]!, target = { row: 40, column: 4 }, delta = { row: target.row - source.from.row, column: target.column - source.from.column };
    const texts = [original, ...original.replies ?? []].map(body => s.getNode(body.bodyNodeId!)!.text), originalHtml = s.getNode(original.bodyNodeId!)!.html;
    const candidate = s.prepare(transaction(s, [{ op: "copyWorkbookComments", sheetId: original.anchor.sheetId, range: source, targetSheetId, target }]), { measureWorkbookLayout: metrics });
    expect(s.snapshot()).toEqual(before); s.commit(candidate);
    const copied = s.getAnnotations().find(annotation => annotation.anchor.kind === "cells" && annotation.anchor.sheetId === targetSheetId && s.getNode(annotation.bodyNodeId!)?.text === texts[0])!;
    expect(copied).toBeDefined(); expect(copied.annotationId).not.toBe(original.annotationId); expect(copied.bodyNodeId).not.toBe(original.bodyNodeId); expect(copied.sourceId).toBeUndefined();
    expect(copied.status).toBe(original.status); expect(copied.author).toEqual(original.author); expect(copied.date).toBe(original.date);
    expect([copied, ...copied.replies ?? []].map(body => s.getNode(body.bodyNodeId!)!.text)).toEqual(texts);
    const binding = s.snapshot().metadata.workbook.comments!.find(binding => binding.annotationId === copied.annotationId)!, old = before.metadata.workbook.comments!.find(binding => binding.annotationId === original.annotationId)!;
    expect(binding.format).toBe(old.format);
    if (binding.format === "classic" && old.format === "classic") {
      expect(binding.nativeShapeId).toBeUndefined(); expect(binding.nativeUid).toBeUndefined(); expect(binding.compatibilityThreadId).toBe(old.compatibilityThreadId);
      for (const key of ["box", "commentPrBox"] as const) if (old[key]) {
        expect(binding[key]).toBeDefined(); expect(binding[key]!.visible).toBe(old[key]!.visible); expect(binding[key]!.cssClass).toBe(old[key]!.cssClass);
      }
      expect(binding.boxSource).toBe(old.boxSource); expect(binding.boxDifferences).toEqual(old.boxDifferences);
    } else if (binding.format === "modern" && old.format === "modern") {
      expect(binding.compatibilityUids).toBeUndefined(); expect(binding.recoveryReasons).toEqual(old.recoveryReasons);
      expect(binding.mentions?.map(mention => [mention.startIndex, mention.length, mention.authorId])).toEqual(old.mentions?.map(mention => [mention.startIndex, mention.length, mention.authorId]));
      for (const [index, reply] of binding.replies.entries()) {
        expect(reply.replyId).toBe(copied.replies![index]!.replyId);
        const parent = old.replies[index]!.parentReplyId;
        expect(reply.parentReplyId).toBe(parent ? copied.replies![original.replies!.findIndex(reply => reply.replyId === parent)]!.replyId : undefined);
        const cell = old.replies[index]!.nativeCell;
        if (cell) expect(reply.nativeCell).toEqual({ row: cell.row + delta.row, column: cell.column + delta.column });
      }
    }
    expect(s.getNode(original.bodyNodeId!)!.html).toBe(originalHtml);
    expect(openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
test("the sole Worker privately copies current comment stories, discards them and retries one commit without leaking a second source model", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/10-modern-threads.xlsx"), { documentId: "worker-copy-comments" }), { kind: "workbook" }));
  const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), original = before.annotations.find(annotation => annotation.anchor.kind === "cells" && annotation.anchor.ranges.length)!;
    if (original.anchor.kind !== "cells") throw Error("Expected native anchor");
    const state = client.state(), source = original.anchor.ranges[0]!, command = { kind: "copy" as const, documentId: state.documentId, revision: state.revision,
      sheetId: original.anchor.sheetId, source, target: { row: 40, column: 4 } };
    const candidate = await client.prepare(command), view = await client.view(candidate);
    const projected = view.cells.snapshot.annotations.filter(annotation => !before.annotations.some(old => old.annotationId === annotation.annotationId));
    expect(projected.length).toBeGreaterThan(0); expect(s.snapshot()).toEqual(before); view.dispose(); await client.discard(candidate); expect(s.snapshot()).toEqual(before);
    await client.commit(await client.prepare(command)); expect(s.history().undoSteps).toBe(1);
    const copied = s.getAnnotations().find(annotation => !before.annotations.some(old => old.annotationId === annotation.annotationId))!;
    expect(s.getNode(copied.bodyNodeId!)!.text).toBe(s.getNode(original.bodyNodeId!)!.text); expect(copied.bodyNodeId).not.toBe(original.bodyNodeId);
    await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
test("orphan compatibility notes retain their missing-thread evidence and recovery state in range and whole-sheet copies", async () => {
  // Hand-authored native compatibility XML, using the existing independent
  // modern input and only the general ZIP primitive; no SDK export baseline.
  const files = new Map(readDocumentZip(readFileSync("examples/xlsx/10-modern-threads.xlsx"))), encoder = new TextEncoder(), decoder = new TextDecoder();
  expect(files.has("xl/comments1.xml")).toBe(false);
  const add = (path: string, close: string, text: string) => { const xml = decoder.decode(files.get(path)!); expect(xml).toContain(close); files.set(path, encoder.encode(xml.replace(close, text + close))); };
  add("[Content_Types].xml", "</Types>", '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>');
  add("xl/worksheets/_rels/sheet1.xml.rels", "</Relationships>", '<Relationship Id="legacyOrphan" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>');
  files.set("xl/comments1.xml", encoder.encode('<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>tc={AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}</author></authors><commentList><comment ref="D4" authorId="0"><text><t>原始兼容缓存</t></text></comment></commentList></comments>'));
  const s = createDocumentSession(openDocument(await importXlsx(writeOpcZip(files), { documentId: "orphan-comment-copy" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), binding = before.metadata.workbook.comments!.find(binding => binding.format === "classic" && binding.recoveryReasons?.includes("orphan-placeholder"))!;
    expect(binding).toBeDefined(); const original = before.annotations.find(annotation => annotation.annotationId === binding.annotationId)!;
    if (original.anchor.kind !== "cells" || !original.anchor.ranges.length || binding.format !== "classic") throw Error("Expected anchored orphan source");
    s.commit(s.prepare(transaction(s, [{ op: "copyWorkbookComments", sheetId: original.anchor.sheetId, range: original.anchor.ranges[0]!, targetSheetId: original.anchor.sheetId, target: { row: 40, column: 4 } }]), { measureWorkbookLayout: metrics }));
    const copied = s.getAnnotations().find(annotation => !before.annotations.some(old => old.annotationId === annotation.annotationId))!;
    expect(copied.status).toBe("needs_reanchor"); expect(s.snapshot().metadata.workbook.comments!.find(value => value.annotationId === copied.annotationId)).toMatchObject({ compatibilityThreadId: binding.compatibilityThreadId, recoveryReasons: binding.recoveryReasons, recoveryStatus: binding.recoveryStatus });
    expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
    s.commit(s.prepare(transaction(s, [{ op: "copyWorkbookSheet", sheetId: original.anchor.sheetId, name: "Copied" }])));
    const sheetId = s.snapshot().metadata.workbook.sheets.at(-1)!.sheetId, orphan = s.getAnnotations().find(annotation => annotation.anchor.kind === "cells" && annotation.anchor.sheetId === sheetId && s.getNode(annotation.bodyNodeId!)?.text === s.getNode(original.bodyNodeId!)!.text)!;
    expect(orphan.status).toBe("needs_reanchor"); expect(s.snapshot().metadata.workbook.comments!.find(value => value.annotationId === orphan.annotationId)).toMatchObject({ compatibilityThreadId: binding.compatibilityThreadId, recoveryReasons: binding.recoveryReasons });
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
  } finally { s.dispose(); }
});
