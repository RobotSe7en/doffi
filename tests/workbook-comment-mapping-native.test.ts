import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
const commit = (s: WorkbookSession, operations: DocumentOperation[]) => s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations }));
async function native() { return createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/10-modern-threads.xlsx"), { documentId: "native-comment-map" }), { kind: "workbook" })); }
const range = (row: number, column: number) => ({ from: { row, column }, to: { row, column } });

test("independently authored modern root/reply cells follow inserted and deleted rows, retain resolved workflow on loss, and undo exact source", async () => {
  const s = await native();
  try {
    const before = s.snapshot(), sheetId = before.metadata.workbook.sheets[0]!.sheetId, root = before.annotations[0]!, body = s.getNode(root.bodyNodeId!)!.html;
    expect(root).toMatchObject({ status: "resolved", anchor: { kind: "cells", sheetId, ranges: [range(1, 1)] } }); expect(before.metadata.workbook.comments![0]).toMatchObject({ replies: [{ nativeCell: { row: 1, column: 1 } }] });
    commit(s, [{ op: "editWorkbookAxis", sheetId, axis: "row", action: "insert", index: 1, count: 2, formatSource: "none" }]);
    expect(s.getAnnotations()[0]!.anchor).toMatchObject({ ranges: [range(3, 1)] }); expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ replies: [{ nativeCell: { row: 3, column: 1 } }] });
    const inserted = s.snapshot(); commit(s, [{ op: "editWorkbookAxis", sheetId, axis: "row", action: "delete", index: 3, count: 1 }]);
    expect(s.getAnnotations()[0]).toMatchObject({ status: "needs_reanchor", anchor: inserted.annotations[0]!.anchor }); expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ recoveryStatus: "resolved", replies: [{ nativeCell: { row: 3, column: 1 } }] });
    expect(s.getNode(root.bodyNodeId!)!.html).toBe(body); expect(openDocument(await readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(inserted.metadata); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});

test("native thread sorting moves root/reply points together and a cross-sheet cut moves all stories through one current-source history step", async () => {
  const s = await native();
  try {
    commit(s, [{ op: "insertWorkbookSheet", name: "Target" }]); const before = s.snapshot(), sheetId = before.metadata.workbook.sheets[0]!.sheetId, other = before.metadata.workbook.sheets[1]!.sheetId, root = before.annotations[0]!;
    const bodies = before.annotations.flatMap(annotation => [annotation, ...annotation.replies ?? []]).flatMap(message => message.bodyNodeId ? [[message.bodyNodeId, s.getNode(message.bodyNodeId)!.html] as const] : []);
    commit(s, [{ op: "reorderWorkbookRows", sheetId, range: { from: { row: 1, column: 0 }, to: { row: 2, column: 3 } }, rows: [2, 1] }]);
    expect(s.getAnnotations()[0]!.anchor).toMatchObject({ ranges: [range(2, 1)] }); expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ replies: [{ nativeCell: { row: 2, column: 1 } }] });
    const sorted = s.snapshot(); commit(s, [{ op: "moveWorkbookRange", sheetId, range: range(2, 1), targetSheetId: other, target: { row: 256, column: 32 } }]);
    expect(s.getAnnotations()[0]!.anchor).toEqual({ kind: "cells", sheetId: other, ranges: [range(256, 32)] }); expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ replies: [{ nativeCell: { row: 256, column: 32 } }] });
    for (const message of [root, ...root.replies ?? []]) expect(s.snapshot().metadata.workbook.stories!.find(story => story.rootNodeId === message.bodyNodeId)!.sheetId).toBe(other);
    for (const [id, html] of bodies) expect(s.getNode(id)!.html).toBe(html); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(sorted.metadata); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
  } finally { s.dispose(); }
});

test("independent inconsistent mentions, duplicate native roots and reply locations remain invalidated provenance under unrelated structural edits", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync("packages/office/tests/fixtures/xlsx-comment-mentions.xlsx"), { documentId: "native-comment-map" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), sheetId = before.metadata.workbook.sheets[0]!.sheetId;
    commit(s, [{ op: "editWorkbookAxis", sheetId, axis: "column", action: "insert", index: 1, count: 2, formatSource: "none" }]);
    for (let index = 0; index < 4; index++) { expect(s.getAnnotations()[index]).toEqual(before.annotations[index]); expect(s.snapshot().metadata.workbook.comments![index]).toEqual(before.metadata.workbook.comments![index]); }
    expect(s.getAnnotations()[4]!.anchor).toMatchObject({ kind: "cells", ranges: [range(4, 6)] }); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().metadata.workbook.commentAuthors).toEqual(before.metadata.workbook.commentAuthors);
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
