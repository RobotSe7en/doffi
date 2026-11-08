import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { AnchorSchema, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
const commit = (s: WorkbookSession, operations: DocumentOperation[]) => s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations }));
const escape = (text: string) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;").replace(/\n/g, "<br>");

test.each(["examples/xlsx/09-classic-notes.xlsx", "examples/xlsx/producer/09-classic-notes.xlsx", "examples/xlsx/10-modern-threads.xlsx"])("native %s edits status, author/time and replies, cleans shared removal, and restores exact message source", async path => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync(path), { documentId: "native-comment-edit" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), original = before.annotations[0]!, body = s.getNode(original.bodyNodeId!)!.html, binding = before.metadata.workbook.comments![0]!, authorId = binding.authorId;
    const others = before.annotations.slice(1).flatMap(annotation => [annotation, ...annotation.replies ?? []]).filter(message => message.bodyNodeId).map(message => [message.bodyNodeId!, s.getNode(message.bodyNodeId!)!.html] as const);
    commit(s, [{ op: "updateWorkbookComment", annotationId: original.annotationId, changes: { status: "closed", ...(original.author ? { author: { ...original.author } } : {}), date: "2026-11-08T02:03:04Z" } }]);
    expect(s.getAnnotations()[0]).toMatchObject({ status: "closed", date: "2026-11-08T02:03:04Z", sourceId: original.sourceId });
    expect(s.snapshot().metadata.workbook.comments![0]!.authorId).toBe(authorId); expect(s.getNode(original.bodyNodeId!)!.html).toBe(body);
    if (binding.format === "modern") {
      commit(s, [{ op: "addWorkbookCommentReply", annotationId: original.annotationId, html: "<p>新的<strong>回复禰󠄀😀</strong></p>", author: { name: "本地复核人" }, date: "2026-11-08T02:04:05Z" }]);
      const reply = s.getAnnotations()[0]!.replies!.at(-1)!; expect(s.getNode(reply.bodyNodeId!)!.text).toBe("新的回复禰󠄀😀");
      commit(s, [{ op: "removeWorkbookCommentReply", annotationId: original.annotationId, replyId: reply.replyId }]); expect(s.getNode(reply.bodyNodeId!)).toBeUndefined(); s.undo(); s.undo();
    }
    const current = s.snapshot(); commit(s, [{ op: "removeAnnotation", annotationId: original.annotationId }]);
    expect(s.getNode(original.bodyNodeId!)).toBeUndefined(); for (const reply of original.replies ?? []) expect(s.getNode(reply.bodyNodeId!)).toBeUndefined();
    for (const [id, html] of others) expect(s.getNode(id)!.html).toBe(html);
    expect(openDocument(await readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(current.metadata); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});

test("independent dual note boxes preserve the unselected native source, differences and UID through selected box edits", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync("packages/office/tests/fixtures/xlsx-dual-note-boxes-conflict.xlsx"), { documentId: "native-comment-edit" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), binding = before.metadata.workbook.comments![0]!; if (binding.format !== "classic" || !binding.box || !binding.commentPrBox) throw Error("Expected independent dual-box source");
    const box = { ...binding.box, visible: !binding.box.visible, altText: "显式修改的备注框" };
    commit(s, [{ op: "setWorkbookNoteBox", annotationId: binding.annotationId, box }]);
    const current = s.snapshot().metadata.workbook.comments![0]!;
    expect(current).toMatchObject({ box, commentPrBox: binding.commentPrBox, boxSource: "vml" }); expect(current).toHaveProperty("boxDifferences", expect.arrayContaining(["visible", "altText"]));
    expect("nativeUid" in current ? current.nativeUid : undefined).toBe(binding.nativeUid); expect(s.snapshot().content).toEqual(before.content); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});

test("native mention offsets move only with exact body proofs, preserve people/provider and conflicting locations, and require explicit repair", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync("packages/office/tests/fixtures/xlsx-comment-mentions.xlsx"), { documentId: "native-comment-edit" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), root = before.annotations[0]!, binding = before.metadata.workbook.comments![0]!; if (binding.format !== "modern") throw Error("Expected modern source");
    const text = s.getNode(root.bodyNodeId!)!.text; expect(text).toBe("请复核预算😀禰󠄀"); const prefix = "头😀";
    commit(s, [{ op: "setWorkbookCommentBody", annotationId: root.annotationId, html: "<p>" + escape(prefix + text) + "</p>", textEdits: [{ from: 0, to: 0, expectedText: "", text: prefix }] }]);
    expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ recoveryReasons: binding.recoveryReasons, recoveryStatus: "resolved", mentions: [{ startIndex: 8, length: 5 }], replies: binding.replies });
    expect(s.snapshot().metadata.workbook.commentAuthors).toEqual(before.metadata.workbook.commentAuthors); expect(s.getAnnotations()[0]!.status).toBe("needs_reanchor");
    const badRoot = s.getAnnotations()[2]!, badBinding = s.snapshot().metadata.workbook.comments![2]!; expect(badBinding).toMatchObject({ recoveryReasons: ["invalid-mention-range"] });
    const recoveredAnchor = AnchorSchema.options[0].parse(badRoot.anchor);
    const current = s.serialize(); expect(() => commit(s, [{ op: "updateWorkbookComment", annotationId: badRoot.annotationId, changes: { anchor: recoveredAnchor, status: "open" } }])).toThrow(); expect(s.serialize()).toEqual(current);
    commit(s, [{ op: "setWorkbookCommentBody", annotationId: badRoot.annotationId, html: "<p>修复后的来源禰󠄀</p>", mentions: [] },
      { op: "updateWorkbookComment", annotationId: badRoot.annotationId, changes: { anchor: recoveredAnchor, status: "open" } }]);
    expect(s.getAnnotations()[2]!.status).toBe("open"); expect(s.snapshot().metadata.workbook.comments![2]).not.toHaveProperty("recoveryReasons");
    expect(openDocument(await readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
  } finally { s.dispose(); }
});

test.each(["updateWorkbookComment", "updateAnnotation"] as const)("%s explicitly repairs native reply positions and a missing root cell without inventing a location", async op => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync("packages/office/tests/fixtures/xlsx-comment-mentions.xlsx"), { documentId: "native-comment-edit" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), first = before.annotations[0]!, unlocated = before.annotations[3]!, sheetId = before.metadata.workbook.sheets[0]!.sheetId;
    expect(before.metadata.workbook.comments![0]).toMatchObject({ replies: [{ nativeCell: { row: 2, column: 1 } }] });
    commit(s, [{ op, annotationId: first.annotationId, changes: { anchor: { kind: "cells", sheetId, ranges: [{ from: { row: 4, column: 3 }, to: { row: 4, column: 3 } }] }, status: "open" } },
      { op, annotationId: unlocated.annotationId, changes: { anchor: { kind: "cells", sheetId, ranges: [{ from: { row: 6, column: 6 }, to: { row: 6, column: 6 } }] }, status: "resolved" } }]);
    expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ replies: [{ nativeCell: { row: 4, column: 3 } }] });
    expect(s.getAnnotations()[0]).toMatchObject({ status: "open", bodyNodeId: first.bodyNodeId }); expect(s.getAnnotations()[3]).toMatchObject({ status: "resolved", bodyNodeId: unlocated.bodyNodeId });
    expect(s.snapshot().metadata.workbook.comments![0]).not.toHaveProperty("recoveryReasons"); expect(s.snapshot().content).toEqual(before.content);
    expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
