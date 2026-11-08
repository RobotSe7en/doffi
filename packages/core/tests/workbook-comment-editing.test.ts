import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, parseContract } from "../src/index.js";
import type { DocumentOperation, WorkbookNoteBox, WorkbookSession } from "../src/index.js";

const sessions: WorkbookSession[] = [], sheetId = "sheet-1", anchor = { kind: "cells" as const, sheetId, ranges: [{ from: { row: 1, column: 1 }, to: { row: 2, column: 2 } }] };
afterEach(() => sessions.splice(0).forEach(session => session.dispose()));
function blank() { const s = createDocumentSession(openDocument(createWorkbook({ documentId: "comment-edit" }), { kind: "workbook" })); sessions.push(s); return s; }
function prepare(s: WorkbookSession, operations: DocumentOperation[]) { const manifest = s.snapshot().manifest; return s.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations }); }
const commit = (s: WorkbookSession, operations: DocumentOperation[]) => s.commit(prepare(s, operations));
const binding = (s: WorkbookSession, id: string) => s.snapshot().metadata.workbook.comments!.find(comment => comment.annotationId === id)!;
function add(s: WorkbookSession, format: "classic" | "modern" = "modern", html = "<p>正文<strong>禰󠄀😀</strong></p>") {
  commit(s, [{ op: "addWorkbookComment", anchor, html, format, author: { name: "王", initials: "W" }, date: "2026-11-08T01:02:03Z" }]); return s.getAnnotations().at(-1)!;
}
const box: WorkbookNoteBox = { anchor: { kind: "one-cell", from: { row: 1, column: 3, offsetXPt: 2, offsetYPt: 3 }, widthPt: 180, heightPt: 90 },
  visible: true, printable: true, locked: true, lockText: false, defaultSize: false, disabled: false, autoFill: false, autoLine: false, autoScale: false, justifyLastLine: false, altText: "备注说明" };

test("classic and modern creation isolate rich bodies, native people, cell ranges and one-step history", () => {
  for (const format of ["classic", "modern"] as const) {
    const s = blank(), before = s.serialize(), candidate = prepare(s, [{ op: "addWorkbookComment", anchor, html: "<p>中文禰󠄀😀&#13;<br>尾</p>", format, author: { name: "王" }, ...(format === "classic" ? { box } : {}) }]);
    expect(s.serialize()).toEqual(before); expect(candidate.preview.annotations).toHaveLength(1); s.commit(candidate);
    const annotation = s.getAnnotations()[0]!; expect(annotation).toMatchObject({ anchor, status: "open", author: { name: "王" } }); expect(s.getNode(annotation.bodyNodeId!)!.text).toBe("中文禰󠄀😀\r\n尾");
    expect(s.snapshot().metadata.workbook.commentAuthors).toMatchObject([{ format, name: "王" }]); expect(binding(s, annotation.annotationId)).toMatchObject({ format });
    expect(s.history().undoSteps).toBe(1); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata.workbook.comments).toBeUndefined(); expect(s.snapshot().manifest.content).toBeUndefined(); s.redo(); expect(s.getAnnotations()[0]!.annotationId).toBe(annotation.annotationId);
  }
});

test("modern replies retain native parent relationships, body identities, authors/timestamps and reparent surviving children on removal", () => {
  const s = blank(), root = add(s), annotationId = root.annotationId;
  commit(s, [{ op: "addWorkbookCommentReply", annotationId, html: "<p>第一回复禰󠄀</p>", author: { name: "李" }, date: "2026-11-08T01:03:00Z" }]);
  const first = s.getAnnotations()[0]!.replies![0]!;
  commit(s, [{ op: "addWorkbookCommentReply", annotationId, parentReplyId: first.replyId, html: "<p>子回复😀</p>", author: { name: "王" } }]);
  const second = s.getAnnotations()[0]!.replies![1]!, original = s.getNode(second.bodyNodeId!)!.html;
  commit(s, [{ op: "updateWorkbookCommentReply", annotationId, replyId: second.replyId, changes: { author: { name: "赵" }, date: "2026-11-08T01:04:00Z", status: "resolved" } },
    { op: "setWorkbookCommentBody", annotationId, replyId: second.replyId, html: "<p><em>修改的子回复</em></p>" }]);
  expect(s.getAnnotations()[0]!.replies![1]).toMatchObject({ replyId: second.replyId, bodyNodeId: second.bodyNodeId, author: { name: "赵" }, status: "resolved" });
  const links = binding(s, annotationId); expect(links).toMatchObject({ format: "modern", replies: [{ replyId: first.replyId }, { replyId: second.replyId, parentReplyId: first.replyId }] });
  const before = s.snapshot(); commit(s, [{ op: "removeWorkbookCommentReply", annotationId, replyId: first.replyId }]);
  expect(s.getNode(first.bodyNodeId!)).toBeUndefined(); expect(s.getNode(second.bodyNodeId!)!.text).toBe("修改的子回复"); expect(binding(s, annotationId)).toMatchObject({ replies: [{ replyId: second.replyId }] });
  expect(binding(s, annotationId)).not.toHaveProperty("replies.0.parentReplyId");
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); s.undo(); expect(s.getNode(second.bodyNodeId!)!.html).toBe(original);
});

test("exact UTF-16 comment edits map preserved mentions; overlap, ambiguous replacements and split variation sequences fail atomically", () => {
  const s = blank(), root = add(s, "modern", "<p>A禰󠄀😀B</p>"), annotationId = root.annotationId, authorId = binding(s, annotationId).authorId!;
  const mention = { mentionId: "mention", authorId, startIndex: 1, length: 5 };
  commit(s, [{ op: "setWorkbookCommentBody", annotationId, html: "<p>A禰󠄀😀B</p>", mentions: [mention] }]); s.clearHistory();
  commit(s, [{ op: "setWorkbookCommentBody", annotationId, html: "<p>头A禰󠄀😀B</p>", textEdits: [{ from: 0, to: 0, expectedText: "", text: "头" }] }]);
  expect(binding(s, annotationId)).toMatchObject({ mentions: [{ ...mention, startIndex: 2 }] }); expect(s.getNode(root.bodyNodeId!)!.text).toBe("头A禰󠄀😀B");
  const before = s.serialize(), history = s.history();
  const invalid: DocumentOperation[] = [
    { op: "setWorkbookCommentBody", annotationId, html: "<p>头A禰󠄀B</p>" },
    { op: "setWorkbookCommentBody", annotationId, html: "<p>头A禰󠄀B</p>", textEdits: [{ from: 5, to: 7, expectedText: "😀", text: "" }] },
    { op: "setWorkbookCommentBody", annotationId, html: "<p>头A禰B</p>", textEdits: [{ from: 3, to: 5, expectedText: "󠄀", text: "" }] },
    { op: "setWorkbookCommentBody", annotationId, html: "<p>头A禰󠄀😀B</p>", textEdits: [{ from: 0, to: 1, expectedText: "错", text: "头" }] },
  ];
  for (const operation of invalid) { expect(() => commit(s, [operation])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history); }
  commit(s, [{ op: "setWorkbookCommentBody", annotationId, html: "<p>头A禰󠄀B</p>", mentions: [] }]); expect(binding(s, annotationId)).not.toHaveProperty("mentions");
  s.undo(); expect(binding(s, annotationId)).toMatchObject({ mentions: [{ startIndex: 2 }] }); s.undo(); expect(binding(s, annotationId)).toMatchObject({ mentions: [mention] });
});

test("status recovery requires an explicit cell anchor and preserves rich bodies across worksheet reanchoring", () => {
  const s = blank(), root = add(s), annotationId = root.annotationId;
  commit(s, [{ op: "insertWorkbookSheet", name: "Other" }]); const other = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
  commit(s, [{ op: "updateWorkbookComment", annotationId, changes: { status: "resolved" } }]);
  commit(s, [{ op: "updateWorkbookComment", annotationId, changes: { status: "needs_reanchor" } }]);
  expect(binding(s, annotationId)).toMatchObject({ recoveryStatus: "resolved" }); const before = s.serialize();
  expect(() => commit(s, [{ op: "updateWorkbookComment", annotationId, changes: { status: "open" } }])).toThrow(); expect(s.serialize()).toEqual(before);
  commit(s, [{ op: "updateWorkbookComment", annotationId, changes: { anchor: { ...anchor, sheetId: other }, status: "open", author: { name: "新作者" }, date: null } }]);
  expect(s.getAnnotations()[0]).toMatchObject({ anchor: { sheetId: other }, status: "open", author: { name: "新作者" }, bodyNodeId: root.bodyNodeId }); expect(s.getAnnotations()[0]).not.toHaveProperty("date");
  expect(binding(s, annotationId)).not.toHaveProperty("recoveryStatus"); expect(s.snapshot().metadata.workbook.stories![0]!.sheetId).toBe(other); expect(s.getNode(root.bodyNodeId!)!.text).toBe("正文禰󠄀😀");
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
});

test("ordinary comment text replacement preserves native run/paragraph source, CRLF, exact mentions and undo", () => {
  const s = blank(), root = add(s, "modern", '<p><span class="cell-default">A&#13;<br></span><strong>禰󠄀😀</strong><em>Z</em></p>'), annotationId = root.annotationId;
  const authorId = binding(s, annotationId).authorId!;
  commit(s, [{ op: "setWorkbookCommentBody", annotationId, html: '<p><span class="cell-default">A&#13;<br></span><strong>禰󠄀😀</strong><em>Z</em></p>', mentions: [{ mentionId: "preserved", authorId, startIndex: 3, length: 5 }] }]);
  const original = s.getNode(root.bodyNodeId!)!, before = s.snapshot(), strongId = original.childNodeIds.flatMap(id => s.getNode(id)!.childNodeIds).find(id => s.getNode(id)!.tagName === "strong")!;
  commit(s, [{ op: "replaceWorkbookCommentText", annotationId, from: 0, to: 1, expectedText: "A", text: "中文" }]);
  expect(s.getNode(root.bodyNodeId!)!.text).toBe("中文\r\n禰󠄀😀Z"); expect(s.getNode(strongId)!.text).toBe("禰󠄀😀"); expect(s.getNode(root.bodyNodeId!)!.html).toContain("<em data-node-id=");
  expect(binding(s, annotationId)).toMatchObject({ mentions: [{ startIndex: 4, length: 5 }] }); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.getNode(root.bodyNodeId!)!.html).toBe(original.html);
});

test("classic note boxes and explicit format conversion preserve body source; replies require modern semantics", () => {
  const s = blank(), root = add(s, "classic"), annotationId = root.annotationId, original = s.getNode(root.bodyNodeId!)!.html;
  commit(s, [{ op: "setWorkbookNoteBox", annotationId, box }]); expect(binding(s, annotationId)).toMatchObject({ box });
  const before = s.serialize(); expect(() => commit(s, [{ op: "addWorkbookCommentReply", annotationId, html: "<p>不能悄悄转换</p>" }])).toThrow(); expect(s.serialize()).toEqual(before);
  commit(s, [{ op: "updateWorkbookComment", annotationId, changes: { format: "modern" } }]); expect(binding(s, annotationId)).toMatchObject({ format: "modern", replies: [] }); expect(s.getNode(root.bodyNodeId!)!.html).toBe(original);
  commit(s, [{ op: "addWorkbookCommentReply", annotationId, html: "<p>回复</p>" }]); const reply = s.getAnnotations()[0]!.replies![0]!;
  expect(() => commit(s, [{ op: "updateWorkbookComment", annotationId, changes: { format: "classic" } }])).toThrow();
  commit(s, [{ op: "removeWorkbookCommentReply", annotationId, replyId: reply.replyId }, { op: "updateWorkbookComment", annotationId, changes: { format: "classic" } }]);
  expect(binding(s, annotationId)).toMatchObject({ format: "classic" }); expect(s.getNode(root.bodyNodeId!)!.html).toBe(original); expect(s.getNode(reply.bodyNodeId!)).toBeUndefined();
});

test("the shared removeAnnotation operation cleans the entire native thread and stories while retaining unrelated comments, source and history", () => {
  const s = blank(), first = add(s); commit(s, [{ op: "addWorkbookCommentReply", annotationId: first.annotationId, html: "<p>回复</p>" }]);
  const reply = s.getAnnotations()[0]!.replies![0]!, second = add(s, "classic"), original = s.getNode(second.bodyNodeId!)!.html, before = s.snapshot();
  commit(s, [{ op: "removeAnnotation", annotationId: first.annotationId }]); expect(s.getNode(first.bodyNodeId!)).toBeUndefined(); expect(s.getNode(reply.bodyNodeId!)).toBeUndefined(); expect(s.getNode(second.bodyNodeId!)!.html).toBe(original);
  expect(s.snapshot().metadata.workbook.comments).toHaveLength(1); expect(s.snapshot().metadata.workbook.stories).toHaveLength(1); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
});

test("equal comment bodies, dates and authors preserve source identities, revisions and redo; unsafe edits and cyclic replies roll back", () => {
  const s = blank(), root = add(s); commit(s, [{ op: "addWorkbookCommentReply", annotationId: root.annotationId, html: "<p>回复</p>" }]); const reply = s.getAnnotations()[0]!.replies![0]!;
  commit(s, [{ op: "setWorkbookCommentBody", annotationId: root.annotationId, html: "<p>临时</p>" }]); s.undo(); const before = s.serialize(), history = s.history();
  expect(commit(s, [{ op: "setWorkbookCommentBody", annotationId: root.annotationId, html: "<p>正文<strong>禰󠄀😀</strong></p>" },
    { op: "updateWorkbookComment", annotationId: root.annotationId, changes: { author: { name: "王", initials: "W" }, date: "2026-11-08T01:02:03Z" } }]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  const bad: DocumentOperation[] = [{ op: "setWorkbookCommentBody", annotationId: root.annotationId, html: '<p onclick="alert(1)">坏</p>' },
    { op: "setWorkbookCommentBody", annotationId: root.annotationId, html: '<p data-node-id="forged">伪造</p>' },
    { op: "updateWorkbookCommentReply", annotationId: root.annotationId, replyId: reply.replyId, changes: { parentReplyId: reply.replyId } },
    { op: "setWorkbookNoteBox", annotationId: root.annotationId, box },
    { op: "setWorkbookCommentBody", annotationId: "foreign", html: "<p>越界</p>" }];
  for (const operation of bad) { expect(() => commit(s, [{ op: "updateWorkbookComment", annotationId: root.annotationId, changes: { status: "resolved" } }, operation])).toThrow(); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history); }
});

test("reversing transient author/date changes in one batch restores source bytes and redo without retaining unused generated people", () => {
  const s = blank(), root = add(s);
  commit(s, [{ op: "updateWorkbookComment", annotationId: root.annotationId, changes: { status: "resolved" } }]); s.undo();
  const before = s.serialize(), history = s.history();
  expect(commit(s, [{ op: "updateWorkbookComment", annotationId: root.annotationId, changes: { author: { name: "临时作者" }, date: null } },
    { op: "updateWorkbookComment", annotationId: root.annotationId, changes: { author: { name: "王", initials: "W" }, date: "2026-11-08T01:02:03Z" } }]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
});

test.each(["updateWorkbookComment", "updateAnnotation"] as const)("%s turns an acknowledged orphan placeholder into a classic note while retaining its UID, body and undo provenance", op => {
  const source = blank(), root = add(source, "classic"), files = new Map(source.serialize()), manifest = source.snapshot().manifest;
  const metadata = parseContract("document-metadata", source.snapshot().metadata); if (!("workbook" in metadata)) throw Error("Expected workbook metadata");
  const binding = metadata.workbook.comments![0]!; if (binding.format !== "classic") throw Error("Expected classic binding");
  binding.compatibilityThreadId = "{11111111-1111-4111-8111-111111111111}"; binding.nativeUid = "{22222222-2222-4222-8222-222222222222}"; binding.recoveryReasons = ["orphan-placeholder"]; binding.recoveryStatus = "open";
  const annotations = parseContract("annotations", source.snapshot().annotations); annotations[0]!.status = "needs_reanchor";
  files.set(manifest.metadata, new TextEncoder().encode(JSON.stringify(metadata))); files.set(manifest.annotations, new TextEncoder().encode(JSON.stringify(annotations)));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); sessions.push(s); const before = s.snapshot();
  commit(s, [{ op, annotationId: root.annotationId, changes: { anchor, status: "open" } }]);
  expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ format: "classic", nativeUid: binding.nativeUid }); expect(s.snapshot().metadata.workbook.comments![0]).not.toHaveProperty("compatibilityThreadId");
  expect(s.getAnnotations()[0]!.status).toBe("open"); expect(s.snapshot().content).toEqual(before.content); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
});
