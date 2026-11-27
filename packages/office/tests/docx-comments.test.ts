import { expect, test } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
function fixture() {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "comments" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "comments",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "正文😀漢\u{e0100}",
        },
        {
          op: "addAnnotation",
          annotation: {
            annotationId: "a",
            documentId: "comments",
            baseRevision: 0,
            status: "open",
            anchor: { kind: "text", nodeId: "paragraph-1", from: 2, to: 4, quote: "😀" },
            instruction: "原意见",
            author: { name: "甲", initials: "A" },
            date: "2026-11-14T10:00:00+08:00",
          },
        },
      ],
    }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "comments",
      baseRevision: 1,
      operations: [
        {
          op: "addWordCommentReply",
          nodeId: "paragraph-1",
          annotationId: "a",
          html: "<p><strong>第一回复</strong></p><p>第二段漢\u{e0100}</p>",
          author: { name: "乙" },
          date: "2026-11-14T03:00:00Z",
        },
      ],
    }),
  );
  return session;
}
test("modern Word threads retain precise anchors, rich replies, authors, UTC dates and durable identities across edits and two native reopens", () => {
  const session = fixture();
  let bytes = exportDocx(session).bytes;
  let durable: string[] | undefined;
  for (let pass = 0; pass < 2; pass++) {
    const files = unzipSync(bytes);
    // Word's loader dispatches these parts by their exact content type. The
    // ms-word prefix belongs to other Office extensions and triggers repair.
    const types = strFromU8(files["[Content_Types].xml"]!);
    for (const part of ["commentsExtended", "commentsIds", "commentsExtensible", "people"])
      expect(types).toContain('PartName="/word/' + part + '.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.' + part + '+xml"');
    const document = strFromU8(files["word/document.xml"]!);
    expect(document).toContain('<w:commentReference w:id="0"/>');
    expect(document).toContain('<w:commentReference w:id="1"/>');
    expect(document.match(/<w:commentRangeStart\b/g)).toHaveLength(1);
    expect(strFromU8(files["word/commentsExtended.xml"]!)).toContain("paraIdParent");
    expect(strFromU8(files["word/commentsIds.xml"]!)).toContain("durableId");
    const model = openDocument(
        importDocx(bytes, { documentId: "reopen", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
      annotation = model.snapshot().annotations[0]!;
    expect(model.snapshot().annotations).toHaveLength(1);
    expect(annotation.anchor.kind === "text" ? annotation.anchor.quote : "").toBe("😀");
    expect(annotation.replies).toHaveLength(1);
    expect(model.getNode(annotation.replies![0]!.bodyNodeId!)!.text).toBe(
      "第一回复第二段漢\u{e0100}",
    );
    expect(annotation.date).toBe("2026-11-14T02:00:00.000Z");
    expect(annotation.replies![0]!.author?.name).toBe("乙");
    const binding = model.snapshot().metadata.word!.comments![0]!,
      ids = [binding.durableId!, binding.replies[0]!.durableId!];
    if (durable) expect(ids).toEqual(durable);
    durable = ids;
    const edit = createDocumentSession(model);
    edit.commit(
      edit.prepare({
        protocolVersion: 1,
        documentId: "reopen",
        baseRevision: 0,
        operations: [
          {
            op: "updateAnnotation",
            annotationId: annotation.annotationId,
            changes: { status: "resolved" },
          },
        ],
      }),
    );
    bytes = exportDocx(edit).bytes;
  }
});
test("unknown modern message properties remain attached to their own extension types after editing", () => {
  const files = unzipSync(exportDocx(fixture()).bytes);
  files["word/commentsExtended.xml"] = strToU8(
    strFromU8(files["word/commentsExtended.xml"]!).replace(
      "<w15:commentEx ",
      '<w15:commentEx w15:future="status" ',
    ),
  );
  files["word/commentsIds.xml"] = strToU8(
    strFromU8(files["word/commentsIds.xml"]!).replace(
      "<w16cid:commentId ",
      '<w16cid:commentId w16cid:future="identity" ',
    ),
  );
  files["word/commentsExtensible.xml"] = strToU8(
    strFromU8(files["word/commentsExtensible.xml"]!).replace(
      "<w16cex:commentExtensible ",
      '<w16cex:commentExtensible w16cex:future="extension" ',
    ),
  );
  const model = openDocument(importDocx(zipSync(files), { documentId: "preserved" }), {
      kind: "document",
    }),
    output = unzipSync(exportDocx(model).bytes);
  for (const [path, value] of [
    ["word/commentsExtended.xml", "status"],
    ["word/commentsIds.xml", "identity"],
    ["word/commentsExtensible.xml", "extension"],
  ]) {
    const xml = strFromU8(output[path!]!);
    expect(xml).toContain('future="' + value + '"');
    expect(xml).not.toContain("urn:doffi:transient:native");
  }
  expect(() => importDocx(exportDocx(model).bytes, { documentId: "again" })).not.toThrow();
});
test("every message in multiple ranged and point threads gets one reference without duplicating the root range", () => {
  const session = fixture();
  session.commit(session.prepare({ protocolVersion: 1, documentId: "comments", baseRevision: session.snapshot().manifest.revision,
    operations: [
      { op: "addWordCommentReply", nodeId: "paragraph-1", annotationId: "a", html: "<p>另一个回复</p>", author: { name: "丙" } },
      { op: "addAnnotation", annotation: { annotationId: "point", documentId: "comments",
        baseRevision: session.snapshot().manifest.revision, status: "open", instruction: "点批注",
        anchor: { kind: "text", nodeId: "paragraph-1", from: 2, to: 2, quote: "" } } },
    ] }));
  session.commit(session.prepare({ protocolVersion: 1, documentId: "comments", baseRevision: session.snapshot().manifest.revision,
    operations: [{ op: "addWordCommentReply", nodeId: "paragraph-1", annotationId: "point", html: "<p>点批注回复</p>" }] }));
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const parts = unzipSync(bytes), document = strFromU8(parts["word/document.xml"]!);
    const references = [...document.matchAll(/<w:commentReference w:id="([^"]+)"\/>/g)].map(match => match[1]);
    expect(references).toHaveLength(5);
    expect(new Set(references).size).toBe(5);
    expect(document.match(/<w:commentRangeStart\b/g)).toHaveLength(2);
    expect(document.match(/<w:commentRangeEnd\b/g)).toHaveLength(2);
    const reopened = openDocument(importDocx(bytes, { documentId: "multiple-threads", unsupportedContent: "reject" }), { kind: "document" });
    expect(reopened.snapshot().annotations.map(item => item.replies?.length ?? 0).sort()).toEqual([1, 2]);
    bytes = exportDocx(reopened).bytes;
  }
});
test("reply parent cycles are malformed data rather than a compatibility recovery retry", () => {
  const files = unzipSync(exportDocx(fixture()).bytes);
  files["word/commentsExtended.xml"] = strToU8(
    strFromU8(files["word/commentsExtended.xml"]!).replace(
      '<w15:commentEx w15:paraId="00000001"',
      '<w15:commentEx w15:paraIdParent="00000002" w15:paraId="00000001"',
    ),
  );
  expect(() => importDocx(zipSync(files), { documentId: "cycle" })).toThrow(/cycle/);
});
