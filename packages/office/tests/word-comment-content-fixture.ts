import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WordContentSemantics } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";

export function wordCommentContentFixture(native = false) {
  const documentId = "comment-copy-source";
  let session = createDocumentSession(
    openDocument(createWordDocument({ documentId }), { kind: "document" }),
  );
  const text = "外重复😀禰\u{e0100}重复尾",
    from = 3,
    to = text.length - 1;
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId,
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text },
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "thread",
        documentId,
        baseRevision: 0,
        status: "resolved",
        anchor: { kind: "text", nodeId: "paragraph-1", from: 1, to, quote: text.slice(1, to) },
        instruction: "根批注",
        author: { name: "甲", initials: "J" },
        date: "2026-11-16T00:00:00Z",
      },
    },
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "second",
        documentId,
        baseRevision: 0,
        status: "open",
        anchor: { kind: "text", nodeId: "paragraph-1", from: to - 2, to, quote: "重复" },
        instruction: "第二处重复",
      },
    },
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "private",
        documentId,
        baseRevision: 0,
        status: "open",
        anchor: { kind: "text", nodeId: "paragraph-1", from: 0, to: 1, quote: "外" },
        instruction: "PRIVATE-COMMENT",
      },
    },
  ]);
  apply([
    {
      op: "setWordCommentBody",
      nodeId: "paragraph-1",
      annotationId: "thread",
      html: "<p><strong>根批注😀</strong><span>链接</span></p>",
    },
  ]);
  apply([
    {
      op: "addWordCommentReply",
      nodeId: session.getAnnotations()[0]!.bodyNodeId!,
      annotationId: "thread",
      html: "<p>第一回复</p>",
      author: { name: "乙" },
      date: "2026-11-16T01:00:00Z",
    },
  ]);
  const root = session.getAnnotations()[0]!,
    reply = root.replies![0]!;
  apply([
    {
      op: "addWordCommentReply",
      nodeId: root.bodyNodeId!,
      annotationId: "thread",
      parentReplyId: reply.replyId,
      html: "<p><em>第二回复</em></p>",
      author: { name: "丙" },
    },
    {
      op: "insertWordField",
      nodeId: session.getNode(root.bodyNodeId!)!.childNodeIds[0]!,
      offset: 0,
      instruction: " PAGE ",
      cachedText: "7",
    },
  ]);
  if (native)
    session = createDocumentSession(
      openDocument(
        importDocx(exportDocx(session).bytes, { documentId, unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    );
  const paragraphId = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
    .map((m) => session.getNode(m[1]!)!)
    .find((n) => n.tagName === "p" && n.text === text)!.nodeId;
  return { session, paragraphId, text, from, to };
}

/** Direct typed-transaction fixture; actual selection and clipboard provenance
 * are exercised independently by the browser cases. */
export function wordCommentContentGraph() {
  const fixture = wordCommentContentFixture(),
    snapshot = fixture.session.snapshot(),
    word = snapshot.metadata.word!;
  const annotations = JSON.parse(
    JSON.stringify(snapshot.annotations.filter((a) => a.annotationId !== "private")),
  ) as NonNullable<WordContentSemantics["annotations"]>;
  for (const annotation of annotations) {
    const anchor = annotation.anchor;
    if (anchor.kind !== "text") throw new Error("Fixture text anchor required");
    const from = Math.max(anchor.from, fixture.from),
      to = Math.min(anchor.to, fixture.to);
    annotation.anchor = {
      kind: "text",
      nodeId: "slice",
      from: from - fixture.from,
      to: to - fixture.from,
      quote: fixture.text.slice(from, to),
    };
  }
  const semantics: WordContentSemantics = {
    annotations,
    comments: word
      .comments!.filter((b) => b.annotationId === "thread")
      .map((b) => JSON.parse(JSON.stringify(b))),
    stories: word.stories
      .filter((s) => s.kind === "comment")
      .map((s) => ({
        ...s,
        kind: "comment" as const,
        html: fixture.session.getNode(s.rootNodeId)!.html,
      })),
    fields: JSON.parse(JSON.stringify(word.fields)),
  };
  return {
    html: '<span data-node-id="slice">' + fixture.text.slice(fixture.from, fixture.to) + "</span>",
    semantics,
  };
}
