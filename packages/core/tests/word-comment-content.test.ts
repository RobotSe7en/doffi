import { expect, test } from "vitest";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { wordCommentContentGraph } from "../../office/tests/word-comment-content-fixture.js";

function target(tracked = false) {
  const documentId = "comment-copy-target",
    session = createDocumentSession(
      openDocument(createWordDocument({ documentId }), { kind: "document" }),
    );
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
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "thread",
        documentId,
        baseRevision: 0,
        status: "open",
        anchor: { kind: "node", nodeId: "paragraph-1" },
        instruction: "接收方批注",
      },
    },
    { op: "setWordReview", nodeId: "document-root", trackChanges: tracked, author: "Receiver" },
  ]);
  return { session, apply };
}
for (const tracked of [false, true])
  for (const block of [false, true])
    test(
      "clipboard comment threads isolate identities and retain native editable bodies " +
        tracked +
        " " +
        block,
      () => {
        const graph = wordCommentContentGraph(),
          { session, apply } = target(tracked),
          before = session.snapshot();
        const op = {
          op: block ? ("pasteWordBlockContent" as const) : ("pasteWordInlineContent" as const),
          nodeId: "paragraph-1",
          from: 1,
          to: 1,
          expectedText: "",
          ...graph,
          ...(block ? { html: '<p data-node-id="copied-p">' + graph.html + "</p>" } : {}),
        };
        apply([op]);
        const pasted = session.snapshot(),
          root = pasted.annotations.find((a) => a.author?.name === "甲")!;
        expect(pasted.annotations).toHaveLength(3);
        expect(root.annotationId).not.toBe("thread");
        expect(root.status).toBe("resolved");
        expect(root.documentId).toBe("comment-copy-target");
        expect(root.replies).toHaveLength(2);
        expect(session.getNode(root.bodyNodeId!)!.text).toBe("7根批注😀链接");
        expect(session.getNode(root.replies![1]!.bodyNodeId!)!.text).toBe("第二回复");
        const binding = pasted.metadata.word!.comments!.find(
          (b) => b.annotationId === root.annotationId,
        )!;
        expect(binding.replies[1]!.parentReplyId).toBe(root.replies![0]!.replyId);
        expect(root.anchor.kind).toBe("text");
        expect(root.anchor.kind === "text" && root.anchor.quote).toBe("😀禰\u{e0100}重复");
        expect(pasted.annotations[0]!.instruction).toBe("接收方批注");
        expect(pasted.metadata.word!.fields).toHaveLength(1);
        expect(() => openDocument(session.serialize(), { kind: "document" })).not.toThrow();
        let bytes = exportDocx(session).bytes;
        for (let pass = 0; pass < 2; pass++) {
          const reopened = openDocument(
            importDocx(bytes, { documentId: "comment-reopen", unsupportedContent: "reject" }),
            { kind: "document" },
          );
          expect(reopened.snapshot().annotations).toHaveLength(3);
          const comment = reopened.snapshot().annotations.find((a) => a.author?.name === "甲")!;
          expect(comment.replies).toHaveLength(2);
          expect(comment.status).toBe("resolved");
          expect(reopened.getNode(comment.bodyNodeId!)!.text).toBe("7根批注😀链接");
          const anchor = comment.anchor;
          expect(
            anchor.kind === "text"
              ? anchor.quote
              : anchor.kind === "text-set"
                ? anchor.ranges.map((r) => r.quote).join("")
                : "",
          ).toBe("😀禰\u{e0100}重复");
          bytes = exportDocx(reopened).bytes;
        }
        session.undo();
        expect(session.snapshot().html).toBe(before.html);
        expect(session.snapshot().annotations.map((a) => a.annotationId)).toEqual(["thread"]);
        session.redo();
        expect(session.snapshot().annotations).toHaveLength(3);
        if (!block) {
          apply([op]);
          const all = session.snapshot().annotations,
            replies = all.flatMap((a) => a.replies ?? []);
          expect(all).toHaveLength(5);
          expect(new Set(all.map((a) => a.annotationId)).size).toBe(5);
          expect(new Set(replies.map((r) => r.replyId)).size).toBe(4);
          expect(
            session.snapshot().metadata.word!.stories.filter((s) => s.kind === "comment"),
          ).toHaveLength(6);
        }
      },
    );

test("comment transport validates anchors, ownership and parent graphs atomically", () => {
  const graph = wordCommentContentGraph(),
    { session, apply } = target(),
    before = session.snapshot();
  const variants = [
    (s: typeof graph.semantics) => {
      s.annotations![0]!.anchor = { kind: "node", nodeId: "paragraph-1" };
    },
    (s: typeof graph.semantics) => {
      const a = s.annotations![0]!;
      if (a.anchor.kind === "text") a.anchor.quote = "错误";
    },
    (s: typeof graph.semantics) => {
      const a = s.annotations![0]!;
      if (a.anchor.kind === "text") a.anchor.to = 1;
    },
    (s: typeof graph.semantics) => {
      s.annotations![0]!.bodyNodeId = "paragraph-1";
    },
    (s: typeof graph.semantics) => {
      s.annotations![0]!.replies![0]!.bodyNodeId = s.annotations![0]!.bodyNodeId!;
    },
    (s: typeof graph.semantics) => {
      s.comments![0]!.replies[0]!.parentReplyId = s.comments![0]!.replies[1]!.replyId;
    },
    (s: typeof graph.semantics) => {
      s.annotations!.push(s.annotations![0]!);
    },
    (s: typeof graph.semantics) => {
      s.comments!.push(s.comments![0]!);
    },
    (s: typeof graph.semantics) => {
      s.stories!.push({
        storyId: "unselected",
        rootNodeId: "extra",
        kind: "comment",
        html: '<div data-node-id="extra" data-office-story="comment"><p>PRIVATE</p></div>',
      });
    },
  ];
  for (const change of variants) {
    const semantics = structuredClone(graph.semantics);
    change(semantics);
    expect(() =>
      apply([
        {
          op: "pasteWordInlineContent",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          html: graph.html,
          semantics,
        },
      ]),
    ).toThrow();
    expect(session.snapshot()).toEqual(before);
  }
  const candidate = session.prepare({
    protocolVersion: 1,
    documentId: before.manifest.documentId,
    baseRevision: before.manifest.revision,
    operations: [
      {
        op: "pasteWordInlineContent",
        nodeId: "paragraph-1",
        from: 0,
        to: 0,
        expectedText: "",
        ...graph,
      },
    ],
  });
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "X" },
  ]);
  expect(() => session.commit(candidate)).toThrow(/revision/i);
});

test("comments created during paste map later edits and deletions in the same candidate", () => {
  const graph = wordCommentContentGraph(),
    { session, apply } = target();
  apply([
    {
      op: "pasteWordBlockContent",
      nodeId: "paragraph-1",
      from: 1,
      to: 1,
      expectedText: "",
      ...graph,
      html: '<p data-node-id="copied-p">' + graph.html + "</p>",
    },
    { op: "reviewWordRevisions", nodeId: "document-root", decision: "accept" },
  ]);
  const root = session.snapshot().annotations.find((a) => a.author?.name === "甲")!;
  expect(root.status).toBe("resolved");
  expect(root.anchor.kind === "text" && session.getNode(root.anchor.nodeId)!.text).toBe(
    "😀禰\u{e0100}重复",
  );
});

test("a verified tracked cut moves existing threads without duplicating their messages", async () => {
  const { wordCommentContentFixture } = await import(
      "../../office/tests/word-comment-content-fixture.js"
    ),
    fixture = wordCommentContentFixture(),
    session = fixture.session,
    graph = wordCommentContentGraph(),
    documentId = session.snapshot().manifest.documentId;
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId,
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  apply([{ op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "Mover" }]);
  // Use bodies/IDs from this exact source, rather than another fixture instance.
  const word = session.snapshot().metadata.word!,
    root = session.getAnnotations().find((a) => a.annotationId === "thread")!;
  graph.semantics.stories = word.stories
    .filter((s) => s.kind === "comment")
    .map((s) => ({ ...s, kind: "comment", html: session.getNode(s.rootNodeId)!.html }));
  graph.semantics.fields = JSON.parse(JSON.stringify(word.fields));
  graph.semantics.comments = JSON.parse(JSON.stringify(word.comments));
  for (const a of graph.semantics.annotations!) {
    const actual = session
      .getAnnotations()
      .find((source) => source.annotationId === a.annotationId)!;
    const anchor = a.anchor;
    Object.assign(a, JSON.parse(JSON.stringify(actual)), { anchor });
  }
  const before = session.snapshot();
  apply([
    {
      op: "cutWordContent",
      nodeId: fixture.paragraphId,
      from: fixture.from,
      to: fixture.to,
      expectedText: fixture.text.slice(fixture.from, fixture.to),
      cutId: "comment-cut",
    },
  ]);
  apply([
    {
      op: "pasteWordInlineContent",
      nodeId: fixture.paragraphId,
      from: 0,
      to: 0,
      expectedText: "",
      ...graph,
      cutId: "comment-cut",
    },
  ]);
  const after = session.snapshot();
  // The larger root range also includes an uncut prefix and remains a separate
  // clipped copy; the fully selected second comment follows its exact move.
  expect(after.annotations.filter((a) => a.instruction === "第二处重复")).toHaveLength(1);
  expect(after.annotations.find((a) => a.annotationId === "second")!.status).toBe("open");
  expect(after.annotations.find((a) => a.annotationId === "thread")!.bodyNodeId).toBe(
    root.bodyNodeId,
  );
  expect(after.metadata.word!.moves).toHaveLength(2);
  session.undo();
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
});
