import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
function setup() {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "thread" }), { kind: "document" }),
  );
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "thread",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  apply([
    {
      op: "replaceText",
      nodeId: "paragraph-1",
      from: 0,
      to: 0,
      expectedText: "",
      text: "正文漢\u{e0100}",
    },
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "a",
        documentId: "thread",
        baseRevision: 0,
        status: "open",
        anchor: { kind: "node", nodeId: "paragraph-1" },
        instruction: "根批注",
      },
    },
  ]);
  return { session, apply };
}
test("Word replies and rich bodies share atomic history and remove their complete owned stories", () => {
  const { session, apply } = setup();
  apply([
    {
      op: "addWordCommentReply",
      nodeId: "paragraph-1",
      annotationId: "a",
      html: "<p>回复漢\u{e0100}</p>",
      author: { name: "乙" },
    },
  ]);
  const reply = session.getAnnotations()[0]!.replies![0]!,
    root = reply.bodyNodeId!;
  expect(session.getNode(root)!.text).toBe("回复漢\u{e0100}");
  apply([
    {
      op: "setWordCommentBody",
      nodeId: root,
      annotationId: "a",
      replyId: reply.replyId,
      html: "<p><strong>修改</strong></p>",
    },
  ]);
  expect(session.getNode(root)!.text).toBe("修改");
  session.undo();
  expect(session.getNode(root)!.text).toBe("回复漢\u{e0100}");
  apply([{ op: "removeAnnotation", annotationId: "a" }]);
  expect(session.getNode(root)).toBeUndefined();
  session.undo();
  expect(session.getNode(root)!.text).toBe("回复漢\u{e0100}");
  expect(() => openDocument(session.serialize(), { kind: "document" })).not.toThrow();
});
test("Word message targets and parent references reject unrelated edits atomically", () => {
  const { session, apply } = setup();
  expect(() =>
    apply([
      {
        op: "addWordCommentReply",
        nodeId: "document-root",
        annotationId: "a",
        html: "<p>越界</p>",
      },
    ]),
  ).toThrow();
  apply([
    { op: "addWordCommentReply", nodeId: "paragraph-1", annotationId: "a", html: "<p>一</p>" },
  ]);
  const first = session.getAnnotations()[0]!.replies![0]!;
  apply([
    {
      op: "addWordCommentReply",
      nodeId: "paragraph-1",
      annotationId: "a",
      parentReplyId: first.replyId,
      html: "<p>二</p>",
    },
  ]);
  const second = session.getAnnotations()[0]!.replies![1]!;
  apply([
    {
      op: "removeWordCommentReply",
      nodeId: first.bodyNodeId!,
      annotationId: "a",
      replyId: first.replyId,
    },
  ]);
  expect(session.snapshot().metadata.word!.comments![0]!.replies).toEqual([
    { replyId: second.replyId },
  ]);
  expect(session.getNode(first.bodyNodeId!)).toBeUndefined();
  const before = session.snapshot();
  expect(() =>
    apply([
      {
        op: "setWordCommentBody",
        nodeId: second.bodyNodeId!,
        annotationId: "a",
        html: "<p>越界</p>",
      },
    ]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});
