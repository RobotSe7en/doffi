import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
function target(tracked = false) {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "block-content" }), { kind: "document" }),
    ),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "block-content",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "ABCD" },
    { op: "formatWordParagraphMark", nodeId: "paragraph-1", properties: { "font-size": "16pt" } },
    { op: "setWordReview", nodeId: "document-root", trackChanges: tracked, author: "Receiver" },
  ]);
  return { session, apply };
}
const content = {
  html: '<div style="font-family:serif"><p data-node-id="p0">FIRST</p><p data-node-id="p1"><span data-node-id="f" data-word-field="simple">9</span></p></div>',
  semantics: {
    fields: [
      { nodeId: "f", instruction: " PAGE ", resultState: "cached" as const, resultNodeIds: ["f"] },
    ],
  },
};
test("block content splits a paragraph while keeping source identities, semantics, marks and a single undo", () => {
  const { session, apply } = target(),
    before = session.snapshot();
  apply([
    {
      op: "pasteWordBlockContent",
      nodeId: "paragraph-1",
      from: 1,
      to: 3,
      expectedText: "BC",
      ...content,
    },
  ]);
  const snapshot = session.snapshot(),
    source = new Set(
      [...before.html.matchAll(/data-node-id="([^"]+)"/g)].map((match) => match[1]!),
    );
  expect(session.getNode("paragraph-1")!.text).toBe("A");
  expect(session.getNode("main-story")!.text).toBe("AFIRST9D");
  expect(snapshot.metadata.word!.fields).toHaveLength(1);
  expect(snapshot.html.match(/data-word-paragraph-mark="true"/g)).toHaveLength(2);
  expect(snapshot.metadata.word!.fields[0]!.nodeId).not.toBe("f");
  expect(source.has("paragraph-1")).toBe(true);
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
});
for (const decision of ["accept", "reject"] as const)
  test(
    "tracked paragraph and field insertion can be reviewed together through transparent format containers: " +
      decision,
    () => {
      const { session, apply } = target(true);
      apply([
        {
          op: "pasteWordBlockContent",
          nodeId: "paragraph-1",
          from: 1,
          to: 3,
          expectedText: "BC",
          ...content,
        },
      ]);
      expect(
        session
          .snapshot()
          .metadata.word!.revisions.filter((item) => item.kind === "paragraphInsert"),
      ).toHaveLength(3);
      const inserted = session.snapshot();
      apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision }]);
      expect(session.getNode("main-story")!.text).toBe(decision === "accept" ? "AFIRST9D" : "ABCD");
      expect(session.snapshot().metadata.word!.fields).toHaveLength(decision === "accept" ? 1 : 0);
      expect(session.snapshot().html.match(/<p\b/g)).toHaveLength(decision === "accept" ? 4 : 1);
      session.undo();
      expect(session.snapshot().html).toBe(inserted.html);
    },
  );
test("a new block paste beside an existing insertion can be rejected independently", () => {
  const { session, apply } = target(true);
  apply([
    {
      op: "trackWordText",
      nodeId: "paragraph-1",
      from: 4,
      to: 4,
      expectedText: "",
      text: "X",
      author: "Receiver",
    },
  ]);
  const original = session
    .snapshot()
    .metadata.word!.revisions.find((item) => item.kind === "insert")!;
  apply([
    {
      op: "pasteWordBlockContent",
      nodeId: "paragraph-1",
      from: 5,
      to: 5,
      expectedText: "",
      ...content,
    },
  ]);
  const inserted = session
    .snapshot()
    .metadata.word!.revisions.find(
      (item) => item.kind === "insert" && item.nodeId !== original.nodeId,
    )!;
  apply([
    {
      op: "reviewWordRevision",
      nodeId: inserted.nodeId,
      revisionId: inserted.revisionId!,
      decision: "reject",
    },
  ]);
  expect(session.getNode("main-story")!.text).toBe("ABCDX");
  expect(session.snapshot().metadata.word!.revisions).toContainEqual(original);
  expect(session.snapshot().metadata.word!.fields).toHaveLength(0);
});
test("invalid block graphs and stale text ranges fail atomically", () => {
  const { session, apply } = target(),
    before = session.snapshot();
  expect(() =>
    apply([
      {
        op: "pasteWordBlockContent",
        nodeId: "paragraph-1",
        from: 1,
        to: 1,
        expectedText: "",
        html: "<table><tr><td><p>X</p></td></tr></table>",
      },
    ]),
  ).toThrow();
  expect(() =>
    apply([
      {
        op: "pasteWordBlockContent",
        nodeId: "paragraph-1",
        from: 1,
        to: 1,
        expectedText: "",
        ...content,
        semantics: { fields: [{ ...content.semantics.fields[0]!, nodeId: "paragraph-1" }] },
      },
    ]),
  ).toThrow();
  expect(() =>
    apply([
      {
        op: "pasteWordBlockContent",
        nodeId: "paragraph-1",
        from: 1,
        to: 2,
        expectedText: "stale",
        ...content,
      },
    ]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});
