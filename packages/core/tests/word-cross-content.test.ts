import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
function fixture(tracked = false, table = false) {
  const documentId = "cross-content",
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
    {
      op: "replaceFragment",
      nodeId: "section-root-1",
      html:
        "<p>AA😀BB</p>" +
        (table
          ? "<table><tr><td><p>LEFT</p></td><td><p>RIGHT</p></td></tr></table>"
          : "<p>CCDD</p>") +
        "<p>EE禰\u{e0100}FF</p>",
    },
  ]);
  const paragraphs = () =>
    [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
      .map((match) => session.getNode(match[1]!)!)
      .filter((node) => node.tagName === "p");
  const first = paragraphs()[0]!.nodeId,
    last = paragraphs().at(-1)!.nodeId;
  apply([
    { op: "setWordReview", nodeId: "document-root", trackChanges: tracked, author: "Receiver" },
  ]);
  return { session, apply, paragraphs, first, last };
}
for (const tracked of [false, true])
  for (const block of [false, true])
    test(
      "cross paragraph replacement preserves both boundaries and grouped review " +
        tracked +
        " " +
        block,
      () => {
        const { session, apply, paragraphs, first, last } = fixture(tracked),
          before = session.snapshot();
        apply([
          {
            op: block ? "pasteWordBlockContent" : "pasteWordInlineContent",
            nodeId: first,
            endNodeId: last,
            from: 2,
            to: 5,
            expectedText: "😀BBCCDDEE禰\u{e0100}",
            html: block ? "<p>NEW</p><p>END</p>" : "<strong>NEW</strong>",
          },
        ]);
        if (tracked) {
          const inserted = session.snapshot();
          for (let pass = 0, bytes = exportDocx(session).bytes; pass < 2; pass++) {
            const reopened = createDocumentSession(
              openDocument(
                importDocx(bytes, {
                  documentId: "cross-content",
                  unsupportedContent: "reject",
                }),
                { kind: "document" },
              ),
            );
            expect(reopened.snapshot().metadata.word!.revisions.length).toBeGreaterThan(0);
            let root = reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!;
            while (root.parentNodeId) root = reopened.getNode(root.parentNodeId)!;
            reopened.commit(
              reopened.prepare({
                protocolVersion: 1,
                documentId: "cross-content",
                baseRevision: reopened.snapshot().manifest.revision,
                operations: [
                  { op: "reviewWordRevisions", nodeId: root.nodeId, decision: "accept" },
                ],
              }),
            );
            expect(
              reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text,
            ).toBe(block ? "AANEWENDFF" : "AANEWFF");
            bytes = exportDocx(
              openDocument(
                importDocx(bytes, { documentId: "cross-content", unsupportedContent: "reject" }),
                { kind: "document" },
              ),
            ).bytes;
          }
          apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision: "accept" }]);
          expect(session.getNode("main-story")!.text).toBe(block ? "AANEWENDFF" : "AANEWFF");
          session.undo();
          expect(session.snapshot().html).toBe(inserted.html);
          apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" }]);
          expect(session.getNode("main-story")!.text).toBe("AA😀BBCCDDEE禰\u{e0100}FF");
          expect(paragraphs()).toHaveLength(3);
          session.undo();
        } else {
          expect(session.getNode("main-story")!.text).toBe(block ? "AANEWENDFF" : "AANEWFF");
          expect(paragraphs()).toHaveLength(block ? 4 : 1);
        }
        session.undo();
        expect(session.snapshot().html).toBe(before.html);
        session.dispose();
      },
    );
test("explicit transfers preserve prefix and suffix anchors without searching repeated text", () => {
  const { session, apply, first, last } = fixture();
  apply([
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "prefix",
        documentId: "cross-content",
        baseRevision: session.snapshot().manifest.revision,
        status: "open",
        instruction: "prefix",
        anchor: { kind: "text", nodeId: first, from: 0, to: 2, quote: "AA" },
      },
    },
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "suffix",
        documentId: "cross-content",
        baseRevision: session.snapshot().manifest.revision,
        status: "open",
        instruction: "suffix",
        anchor: { kind: "text", nodeId: last, from: 5, to: 7, quote: "FF" },
      },
    },
  ]);
  const before = session.snapshot();
  const receipt = apply([
    {
      op: "pasteWordInlineContent",
      nodeId: first,
      endNodeId: last,
      from: 2,
      to: 5,
      expectedText: "😀BBCCDDEE禰\u{e0100}",
      html: "NEW",
    },
  ]);
  expect(session.getAnnotations().map((item) => item.status)).toEqual(["open", "open"]);
  expect(session.getAnnotations().map((item) => item.anchor)).toEqual([
    { kind: "text", nodeId: last, from: 0, to: 2, quote: "AA" },
    { kind: "text", nodeId: last, from: 5, to: 7, quote: "FF" },
  ]);
  expect(receipt.mapping.textTransfers).toBeTruthy();
  session.undo();
  expect(session.snapshot().annotations.map((item) => item.anchor)).toEqual(
    before.annotations.map((item) => item.anchor),
  );
  session.redo();
  expect(session.getAnnotations().every((item) => item.status === "open")).toBe(true);
  session.dispose();
});
test("cross paragraph candidates reject stale, reversed and split Unicode endpoints atomically", () => {
  const { session, apply, first, last } = fixture();
  const before = session.snapshot();
  for (const change of [
    { expectedText: "stale" },
    { nodeId: last, endNodeId: first },
    { from: 3 },
    { to: 4 },
  ]) {
    expect(() =>
      apply([
        {
          op: "pasteWordBlockContent",
          nodeId: first,
          endNodeId: last,
          from: 2,
          to: 5,
          expectedText: "😀BBCCDDEE禰\u{e0100}",
          html: "<p>NEW</p>",
          ...change,
        },
      ]),
    ).toThrow();
    expect(session.snapshot()).toEqual(before);
  }
  session.dispose();
});
for (const tracked of [false, true])
  test(
    "whole selected tables and partial cell boundaries remain structurally valid " + tracked,
    () => {
      const { session, apply, first, last } = fixture(tracked, true);
      apply([
        {
          op: "pasteWordInlineContent",
          nodeId: first,
          endNodeId: last,
          from: 2,
          to: 5,
          expectedText: "😀BBLEFTRIGHTEE禰\u{e0100}",
          html: "NEW",
        },
      ]);
      if (tracked)
        apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision: "accept" }]);
      expect(session.snapshot().html).not.toContain("<table");
      expect(session.getNode("main-story")!.text).toBe("AANEWFF");
      expect(() => openDocument(session.serialize(), { kind: "document" })).not.toThrow();
      session.dispose();
    },
  );
test("partial table selection keeps the table grid and clears only selected cell content", () => {
  const { session, apply, paragraphs, last } = fixture(false, true),
    first = paragraphs().find((node) => node.text === "LEFT")!.nodeId;
  apply([
    {
      op: "pasteWordBlockContent",
      nodeId: first,
      endNodeId: last,
      from: 2,
      to: 5,
      expectedText: "FTRIGHTEE禰\u{e0100}",
      html: "<p>NEW</p>",
    },
  ]);
  expect(session.snapshot().html).toContain("<table");
  expect(session.getNode("main-story")!.text).toBe("AA😀BBLENEWFF");
  expect(() => openDocument(session.serialize(), { kind: "document" })).not.toThrow();
  session.dispose();
});
