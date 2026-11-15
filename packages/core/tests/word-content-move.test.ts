import { expect, test } from "vitest";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentOperation, ContentDocumentModel } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";

function setup(html = "<p>AMOVEZ</p><p>TARGET</p>") {
  const documentId = "clipboard-move",
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
    { op: "replaceFragment", nodeId: "section-root-1", html },
    { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "Receiver" },
  ]);
  const paragraphs = () =>
    [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
      .map((m) => session.getNode(m[1]!)!)
      .filter((n) => n.tagName === "p");
  return { session, apply, paragraphs };
}
function root(model: ContentDocumentModel) {
  let node = model.getNode(model.snapshot().metadata.word!.stories[0]!.rootNodeId)!;
  while (node.parentNodeId) node = model.getNode(node.parentNodeId)!;
  return node.nodeId;
}
function text(model: ContentDocumentModel) {
  return model.getNode(model.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text;
}
for (const block of [false, true])
  for (const decision of ["accept", "reject"] as const)
    test(
      "typed cut creates paired editable move with grouped review and exact anchor " +
        block +
        " " +
        decision,
      () => {
        const { session, apply, paragraphs } = setup(
          block ? "<p>ALEFT</p><p>MID</p><p>RIGHTZ</p><p>TARGET</p>" : undefined,
        );
        const ps = paragraphs(),
          first = ps[0]!,
          last = ps[block ? 2 : 0]!,
          target = ps.at(-1)!,
          cutId = "test-cut";
        apply([
          {
            op: "addAnnotation",
            annotation: {
              annotationId: "moved",
              documentId: "clipboard-move",
              baseRevision: session.snapshot().manifest.revision,
              status: "open",
              instruction: "moved",
              anchor: {
                kind: "text",
                nodeId: first.nodeId,
                from: 1,
                to: block ? 5 : 5,
                quote: block ? "LEFT" : "MOVE",
              },
            },
          },
        ]);
        const before = session.snapshot();
        apply([
          {
            op: "cutWordContent",
            nodeId: first.nodeId,
            endNodeId: last.nodeId,
            from: 1,
            to: 5,
            expectedText: block ? "LEFTMIDRIGHT" : "MOVE",
            cutId,
          },
        ]);
        const cut = session.snapshot();
        apply([
          {
            op: block ? "pasteWordBlockContent" : "pasteWordInlineContent",
            nodeId: target.nodeId,
            from: 3,
            to: 3,
            expectedText: "",
            html: block ? "<p>LEFT</p><p>MID</p><p>RIGHT</p>" : "<strong>MOVE</strong>",
            cutId,
          },
        ]);
        const moved = session.snapshot(),
          word = moved.metadata.word!;
        expect(word.moves).toHaveLength(2);
        expect(new Set(word.revisions.filter((r) => r.moveId).map((r) => r.moveId)).size).toBe(1);
        const annotation = session.getAnnotations()[0]!;
        expect(annotation.status).toBe("open");
        expect(annotation.anchor.kind).toBe("text");
        if (annotation.anchor.kind === "text")
          expect(
            session
              .getNode(annotation.anchor.nodeId)!
              .text.slice(annotation.anchor.from, annotation.anchor.to),
          ).toBe(block ? "LEFT" : "MOVE");
        // Reviewing either endpoint reaches the complete structural move.
        const revision = word.revisions.find((r) => r.kind === "moveTo")!;
        apply([
          {
            op: "reviewWordRevision",
            nodeId: revision.nodeId,
            revisionId: revision.revisionId ?? revision.nodeId,
            decision,
          },
        ]);
        expect(text(session)).toBe(
          decision === "accept"
            ? block
              ? "AZTARLEFTMIDRIGHTGET"
              : "AZTARMOVEGET"
            : block
              ? "ALEFTMIDRIGHTZTARGET"
              : "AMOVEZTARGET",
        );
        expect(session.snapshot().metadata.word!.moves).toHaveLength(0);
        session.undo();
        expect(session.snapshot().html).toBe(moved.html);
        let bytes = exportDocx(session).bytes;
        for (let pass = 0; pass < 2; pass++) {
          const reopened = createDocumentSession(
            openDocument(
              importDocx(bytes, { documentId: "clipboard-move", unsupportedContent: "reject" }),
              { kind: "document" },
            ),
          );
          expect(reopened.snapshot().metadata.word!.moves).toHaveLength(2);
          const r = reopened
            .snapshot()
            .metadata.word!.revisions.find((r) => r.kind === "moveFrom")!;
          reopened.commit(
            reopened.prepare({
              protocolVersion: 1,
              documentId: "clipboard-move",
              baseRevision: reopened.snapshot().manifest.revision,
              operations: [
                {
                  op: "reviewWordRevision",
                  nodeId: r.nodeId,
                  revisionId: r.revisionId ?? r.nodeId,
                  decision,
                },
              ],
            }),
          );
          expect(text(reopened)).toBe(
            decision === "accept"
              ? block
                ? "AZTARLEFTMIDRIGHTGET"
                : "AZTARMOVEGET"
              : block
                ? "ALEFTMIDRIGHTZTARGET"
                : "AMOVEZTARGET",
          );
          bytes = exportDocx(
            openDocument(
              importDocx(bytes, { documentId: "clipboard-move", unsupportedContent: "reject" }),
              { kind: "document" },
            ),
          ).bytes;
        }
        session.undo();
        expect(session.snapshot().html).toBe(cut.html);
        session.undo();
        expect(session.snapshot().html).toBe(before.html);
        session.redo();
        session.redo();
        expect(session.snapshot().html).toBe(moved.html);
      },
    );

for (const own of [false, true])
  for (const decision of ["accept", "reject"] as const)
    test(
      "block replacement partitions pending insertion and deletion histories " +
        own +
        " " +
        decision,
      () => {
        const { session, apply, paragraphs } = setup("<p>ABCDEF</p>"),
          p = paragraphs()[0]!.nodeId;
        apply([
          {
            op: "trackWordText",
            nodeId: p,
            from: 1,
            to: 1,
            expectedText: "",
            text: "Q",
            author: own ? "Receiver" : "Other",
          },
        ]);
        apply([
          {
            op: "trackWordText",
            nodeId: p,
            from: 3,
            to: 4,
            expectedText: "C",
            text: "",
            author: "Other",
          },
        ]);
        const before = session.snapshot();
        apply([
          {
            op: "pasteWordBlockContent",
            nodeId: p,
            from: 0,
            to: 7,
            expectedText: "AQBCDEF",
            html: "<p>NEW</p>",
          },
        ]);
        const pasted = session.snapshot();
        expect(pasted.metadata.word!.revisions.some((r) => r.author === "Other")).toBe(true);
        let bytes = exportDocx(session).bytes;
        for (let pass = 0; pass < 2; pass++) {
          const reopened = createDocumentSession(
            openDocument(
              importDocx(bytes, { documentId: "clipboard-move", unsupportedContent: "reject" }),
              { kind: "document" },
            ),
          );
          reopened.commit(
            reopened.prepare({
              protocolVersion: 1,
              documentId: "clipboard-move",
              baseRevision: reopened.snapshot().manifest.revision,
              operations: [{ op: "reviewWordRevisions", nodeId: root(reopened), decision }],
            }),
          );
          expect(text(reopened)).toBe(decision === "accept" ? "NEW" : "ABCDEF");
          bytes = exportDocx(
            openDocument(
              importDocx(bytes, { documentId: "clipboard-move", unsupportedContent: "reject" }),
              { kind: "document" },
            ),
          ).bytes;
        }
        apply([{ op: "reviewWordRevisions", nodeId: root(session), decision }]);
        expect(text(session)).toBe(decision === "accept" ? "NEW" : "ABCDEF");
        session.undo();
        expect(session.snapshot().html).toBe(pasted.html);
        session.undo();
        expect(session.snapshot().html).toBe(before.html);
      },
    );

test("a changed clipboard and consumed cut never fabricate a move", () => {
  const { session, apply, paragraphs } = setup(),
    [source, target] = paragraphs();
  apply([
    {
      op: "cutWordContent",
      nodeId: source!.nodeId,
      from: 1,
      to: 5,
      expectedText: "MOVE",
      cutId: "cut-changed",
    },
  ]);
  apply([
    {
      op: "pasteWordInlineContent",
      nodeId: target!.nodeId,
      from: 0,
      to: 0,
      expectedText: "",
      html: "DIFFERENT",
      cutId: "cut-changed",
    },
  ]);
  expect(session.snapshot().metadata.word!.moves ?? []).toHaveLength(0);
  apply([
    {
      op: "pasteWordInlineContent",
      nodeId: target!.nodeId,
      from: 0,
      to: 0,
      expectedText: "",
      html: "MOVE",
      cutId: "cut-changed",
    },
  ]);
  expect(session.snapshot().metadata.word!.moves).toHaveLength(2);
  apply([
    {
      op: "pasteWordInlineContent",
      nodeId: target!.nodeId,
      from: 0,
      to: 0,
      expectedText: "",
      html: "MOVE",
      cutId: "cut-changed",
    },
  ]);
  expect(session.snapshot().metadata.word!.moves).toHaveLength(2);
});

test("block paste splits an existing move without changing its original ownership", () => {
  const { session, apply, paragraphs } = setup(),
    [source, target] = paragraphs();
  apply([
    {
      op: "cutWordContent",
      nodeId: source!.nodeId,
      from: 1,
      to: 5,
      expectedText: "MOVE",
      cutId: "existing-move",
    },
  ]);
  apply([
    {
      op: "pasteWordInlineContent",
      nodeId: target!.nodeId,
      from: 3,
      to: 3,
      expectedText: "",
      html: "MOVE",
      cutId: "existing-move",
    },
  ]);
  const before = session.snapshot();
  apply([{ op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "Second" }]);
  apply([
    {
      op: "pasteWordBlockContent",
      nodeId: target!.nodeId,
      from: 4,
      to: 6,
      expectedText: "OV",
      html: "<div><p>NEW</p></div>",
    },
  ]);
  expect(
    session
      .snapshot()
      .metadata.word!.revisions.filter((r) => r.kind === "moveTo")
      .every((r) => r.author === "Receiver"),
  ).toBe(true);
  const second = session
    .snapshot()
    .metadata.word!.revisions.find((r) => r.author === "Second" && r.kind === "insert")!;
  apply([
    {
      op: "reviewWordRevision",
      nodeId: second.nodeId,
      revisionId: second.revisionId ?? second.nodeId,
      decision: "reject",
    },
  ]);
  expect(text(session)).toBe("AMOVEZTARMOVEGET");
  expect(session.snapshot().metadata.word!.moves).toHaveLength(2);
  expect(() =>
    importDocx(exportDocx(session).bytes, {
      documentId: "clipboard-move",
      unsupportedContent: "reject",
    }),
  ).not.toThrow();
  session.undo();
  session.undo();
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
});

test("cross replacement cancels a complete own pending table insertion and undo restores its grid", () => {
  const { session, apply, paragraphs } = setup("<p>A</p><p>Z</p>"),
    [first, last] = paragraphs();
  apply([{ op: "insertWordTable", nodeId: first!.nodeId, position: "after", rows: 2, columns: 2 }]);
  expect(session.snapshot().metadata.word!.revisions.some((r) => r.kind === "rowInsert")).toBe(
    true,
  );
  const before = session.snapshot();
  apply([
    {
      op: "pasteWordBlockContent",
      nodeId: first!.nodeId,
      endNodeId: last!.nodeId,
      from: 1,
      to: 0,
      expectedText: "",
      html: "<p>NEW</p>",
    },
  ]);
  expect(session.snapshot().metadata.word!.tables).toHaveLength(0);
  apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" }]);
  expect(text(session)).toBe("AZ");
  session.undo();
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  expect(session.snapshot().metadata.word!.tables).toEqual(before.metadata.word!.tables);
});

for (const decision of ["accept", "reject"] as const)
  test("empty native tables keep structural move review after reopen " + decision, () => {
    const { session, apply, paragraphs } = setup("<p>A</p><p>Z</p><p>TARGET</p>"),
      [first, last, target] = paragraphs();
    apply([
      { op: "setWordReview", nodeId: "document-root", trackChanges: false },
      { op: "insertWordTable", nodeId: first!.nodeId, position: "after", rows: 1, columns: 1 },
    ]);
    const table = session.snapshot().metadata.word!.tables[0]!,
      html = session.getNode(table.nodeId)!.html;
    apply([{ op: "setWordReview", nodeId: "document-root", trackChanges: true }]);
    apply([
      {
        op: "cutWordContent",
        nodeId: first!.nodeId,
        endNodeId: last!.nodeId,
        from: 1,
        to: 0,
        expectedText: "",
        cutId: "empty-table-move",
      },
    ]);
    apply([
      {
        op: "pasteWordBlockContent",
        nodeId: target!.nodeId,
        from: 3,
        to: 3,
        expectedText: "",
        html,
        semantics: { tables: [JSON.parse(JSON.stringify(table))] },
        cutId: "empty-table-move",
      },
    ]);
    expect(session.snapshot().metadata.word!.moves).toHaveLength(2);
    const reopened = createDocumentSession(
      openDocument(
        importDocx(exportDocx(session).bytes, {
          documentId: "clipboard-move",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    );
    const revision = reopened
      .snapshot()
      .metadata.word!.revisions.find((r) => r.kind === "rowDelete")!;
    expect(revision.moveId).toBeTruthy();
    reopened.commit(
      reopened.prepare({
        protocolVersion: 1,
        documentId: "clipboard-move",
        baseRevision: reopened.snapshot().manifest.revision,
        operations: [
          {
            op: "reviewWordRevision",
            nodeId: revision.nodeId,
            revisionId: revision.revisionId ?? revision.nodeId,
            decision,
          },
        ],
      }),
    );
    expect(reopened.snapshot().metadata.word!.tables).toHaveLength(1);
    expect(reopened.snapshot().metadata.word!.revisions).toHaveLength(0);
    expect(text(reopened)).toBe("AZTARGET");
  });
