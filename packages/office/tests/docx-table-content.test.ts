import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { unzipSync, strFromU8 } from "fflate";
import {
  wordTableContentFixture,
  nestedTableContentFixture,
} from "./word-table-content-fixture.js";
for (const tracked of [false, true])
  test(
    "copied native merged tables retain editable continuation formats across export and review " +
      tracked,
    () => {
      const fragment = wordTableContentFixture(),
        session = createDocumentSession(
          openDocument(createWordDocument({ documentId: "native-table-content" }), {
            kind: "document",
          }),
        );
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "native-table-content",
          baseRevision: 0,
          operations: [
            {
              op: "replaceText",
              nodeId: "paragraph-1",
              from: 0,
              to: 0,
              expectedText: "",
              text: "AB",
            },
            {
              op: "setWordReview",
              nodeId: "document-root",
              trackChanges: tracked,
              author: "Receiver",
            },
            {
              op: "pasteWordBlockContent",
              nodeId: "paragraph-1",
              from: 1,
              to: 1,
              expectedText: "",
              ...fragment,
            },
          ],
        }),
      );
      let bytes = exportDocx(session).bytes;
      for (let pass = 0; pass < 3; pass++) {
        const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
        expect(xml.toLowerCase()).toContain("ffff88");
        expect(xml).toContain("880000");
        const model = openDocument(
          importDocx(bytes, { documentId: "table-content-reopen", unsupportedContent: "reject" }),
          { kind: "document" },
        );
        expect(model.snapshot().metadata.word!.tables).toHaveLength(1);
        expect(
          model.snapshot().metadata.word!.tables[0]!.continuationFormats![0]!.textDirection,
        ).toBe("tbRl");
        expect(
          model.snapshot().metadata.word!.revisions.filter((item) => item.kind === "rowInsert"),
        ).toHaveLength(tracked ? 2 : 0);
        bytes = exportDocx(model).bytes;
      }
      if (tracked)
        for (const decision of ["accept", "reject"] as const) {
          const review = createDocumentSession(
              openDocument(
                importDocx(bytes, {
                  documentId: "table-content-review",
                  unsupportedContent: "reject",
                }),
                { kind: "document" },
              ),
            ),
            root = review.getNode(
              review.snapshot().metadata.word!.stories[0]!.rootNodeId,
            )!.parentNodeId!;
          review.commit(
            review.prepare({
              protocolVersion: 1,
              documentId: "table-content-review",
              baseRevision: 0,
              operations: [{ op: "reviewWordRevisions", nodeId: root, decision }],
            }),
          );
          const model = openDocument(
            importDocx(exportDocx(review).bytes, {
              documentId: "table-content-reviewed",
              unsupportedContent: "reject",
            }),
            { kind: "document" },
          );
          expect(model.snapshot().metadata.word!.tables).toHaveLength(
            decision === "accept" ? 1 : 0,
          );
          if (decision === "reject") {
            expect(
              model.getNode(model.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text,
            ).toBe("AB");
            expect(model.snapshot().html.match(/<p\b/g)).toHaveLength(1);
          }
        }
    },
  );

for (const decision of ["accept", "reject"] as const)
  test("nested table paste stays editable through native review and undo: " + decision, () => {
    const documentId = "nested-table-content";
    const session = createDocumentSession(
      openDocument(createWordDocument({ documentId }), { kind: "document" }),
    );
    const apply = (operations: import("@doffi/core").DocumentOperation[]) =>
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
      { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "Receiver" },
    ]);
    const before = session.snapshot();
    apply([
      {
        op: "pasteWordBlockContent",
        nodeId: "paragraph-1",
        from: 1,
        to: 1,
        expectedText: "",
        ...wordTableContentFixture(nestedTableContentFixture()),
      },
    ]);
    expect(session.snapshot().metadata.word!.tables.map((table) => table.grid)).toEqual([
      [240],
      [120],
    ]);
    expect(
      session.snapshot().metadata.word!.revisions.filter((item) => item.kind === "rowInsert"),
    ).toHaveLength(4);
    const outer = session.snapshot().metadata.word!.tables[0]!.nodeId;
    const innerParagraph = [...session.getNode(outer)!.html.matchAll(/data-node-id="([^"]+)"/g)]
      .map((match) => session.getNode(match[1]!)!)
      .find((node) => node.tagName === "p" && node.text === "INNER-B")!.nodeId;
    apply([
      {
        op: "trackWordText",
        nodeId: innerParagraph,
        from: 7,
        to: 7,
        expectedText: "",
        text: "EDIT",
        author: "Receiver",
      },
    ]);
    expect(session.getNode(innerParagraph)!.text).toBe("INNER-BEDIT");
    const edited = session.snapshot();
    let bytes = exportDocx(session).bytes;
    for (let pass = 0; pass < 2; pass++) {
      const reopened = createDocumentSession(
        openDocument(importDocx(bytes, { documentId, unsupportedContent: "reject" }), {
          kind: "document",
        }),
      );
      expect(reopened.snapshot().metadata.word!.tables).toHaveLength(2);
      expect(
        reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text,
      ).toContain("INNER-BEDIT");
      const root = reopened.getNode(
        reopened.snapshot().metadata.word!.stories[0]!.rootNodeId,
      )!.parentNodeId!;
      reopened.commit(
        reopened.prepare({
          protocolVersion: 1,
          documentId,
          baseRevision: 0,
          operations: [{ op: "reviewWordRevisions", nodeId: root, decision }],
        }),
      );
      const reviewed = openDocument(
        importDocx(exportDocx(reopened).bytes, { documentId, unsupportedContent: "reject" }),
        { kind: "document" },
      );
      expect(reviewed.snapshot().metadata.word!.tables).toHaveLength(decision === "accept" ? 2 : 0);
      const story = reviewed.getNode(reviewed.snapshot().metadata.word!.stories[0]!.rootNodeId)!;
      if (decision === "reject") {
        expect(story.text).toBe("AB");
        expect(story.html.match(/<p\b/g)).toHaveLength(1);
      } else expect(story.text).toBe("AOUTERINNER-AINNER-BEDITTAILB");
      reopened.dispose();
      bytes = exportDocx(
        openDocument(importDocx(bytes, { documentId, unsupportedContent: "reject" }), {
          kind: "document",
        }),
      ).bytes;
    }
    const inserted = session
      .snapshot()
      .metadata.word!.revisions.find((item) => item.kind === "rowInsert")!;
    apply([
      {
        op: "reviewWordRevision",
        nodeId: inserted.nodeId,
        revisionId: inserted.revisionId!,
        decision,
      },
    ]);
    expect(session.snapshot().metadata.word!.tables).toHaveLength(decision === "accept" ? 2 : 0);
    expect(session.snapshot().metadata.word!.revisions).toEqual([]);
    session.undo();
    expect(session.snapshot().html).toBe(edited.html);
    session.undo();
    expect(session.getNode(innerParagraph)!.text).toBe("INNER-B");
    session.undo();
    expect(session.snapshot().html).toBe(before.html);
    expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
    session.dispose();
  });
