import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { wordInlineContentFixture } from "./word-inline-content-fixture.js";
for (const tracked of [false, true])
  test(
    "pasted inline field, formula and control remain editable across native export and reopens " +
      tracked,
    () => {
      const fragment = wordInlineContentFixture(),
        session = createDocumentSession(
          openDocument(createWordDocument({ documentId: "portable-inline-target" }), {
            kind: "document",
          }),
        );
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "portable-inline-target",
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
              author: "Target",
            },
            {
              op: "pasteWordInlineContent",
              nodeId: "paragraph-1",
              from: 1,
              to: 1,
              expectedText: "",
              html: fragment.html,
              semantics: fragment.semantics,
            },
          ],
        }),
      );
      let bytes = exportDocx(session).bytes;
      for (let pass = 0; pass < 3; pass++) {
        const model = openDocument(
            importDocx(bytes, {
              documentId: "portable-inline-reopened",
              unsupportedContent: "reject",
            }),
            { kind: "document" },
          ),
          word = model.snapshot().metadata.word!;
        expect(word.fields).toHaveLength(1);
        expect(word.equations).toHaveLength(1);
        expect(word.controls).toHaveLength(1);
        expect(word.hyperlinks).toHaveLength(1);
        expect(word.hyperlinks[0]!.target).toBe("https://example.com/source");
        expect(word.fields[0]!.instruction.trim()).toBe("PAGE");
        expect(word.equations![0]!.nodes.some((node) => node.kind === "fraction")).toBe(true);
        expect(word.controls![0]!.checked).toBe(true);
        if (tracked) expect(word.revisions.some((item) => item.kind === "insert")).toBe(true);
        bytes = exportDocx(model).bytes;
      }
      if (tracked)
        for (const decision of ["accept", "reject"] as const) {
          const review = createDocumentSession(
              openDocument(
                importDocx(bytes, {
                  documentId: "portable-reviewed",
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
              documentId: "portable-reviewed",
              baseRevision: 0,
              operations: [{ op: "reviewWordRevisions", nodeId: root, decision }],
            }),
          );
          const reopened = openDocument(
            importDocx(exportDocx(review).bytes, {
              documentId: "portable-reviewed-reopened",
              unsupportedContent: "reject",
            }),
            { kind: "document" },
          );
          expect(reopened.snapshot().metadata.word!.fields).toHaveLength(
            decision === "accept" ? 1 : 0,
          );
          expect(reopened.snapshot().metadata.word!.equations ?? []).toHaveLength(
            decision === "accept" ? 1 : 0,
          );
          expect(reopened.snapshot().metadata.word!.controls ?? []).toHaveLength(
            decision === "accept" ? 1 : 0,
          );
          expect(reopened.snapshot().metadata.word!.hyperlinks).toHaveLength(
            decision === "accept" ? 1 : 0,
          );
        }
    },
  );
test("empty cached fields keep a structural insertion boundary through native reopen and reject", () => {
  const fragment = wordInlineContentFixture(true),
    session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "portable-empty" }), { kind: "document" }),
    );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "portable-empty",
      baseRevision: 0,
      operations: [
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "Target" },
        {
          op: "pasteWordInlineContent",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          html: fragment.html,
          semantics: fragment.semantics,
        },
      ],
    }),
  );
  const review = createDocumentSession(
    openDocument(
      importDocx(exportDocx(session).bytes, {
        documentId: "portable-empty-reopened",
        unsupportedContent: "reject",
      }),
      { kind: "document" },
    ),
  );
  expect(review.snapshot().metadata.word!.fields).toHaveLength(1);
  expect(review.snapshot().metadata.word!.revisions.some((item) => item.kind === "insert")).toBe(
    true,
  );
  review.commit(
    review.prepare({
      protocolVersion: 1,
      documentId: "portable-empty-reopened",
      baseRevision: 0,
      operations: [
        {
          op: "reviewWordRevisions",
          nodeId: review.getNode(review.snapshot().metadata.word!.stories[0]!.rootNodeId)!
            .parentNodeId!,
          decision: "reject",
        },
      ],
    }),
  );
  expect(review.snapshot().metadata.word!.fields).toHaveLength(0);
  exportDocx(review);
});
