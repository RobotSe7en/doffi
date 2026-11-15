import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { unzipSync, strFromU8 } from "fflate";
import { wordTableContentFixture } from "./word-table-content-fixture.js";
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
