import { expect, test } from "vitest";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { strFromU8, unzipSync } from "fflate";
import { wordListContentFixture } from "./word-list-content-fixture.js";

for (const tableOnly of [false, true])
  for (const tracked of [false, true])
    test(
      "copied numbering and continuation references remain editable over strict native reopen: " +
        tableOnly +
        " " +
        tracked,
      () => {
        const documentId = "native-list-paste";
        const session = createDocumentSession(
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
            op: "replaceText",
            nodeId: "paragraph-1",
            from: 0,
            to: 0,
            expectedText: "",
            text: "AB",
          },
          {
            op: "setWordNumbering",
            nodeId: "document-root",
            numberingId: "1",
            levels: [
              {
                level: 0,
                start: 91,
                format: "decimal",
                text: "%1.",
                suffix: "tab",
                cssClass: "receiver-marker",
              },
            ],
          },
          {
            op: "setWordReview",
            nodeId: "document-root",
            trackChanges: tracked,
            author: "Receiver",
          },
        ]);
        const before = session.snapshot();
        const fragment = wordListContentFixture(tableOnly);
        expect(fragment.semantics.numbering!.map((item) => item.numberingId)).toEqual(
          tableOnly ? ["2"] : ["1", "2"],
        );
        apply([
          {
            op: "pasteWordBlockContent",
            nodeId: "paragraph-1",
            from: 1,
            to: 1,
            expectedText: "",
            ...fragment,
          },
        ]);
        expect(session.snapshot().metadata.word!.numbering[0]).toEqual(
          before.metadata.word!.numbering[0],
        );
        if (tableOnly) {
          const number =
            session.snapshot().metadata.word!.tables[0]!.continuationFormats![0]!
              .emptyParagraphs![0]!.numbering!;
          expect(number.numberingId).not.toBe("2");
          expect(
            session
              .snapshot()
              .metadata.word!.numbering.some((item) => item.numberingId === number.numberingId),
          ).toBe(true);
        }
        let bytes = exportDocx(session).bytes;
        for (let pass = 0; pass < 3; pass++) {
          const packageFiles = unzipSync(bytes);
          const numberingXml = strFromU8(packageFiles["word/numbering.xml"]!);
          expect(numberingXml).toContain("Liberation Serif");
          expect(numberingXml).toContain("Liberation Mono");
          expect(numberingXml).toContain("FandolFang");
          expect(numberingXml).toContain('w:val="10"');
          expect(numberingXml).toContain('<w:lvlRestart w:val="1"');
          expect(numberingXml).not.toContain("PRIVATE-MARKER");
          expect(strFromU8(packageFiles["word/document.xml"]!)).not.toContain("PRIVATE-BODY");
          const model = openDocument(
            importDocx(bytes, { documentId, unsupportedContent: "reject" }),
            { kind: "document" },
          );
          expect(model.snapshot().metadata.word!.numbering).toHaveLength(tableOnly ? 2 : 3);
          expect(model.snapshot().metadata.word!.numbering.at(-1)!.levels[0]!.start).toBe(10);
          bytes = exportDocx(model).bytes;
        }
        for (const decision of ["accept", "reject"] as const) {
          const review = createDocumentSession(
            openDocument(importDocx(bytes, { documentId, unsupportedContent: "reject" }), {
              kind: "document",
            }),
          );
          const root = review.getNode(
            review.snapshot().metadata.word!.stories[0]!.rootNodeId,
          )!.parentNodeId!;
          review.commit(
            review.prepare({
              protocolVersion: 1,
              documentId,
              baseRevision: 0,
              operations: [{ op: "reviewWordRevisions", nodeId: root, decision }],
            }),
          );
          if (!tableOnly && (!tracked || decision === "accept")) {
            const paragraph = [...review.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
              .map((match) => review.getNode(match[1]!)!)
              .find((node) => node.tagName === "p" && node.text === "RESTART")!;
            review.commit(
              review.prepare({
                protocolVersion: 1,
                documentId,
                baseRevision: review.snapshot().manifest.revision,
                operations: [
                  {
                    op: "replaceText",
                    nodeId: paragraph.nodeId,
                    from: 7,
                    to: 7,
                    expectedText: "",
                    text: "EDIT",
                  },
                ],
              }),
            );
          }
          if (tableOnly && (!tracked || decision === "accept")) {
            const tableId = review.snapshot().metadata.word!.tables[0]!.nodeId;
            const numberingId =
              review.snapshot().metadata.word!.tables[0]!.continuationFormats![0]!
                .emptyParagraphs![0]!.numbering!.numberingId;
            review.commit(
              review.prepare({
                protocolVersion: 1,
                documentId,
                baseRevision: review.snapshot().manifest.revision,
                operations: [
                  { op: "setWordReview", nodeId: root, trackChanges: false },
                  {
                    op: "editWordTable",
                    nodeId: tableId,
                    change: { kind: "splitCell", row: 0, column: 0 },
                  },
                ],
              }),
            );
            const paragraph = [...review.getNode(tableId)!.html.matchAll(/data-node-id="([^"]+)"/g)]
              .map((match) => review.getNode(match[1]!)!)
              .find(
                (node) =>
                  node.tagName === "p" && node.attributes["data-word-numbering"] === numberingId,
              )!;
            expect(paragraph).toBeDefined();
            expect(paragraph.attributes["data-word-level"]).toBe("1");
            review.commit(
              review.prepare({
                protocolVersion: 1,
                documentId,
                baseRevision: review.snapshot().manifest.revision,
                operations: [
                  {
                    op: "replaceText",
                    nodeId: paragraph.nodeId,
                    from: 0,
                    to: 0,
                    expectedText: "",
                    text: "CONTINUATION-EDIT",
                  },
                ],
              }),
            );
          }
          const reopened = openDocument(
            importDocx(exportDocx(review).bytes, { documentId, unsupportedContent: "reject" }),
            { kind: "document" },
          );
          const story = reopened.getNode(
            reopened.snapshot().metadata.word!.stories[0]!.rootNodeId,
          )!;
          if (tracked && decision === "reject") {
            expect(story.text).toBe("AB");
            expect(story.html.match(/<p\b/g)).toHaveLength(1);
          } else expect(story.text).toContain(tableOnly ? "MERGED" : "RESTARTEDIT");
          review.dispose();
        }
        session.undo();
        expect(session.snapshot().html).toBe(before.html);
        expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
        session.dispose();
      },
    );
