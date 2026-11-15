import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WordContentSemantics } from "@doffi/core";
const html =
  '<div><table data-node-id="t" class="foreign-table" style="width:144pt;border-collapse:collapse"><tbody data-node-id="body"><tr data-node-id="r0"><td data-node-id="c0" rowspan="2"><p data-node-id="p" class="foreign-p">MERGED</p></td><td data-node-id="c1"><p data-node-id="p1">ONE</p></td></tr><tr data-node-id="r1"><td data-node-id="c2"><p data-node-id="p2">TWO</p></td></tr></tbody></table></div>';
const semantics: WordContentSemantics = {
  tables: [
    {
      nodeId: "t",
      grid: [72, 72],
      styleContexts: [
        { nodeId: "t", kind: "table", cssClass: "foreign-table" },
        { nodeId: "p", kind: "paragraph", cssClass: "foreign-p" },
      ],
      continuationFormats: [
        {
          row: 1,
          column: 0,
          cellCssClass: "continuation-cell",
          emptyParagraphs: [
            {
              paragraphCssClass: "continuation-p",
              runCssClass: "continuation-mark",
              runs: [{ cssClass: "continuation-run" }],
            },
          ],
        },
      ],
    },
  ],
  tabs: [
    {
      paragraphNodeId: "p",
      defaultInterval: 36,
      stops: [{ position: 18, alignment: "center", leader: "dot" }],
    },
  ],
  conditionalFormats: [{ nodeId: "c0", format: { firstRow: false } }],
  extraClasses: {
    "continuation-cell": { "background-color": "#ffff88", "padding-left": "9pt" },
    "continuation-p": { "text-align": "right" },
    "continuation-mark": { "font-size": "14pt" },
    "continuation-run": { color: "#123456" },
  },
};
for (const tracked of [false, true])
  test(
    "whole-table paste owns its grid, continuation formats and typed row review in one transaction " +
      tracked,
    () => {
      const session = createDocumentSession(
          openDocument(createWordDocument({ documentId: "table-content" }), { kind: "document" }),
        ),
        apply = (operations: DocumentOperation[]) =>
          session.commit(
            session.prepare({
              protocolVersion: 1,
              documentId: "table-content",
              baseRevision: session.snapshot().manifest.revision,
              operations,
            }),
          );
      apply([
        { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
        { op: "setWordReview", nodeId: "document-root", trackChanges: tracked, author: "Receiver" },
      ]);
      const before = session.snapshot();
      apply([
        {
          op: "pasteWordBlockContent",
          nodeId: "paragraph-1",
          from: 1,
          to: 1,
          expectedText: "",
          html,
          semantics,
        },
      ]);
      const word = session.snapshot().metadata.word!,
        table = word.tables[0]!;
      expect(table.grid).toEqual([72, 72]);
      expect(table.nodeId).not.toBe("t");
      expect(table.continuationFormats![0]!.cellCssClass).not.toBe("continuation-cell");
      expect(session.getNode(table.nodeId)!.html).not.toMatch(/foreign-table|foreign-p/);
      expect(word.tabs[0]!.paragraphNodeId).not.toBe("p");
      expect(word.conditionalFormats![0]!.format.firstRow).toBe(false);
      expect(word.revisions.filter((item) => item.kind === "rowInsert")).toHaveLength(
        tracked ? 2 : 0,
      );
      expect(() => openDocument(session.serialize(), { kind: "document" })).not.toThrow();
      if (tracked) {
        apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" }]);
        expect(session.getNode("main-story")!.text).toBe("AB");
        expect(session.snapshot().metadata.word!.tables).toHaveLength(0);
        session.undo();
      }
      session.undo();
      expect(session.snapshot().html).toBe(before.html);
    },
  );
test("missing or injected continuation CSS references reject before changing source or target styles", () => {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "bad-table-content" }), { kind: "document" }),
    ),
    before = session.snapshot();
  expect(() =>
    session.prepare({
      protocolVersion: 1,
      documentId: "bad-table-content",
      baseRevision: 0,
      operations: [
        {
          op: "pasteWordBlockContent",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          html,
          semantics: { ...semantics, extraClasses: {} },
        },
      ],
    }),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});
