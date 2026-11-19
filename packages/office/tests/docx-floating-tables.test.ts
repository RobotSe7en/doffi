import { expect, test } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  createDocumentSession,
  openDocument,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
} from "@doffi/core";
import type { ContentDocumentModel, DocumentOperation, SourceStyleNode } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { floatingTableFixture, floatingHeaderFixture } from "./floating-table-fixture.js";
function position(model: ContentDocumentModel, id: string) {
  const resolver = new SourceStyleResolver(
      model.snapshot().styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    project = (id: string): SourceStyleNode => {
      const node = model.getNode(id)!;
      return {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
    };
  return resolver.get(project(id));
}
for (const raw of [
  '<w:tblpPr w:horzAnchor="page" w:vertAnchor="margin" w:tblpX="1441" w:tblpY="1" w:leftFromText="120" w:rightFromText="180"/><w:tblOverlap w:val="never"/>',
  "<w:tblpPr/>",
  '<w:tblpPr w:tblpX="-240" w:tblpY="1.2pt" w:tblpXSpec="outside" w:tblpYSpec="center"/>',
])
  test(
    "floating table positioning remains editable and preserves native anchors and explicit empty overrides: " +
      raw,
    () => {
      let bytes = floatingTableFixture(raw);
      for (let pass = 0; pass < 3; pass++) {
        const model = openDocument(
            importDocx(bytes, { documentId: "floating-table", unsupportedContent: "reject" }),
            { kind: "document" },
          ),
          info = model.snapshot().metadata.word!.tables[0]!;
        expect(model.getNode(info.nodeId)!.text).toBe("CELL");
        expect(position(model, info.nodeId)["--word-table-floating"]).toBe(
          raw === "<w:tblpPr/>" ? "0" : "1",
        );
        expect(
          model
            .snapshot()
            .metadata.word!.nativeContent?.fragments.some((item) => item.kind !== "property") ??
            false,
        ).toBe(false);
        bytes = exportDocx(model).bytes;
        const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
        expect(xml).toContain("<w:tblpPr");
        if (raw.includes("never")) expect(xml).toContain('<w:tblOverlap w:val="never"/>');
      }
    },
  );
test("position formatting shares the native table review and undo boundary", () => {
  const session = createDocumentSession(
      openDocument(
        importDocx(floatingTableFixture(), {
          documentId: "floating-edit",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    word = session.snapshot().metadata.word!,
    table = word.tables[0]!.nodeId,
    root = session.getNode(word.stories[0]!.rootNodeId)!.parentNodeId!,
    before = session.snapshot(),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "floating-edit",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
    {
      op: "formatWordTable",
      nodeId: table,
      properties: {
        "--word-table-position-h-anchor": "page",
        "--word-table-position-x": "72.05pt",
        "--word-table-position-x-align": "none",
        "--word-table-overlap": "never",
      },
    },
  ]);
  const reopened = createDocumentSession(
      openDocument(
        importDocx(exportDocx(session).bytes, {
          documentId: "floating-review",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    info = reopened.snapshot().metadata.word!.tables[0]!;
  expect(position(reopened, info.nodeId)["--word-table-position-x"]).toBe("72.05pt");
  reopened.commit(
    reopened.prepare({
      protocolVersion: 1,
      documentId: "floating-review",
      baseRevision: 0,
      operations: [
        {
          op: "reviewWordRevisions",
          nodeId: reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!
            .parentNodeId!,
          decision: "reject",
        },
      ],
    }),
  );
  expect(position(reopened, info.nodeId)["--word-table-position-x"]).toBe("0.05pt");
  openDocument(
    importDocx(exportDocx(reopened).bytes, {
      documentId: "floating-rejected",
      unsupportedContent: "reject",
    }),
    { kind: "document" },
  );
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  session.redo();
  expect(position(session, table)["--word-table-position-x"]).toBe("72.05pt");
});
test("malformed native anchors and negative text distances are rejected", () => {
  for (const raw of [
    '<w:tblpPr w:horzAnchor="bad"/>',
    '<w:tblpPr w:tblpYSpec="bad"/>',
    '<w:tblpPr w:leftFromText="-1"/>',
    '<w:tblOverlap w:val="bad"/>',
  ])
    expect(() => importDocx(floatingTableFixture(raw), { documentId: "bad-floating" })).toThrow();
});

test("partial direct positioning overrides retain the complete inherited native position group", () => {
  const files = unzipSync(floatingTableFixture()),
    raw =
      '<w:tblpPr w:horzAnchor="text" w:vertAnchor="text" w:tblpX="1" w:tblpY="1" w:rightFromText="120" w:bottomFromText="120"/>';
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace(raw, '<w:tblStyle w:val="FloatingStyle"/>'),
  );
  files["word/styles.xml"] = strToU8(
    strFromU8(files["word/styles.xml"]!).replace(
      "</w:styles>",
      '<w:style w:type="table" w:styleId="FloatingStyle"><w:tblPr>' +
        raw +
        "</w:tblPr></w:style></w:styles>",
    ),
  );
  const session = createDocumentSession(
      openDocument(
        importDocx(zipSync(files), { documentId: "floating-style", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    ),
    id = session.snapshot().metadata.word!.tables[0]!.nodeId;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "floating-style",
      baseRevision: 0,
      operations: [
        { op: "formatWordTable", nodeId: id, properties: { "--word-table-position-x": "0.05pt" } },
      ],
    }),
  );
  expect(strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!)).toContain(
    "<w:tblpPr",
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "floating-style",
      baseRevision: 1,
      operations: [
        { op: "formatWordTable", nodeId: id, properties: { "--word-table-position-x": "48.05pt" } },
      ],
    }),
  );
  const root = session.getNode(
    session.snapshot().metadata.word!.stories[0]!.rootNodeId,
  )!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "floating-style",
      baseRevision: 2,
      operations: [
        { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
        { op: "formatWordTable", nodeId: id, properties: { "--word-table-position-y": "24.05pt" } },
      ],
    }),
  );
  const bytes = exportDocx(session).bytes,
    reopened = openDocument(
      importDocx(bytes, { documentId: "floating-style-reopened", unsupportedContent: "reject" }),
      { kind: "document" },
    ),
    info = reopened.snapshot().metadata.word!.tables[0]!;
  expect(position(reopened, info.nodeId)["--word-table-position-x"]).toBe("48.05pt");
  expect(position(reopened, info.nodeId)["--word-table-position-v-anchor"]).toBe("text");
  expect(position(reopened, info.nodeId)["--word-table-position-right-distance"]).toBe("6pt");
  const review = createDocumentSession(reopened);
  review.commit(
    review.prepare({
      protocolVersion: 1,
      documentId: "floating-style-reopened",
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
  expect(position(review, info.nodeId)["--word-table-position-x"]).toBe("48.05pt");
  expect(position(review, info.nodeId)["--word-table-position-y"]).toBe("0.05pt");
  exportDocx(review);
});

test("native floating header tables preserve their story ownership and editable table through reopens", () => {
  let bytes = floatingHeaderFixture();
  for (let pass = 0; pass < 3; pass++) {
    const model = openDocument(
      importDocx(bytes, { documentId: "floating-header", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(model.snapshot().metadata.word!.tables).toHaveLength(1);
    expect(model.snapshot().metadata.word!.stories.some((story) => story.kind === "header")).toBe(
      true,
    );
    expect(
      position(model, model.snapshot().metadata.word!.tables[0]!.nodeId)["--word-table-position-x"],
    ).toBe("72.05pt");
    bytes = exportDocx(model).bytes;
  }
});
