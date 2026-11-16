import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { importDocx, importDocxWithReport, exportDocx } from "@doffi/office";
import { strFromU8, unzipSync } from "fflate";
import { dynamicGridFixture as fixture } from "./dynamic-grid-fixture.js";
import { strToU8, zipSync } from "fflate";
import { rowGapFixture } from "./row-gap-fixture.js";
import { continuationCellFixture } from "./continuation-cell-fixture.js";

for (const [grid, span] of [
  ["", 1],
  ["<w:gridCol/>", 1],
  ['<w:gridCol w:w="0"/>', 1],
  ['<w:gridCol w:w="1440"/>', 2],
] as const)
  test(
    "valid inferred or expanded native grids remain importable and roundtrip without losing their table " +
      grid +
      " " +
      span,
    () => {
      const input = fixture(grid, span),
        { files, report } = importDocxWithReport(input, { documentId: "dynamic" });
      expect(report.preserved.some((item) => item.kind === "block")).toBe(false);
      expect(() =>
        importDocx(input, { documentId: "strict", unsupportedContent: "reject" }),
      ).not.toThrow();
      const session = createDocumentSession(openDocument(files, { kind: "document" })),
        nodes = [session.snapshot().metadata.word!.stories[0]!.rootNodeId];
      for (let index = 0; index < nodes.length; index++)
        nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
      const paragraph = nodes.find(
        (id) => session.getNode(id)!.tagName === "p" && session.getNode(id)!.text === "AFTER",
      )!;
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "dynamic",
          baseRevision: 0,
          operations: [
            {
              op: "replaceText",
              nodeId: paragraph,
              from: 0,
              to: 0,
              expectedText: "",
              text: "EDIT",
            },
          ],
        }),
      );
      let bytes = exportDocx(session).bytes;
      for (let pass = 0; pass < 2; pass++) {
        const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
        expect(xml).toContain("DYNAMIC");
        expect(xml).toContain("EDITAFTER");
        expect(xml).not.toContain("transient:native");
        const model = openDocument(importDocx(bytes, { documentId: "dynamic-reopen" }), {
          kind: "document",
        });
        expect(model.snapshot().metadata.word!.tables).toHaveLength(1);
        bytes = exportDocx(model).bytes;
      }
      session.undo();
      expect(session.getNode(paragraph)!.text).toBe("AFTER");
    },
  );
test("universal grid and preferred width units compile to editable point geometry and reopen strictly", () => {
  const model = openDocument(
    importDocx(fixture('<w:gridCol w:w="1in"/>', 1, "72pt", "100%", "pct"), {
      documentId: "units",
      unsupportedContent: "reject",
    }),
    { kind: "document" },
  );
  expect(model.snapshot().metadata.word!.tables[0]!.grid).toEqual([72]);
  expect(model.snapshot().html).toContain("width:72pt");
  expect(model.snapshot().html).toContain("width:100%");
  let bytes = exportDocx(model).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const reopened = openDocument(
      importDocx(bytes, { documentId: "units-reopen", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(reopened.snapshot().metadata.word!.tables[0]!.grid).toEqual([72]);
    bytes = exportDocx(reopened).bytes;
  }
});
test("negative and malformed grid widths still reject instead of entering native recovery", () => {
  for (const raw of ["-1", "broken"])
    expect(() =>
      importDocx(fixture('<w:gridCol w:w="' + raw + '"/>'), { documentId: "malformed-grid" }),
    ).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});

for (const input of [rowGapFixture(), continuationCellFixture()])
  test(
    "inferred omitted and vertically merged grids support internal editing, resizing and exact undo " +
      input.length,
    () => {
      const parts = unzipSync(input);
      parts["word/document.xml"] = strToU8(
        strFromU8(parts["word/document.xml"]!).replace(
          /<w:tblGrid>[\s\S]*?<\/w:tblGrid>/,
          "<w:tblGrid/>",
        ),
      );
      const session = createDocumentSession(
        openDocument(
          importDocx(zipSync(parts), { documentId: "infer-edit", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
      );
      const table = session.snapshot().metadata.word!.tables[0]!,
        before = session.snapshot(),
        root = session.getNode(
          session.snapshot().metadata.word!.stories[0]!.rootNodeId,
        )!.parentNodeId!;
      const apply = (operations: import("@doffi/core").DocumentOperation[]) =>
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: "infer-edit",
            baseRevision: session.snapshot().manifest.revision,
            operations,
          }),
        );
      const stale = session.prepare({
        protocolVersion: 1,
        documentId: "infer-edit",
        baseRevision: 0,
        operations: [
          {
            op: "formatWordTable",
            nodeId: table.nodeId,
            properties: { width: "80%", "table-layout": "auto" },
          },
        ],
      });
      apply([
        {
          op: "editWordTable",
          nodeId: table.nodeId,
          change: { kind: "insertRows", at: 1, count: 1 },
        },
        {
          op: "editWordTable",
          nodeId: table.nodeId,
          change: { kind: "resize", grid: table.grid.map((width, index) => width + index + 1) },
        },
        {
          op: "formatWordTable",
          nodeId: table.nodeId,
          properties: { "table-layout": "auto", width: "80%" },
        },
      ]);
      expect(() => session.commit(stale)).toThrowError(
        expect.objectContaining({ code: "RevisionConflict" }),
      );
      let bytes = exportDocx(session).bytes;
      for (let pass = 0; pass < 2; pass++) {
        const reopened = openDocument(
          importDocx(bytes, { documentId: "infer-reopen", unsupportedContent: "reject" }),
          { kind: "document" },
        );
        expect(reopened.snapshot().metadata.word!.tables[0]!.grid.length).toBe(table.grid.length);
        expect(reopened.snapshot().html).toContain("width:80%");
        bytes = exportDocx(reopened).bytes;
      }
      session.undo();
      expect(session.snapshot().html).toBe(before.html);
      expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
      apply([
        { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
        {
          op: "editWordTable",
          nodeId: table.nodeId,
          change: { kind: "resize", grid: table.grid.map((w) => w + 5) },
        },
      ]);
      const output = exportDocx(session).bytes;
      const reviewed = createDocumentSession(
        openDocument(
          importDocx(output, { documentId: "infer-review", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
      );
      reviewed.commit(
        reviewed.prepare({
          protocolVersion: 1,
          documentId: "infer-review",
          baseRevision: 0,
          operations: [
            {
              op: "reviewWordRevisions",
              nodeId: reviewed.getNode(reviewed.snapshot().metadata.word!.stories[0]!.rootNodeId)!
                .parentNodeId!,
              decision: "reject",
            },
          ],
        }),
      );
      expect(reviewed.snapshot().metadata.word!.tables[0]!.grid).toEqual(table.grid);
    },
  );
test("auto and nil preferred widths retain their native type and excessive inferred spans fail before allocation", () => {
  for (const type of ["auto", "nil"]) {
    let bytes = fixture("", 1, "1440", "0", type);
    for (let pass = 0; pass < 2; pass++) {
      const model = openDocument(
        importDocx(bytes, { documentId: "auto-grid", unsupportedContent: "reject" }),
        { kind: "document" },
      );
      expect(model.snapshot().metadata.word!.tables).toHaveLength(1);
      bytes = exportDocx(model).bytes;
      expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).toContain('w:type="' + type + '"');
    }
  }
  expect(() => importDocx(fixture("", 1000000), { documentId: "grid-budget" })).toThrowError(
    expect.objectContaining({ code: "ResourceLimitExceeded" }),
  );
});

test("missing-grid fixed tables use their section column and reject repeated singleton grid definitions", () => {
  const parts = unzipSync(fixture("", 1, "0", "0", "auto"));
  parts["word/document.xml"] = strToU8(
    strFromU8(parts["word/document.xml"]!).replace(
      "</w:body>",
      '<w:sectPr><w:pgSz w:w="6000" w:h="6000"/><w:pgMar w:top="480" w:right="480" w:bottom="480" w:left="480" w:header="0" w:footer="0" w:gutter="0"/><w:cols w:num="2" w:space="480"/></w:sectPr></w:body>',
    ),
  );
  const model = openDocument(
    importDocx(zipSync(parts), { documentId: "column-inference", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  expect(model.snapshot().metadata.word!.tables[0]!.grid).toEqual([114]);
  parts["word/document.xml"] = strToU8(
    strFromU8(parts["word/document.xml"]!).replace("<w:tblGrid>", "<w:tblGrid/><w:tblGrid>"),
  );
  expect(() => importDocx(zipSync(parts), { documentId: "repeated-grid" })).toThrowError(
    expect.objectContaining({ code: "InvalidDocument" }),
  );
});
