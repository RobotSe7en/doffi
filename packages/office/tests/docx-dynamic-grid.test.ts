import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { importDocx, importDocxWithReport, exportDocx } from "@doffi/office";
import { strFromU8, unzipSync } from "fflate";
import { mutateDocx, body, styles } from "./regression-fixtures.js";
function fixture(grid: string, span = 1, cellWidth = "1440", tableWidth = "2880", type = "dxa") {
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/document.xml": (source) =>
      body(
        source,
        '<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:type="' +
          type +
          '" w:w="' +
          tableWidth +
          '"/></w:tblPr><w:tblGrid>' +
          grid +
          '</w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="' +
          cellWidth +
          '"/><w:gridSpan w:val="' +
          span +
          '"/></w:tcPr><w:p><w:r><w:t>DYNAMIC</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>AFTER</w:t></w:r></w:p>',
      ),
  });
}
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
      expect(report.preserved.some((item) => item.kind === "block")).toBe(true);
      expect(() =>
        importDocx(input, { documentId: "strict", unsupportedContent: "reject" }),
      ).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
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
        expect(model.snapshot().metadata.word!.tables).toHaveLength(0);
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
