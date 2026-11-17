import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { importDocxWithReport, exportDocx } from "@doffi/office";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { drawingSizeFixture } from "./drawing-size-fixture.js";
const refs = ["page", "margin", "insideMargin", "outsideMargin", "leftMargin", "rightMargin"];
for (const [index, widthFrom] of refs.entries())
  test("native relative dimensions remain editable across two strict reopens: " + widthFrom, () => {
    const heightFrom = [
      "page",
      "margin",
      "insideMargin",
      "outsideMargin",
      "topMargin",
      "bottomMargin",
    ][index]!;
    let bytes = drawingSizeFixture({ widthFrom, heightFrom, width: "125123", height: "50000" });
    for (let pass = 0; pass < 2; pass++) {
      const imported = importDocxWithReport(bytes, {
        documentId: "size-native",
        unsupportedContent: "reject",
      });
      expect(imported.report.preserved).toEqual([]);
      const model = openDocument(imported.files, { kind: "document" }),
        drawing = model.snapshot().metadata.word!.drawings![0]!;
      expect(drawing.geometry).toMatchObject({ width: 160, height: 100 });
      expect(drawing.layout).toMatchObject({
        relativeWidth: { relativeFrom: widthFrom, percent: 125.123 },
        relativeHeight: { relativeFrom: heightFrom, percent: 50 },
      });
      bytes = exportDocx(model).bytes;
      const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
      expect(xml).toContain("125123");
      expect(xml).toContain('mc:Ignorable="wp14"');
    }
  });
for (const width of ["0", "abs"])
  test("disabled or absent relative dimensions retain their native declaration " + width, () => {
    const model = openDocument(
      importDocxWithReport(drawingSizeFixture({ width, height: "abs" }), {
        documentId: "size-zero",
        unsupportedContent: "reject",
      }).files,
      { kind: "document" },
    );
    expect(model.snapshot().metadata.word!.drawings![0]!.layout.relativeWidth?.percent).toBe(
      width === "0" ? 0 : undefined,
    );
    const xml = strFromU8(unzipSync(exportDocx(model).bytes)["word/document.xml"]!);
    expect(xml.includes("sizeRelH")).toBe(width === "0");
    expect(xml).not.toContain("sizeRelV");
  });
for (const value of ["-1", "NaN", "1.25", "2147483648", ""])
  test("invalid native percentage fails rather than becoming a preserved object: " + value, () => {
    expect(() =>
      importDocxWithReport(drawingSizeFixture({ width: value }), { documentId: "bad-size" }),
    ).toThrow();
  });
test("missing, duplicate, invalid-reference and inline relative sizes fail; unknown properties remain preserved", () => {
  const original = unzipSync(drawingSizeFixture()),
    source = strFromU8(original["word/document.xml"]!);
  for (const raw of [
    source.replace(/<wp14:pctWidth>[\s\S]*?<\/wp14:pctWidth>/, ""),
    source.replace(/(<wp14:sizeRelH[\s\S]*?<\/wp14:sizeRelH>)/, "$1$1"),
    source.replace('relativeFrom="page"><wp14:pctWidth>', 'relativeFrom="column"><wp14:pctWidth>'),
    source.replaceAll("wp:anchor", "wp:inline"),
  ]) {
    const parts = { ...original, "word/document.xml": strToU8(raw) };
    expect(() => importDocxWithReport(zipSync(parts), { documentId: "malformed-size" })).toThrow();
  }
  const parts = {
    ...original,
    "word/document.xml": strToU8(
      source.replace("<wp14:pctWidth>", '<wp14:pctWidth future="keep">'),
    ),
  };
  const imported = importDocxWithReport(zipSync(parts), { documentId: "unknown-size" });
  expect(imported.report.preserved.length).toBeGreaterThan(0);
  const xml = strFromU8(
    unzipSync(exportDocx(openDocument(imported.files, { kind: "document" })).bytes)[
      "word/document.xml"
    ]!,
  );
  expect(xml).toContain('future="keep"');
  const session = createDocumentSession(
      openDocument(
        importDocxWithReport(drawingSizeFixture(), {
          documentId: "size-undo",
          unsupportedContent: "reject",
        }).files,
        { kind: "document" },
      ),
    ),
    before = session.snapshot(),
    d = before.metadata.word!.drawings![0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "size-undo",
      baseRevision: 0,
      operations: [
        {
          op: "setWordDrawingLayout",
          nodeId: d.nodeId,
          layout: {
            ...d.layout,
            wrapPolygon: d.layout.wrapPolygon?.map((point) => ({ ...point })),
            relativeWidth: { relativeFrom: "margin", percent: 75 },
          },
        },
      ],
    }),
  );
  expect(strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!)).toContain("75000");
  session.undo();
  expect(session.snapshot().metadata).toEqual(before.metadata);
});
