import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocxWithReport } from "@doffi/office";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { drawingWrapFixture, notchPolygon } from "./drawing-wrap-fixture.js";
for (const wrap of ["square", "tight", "through", "topAndBottom", "none"] as const)
  test("independent native wrap source stays editable through two strict reopens: " + wrap, () => {
    let bytes = drawingWrapFixture({
      wrap,
      side: "largest",
      inCell: true,
      layoutInCell: false,
      allowOverlap: false,
      rotation: 25,
    });
    for (let pass = 0; pass < 2; pass++) {
      const imported = importDocxWithReport(bytes, {
        documentId: "wrap-native",
        unsupportedContent: "reject",
      });
      expect(imported.report.preserved).toEqual([]);
      const model = openDocument(imported.files, { kind: "document" }),
        drawing = model.snapshot().metadata.word!.drawings![0]!;
      expect(drawing.layout).toMatchObject({
        mode: "floating",
        wrap,
        allowOverlap: false,
        layoutInCell: false,
      });
      expect(drawing.geometry.rotation).toBe(25);
      if (wrap === "tight" || wrap === "through") {
        expect(drawing.layout.wrapPolygon).toEqual(notchPolygon);
        expect(drawing.layout.wrapPolygonEdited).toBe(true);
      }
      if (["square", "tight", "through"].includes(wrap)) {
        expect(drawing.layout.wrapSide).toBe("largest");
        expect(drawing.layout.distances).toEqual({ top: 2, right: 4, bottom: 3, left: 3 });
      }
      bytes = exportDocx(model).bytes;
      expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).not.toContain("sdk-word-");
    }
  });
test("wrap editing shares atomic history and rejects invalid or stale batches", () => {
  const imported = importDocxWithReport(drawingWrapFixture({ wrap: "tight" }), {
      documentId: "wrap-edit",
      unsupportedContent: "reject",
    }),
    session = createDocumentSession(openDocument(imported.files, { kind: "document" })),
    before = session.snapshot(),
    drawing = before.metadata.word!.drawings![0]!,
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "wrap-edit",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      ),
    layout = {
      ...drawing.layout,
      wrap: "through" as const,
      wrapSide: "right" as const,
      layoutInCell: false,
      allowOverlap: false,
      wrapPolygonEdited: true,
      wrapPolygon: [
        { x: -100, y: 0 },
        { x: 22000, y: 0 },
        { x: 10800, y: 21600 },
      ],
    },
    stale = session.prepare({
      protocolVersion: 1,
      documentId: "wrap-edit",
      baseRevision: 0,
      operations: [{ op: "setWordDrawingLayout", nodeId: drawing.nodeId, layout }],
    });
  apply([{ op: "setWordDrawingLayout", nodeId: drawing.nodeId, layout }]);
  expect(() => session.commit(stale)).toThrowError(
    expect.objectContaining({ code: "RevisionConflict" }),
  );
  const current = session.snapshot();
  expect(() =>
    apply([
      {
        op: "setWordDrawingGeometry",
        nodeId: drawing.nodeId,
        geometry: { ...drawing.geometry, width: 200 },
      },
      {
        op: "setWordDrawingLayout",
        nodeId: drawing.nodeId,
        layout: { ...layout, wrapPolygon: [{ x: 0, y: 0 }] },
      },
    ]),
  ).toThrow();
  expect(session.snapshot()).toEqual(current);
  session.undo();
  expect(session.snapshot().metadata).toEqual(before.metadata);
  session.redo();
  expect(session.snapshot().metadata.word!.drawings![0]!.layout).toEqual(layout);
  const reopened = openDocument(
    importDocxWithReport(exportDocx(session).bytes, {
      documentId: "wrap-edit",
      unsupportedContent: "reject",
    }).files,
    { kind: "document" },
  );
  expect(reopened.snapshot().metadata.word!.drawings![0]!.layout).toMatchObject(layout);
});
test("wrap-level distances reject negative native values without partial import", () => {
  const parts = unzipSync(drawingWrapFixture({ wrap: "through" }));
  parts["word/document.xml"] = strToU8(
    strFromU8(parts["word/document.xml"]!).replace('distL="38100"', 'distL="-1"'),
  );
  expect(() => importDocxWithReport(zipSync(parts), { documentId: "negative-distance" })).toThrow();
});

for (const wrap of ["tight", "through"] as const)
  for (const malformed of ["missing", "duplicate", "short"] as const)
    test("malformed native wrap boundaries fail atomically: " + wrap + " " + malformed, () => {
      const parts = unzipSync(drawingWrapFixture({ wrap }));
      const source = strFromU8(parts["word/document.xml"]!),
        polygon = source.match(/<wp:wrapPolygon[\s\S]*?<\/wp:wrapPolygon>/)![0];
      parts["word/document.xml"] = strToU8(
        source.replace(
          polygon,
          malformed === "missing"
            ? ""
            : malformed === "duplicate"
              ? polygon + polygon
              : '<wp:wrapPolygon><wp:start x="0" y="0"/><wp:lineTo x="21600" y="21600"/></wp:wrapPolygon>',
        ),
      );
      expect(() =>
        importDocxWithReport(zipSync(parts), { documentId: "invalid-boundary" }),
      ).toThrow();
    });
