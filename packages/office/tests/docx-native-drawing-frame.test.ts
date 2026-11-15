import { expect, test } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { openDocument, createDocumentSession } from "@doffi/core";
import { exportDocx, importDocx, importDocxWithReport } from "@doffi/office";
import { nativeDrawingFrameFixture } from "./native-drawing-frame-fixture.js";
for (const kind of ["size", "position"] as const)
  test(
    "unmodeled drawing " +
      kind +
      " extensions retain their complete frame while surrounding source stays editable",
    () => {
      const original = nativeDrawingFrameFixture(kind);
      expect(() =>
        importDocx(original, { documentId: "strict-native-frame", unsupportedContent: "reject" }),
      ).toThrow(/frame extension|position extension/);
      const imported = importDocxWithReport(original, { documentId: "native-frame" }),
        session = createDocumentSession(openDocument(imported.files, { kind: "document" })),
        snapshot = session.snapshot(),
        fragment = snapshot.metadata.word!.nativeContent!.fragments.find(
          (item) => item.kind !== "property",
        )!,
        nodes = [snapshot.metadata.word!.stories[0]!.rootNodeId];
      for (let i = 0; i < nodes.length; i++)
        nodes.push(...session.getNode(nodes[i]!)!.childNodeIds);
      const paragraph = nodes.find(
        (id) =>
          session.getNode(id)!.text === "AFTER" && session.getNode(id)!.childNodeIds.length === 0,
      )!;
      expect(imported.report.preserved.length).toBeGreaterThan(0);
      expect(session.getNode(fragment.nodeId!)!.text).toBe("\ufffc");
      const before = session.snapshot(),
        length = session.getNode(paragraph)!.text.length;
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "native-frame",
          baseRevision: 0,
          operations: [
            {
              op: "replaceText",
              nodeId: paragraph,
              from: length,
              to: length,
              expectedText: "",
              text: "EDIT",
            },
          ],
        }),
      );
      let bytes = exportDocx(session).bytes;
      for (let pass = 0; pass < 3; pass++) {
        const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
        expect(xml).toContain(kind === "size" ? "sizeRelH" : "pctPosHOffset");
        expect(xml).toContain("50000");
        expect(xml).toContain("EDIT");
        const model = openDocument(importDocx(bytes, { documentId: "native-frame-reopened" }), {
          kind: "document",
        });
        expect(
          model
            .snapshot()
            .metadata.word!.nativeContent!.fragments.filter((item) => item.kind !== "property"),
        ).toHaveLength(1);
        bytes = exportDocx(model).bytes;
      }
      session.undo();
      expect(session.snapshot().html).toBe(before.html);
      session.redo();
      expect(session.getNode(paragraph)!.text).toContain("EDIT");
    },
  );
test("more than one thousand identical drawing extensions recover in one bounded native collection", () => {
  const imported = importDocxWithReport(nativeDrawingFrameFixture("size", 1100), {
      documentId: "native-frame-batch",
    }),
    model = openDocument(imported.files, { kind: "document" }),
    fragments = model
      .snapshot()
      .metadata.word!.nativeContent!.fragments.filter((item) => item.kind !== "property");
  expect(fragments).toHaveLength(1100);
  expect(new Set(fragments.map((item) => item.assetPath)).size).toBe(1);
  const bytes = exportDocx(model).bytes,
    xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
  expect(xml.match(/sizeRelH/g)).toHaveLength(2200);
  expect(
    openDocument(importDocx(bytes, { documentId: "native-frame-batch-reopened" }), {
      kind: "document",
    })
      .snapshot()
      .metadata.word!.nativeContent!.fragments.filter((item) => item.kind !== "property"),
  ).toHaveLength(1100);
});
