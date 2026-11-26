import { expect, test } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocumentSession, openDocument } from "@doffi/core";
import { exportDocx, importDocx, importDocxWithReport } from "@doffi/office";
import { nativeDrawingFrameFixture } from "./native-drawing-frame-fixture.js";

const variants: [string, (xml: string) => string, string][] = [
  ["negative effect extent", (s) => s.replace(/(<wp:effectExtent[^>]* l=")0/, "$1-12700"), 'l="-12700"'],
  [
    "frame unknown attribute",
    (s) => s.replace("<wp:anchor ", '<wp:anchor dataNative="frame-value" '),
    'dataNative="frame-value"',
  ],
  [
    "extent unknown attribute",
    (s) => s.replace("<wp:extent ", '<wp:extent dataNative="extent-value" '),
    'dataNative="extent-value"',
  ],
  [
    "frame locks",
    (s) =>
      s.replace(
        "<wp:cNvGraphicFramePr/>",
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noMove="1"/></wp:cNvGraphicFramePr>',
      ),
    'noMove="1"',
  ],
  [
    "nested effect attribute",
    (s) =>
      s.replace(
        "</wps:spPr>",
        '<a:effectLst><a:outerShdw blurRad="0" dist="0" dir="0" nativeEffect="value"><a:srgbClr val="112233"/></a:outerShdw></a:effectLst></wps:spPr>',
      ),
    'nativeEffect="value"',
  ],
  [
    "nested color attribute",
    (s) =>
      s.replace(
        "</wps:spPr>",
        '<a:solidFill><a:srgbClr val="112233" nativeColor="value"/></a:solidFill></wps:spPr>',
      ),
    'nativeColor="value"',
  ],
  [
    "nested transform attribute",
    (s) => s.replace("<a:off ", '<a:off nativeOffset="value" '),
    'nativeOffset="value"',
  ],
  [
    "nonvisual child",
    (s) =>
      s.replace(/<wps:cNvSpPr[^>]*\/>|<wps:cNvSpPr[^>]*>[\s\S]*?<\/wps:cNvSpPr>/, '<wps:cNvSpPr><a:spLocks noSelect="1"/></wps:cNvSpPr>'),
    'noSelect="1"',
  ],
  [
    "text body attribute",
    (s) => s.replace("<wps:bodyPr ", '<wps:bodyPr anchorCtr="1" '),
    'anchorCtr="1"',
  ],
  ["simple position", (s) => s.replace('simplePos="0"', 'simplePos="1"'), 'simplePos="1"'],
  [
    "graphic wrapper attribute",
    (s) => s.replace("<a:graphic>", '<a:graphic nativeGraphic="value">'),
    'nativeGraphic="value"',
  ],
];
for (const [name, mutate, expected] of variants)
  test(`unmodeled ${name} is diagnosed, strictly rejected and preserved through editing and reopen`, () => {
    const files = unzipSync(nativeDrawingFrameFixture("none")),
      original = strFromU8(files["word/document.xml"]!),
      changed = mutate(original);
    expect(changed).not.toBe(original);
    files["word/document.xml"] = strToU8(changed);
    const bytes = zipSync(files);
    expect(() =>
      importDocx(bytes, { documentId: "strict-drawing", unsupportedContent: "reject" }),
    ).toThrow();
    const imported = importDocxWithReport(bytes, { documentId: "preserved-drawing" }),
      session = createDocumentSession(openDocument(imported.files, { kind: "document" })),
      atom = session
        .snapshot()
        .metadata.word!.nativeContent!.fragments.find((f) => f.kind !== "property")!;
    expect(imported.report.preserved.length).toBeGreaterThan(0);
    expect(session.snapshot().metadata.word!.drawings ?? []).toHaveLength(0);
    const nodes = [session.snapshot().metadata.word!.stories[0]!.rootNodeId];
    for (let i = 0; i < nodes.length; i++) nodes.push(...session.getNode(nodes[i]!)!.childNodeIds);
    const after = nodes.find(
      (id) =>
        session.getNode(id)!.text === "AFTER" && session.getNode(id)!.childNodeIds.length === 0,
    )!;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "preserved-drawing",
        baseRevision: 0,
        operations: [
          { op: "replaceText", nodeId: after, from: 0, to: 0, expectedText: "", text: "EDIT-" },
        ],
      }),
    );
    let output = exportDocx(session).bytes;
    for (let pass = 0; pass < 2; pass++) {
      const xml = strFromU8(unzipSync(output)["word/document.xml"]!);
      expect(xml).toContain(expected);
      expect(xml).toContain("EDIT-AFTER");
      const reopened = openDocument(importDocx(output, { documentId: "coverage-reopen" }), {
        kind: "document",
      });
      expect(
        reopened
          .snapshot()
          .metadata.word!.nativeContent!.fragments.filter((f) => f.kind !== "property"),
      ).toHaveLength(1);
      output = exportDocx(reopened).bytes;
    }
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "preserved-drawing",
        baseRevision: 1,
        operations: [{ op: "removeNode", nodeId: atom.nodeId! }],
      }),
    );
    expect(strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!)).not.toContain(
      expected,
    );
    session.undo();
    expect(strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!)).toContain(
      expected,
    );
  });

test("malformed known drawing dimensions still fail compatibility import", () => {
  const files = unzipSync(nativeDrawingFrameFixture("none"));
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace(/(<wp:effectExtent[^>]* l=")0/, "$1invalid"),
  );
  expect(() => importDocx(zipSync(files), { documentId: "invalid-effect-extent" })).toThrow(
    /numeric/,
  );
});
