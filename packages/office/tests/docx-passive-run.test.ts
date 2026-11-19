import { expect, test } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocumentSession, openDocument } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { body, mutateDocx, styles } from "./regression-fixtures.js";

const resets =
  '<w:noProof/><w:outline w:val="0"/><w:shadow w:val="false"/><w:emboss w:val="0"/><w:imprint w:val="0"/><w:vanish w:val="0"/><w:webHidden w:val="0"/><w:rtl w:val="0"/><w:cs w:val="0"/><w:specVanish w:val="0"/><w:w w:val="100"/><w:position w:val="0"/><w:em w:val="none"/><w:effect w:val="none"/>';
function fixture(activeStyle = false) {
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>' +
          (activeStyle
            ? '<w:style w:type="character" w:styleId="Effects"><w:rPr><w:outline/><w:shadow/><w:emboss/><w:vanish/></w:rPr></w:style>'
            : ""),
      ),
    "word/document.xml": (source) =>
      body(
        source,
        "<w:p><w:r><w:rPr>" +
          (activeStyle ? '<w:rStyle w:val="Effects"/>' : "") +
          resets +
          "</w:rPr><w:t>保留原生属性</w:t></w:r></w:p>",
      ),
  });
}
function textParagraph(session: ReturnType<typeof createDocumentSession>) {
  const ids = [
    session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
  ];
  for (let index = 0; index < ids.length; index++) {
    const node = session.getNode(ids[index]!)!;
    if (node.tagName === "p") return node.nodeId;
    ids.push(...node.childNodeIds);
  }
  throw new Error("Missing paragraph");
}
function expectResets(xml: string): void {
  for (const name of [
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "vanish",
    "webHidden",
    "rtl",
    "cs",
    "specVanish",
  ])
    expect(xml).toMatch(new RegExp("<w:" + name + '\\b[^>]*w:val="0"'));
  expect(xml).toMatch(/<w:noProof\b[^>]*w:val="1"/);
  for (const [name, value] of [
    ["w", "100"],
    ["position", "0"],
    ["em", "none"],
    ["effect", "none"],
  ])
    expect(xml).toMatch(new RegExp("<w:" + name + '\\b[^>]*w:val="' + value + '"'));
}
test("proofing state and explicit passive run resets survive strict reads, text edits, undo and two native reopens", () => {
  const session = createDocumentSession(
      openDocument(importDocx(fixture(), { documentId: "passive", unsupportedContent: "reject" }), {
        kind: "document",
      }),
    ),
    paragraph = textParagraph(session);
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "passive",
      baseRevision: 0,
      operations: [
        { op: "replaceText", nodeId: paragraph, from: 0, to: 0, expectedText: "", text: "编辑" },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    expectResets(strFromU8(unzipSync(bytes)["word/document.xml"]!));
    const model = openDocument(
      importDocx(bytes, { documentId: "passive-reopen", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(model.getNode(model.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text).toContain(
      "编辑保留原生属性",
    );
    bytes = exportDocx(model).bytes;
  }
  session.undo();
  expect(session.getNode(paragraph)!.text).toBe("保留原生属性");
  expectResets(strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!));
});
test("explicit false run effects still cancel preserved active style effects after supported text formatting", () => {
  const session = createDocumentSession(
      openDocument(importDocx(fixture(true), { documentId: "passive-style" }), {
        kind: "document",
      }),
    ),
    paragraph = textParagraph(session);
  expect(
    session
      .snapshot()
      .metadata.word!.nativeContent!.fragments.some((fragment) => fragment.styleId === "Effects"),
  ).toBe(true);
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "passive-style",
      baseRevision: 0,
      operations: [
        {
          op: "formatWordText",
          nodeId: paragraph,
          from: 0,
          to: 6,
          expectedText: "保留原生属性",
          properties: { color: "#224466" },
        },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const native = unzipSync(bytes);
    expectResets(strFromU8(native["word/document.xml"]!));
    expect(strFromU8(native["word/styles.xml"]!)).toContain("outline");
    const model = openDocument(importDocx(bytes, { documentId: "passive-style-reopen" }), {
      kind: "document",
    });
    bytes = exportDocx(model).bytes;
  }
});
test("invalid native Boolean values remain malformed input rather than being treated as passive false values", () => {
  const native = unzipSync(fixture());
  native["word/document.xml"] = strToU8(
    strFromU8(native["word/document.xml"]!).replace("<w:noProof/>", '<w:noProof w:val="maybe"/>'),
  );
  expect(() => importDocx(zipSync(native), { documentId: "invalid-passive" })).toThrowError(
    expect.objectContaining({ code: "InvalidDocument" }),
  );
});
