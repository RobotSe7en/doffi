import { expect, test } from "vitest";
import { unzipSync, strFromU8 } from "fflate";
import {
  createWordDocument,
  createDocumentSession,
  openDocument,
  getWordOutline,
  getWordTextDefaults,
} from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
test("partial repeated direct and named character formats compile one earliest history per run and review after two native reopens", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "format-merge" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "format-merge",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "ABCDEF",
        },
        {
          op: "setWordStyle",
          nodeId: "document-root",
          style: { styleId: "Emphasis", name: "强调", kind: "character", cssClass: "emphasis" },
          runProperties: { "font-style": "italic" },
        },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "甲" },
        {
          op: "formatWordText",
          nodeId: "paragraph-1",
          from: 0,
          to: 6,
          expectedText: "ABCDEF",
          properties: { "font-weight": "700" },
        },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "乙" },
        {
          op: "applyWordCharacterStyle",
          nodeId: "paragraph-1",
          from: 1,
          to: 4,
          expectedText: "BCD",
          styleId: "Emphasis",
        },
        {
          op: "formatWordText",
          nodeId: "paragraph-1",
          from: 2,
          to: 5,
          expectedText: "CDE",
          properties: { color: "#cc0000" },
        },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const model = openDocument(
      importDocx(bytes, { documentId: "merged-native", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(
      model.snapshot().metadata.word!.revisions.filter((item) => item.property === "run"),
    ).toHaveLength(5);
    bytes = exportDocx(model).bytes;
  }
  const reopened = createDocumentSession(
      openDocument(
        importDocx(bytes, { documentId: "merged-native", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    ),
    story = reopened.snapshot().metadata.word!.stories[0]!,
    root = reopened.getNode(story.rootNodeId)!.parentNodeId!;
  reopened.commit(
    reopened.prepare({
      protocolVersion: 1,
      documentId: "merged-native",
      baseRevision: 0,
      operations: [{ op: "reviewWordRevisions", nodeId: root, decision: "reject" }],
    }),
  );
  const xml = strFromU8(unzipSync(exportDocx(reopened).bytes)["word/document.xml"]!);
  expect(xml).not.toContain("PrChange");
  expect(xml).not.toContain("<w:b ");
  expect(xml).not.toContain('w:val="Emphasis"');
  expect(xml).not.toContain("CC0000");
  expect(reopened.getNode(story.rootNodeId)!.text).toBe("ABCDEF");
});
test("editable style definitions preserve names, presentation flags, inheritance and nine-level outline through native reopens", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "styles" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "styles",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "自定义大纲",
        },
        {
          op: "setWordStyle",
          nodeId: "document-root",
          style: {
            styleId: "Custom",
            name: "中文样式",
            aliases: "别名",
            kind: "paragraph",
            cssClass: "custom",
            basedOn: "Normal",
            nextStyleId: "Normal",
            default: true,
            customStyle: true,
            outlineLevel: 8,
            qFormat: true,
            semiHidden: false,
            unhideWhenUsed: true,
            locked: false,
            uiPriority: 12,
          },
          runProperties: { "font-size": "17pt", "font-weight": "700" },
          paragraphProperties: { "margin-top": "10pt" },
        },
        { op: "setWordParagraph", nodeId: "paragraph-1", styleId: "Custom" },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const model = openDocument(
        importDocx(bytes, { documentId: "native", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
      definition = model
        .snapshot()
        .metadata.word!.styles.find((style) => style.styleId === "Custom")!;
    expect(definition).toMatchObject({
      name: "中文样式",
      aliases: "别名",
      default: true,
      customStyle: true,
      outlineLevel: 8,
      nextStyleId: "Normal",
      qFormat: true,
      unhideWhenUsed: true,
      uiPriority: 12,
    });
    expect(getWordOutline(model)[0]!.level).toBe(9);
    expect(getWordTextDefaults(model, getWordOutline(model)[0]!.nodeId)["font-size"]).toBe("17pt");
    bytes = exportDocx(model).bytes;
    expect(strFromU8(unzipSync(bytes)["word/styles.xml"]!)).toContain("中文样式");
  }
});
test("new paragraph and named character formatting histories reopen and reject through one review boundary", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "tracked-styles" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "tracked-styles",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "前文字后",
        },
        {
          op: "setWordStyle",
          nodeId: "document-root",
          style: { styleId: "Emphasis", name: "强调", kind: "character", cssClass: "emphasis" },
          runProperties: { "font-style": "italic" },
        },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "甲" },
        { op: "setWordParagraph", nodeId: "paragraph-1", properties: { "text-align": "right" } },
        {
          op: "applyWordCharacterStyle",
          nodeId: "paragraph-1",
          from: 1,
          to: 3,
          expectedText: "文字",
          styleId: "Emphasis",
        },
      ],
    }),
  );
  const bytes = exportDocx(session).bytes,
    xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
  expect(xml).toContain("pPrChange");
  expect(xml).toContain("rPrChange");
  const reopened = createDocumentSession(
    openDocument(
      importDocx(bytes, { documentId: "tracked-reopen", unsupportedContent: "reject" }),
      { kind: "document" },
    ),
  );
  expect(reopened.snapshot().metadata.word!.revisions).toHaveLength(2);
  const root = reopened.getNode(
    reopened.snapshot().metadata.word!.stories[0]!.rootNodeId,
  )!.parentNodeId!;
  reopened.commit(
    reopened.prepare({
      protocolVersion: 1,
      documentId: "tracked-reopen",
      baseRevision: 0,
      operations: [{ op: "reviewWordRevisions", nodeId: root, decision: "reject" }],
    }),
  );
  const output = strFromU8(unzipSync(exportDocx(reopened).bytes)["word/document.xml"]!);
  expect(output).not.toContain("PrChange");
  expect(output).not.toContain('w:val="right"');
  expect(output).not.toContain('w:val="Emphasis"');
  expect(reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text).toBe(
    "前文字后",
  );
});

test("run language roles and complex-script fonts remain editable across native reopens and reject formatting correctly", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "language-roles" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "language-roles",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "Latin 漢 العربية",
        },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "甲" },
        {
          op: "formatWordText",
          nodeId: "paragraph-1",
          from: 0,
          to: 15,
          expectedText: "Latin 漢 العربية",
          properties: {
            "--word-language-val": '"en-GB"',
            "--word-language-eastAsia": '"ja-JP"',
            "--word-language-bidi": '"ar-SA"',
            "--word-font-size-cs": "16pt",
            "--word-bold-cs": "1",
          },
        },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const model = openDocument(
        importDocx(bytes, { documentId: "language-native", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
      xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(xml).toContain('w:eastAsia="ja-JP"');
    expect(xml).toContain('w:bidi="ar-SA"');
    expect(xml).toContain('w:val="en-GB"');
    expect(model.snapshot().html).toContain("--word-font-size-cs:16pt");
    bytes = exportDocx(model).bytes;
  }
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "language-roles",
      baseRevision: 1,
      operations: [{ op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" }],
    }),
  );
  expect(strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!)).not.toContain(
    "ja-JP",
  );
});
