import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { strFromU8, unzipSync } from "fflate";
import { body, mutateDocx, styles, wordNamespace } from "./regression-fixtures.js";

const names = [
  "kinsoku",
  "wordWrap",
  "autoSpaceDE",
  "autoSpaceDN",
  "overflowPunct",
  "topLinePunct",
] as const;
const keys = [
  "--word-kinsoku",
  "--word-word-wrap",
  "--word-auto-space-de",
  "--word-auto-space-dn",
  "--word-overflow-punctuation",
  "--word-top-line-punctuation",
] as const;
function fixture(properties: string, shared = "", settings?: string) {
  return mutateDocx("01-text-numbering.docx", {
    ...(settings === undefined
      ? {}
      : {
          "word/settings.xml": () =>
            '<w:settings xmlns:w="' + wordNamespace + '">' + settings + "</w:settings>",
        }),
    "word/styles.xml": () =>
      styles(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:pPr>' +
          shared +
          "</w:pPr></w:style>",
      ),
    "word/document.xml": (source) =>
      body(
        source,
        '<w:p><w:pPr><w:pStyle w:val="Normal"/>' +
          properties +
          "</w:pPr><w:r><w:t>中文Latin数字123</w:t></w:r></w:p>",
      ),
  });
}
function current(bytes: Uint8Array, unsupportedContent: "reject" | "preserve" = "reject") {
  const session = createDocumentSession(
      openDocument(importDocx(bytes, { documentId: "typography", unsupportedContent }), {
        kind: "document",
      }),
    ),
    ids = [session.snapshot().metadata.word!.stories[0]!.rootNodeId];
  for (let index = 0; index < ids.length; index++)
    ids.push(...session.getNode(ids[index]!)!.childNodeIds);
  return {
    session,
    paragraph: ids.find((id) => session.getNode(id)!.tagName === "p")!,
    root: session.getNode(ids[0]!)!.parentNodeId!,
  };
}
test("East Asian paragraph switches remain editable and retain earliest tracked values through two native reopens", () => {
  const { session, paragraph, root } = current(
      fixture(names.map((name) => "<w:" + name + ' w:val="0"/>').join("")),
    ),
    before = session.snapshot();
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "typography",
      baseRevision: 0,
      operations: [
        { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
        {
          op: "setWordParagraph",
          nodeId: paragraph,
          properties: Object.fromEntries(keys.map((key) => [key, "1"])),
        },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    for (const name of names) {
      expect(xml).toMatch(new RegExp("<w:" + name + ' w:val="1"'));
      expect(xml).toMatch(new RegExp("<w:" + name + ' w:val="0"'));
    }
    const next = current(bytes),
      revision = next.session.snapshot().metadata.word!.revisions[0]!;
    expect(revision.property).toBe("paragraph");
    next.session.commit(
      next.session.prepare({
        protocolVersion: 1,
        documentId: "typography",
        baseRevision: 0,
        operations: [
          {
            op: "reviewWordRevision",
            nodeId: revision.nodeId,
            revisionId: revision.revisionId!,
            decision: "reject",
          },
        ],
      }),
    );
    for (const key of keys)
      expect(next.session.getNode(next.paragraph)!.attributes.style).toContain(key + ":0");
    current(exportDocx(next.session).bytes);
    bytes = exportDocx(current(bytes).session).bytes;
  }
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
});
test("shared paragraph switches retain associations while direct overrides remain independent", () => {
  const { session, paragraph } = current(
    fixture('<w:autoSpaceDN w:val="0"/>', names.map((name) => "<w:" + name + "/>").join("")),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "typography",
      baseRevision: 0,
      operations: [
        { op: "setWordParagraph", nodeId: paragraph, properties: { "margin-left": "12pt" } },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const files = unzipSync(bytes),
      shared = strFromU8(files["word/styles.xml"]!),
      xml = strFromU8(files["word/document.xml"]!);
    for (const name of names) expect(shared).toContain("<w:" + name + ' w:val="1"');
    expect(xml).toContain('<w:autoSpaceDN w:val="0"');
    bytes = exportDocx(current(bytes).session).bytes;
  }
});
for (const name of names)
  test("invalid typography Boolean remains malformed: " + name, () => {
    expect(() =>
      importDocx(fixture("<w:" + name + ' w:val="maybe"/>'), { documentId: "invalid" }),
    ).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  });

test("document typography settings are editable, undoable and ordered across strict reopens", () => {
  const { session, root } = current(
    fixture(
      "",
      "",
      '<w:characterSpacingControl w:val="compressPunctuation"/><w:strictFirstAndLastChars/>' +
        '<w:noLineBreaksAfter w:lang="ja-JP" w:val="（"/><w:noLineBreaksBefore w:lang="zh-CN" w:val="丁&amp;𠀀&#9;"/>' +
        "<w:compat><w:doNotWrapTextWithPunct/><w:useAltKinsokuLineBreakRules/></w:compat>",
    ),
    "preserve",
  );
  expect(session.snapshot().metadata.word!.typography).toEqual({
    characterSpacingControl: "compressPunctuation",
    strictFirstAndLastChars: true,
    noLineBreaksAfter: { language: "ja-JP", characters: "（" },
    noLineBreaksBefore: { language: "zh-CN", characters: "丁&𠀀\t" },
  });
  expect(session.snapshot().metadata.word!.doNotWrapTextWithPunct).toBe(true);
  const typography = {
    characterSpacingControl: "compressPunctuationAndJapaneseKana" as const,
    strictFirstAndLastChars: false,
    noLineBreaksBefore: { language: "zh-Hant", characters: "" },
  };
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "typography",
      baseRevision: 0,
      operations: [
        { op: "setWordOptions", nodeId: root, typography, doNotWrapTextWithPunct: false },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const xml = strFromU8(unzipSync(bytes)["word/settings.xml"]!),
      next = current(bytes, "preserve").session;
    current(bytes);
    expect(next.snapshot().metadata.word!.typography).toEqual(typography);
    expect(next.snapshot().metadata.word!.doNotWrapTextWithPunct).toBe(false);
    expect(xml).not.toContain("noLineBreaksAfter");
    expect(xml).toContain('<w:doNotWrapTextWithPunct w:val="0"');
    expect(xml).toContain("<w:useAltKinsokuLineBreakRules");
    expect(xml.indexOf("<w:characterSpacingControl")).toBeLessThan(
      xml.indexOf("<w:strictFirstAndLastChars"),
    );
    expect(xml.indexOf("<w:strictFirstAndLastChars")).toBeLessThan(
      xml.indexOf("<w:noLineBreaksBefore"),
    );
    bytes = exportDocx(next).bytes;
  }
  session.undo();
  expect(session.snapshot().metadata.word!.typography?.noLineBreaksAfter?.language).toBe("ja-JP");
  session.redo();
  expect(session.snapshot().metadata.word!.typography).toEqual(typography);
});

test.each([
  "<w:characterSpacingControl/>",
  '<w:characterSpacingControl w:val="unknown"/>',
  '<w:strictFirstAndLastChars w:val="unknown"/>',
  '<w:noLineBreaksBefore w:lang="zh-CN"/>',
  '<w:noLineBreaksAfter w:val="（"/>',
  '<w:noLineBreaksAfter w:lang="zh-CN" w:val="（"/><w:noLineBreaksAfter w:lang="ja-JP" w:val="（"/>',
  '<w:compat><w:doNotWrapTextWithPunct w:val="unknown"/></w:compat>',
])("rejects malformed East Asian settings (%s)", (settings) => {
  expect(() => importDocx(fixture("", "", settings), { documentId: "invalid" })).toThrowError(
    expect.objectContaining({ code: "InvalidDocument" }),
  );
});
