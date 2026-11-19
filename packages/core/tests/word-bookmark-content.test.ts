import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WordInlineContentSemantics } from "@doffi/core";
import { remapWordFieldBookmarks } from "../src/session/word-field-bookmarks.js";
const html =
  '<span data-node-id="b0" data-word-bookmark="start"></span><span data-node-id="text">SOURCE</span><span data-node-id="b1" data-word-bookmark="end"></span><span data-node-id="link">GO</span><span data-node-id="field" data-word-field="simple">SOURCE</span>';
const semantics: WordInlineContentSemantics = {
  bookmarks: [{ bookmarkId: "bookmark", name: "Anchor", startNodeId: "b0", endNodeId: "b1" }],
  hyperlinks: [{ nodeId: "link", bookmark: "Anchor" }],
  fields: [
    {
      nodeId: "field",
      instruction: ' REF "Anchor" \\h ',
      resultState: "cached",
      resultNodeIds: ["field"],
    },
  ],
};
function target(tracked = false) {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "bookmark-content" }), { kind: "document" }),
    ),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "bookmark-content",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    {
      op: "replaceText",
      nodeId: "paragraph-1",
      from: 0,
      to: 0,
      expectedText: "",
      text: "ORIGINAL",
    },
    {
      op: "insertWordBookmark",
      nodeId: "paragraph-1",
      from: 0,
      to: 8,
      expectedText: "ORIGINAL",
      name: "Anchor",
    },
    { op: "setWordReview", nodeId: "document-root", trackChanges: tracked, author: "Receiver" },
  ]);
  return { session, apply };
}
for (const tracked of [false, true])
  test(
    "copied bookmark references remain independent of conflicting target names and restore with undo " +
      tracked,
    () => {
      const { session, apply } = target(tracked),
        before = session.snapshot();
      apply([
        {
          op: "pasteWordInlineContent",
          nodeId: "paragraph-1",
          from: 8,
          to: 8,
          expectedText: "",
          html,
          semantics,
        },
      ]);
      const word = session.snapshot().metadata.word!,
        copy = word.bookmarks.find((item) => item.name !== "Anchor")!;
      expect(copy.name).toMatch(/^Copy_[a-f0-9]{32}$/);
      expect(word.fields[0]!.instruction).toBe(' REF "' + copy.name + '" \\h ');
      expect(word.hyperlinks[0]!.bookmark).toBe(copy.name);
      expect(copy.startNodeId).not.toBe("b0");
      if (tracked) {
        apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" }]);
        expect(session.snapshot().metadata.word!.bookmarks.map((item) => item.name)).toEqual([
          "Anchor",
        ]);
        expect(session.snapshot().metadata.word!.hyperlinks).toHaveLength(0);
        session.undo();
      }
      session.undo();
      expect(session.snapshot().html).toBe(before.html);
      expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
    },
  );
test("references whose bookmark was not selected never bind an unrelated target bookmark", () => {
  const { session, apply } = target();
  apply([
    {
      op: "pasteWordInlineContent",
      nodeId: "paragraph-1",
      from: 8,
      to: 8,
      expectedText: "",
      html: '<span data-node-id="link">GO</span><span data-node-id="field" data-word-field="simple">CACHE</span>',
      semantics: { hyperlinks: semantics.hyperlinks, fields: semantics.fields },
    },
  ]);
  const word = session.snapshot().metadata.word!;
  expect(word.bookmarks.map((item) => item.name)).toEqual(["Anchor"]);
  expect(word.hyperlinks[0]!.bookmark).not.toBe("Anchor");
  expect(word.fields[0]!.instruction).toContain(word.hyperlinks[0]!.bookmark!);
  expect(() => openDocument(session.serialize(), { kind: "document" })).not.toThrow();
});
test("foreign bookmark endpoints reject the entire semantic content transaction", () => {
  const { session, apply } = target(),
    before = session.snapshot();
  expect(() =>
    apply([
      {
        op: "pasteWordInlineContent",
        nodeId: "paragraph-1",
        from: 8,
        to: 8,
        expectedText: "",
        html,
        semantics: {
          ...semantics,
          bookmarks: [{ ...semantics.bookmarks![0]!, endNodeId: "paragraph-1" }],
        },
      },
    ]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});
test("field bookmark operands preserve quoting, nested offsets, external targets and quoted switch-like text", () => {
  const instruction = ' REF "Anchor" \\* MERGEFORMAT ',
    field = {
      nodeId: "f",
      instruction,
      resultState: "cached" as const,
      instructionFields: [{ fieldNodeId: "nested", offset: instruction.length }],
    };
  remapWordFieldBookmarks(field, () => "LongerCopy");
  expect(field.instruction).toBe(' REF "LongerCopy" \\* MERGEFORMAT ');
  expect(field.instructionFields[0]!.offset).toBe(field.instruction.length);
  for (const [instruction, expected] of [
    [' HYPERLINK \\o "tip \\l Other" \\l "Anchor" ', ' HYPERLINK \\o "tip \\l Other" \\l "Copy" '],
    [
      ' HYPERLINK "https://example.com/file.docx" \\l "Anchor" ',
      ' HYPERLINK "https://example.com/file.docx" \\l "Anchor" ',
    ],
    [' TOC \\t "Heading \\b Other,1" \\b Anchor ', ' TOC \\t "Heading \\b Other,1" \\b Copy '],
    [" STYLEREF Anchor ", " STYLEREF Anchor "],
  ]) {
    const field = { nodeId: "f", instruction: instruction!, resultState: "cached" as const };
    remapWordFieldBookmarks(field, () => "Copy");
    expect(field.instruction).toBe(expected);
  }
});
