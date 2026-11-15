import { expect, test } from "vitest";
import {
  createDocumentSession,
  createWordDocument,
  openDocument,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
  WordInlineContentSemanticsSchema,
} from "@doffi/core";
import type { DocumentOperation, WordContentSemantics } from "@doffi/core";

const semantics: WordContentSemantics = {
  numbering: [
    {
      numberingId: "1",
      levels: [
        {
          level: 0,
          start: 5,
          format: "upperRoman",
          text: "%1)",
          suffix: "tab",
          cssClass: "marker-main",
        },
        {
          level: 1,
          start: 2,
          restart: 1,
          format: "lowerLetter",
          text: "%1.%2)",
          suffix: "space",
          cssClass: "marker-child",
        },
      ],
    },
    {
      numberingId: "2",
      levels: [
        {
          level: 0,
          start: 10,
          format: "decimal",
          text: "%1.",
          suffix: "nothing",
          cssClass: "marker-main",
        },
      ],
    },
  ],
  extraClasses: {
    "marker-main": { "font-family": '"Liberation Serif"', color: "#880000" },
    "marker-child": { "font-family": '"Liberation Mono"', "font-size": "14pt" },
  },
};
const html =
  '<div><p data-node-id="first" data-word-numbering="1" data-word-level="0">FIRST</p><p data-node-id="child" data-word-numbering="1" data-word-level="1">CHILD</p><p data-node-id="restart" data-word-numbering="2" data-word-level="0">RESTART</p></div>';
function target(tracked = false) {
  const documentId = "list-content";
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId }), { kind: "document" }),
  );
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId,
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
    {
      op: "setWordNumbering",
      nodeId: "document-root",
      numberingId: "1",
      levels: [
        {
          level: 0,
          start: 91,
          format: "decimal",
          text: "%1.",
          suffix: "tab",
          cssClass: "receiver-marker",
        },
      ],
    },
    { op: "setWordReview", nodeId: "document-root", trackChanges: tracked, author: "Receiver" },
  ]);
  const paragraphs = () =>
    [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
      .map((match) => session.getNode(match[1]!)!)
      .filter((node) => node.tagName === "p");
  return { session, apply, paragraphs };
}
for (const tracked of [false, true])
  test(
    "native numbering references and marker CSS are isolated with one paste undo: " + tracked,
    () => {
      const { session, apply, paragraphs } = target(tracked);
      const before = session.snapshot();
      apply([
        {
          op: "pasteWordBlockContent",
          nodeId: "paragraph-1",
          from: 1,
          to: 1,
          expectedText: "",
          html,
          semantics,
        },
      ]);
      const definitions = session.snapshot().metadata.word!.numbering;
      expect(definitions[0]).toEqual(before.metadata.word!.numbering[0]);
      expect(definitions).toHaveLength(3);
      expect(definitions[1]!.numberingId).not.toBe("1");
      expect(definitions[2]!.numberingId).not.toBe("2");
      expect(
        definitions[1]!.levels.map((level) => [level.level, level.start, level.restart]),
      ).toEqual([
        [0, 5, undefined],
        [1, 2, 1],
      ]);
      const first = paragraphs().find((node) => node.text === "FIRST")!;
      const child = paragraphs().find((node) => node.text === "CHILD")!;
      expect(first.attributes["data-word-numbering"]).toBe(definitions[1]!.numberingId);
      expect(child.attributes["data-word-numbering"]).toBe(definitions[1]!.numberingId);
      const css = new SourceStyleResolver(
        session.snapshot().styles.map((sheet) => sheet.css),
        DEFAULT_DOCUMENT_READ_LIMITS,
      );
      expect(css.specifiedClass(definitions[1]!.levels[1]!.cssClass)).toMatchObject({
        "font-family": '"Liberation Mono"',
        "font-size": "14pt",
      });
      expect(() => openDocument(session.serialize(), { kind: "document" })).not.toThrow();
      apply([
        {
          op: "setWordList",
          nodeId: first.nodeId,
          numberingId: definitions[1]!.numberingId,
          level: 1,
        },
      ]);
      session.undo();
      if (tracked) {
        apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" }]);
        expect(session.getNode("main-story")!.text).toBe("AB");
        expect(paragraphs()).toHaveLength(1);
        session.undo();
      }
      session.undo();
      expect(session.snapshot().html).toBe(before.html);
      expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
      session.dispose();
    },
  );
for (const decision of ["accept", "reject"] as const)
  test(
    "nested HTML list blocks normalize into editable Word paragraphs and grouped review: " +
      decision,
    () => {
      const { session, apply, paragraphs } = target(true);
      apply([
        {
          op: "pasteWordBlockContent",
          nodeId: "paragraph-1",
          from: 1,
          to: 1,
          expectedText: "",
          html: '<ol start="4" type="A"><li>ONE<ul><li>SUB</li></ul><p>TAIL</p></li><li value="9">TWO</li><li>THREE</li></ol>',
        },
      ]);
      expect(session.snapshot().html).not.toMatch(/<(?:ol|ul|li)\b/);
      const lookup = (text: string) => {
        const paragraph = paragraphs().find((node) => node.text === text)!;
        return session
          .snapshot()
          .metadata.word!.numbering.find(
            (item) => item.numberingId === paragraph.attributes["data-word-numbering"],
          )!;
      };
      expect(lookup("ONE").levels[0]).toMatchObject({ start: 4, format: "upperLetter", level: 0 });
      expect(lookup("TWO").levels[0]).toMatchObject({ start: 9, format: "upperLetter" });
      expect(lookup("THREE").numberingId).toBe(lookup("TWO").numberingId);
      expect(lookup("SUB").levels[0]).toMatchObject({ format: "bullet", level: 1, text: "•" });
      const inserted = session.snapshot();
      apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision }]);
      expect(session.getNode("main-story")!.text).toBe(
        decision === "accept" ? "AONESUBTAILTWOTHREEB" : "AB",
      );
      expect(paragraphs()).toHaveLength(decision === "accept" ? 7 : 1);
      session.undo();
      expect(session.snapshot().html).toBe(inserted.html);
      session.dispose();
    },
  );
test("missing, conflicting or duplicate list graphs reject atomically and inline content cannot carry list definitions", () => {
  const { session, apply } = target();
  const before = session.snapshot();
  for (const graph of [
    undefined,
    { ...semantics, numbering: [] },
    { ...semantics, extraClasses: {} },
    { ...semantics, numbering: [semantics.numbering![0]!, semantics.numbering![0]!] },
  ]) {
    expect(() =>
      apply([
        {
          op: "pasteWordBlockContent",
          nodeId: "paragraph-1",
          from: 1,
          to: 1,
          expectedText: "",
          html,
          semantics: graph,
        },
      ]),
    ).toThrow();
    expect(session.snapshot()).toEqual(before);
  }
  for (const source of [
    "<ol reversed><li>X</li></ol>",
    '<ol start="-1"><li>X</li></ol>',
    '<p data-word-numbering="1" data-word-level="0">X</p>',
  ]) {
    expect(() =>
      apply([
        {
          op: "pasteWordBlockContent",
          nodeId: "paragraph-1",
          from: 1,
          to: 1,
          expectedText: "",
          html: source,
        },
      ]),
    ).toThrow();
    expect(session.snapshot()).toEqual(before);
  }
  expect(() => WordInlineContentSemanticsSchema.parse(semantics)).toThrow();
  expect(() =>
    apply([
      {
        op: "pasteWordInlineContent",
        nodeId: "paragraph-1",
        from: 1,
        to: 1,
        expectedText: "",
        html: '<span data-word-numbering="1" data-word-level="0">X</span>',
      },
    ]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
  const candidate = session.prepare({
    protocolVersion: 1,
    documentId: "list-content",
    baseRevision: session.snapshot().manifest.revision,
    operations: [
      {
        op: "pasteWordBlockContent",
        nodeId: "paragraph-1",
        from: 1,
        to: 1,
        expectedText: "",
        html,
        semantics,
      },
    ],
  });
  expect(session.snapshot()).toEqual(before);
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "CURRENT" },
  ]);
  const current = session.snapshot();
  expect(() => session.commit(candidate)).toThrowError(
    expect.objectContaining({ code: "RevisionConflict" }),
  );
  expect(session.snapshot()).toEqual(current);
  session.dispose();
});
