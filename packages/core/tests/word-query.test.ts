import { expect, test } from "vitest";
import {
  createWordDocument,
  createDocumentSession,
  openDocument,
  findWordText,
  getWordOutline,
  getWordTextStatistics,
  wordTextReplacementOperations,
} from "@doffi/core";
function document() {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "query" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "query",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html: "<h1>标题😀</h1><p>Office office办公 OfficeX</p><p>office 😀 office</p>",
        },
      ],
    }),
  );
  return session;
}
test("Word search preserves UTF-16 offsets and supports case and word boundaries", () => {
  const model = document();
  expect(findWordText(model, "office")).toHaveLength(5);
  expect(findWordText(model, "office", { matchCase: true })).toHaveLength(3);
  expect(findWordText(model, "Office", { wholeWord: true })).toHaveLength(3);
  const emoji = findWordText(model, "😀");
  expect(emoji).toHaveLength(2);
  expect(emoji[0]!.to - emoji[0]!.from).toBe(2);
  expect(findWordText(model, "")).toEqual([]);
});
test("whole replacement shares one history step and title navigation is derived from source", () => {
  const session = document(),
    before = session.snapshot(),
    matches = findWordText(session, "office");
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "query",
      baseRevision: 1,
      operations: [...matches]
        .reverse()
        .map((match) => ({ op: "replaceText" as const, ...match, text: "Word" })),
    }),
  );
  expect(findWordText(session, "office")).toEqual([]);
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  expect(getWordOutline(session).map((h) => [h.level, h.text])).toEqual([[1, "标题😀"]]);
  expect(getWordTextStatistics(session).paragraphs).toBe(3);
});
test("literal regex punctuation is searched without interpretation or paragraph crossing", () => {
  const session = document();
  expect(findWordText(session, "Office.*")).toEqual([]);
  expect(findWordText(session, "OfficeXoffice")).toEqual([]);
});
test("final-view matches across hidden deletions replace only visible source pieces and keep review, undo and stale checks exact", () => {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "visible-query" }), { kind: "document" }),
    ),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "visible-query",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "abXcd" },
    {
      op: "trackWordText",
      nodeId: "paragraph-1",
      from: 2,
      to: 3,
      expectedText: "X",
      text: "",
      author: "old",
    },
  ]);
  const matches = findWordText(session, "abcd", { revisionView: "final" });
  expect(matches).toHaveLength(1);
  expect(matches[0]!.ranges?.map((range) => [range.from, range.to, range.expectedText])).toEqual([
    [0, 2, "ab"],
    [3, 5, "cd"],
  ]);
  expect(findWordText(session, "abcd", { revisionView: "original" })).toEqual([]);
  expect(getWordTextStatistics(session, { revisionView: "final" }).characters).toBe(4);
  expect(getWordTextStatistics(session, { revisionView: "original" }).characters).toBe(5);
  expect(() =>
    wordTextReplacementOperations(
      session,
      findWordText(session, "X", { revisionView: "original" }),
      "Q",
    ),
  ).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  apply([{ op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "new" }]);
  const before = session.snapshot(),
    old = before.metadata.word!.revisions[0]!;
  apply([...wordTextReplacementOperations(session, matches, "Z")]);
  expect(session.getNode(old.nodeId)!.text).toBe("X");
  expect(findWordText(session, "Z", { revisionView: "final" })).toHaveLength(1);
  expect(findWordText(session, "abXcd", { revisionView: "original" })).toHaveLength(1);
  expect(() => wordTextReplacementOperations(session, matches, "Q")).toThrowError(
    expect.objectContaining({ code: "RevisionConflict" }),
  );
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  expect(() => wordTextReplacementOperations(session, [...matches, ...matches], "Q")).toThrowError(
    expect.objectContaining({ code: "InvalidOperation" }),
  );
});
test("original queries omit inserted text and retain UTF-16 variation boundaries and original outline formatting", () => {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "original-query" }), { kind: "document" }),
    ),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "original-query",
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
      text: "前𠀀\u{e0100}后",
    },
    {
      op: "trackWordText",
      nodeId: "paragraph-1",
      from: 5,
      to: 5,
      expectedText: "",
      text: "新",
      author: "new",
    },
    { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "new" },
    { op: "setWordParagraph", nodeId: "paragraph-1", outlineLevel: 2 },
  ]);
  expect(findWordText(session, "前𠀀\u{e0100}后", { revisionView: "original" })).toHaveLength(1);
  expect(findWordText(session, "𠀀", { revisionView: "final" })).toEqual([]);
  expect(findWordText(session, "𠀀\u{e0100}", { revisionView: "final" })[0]!.to).toBe(5);
  expect(getWordOutline(session, { revisionView: "original" })).toEqual([]);
  expect(getWordOutline(session, { revisionView: "final" })[0]!.level).toBe(3);
});
test("hidden paragraph marks join query text and statistics using the surviving paragraph format", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "marks-query" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "marks-query",
      baseRevision: 0,
      operations: [{ op: "replaceFragment", nodeId: "section-root-1", html: "<h1>A</h1><p>B</p>" }],
    }),
  );
  const snapshot = session.snapshot(),
    files = new Map(session.serialize()),
    metadata = JSON.parse(JSON.stringify(snapshot.metadata)),
    first = session.getNode("section-root-1")!.childNodeIds[0]!;
  metadata.word.revisions.push({
    nodeId: first,
    sourceId: "paragraph-mark",
    kind: "paragraphDelete",
  });
  files.set(snapshot.manifest.metadata, new TextEncoder().encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: "document" }),
    match = findWordText(model, "AB", { revisionView: "final" })[0]!;
  expect(match.ranges).toHaveLength(2);
  expect(getWordTextStatistics(model, { revisionView: "final" }).paragraphs).toBe(1);
  expect(getWordTextStatistics(model, { revisionView: "original" }).paragraphs).toBe(2);
  expect(getWordOutline(model, { revisionView: "final" })).toEqual([]);
  expect(getWordOutline(model, { revisionView: "original" })[0]!.text).toBe("A");
});
test("deleted table rows are absent from final queries and counts while original queries retain their text", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "row-query" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "row-query",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html: "<table><tr><td><p>ROW-A</p></td></tr><tr><td><p>ROW-B</p></td></tr></table>",
        },
      ],
    }),
  );
  const table = session.getNode(session.getNode("section-root-1")!.childNodeIds[0]!)!,
    row = session.getNode(table.childNodeIds[0]!)!.childNodeIds[0]!,
    snapshot = session.snapshot(),
    files = new Map(session.serialize()),
    metadata = JSON.parse(JSON.stringify(snapshot.metadata));
  metadata.word.revisions.push({ nodeId: row, sourceId: "row-delete", kind: "rowDelete" });
  files.set(snapshot.manifest.metadata, new TextEncoder().encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: "document" });
  expect(findWordText(model, "ROW-A", { revisionView: "final" })).toEqual([]);
  expect(findWordText(model, "ROW-A", { revisionView: "original" })).toHaveLength(1);
  expect(getWordTextStatistics(model, { revisionView: "final" }).paragraphs).toBe(1);
  expect(getWordTextStatistics(model, { revisionView: "original" }).paragraphs).toBe(2);
});
