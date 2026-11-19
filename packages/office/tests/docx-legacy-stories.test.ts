import { expect, test } from "vitest";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentOperation, WordDocumentMetadata } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";

function fixture(withWord: boolean) {
  const files = new Map(createWordDocument({ documentId: "legacy" })),
    decode = new TextDecoder(),
    encode = new TextEncoder();
  const metadata = JSON.parse(decode.decode(files.get("metadata.json"))) as WordDocumentMetadata;
  metadata.sections[0]!.rootNodeId = "document-root";
  if (withWord) metadata.word!.stories[0]!.rootNodeId = "document-root";
  else delete metadata.word;
  files.set("metadata.json", encode.encode(JSON.stringify(metadata)));
  files.set(
    "content.html",
    encode.encode(
      '<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body><article data-node-id="document-root" data-office-kind="document"' +
        (withWord ? ' data-office-story="main"' : "") +
        '><p data-node-id="paragraph-1">前😀后</p><p data-node-id="paragraph-2">第二段</p></article></body></html>',
    ),
  );
  const session = createDocumentSession(openDocument(files, { kind: "document" }));
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "legacy",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  return { session, apply };
}
const geometry = { x: 0, y: 0, width: 80, height: 40, rotation: 0, flipH: false, flipV: false };

for (const withWord of [false, true])
  for (const operation of [
    { op: "insertWordNote", nodeId: "paragraph-1", offset: 3, kind: "footnote", text: "注释" },
    {
      op: "setWordNoteSeparator",
      nodeId: "document-root",
      kind: "endnote",
      type: "continuationNotice",
      text: "下页继续",
    },
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 3,
      geometry,
      layout: { mode: "inline" },
      object: { kind: "textBox", text: "框内文字" },
    },
    {
      op: "insertWordStory",
      nodeId: "document-root",
      kind: "header",
      variant: "default",
      html: "<p>页眉</p>",
    },
  ] satisfies DocumentOperation[])
    test(`legacy ${withWord ? "Word" : "HTML"} main roots create ${operation.op} without nested stories and keep same-batch section targets`, () => {
      const { session, apply } = fixture(withWord),
        before = session.snapshot();
      apply([
        operation,
        { op: "setWordSection", nodeId: "document-root", changes: { pageNumberStart: 7 } },
      ]);
      const current = session.snapshot(),
        main = current.metadata.word!.stories.find((story) => story.kind === "main")!;
      expect(main.rootNodeId).not.toBe("document-root");
      expect(session.getNode("paragraph-1")!.parentNodeId).toBe(main.rootNodeId);
      expect(session.getNode("paragraph-2")!.text).toBe("第二段");
      expect(current.metadata.sections![0]!).toMatchObject({
        sectionId: "section-1",
        rootNodeId: main.rootNodeId,
        pageNumberStart: 7,
      });
      expect(openDocument(session.serialize(), { kind: "document" }).snapshot()).toEqual(current);
      let bytes = exportDocx(session).bytes;
      for (let pass = 0; pass < 2; pass++) {
        const reopened = openDocument(
          importDocx(bytes, { documentId: "reopen", unsupportedContent: "reject" }),
          { kind: "document" },
        );
        expect(reopened.snapshot().metadata.sections![0]!.pageNumberStart).toBe(7);
        expect(reopened.snapshot().metadata.word!.stories).toHaveLength(
          operation.op === "insertWordNote" || operation.op === "setWordNoteSeparator" ? 4 : 2,
        );
        expect(
          reopened.getNode(
            reopened.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!
              .rootNodeId,
          )!.text,
        ).toBe(session.getNode(main.rootNodeId)!.text);
        bytes = exportDocx(reopened).bytes;
      }
      session.undo();
      expect(session.snapshot().html).toBe(before.html);
      expect(session.snapshot().metadata).toEqual(before.metadata);
      session.redo();
      expect(session.snapshot().html).toBe(current.html);
      expect(session.snapshot().metadata).toEqual(current.metadata);
    });

test("legacy section roots normalize once across notes, comments, text boxes and section splits in one transaction", () => {
  const { session, apply } = fixture(true);
  apply([
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "a",
        documentId: "legacy",
        baseRevision: 0,
        status: "open",
        anchor: { kind: "node", nodeId: "paragraph-2" },
        instruction: "意见",
      },
    },
    { op: "addWordCommentReply", nodeId: "paragraph-2", annotationId: "a", html: "<p>回复</p>" },
    { op: "insertWordNote", nodeId: "paragraph-1", offset: 3, kind: "endnote", text: "尾注" },
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 3,
      geometry,
      layout: { mode: "inline" },
      object: { kind: "shape", preset: "rect", text: "形状文字" },
    },
    { op: "splitWordSection", nodeId: "paragraph-1", breakType: "nextPage" },
    { op: "setWordSection", nodeId: "document-root", changes: { pageNumberStart: 9 } },
    {
      op: "insertWordStory",
      nodeId: "document-root",
      kind: "footer",
      variant: "first",
      html: "<p>页脚</p>",
    },
  ]);
  const word = session.snapshot().metadata.word!,
    main = word.stories.find((story) => story.kind === "main")!;
  expect(word.stories).toHaveLength(5);
  expect(session.snapshot().metadata.sections).toHaveLength(2);
  expect(session.snapshot().metadata.sections![0]!.rootNodeId).not.toBe(main.rootNodeId);
  expect(session.getNode("paragraph-2")!.parentNodeId).toBe(
    session.snapshot().metadata.sections![1]!.rootNodeId,
  );
  const model = openDocument(
    importDocx(exportDocx(session).bytes, { documentId: "combined", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  expect(model.snapshot().metadata.sections).toHaveLength(2);
  expect(model.snapshot().annotations[0]!.replies).toHaveLength(1);
  expect(
    model.snapshot().metadata.word!.notes?.filter((note) => !note.type || note.type === "normal"),
  ).toHaveLength(1);
  expect(model.snapshot().metadata.word!.drawings).toHaveLength(1);
  session.undo();
  expect(session.snapshot().metadata.word!.stories).toHaveLength(1);
  expect(session.getNode("paragraph-1")!.parentNodeId).toBe("document-root");
});

test("legacy section formatting history follows a relocated container and invalid batches remain atomic", () => {
  const { session, apply } = fixture(true);
  apply([
    { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "A" },
    { op: "setWordSection", nodeId: "document-root", changes: { pageNumberStart: 3 } },
  ]);
  apply([
    { op: "insertWordNote", nodeId: "paragraph-1", offset: 3, kind: "footnote", text: "注释" },
  ]);
  const section = session.snapshot().metadata.sections![0]!,
    revision = session
      .snapshot()
      .metadata.word!.revisions.find((revision) => revision.property === "section")!;
  expect(revision.nodeId).toBe(section.rootNodeId);
  const reopened = createDocumentSession(
      openDocument(
        importDocx(exportDocx(session).bytes, {
          documentId: "history",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    history = reopened
      .snapshot()
      .metadata.word!.revisions.find((revision) => revision.property === "section")!;
  reopened.commit(
    reopened.prepare({
      protocolVersion: 1,
      documentId: "history",
      baseRevision: 0,
      operations: [
        {
          op: "reviewWordRevision",
          nodeId: history.nodeId,
          revisionId: history.revisionId!,
          decision: "reject",
        },
      ],
    }),
  );
  expect(reopened.snapshot().metadata.sections![0]!.pageNumberStart).toBeUndefined();
  const before = session.snapshot();
  expect(() =>
    apply([
      { op: "setWordNoteSeparator", nodeId: "document-root", kind: "footnote", type: "separator" },
      { op: "setWordSection", nodeId: section.rootNodeId, changes: { gutter: 10000 } },
    ]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});
