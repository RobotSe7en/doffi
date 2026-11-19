import { expect, test } from "vitest";
import { createWordDocument, openDocument, createDocumentSession } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
function create() {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "pages" }), { kind: "document" }),
  );
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "pages",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  return { session, apply };
}
test("page geometry changes and section splits retain source identities and one-step history", () => {
  const { session, apply } = create();
  apply([
    {
      op: "setWordSection",
      nodeId: "section-root-1",
      changes: {
        pageSetup: {
          unit: "pt",
          width: 792,
          height: 612,
          margins: { top: 50, right: 60, bottom: 70, left: 80 },
        },
        pageNumberStart: 3,
        pageNumberFormat: "upperRoman",
      },
    },
  ]);
  expect(session.snapshot().metadata.sections![0]!.pageSetup!.width).toBe(792);
  apply([{ op: "splitWordSection", nodeId: "paragraph-1", breakType: "oddPage" }]);
  expect(session.snapshot().metadata.sections).toHaveLength(2);
  expect(session.snapshot().metadata.sections![1]!.pageNumberStart).toBeUndefined();
  expect(session.getNode("paragraph-1")).toBeDefined();
  session.undo();
  expect(session.snapshot().metadata.sections).toHaveLength(1);
  session.redo();
  expect(openDocument(session.serialize(), { kind: "document" }).snapshot()).toEqual(
    session.snapshot(),
  );
});
test("headers and footers remain separate rich source stories and page fields use typed metadata", () => {
  const { session, apply } = create();
  apply([
    {
      op: "insertWordStory",
      nodeId: "section-root-1",
      kind: "header",
      variant: "first",
      html: "<p>页眉😀</p>",
    },
    {
      op: "insertWordStory",
      nodeId: "section-root-1",
      kind: "footer",
      variant: "default",
      html: "<p></p>",
    },
  ]);
  const word = session.snapshot().metadata.word!,
    footer = word.stories.find((story) => story.kind === "footer")!,
    paragraph = session.getNode(footer.rootNodeId)!.childNodeIds[0]!;
  apply([{ op: "insertWordField", nodeId: paragraph, instruction: " PAGE ", cachedText: "1" }]);
  expect(session.snapshot().metadata.word!.fields).toHaveLength(1);
  expect(session.snapshot().metadata.sections![0]!.differentFirstPage).toBe(true);
  expect(session.getNode("paragraph-1")!.text).toBe("");
  const before = session.snapshot();
  expect(() =>
    apply([
      {
        op: "setWordSection",
        nodeId: "section-root-1",
        changes: {
          pageSetup: {
            unit: "pt",
            width: 100,
            height: 100,
            margins: { top: 80, right: 80, bottom: 80, left: 80 },
          },
        },
      },
    ]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});
test("unlinking a header clones source IDs and its field/table/tab semantics in one undo step", () => {
  const { session, apply } = create();
  apply([
    {
      op: "insertWordStory",
      nodeId: "section-root-1",
      kind: "header",
      variant: "default",
      html: "<p>页眉</p>",
    },
  ]);
  const story = session.snapshot().metadata.word!.stories.find((story) => story.kind === "header")!,
    paragraph = session.getNode(story.rootNodeId)!.childNodeIds[0]!;
  apply([
    { op: "insertWordField", nodeId: paragraph, instruction: "PAGE" },
    {
      op: "setWordTabs",
      nodeId: paragraph,
      defaultInterval: 36,
      stops: [{ position: 72, alignment: "right", leader: "dot" }],
    },
    { op: "splitWordSection", nodeId: "paragraph-1", breakType: "nextPage" },
  ]);
  const next = session.snapshot().metadata.sections![1]!;
  apply([
    { op: "cloneWordStory", nodeId: next.rootNodeId, storyId: story.storyId, variant: "default" },
  ]);
  const snapshot = session.snapshot(),
    clone = snapshot.metadata.word!.stories.find(
      (item) => item.storyId === snapshot.metadata.sections![1]!.headers!.default,
    )!;
  expect(clone.rootNodeId).not.toBe(story.rootNodeId);
  expect(session.getNode(clone.rootNodeId)!.text).toBe(session.getNode(story.rootNodeId)!.text);
  expect(snapshot.metadata.word!.fields).toHaveLength(2);
  expect(snapshot.metadata.word!.tabs).toHaveLength(2);
  expect(openDocument(session.serialize(), { kind: "document" }).snapshot()).toEqual(snapshot);
  session.undo();
  expect(session.snapshot().metadata.word!.fields).toHaveLength(1);
});
test("invalid gutter and duplicate tabs cannot alter the accepted revision", () => {
  const { session, apply } = create(),
    before = session.snapshot();
  expect(() =>
    apply([{ op: "setWordSection", nodeId: "section-root-1", changes: { gutter: 10000 } }]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
  expect(() =>
    apply([
      {
        op: "setWordTabs",
        nodeId: "paragraph-1",
        defaultInterval: 36,
        stops: [
          { position: 24, alignment: "left", leader: "none" },
          { position: 24, alignment: "right", leader: "dot" },
        ],
      },
    ]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});
test("paragraph-scoped proposals cannot split the containing section or copy an unread story", () => {
  const { session, apply } = create();
  apply([
    {
      op: "insertWordStory",
      nodeId: "section-root-1",
      kind: "header",
      variant: "default",
      html: "<p>独立内容</p>",
    },
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "paragraph-scope",
        documentId: "pages",
        baseRevision: 0,
        status: "open",
        anchor: { kind: "node", nodeId: "paragraph-1" },
        instruction: "修改段落",
      },
    },
  ]);
  const request = session.buildInteractionRequest({ annotationIds: ["paragraph-scope"] }),
    proposal = {
      protocolVersion: 1 as const,
      proposalId: "scoped",
      summary: "布局操作",
      requestId: request.requestId,
      documentId: "pages",
      baseRevision: request.baseRevision,
      annotationIds: ["paragraph-scope"],
      operations: [
        { op: "splitWordSection" as const, nodeId: "paragraph-1", breakType: "nextPage" as const },
      ],
    };
  expect(() => session.prepareProposal(request, proposal)).toThrowError(
    expect.objectContaining({ code: "InvalidProposal" }),
  );
  apply([
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "section-scope",
        documentId: "pages",
        baseRevision: session.snapshot().manifest.revision,
        status: "open",
        anchor: { kind: "node", nodeId: "section-root-1" },
        instruction: "修改本节",
      },
    },
  ]);
  const next = session.buildInteractionRequest({ annotationIds: ["section-scope"] });
  expect(() =>
    session.prepareProposal(next, {
      ...proposal,
      requestId: next.requestId,
      baseRevision: next.baseRevision,
      annotationIds: ["section-scope"],
      operations: [
        {
          op: "cloneWordStory",
          nodeId: "section-root-1",
          storyId: session
            .snapshot()
            .metadata.word!.stories.find((story) => story.kind === "header")!.storyId,
          variant: "default",
        },
      ],
    }),
  ).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
});
