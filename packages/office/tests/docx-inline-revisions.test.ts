import { expect, test } from "vitest";
import { readFileSync } from "node:fs";
import {
  createWordDocument,
  createDocumentSession,
  openDocument,
  createAssetDescriptor,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
} from "@doffi/core";
import type { ContentDocumentModel, DocumentOperation, SourceStyleNode } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
function setup() {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "inline-revision" }), { kind: "document" }),
    ),
    apply = (operations: DocumentOperation[], options?: Parameters<typeof session.prepare>[1]) =>
      session.commit(
        session.prepare(
          {
            protocolVersion: 1,
            documentId: "inline-revision",
            baseRevision: session.snapshot().manifest.revision,
            operations,
          },
          options,
        ),
      );
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
    { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "A" },
  ]);
  return { session, apply };
}
function appearance(model: ContentDocumentModel, text: string) {
  const word = model.snapshot().metadata.word!,
    css = new SourceStyleResolver(
      model.snapshot().styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
      new Set(word.revisions.map((item) => item.nodeId)),
    );
  let result: Readonly<Record<string, string>> | undefined;
  const visit = (id: string, parent?: SourceStyleNode): void => {
    const node = model.getNode(id)!,
      source = { tag: node.tagName, attrs: node.attributes, parent };
    if (node.text === text) result = css.get(source);
    for (const child of node.childNodeIds) visit(child, source);
  };
  visit(word.stories[0]!.rootNodeId);
  return result!;
}
for (const decision of ["accept", "reject"] as const)
  test(
    "formatted inline replacement keeps its own styles and native accept/reject through two reopens: " +
      decision,
    () => {
      const { session, apply } = setup();
      apply([
        {
          op: "trackWordInlineContent",
          nodeId: "paragraph-1",
          from: 1,
          to: 2,
          expectedText: "B",
          html: '<strong style="font-size:24pt;color:#880000">X</strong><em>Y</em>',
        },
      ]);
      expect(session.getNode("paragraph-1")!.text).toBe("ABXY");
      expect(appearance(session, "X")).toMatchObject({
        "font-size": "24pt",
        color: "#880000",
        "font-weight": "bold",
      });
      expect(appearance(session, "Y")["font-style"]).toBe("italic");
      const before = session.snapshot(),
        grouped = createDocumentSession(session),
        group = grouped.snapshot().metadata.word!.revisions[0]!;
      grouped.commit(
        grouped.prepare({
          protocolVersion: 1,
          documentId: "inline-revision",
          baseRevision: session.snapshot().manifest.revision,
          operations: [
            {
              op: "reviewWordRevision",
              nodeId: group.nodeId,
              revisionId: group.revisionId ?? group.nodeId,
              decision,
            },
          ],
        }),
      );
      expect(grouped.getNode("paragraph-1")!.text).toBe(decision === "accept" ? "AXY" : "AB");
      expect(grouped.snapshot().metadata.word!.revisions).toHaveLength(0);
      let bytes = exportDocx(session).bytes;
      for (let i = 0; i < 2; i++) {
        const model = openDocument(
          importDocx(bytes, { documentId: "inline-native", unsupportedContent: "reject" }),
          { kind: "document" },
        );
        expect(appearance(model, "X")).toMatchObject({
          "font-size": "24pt",
          color: "#880000",
          "font-weight": "bold",
        });
        const reviewer = createDocumentSession(model);
        reviewer.commit(
          reviewer.prepare({
            protocolVersion: 1,
            documentId: "inline-native",
            baseRevision: 0,
            operations: [
              {
                op: "reviewWordRevisions",
                nodeId: reviewer.getNode(reviewer.snapshot().metadata.word!.stories[0]!.rootNodeId)!
                  .parentNodeId!,
                decision,
              },
            ],
          }),
        );
        expect(
          reviewer.getNode(reviewer.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text,
        ).toBe(decision === "accept" ? "AXY" : "AB");
        openDocument(
          importDocx(exportDocx(reviewer).bytes, {
            documentId: "reviewed-inline",
            unsupportedContent: "reject",
          }),
          { kind: "document" },
        );
        bytes = exportDocx(model).bytes;
      }
      session.undo();
      expect(session.getNode("paragraph-1")!.text).toBe("AB");
      session.redo();
      expect(session.snapshot().html).toBe(before.html);
      expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
    },
  );
test("editing an owned pending formatted insertion replaces its text and image within one original revision", () => {
  const { session, apply } = setup(),
    bytes = new Uint8Array(
      readFileSync(new URL("../../../examples/fidelity/assets/sample.png", import.meta.url)),
    ),
    descriptor = createAssetDescriptor("assets/inline.png", "image/png", bytes);
  apply(
    [
      {
        op: "trackWordInlineContent",
        nodeId: "paragraph-1",
        from: 1,
        to: 1,
        expectedText: "",
        html: '<span style="font-size:24pt">X<img src="assets/inline.png" style="width:12pt;height:8pt"></span>',
      },
    ],
    { assets: [{ descriptor, bytes }] },
  );
  const revision = session.snapshot().metadata.word!.revisions[0]!,
    before = session.snapshot();
  expect(session.getNode("paragraph-1")!.text).toBe("AX\ufffcB");
  const native = openDocument(
    importDocx(exportDocx(session).bytes, {
      documentId: "image-inline",
      unsupportedContent: "reject",
    }),
    { kind: "document" },
  );
  expect(
    native.snapshot().metadata.word!.drawings!.filter((item) => item.kind === "image"),
  ).toHaveLength(1);
  for (const decision of ["accept", "reject"] as const) {
    const reviewer = createDocumentSession(native),
      item = reviewer.snapshot().metadata.word!.revisions[0]!;
    reviewer.commit(
      reviewer.prepare({
        protocolVersion: 1,
        documentId: "image-inline",
        baseRevision: 0,
        operations: [
          {
            op: "reviewWordRevision",
            nodeId: item.nodeId,
            revisionId: item.revisionId ?? item.nodeId,
            decision,
          },
        ],
      }),
    );
    openDocument(
      importDocx(exportDocx(reviewer).bytes, {
        documentId: "image-reviewed",
        unsupportedContent: "reject",
      }),
      { kind: "document" },
    );
  }
  apply([
    {
      op: "trackWordInlineContent",
      nodeId: "paragraph-1",
      from: 1,
      to: 3,
      expectedText: "X\ufffc",
      html: "<u>Z</u>",
    },
  ]);
  expect(session.getNode("paragraph-1")!.text).toBe("AZB");
  expect(session.snapshot().metadata.word!.revisions).toEqual([revision]);
  expect(session.snapshot().manifest.assets).toHaveLength(0);
  expect(appearance(session, "Z")["text-decoration-line"]).toBe("underline");
  openDocument(
    importDocx(exportDocx(session).bytes, {
      documentId: "continued-inline",
      unsupportedContent: "reject",
    }),
    { kind: "document" },
  );
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  expect(session.snapshot().manifest.assets).toEqual(before.manifest.assets);
});
test("invalid inline structure and Unicode replacement boundaries fail without changing source or revisions", () => {
  const { session, apply } = setup(),
    before = session.snapshot();
  for (const html of [
    "<p>BLOCK</p>",
    "<script>execute()</script>",
    '<span data-node-id="forged">X</span>',
  ])
    expect(() =>
      apply([
        {
          op: "trackWordInlineContent",
          nodeId: "paragraph-1",
          from: 1,
          to: 2,
          expectedText: "B",
          html,
        },
      ]),
    ).toThrow();
  expect(() =>
    apply([
      {
        op: "trackWordInlineContent",
        nodeId: "paragraph-1",
        from: 0,
        to: 1,
        expectedText: "WRONG",
        html: "X",
      },
    ]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});
