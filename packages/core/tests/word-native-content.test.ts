import { expect, test } from "vitest";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { nativeContentGraph } from "../../office/tests/word-native-content-fixture.js";

function fixture(tracked = false) {
  const documentId = "native-receiver",
    session = createDocumentSession(
      openDocument(createWordDocument({ documentId }), { kind: "document" }),
    ),
    graph = nativeContentGraph(),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare(
          {
            protocolVersion: 1,
            documentId,
            baseRevision: session.snapshot().manifest.revision,
            operations,
          },
          operations.some((op) => op.op === "pasteWordBlockContent")
            ? { assets: graph.assets }
            : undefined,
        ),
      );
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
    { op: "setWordReview", nodeId: "document-root", trackChanges: tracked, author: "Receiver" },
  ]);
  const paste = {
    op: "pasteWordBlockContent" as const,
    nodeId: "paragraph-1",
    from: 1,
    to: 1,
    expectedText: "",
    html: graph.html,
    semantics: graph.semantics,
  };
  return { session, graph, apply, paste };
}
for (const decision of ["accept", "reject"] as const)
  test("opaque block insertion shares atomic review and undo " + decision, () => {
    const { session, apply, paste } = fixture(true),
      before = session.snapshot();
    apply([paste]);
    const pasted = session.snapshot(),
      revision = pasted.metadata.word!.revisions.find((r) => r.kind === "nativeInsert")!;
    expect(revision).toBeDefined();
    expect(session.getNode(revision.nodeId)!.text).toBe("\ufffc");
    apply([
      {
        op: "reviewWordRevision",
        nodeId: revision.nodeId,
        revisionId: revision.revisionId!,
        decision,
      },
    ]);
    expect(session.snapshot().metadata.word!.revisions).toHaveLength(0);
    expect(session.getNode("document-root")!.text).toBe(decision === "reject" ? "AB" : "A\ufffcB");
    session.undo();
    expect(session.snapshot().html).toBe(pasted.html);
    session.undo();
    expect(session.snapshot().html).toBe(before.html);
  });

test("native clipboard validates the complete declared dependency graph and stale candidates without changing source", () => {
  const { session, graph, paste, apply } = fixture(),
    before = session.snapshot();
  for (const change of ["missing", "surplus", "root", "foreign-node"] as const) {
    const semantics = structuredClone(graph.semantics),
      native = semantics.nativeContent!;
    if (change === "missing") native.fragments[0]!.dependencyParts = ["customXml/missing.xml"];
    if (change === "surplus")
      native.parts.push({
        sourcePart: "customXml/private.xml",
        assetPath: graph.assets[0]!.descriptor.path,
        contentType: "application/xml",
      });
    if (change === "root")
      native.relationships.push({
        sourcePart: "word/document.xml",
        id: "private",
        type: "urn:private",
        target: "https://example.invalid",
        external: true,
        root: true,
      });
    if (change === "foreign-node") native.fragments[0]!.nodeId = "paragraph-1";
    expect(() => apply([{ ...paste, semantics }])).toThrow();
    expect(session.snapshot()).toEqual(before);
  }
  const candidate = session.prepare(
    {
      protocolVersion: 1,
      documentId: before.manifest.documentId,
      baseRevision: before.manifest.revision,
      operations: [paste],
    },
    { assets: graph.assets },
  );
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 1, expectedText: "A", text: "C" },
  ]);
  const edited = session.snapshot();
  expect(() => session.commit(candidate)).toThrow();
  expect(session.snapshot()).toEqual(edited);
});

test("repeated native clipboard inserts isolate part scopes and release only removed selected resources", () => {
  const { session, graph, apply, paste } = fixture();
  apply([paste]);
  const first = session.snapshot().metadata.word!.nativeContent!.fragments[0]!;
  apply([{ ...paste, from: 0, to: 0 }]);
  const fragments = session.snapshot().metadata.word!.nativeContent!.fragments;
  expect(new Set(fragments.map((f) => f.sourcePart)).size).toBe(2);
  expect(new Set(fragments.map((f) => f.nativeId)).size).toBe(2);
  apply([{ op: "removeNode", nodeId: first.nodeId! }]);
  expect(session.serialize().has(graph.assets[0]!.descriptor.path)).toBe(true);
  apply([{ op: "removeNode", nodeId: fragments[1]!.nodeId! }]);
  expect(session.serialize().has(graph.assets[0]!.descriptor.path)).toBe(false);
  session.undo();
  expect(session.serialize().get(graph.assets[0]!.descriptor.path)).toEqual(graph.assets[0]!.bytes);
});
