import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import {
  wordNoteContentFixture,
  wordNoteContentGraph,
} from "../../office/tests/word-note-content-fixture.js";

for (const native of [false, true])
  for (const tracked of [false, true])
    for (const block of [false, true])
      test(
        "footnote and endnote clipboard owns only selected references and stories " +
          native +
          " " +
          tracked +
          " " +
          block,
        () => {
          const source = wordNoteContentFixture(native),
            graph = wordNoteContentGraph(source.session, [source.footId, source.endId]);
          expect(JSON.stringify(graph)).not.toMatch(/PRIVATE-/);
          const documentId = "note-copy-target",
            session = createDocumentSession(
              openDocument(createWordDocument({ documentId }), { kind: "document" }),
            );
          const paths = new Set(
            [
              ...(graph.semantics.stories ?? [])
                .map((s) => s.html)
                .join("")
                .matchAll(/src="([^"]+)"/g),
            ].map((m) =>
              decodeURIComponent(
                new URL(
                  m[1]!,
                  "https://doffi.invalid/" + source.session.snapshot().manifest.entrypoint,
                ).pathname.slice(1),
              ),
            ),
          );
          const assets = source.session
            .snapshot()
            .manifest.assets.filter((a) => paths.has(a.path))
            .map((descriptor) => ({
              descriptor,
              bytes: source.session.serialize().get(descriptor.path)!,
            }));
          const apply = (operations: DocumentOperation[]) =>
            session.commit(
              session.prepare(
                {
                  protocolVersion: 1,
                  documentId,
                  baseRevision: session.snapshot().manifest.revision,
                  operations,
                },
                operations.some(
                  (op) => op.op === "pasteWordInlineContent" || op.op === "pasteWordBlockContent",
                )
                  ? { assets }
                  : undefined,
              ),
            );
          apply([
            {
              op: "replaceText",
              nodeId: "paragraph-1",
              from: 0,
              to: 0,
              expectedText: "",
              text: "AB",
            },
            {
              op: "insertWordNote",
              nodeId: "paragraph-1",
              offset: 1,
              kind: "footnote",
              text: "RECEIVER-NOTE",
            },
            {
              op: "setWordReview",
              nodeId: "document-root",
              trackChanges: tracked,
              author: "Receiver",
            },
          ]);
          const before = session.snapshot(),
            oldNote = before.metadata.word!.notes![0]!;
          const operation = {
            op: block ? ("pasteWordBlockContent" as const) : ("pasteWordInlineContent" as const),
            nodeId: "paragraph-1",
            from: 0,
            to: 0,
            expectedText: "",
            html: block ? "<p>" + graph.html + "</p>" : graph.html,
            semantics: graph.semantics,
          };
          apply([operation]);
          const once = session.snapshot(),
            word = once.metadata.word!;
          expect(word.notes).toHaveLength(3);
          expect(word.noteReferences).toHaveLength(3);
          const pasted = word.notes!.filter((n) => n.noteId !== oldNote.noteId);
          expect(
            pasted.every(
              (n) =>
                !graph.semantics.notes!.some(
                  (s) => s.noteId === n.noteId || s.storyId === n.storyId,
                ),
            ),
          ).toBe(true);
          expect(word.noteReferences!.some((r) => r.customMark === "*")).toBe(true);
          expect(word.fields).toHaveLength(2);
          expect(word.tables).toHaveLength(1);
          expect(word.numbering).toHaveLength(1);
          expect(word.drawings).toHaveLength(2);
          expect(word.hyperlinks[0]!.target).toBe("https://example.com/note");
          expect(
            word.tabs.some((t) => t.stops.some((s) => s.position === 72 && s.leader === "dot")),
          ).toBe(true);
          apply([operation]);
          expect(session.snapshot().metadata.word!.notes).toHaveLength(5);
          expect(new Set(session.snapshot().metadata.word!.notes!.map((n) => n.storyId)).size).toBe(
            5,
          );
          session.undo();
          expect(session.snapshot().html).toBe(once.html);
          let bytes = exportDocx(session).bytes;
          for (let pass = 0; pass < 2; pass++) {
            const reopened = createDocumentSession(
              openDocument(importDocx(bytes, { documentId, unsupportedContent: "reject" }), {
                kind: "document",
              }),
            );
            expect(
              reopened
                .snapshot()
                .metadata.word!.notes!.filter((n) => (n.type ?? "normal") === "normal"),
            ).toHaveLength(3);
            expect(reopened.snapshot().metadata.word!.fields).toHaveLength(2);
            expect(
              reopened.snapshot().metadata.word!.noteReferences!.some((r) => r.customMark === "*"),
            ).toBe(true);
            expect(reopened.snapshot().html).toContain("NOTE-CELL");
            if (tracked) {
              let root = reopened.getNode(
                reopened.snapshot().metadata.word!.stories[0]!.rootNodeId,
              )!;
              while (root.parentNodeId) root = reopened.getNode(root.parentNodeId)!;
              reopened.commit(
                reopened.prepare({
                  protocolVersion: 1,
                  documentId,
                  baseRevision: reopened.snapshot().manifest.revision,
                  operations: [
                    { op: "reviewWordRevisions", nodeId: root.nodeId, decision: "reject" },
                  ],
                }),
              );
              const current = reopened.snapshot().metadata.word!;
              expect(current.notes!.filter((n) => (n.type ?? "normal") === "normal")).toHaveLength(
                1,
              );
              expect(current.fields).toHaveLength(0);
              expect(current.tables).toHaveLength(0);
              expect(current.drawings ?? []).toHaveLength(0);
            }
            bytes = exportDocx(
              openDocument(importDocx(bytes, { documentId, unsupportedContent: "reject" }), {
                kind: "document",
              }),
            ).bytes;
          }
          if (tracked) {
            apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" }]);
            expect(session.snapshot().metadata.word!.notes).toHaveLength(1);
            expect(session.snapshot().manifest.assets).toEqual(before.manifest.assets);
            expect(
              session.snapshot().metadata.word!.stories.filter((s) => s.kind === "textBox"),
            ).toHaveLength(0);
            session.undo();
          }
          session.undo();
          expect(session.snapshot().html).toBe(before.html);
          session.redo();
          expect(session.snapshot().metadata.word!.notes).toHaveLength(3);
        },
      );

test("note transport identities cannot resolve against receiver notes and malformed graphs are atomic", () => {
  const source = wordNoteContentFixture(),
    graph = wordNoteContentGraph(source.session, [source.endId]),
    documentId = "note-invalid",
    session = createDocumentSession(
      openDocument(createWordDocument({ documentId }), { kind: "document" }),
    );
  const before = session.snapshot();
  for (const mutate of [
    (g: typeof graph) => {
      g.semantics.stories = [];
    },
    (g: typeof graph) => {
      g.semantics.noteReferences![0]!.noteId = "receiver-note";
    },
    (g: typeof graph) => {
      g.semantics.noteReferences!.push({ ...g.semantics.noteReferences![0]! });
    },
    (g: typeof graph) => {
      g.semantics.notes![0]!.kind = "footnote";
    },
  ]) {
    const copy = JSON.parse(JSON.stringify(graph)) as typeof graph;
    mutate(copy);
    expect(() =>
      session.prepare({
        protocolVersion: 1,
        documentId,
        baseRevision: 0,
        operations: [
          {
            op: "pasteWordInlineContent",
            nodeId: "paragraph-1",
            from: 0,
            to: 0,
            expectedText: "",
            ...copy,
          },
        ],
      }),
    ).toThrow();
    expect(session.snapshot()).toEqual(before);
  }
});

test("paste and reject in the same candidate leaves no newly owned note story", () => {
  const source = wordNoteContentFixture(),
    graph = wordNoteContentGraph(source.session, [source.endId]),
    documentId = "note-batch",
    session = createDocumentSession(
      openDocument(createWordDocument({ documentId }), { kind: "document" }),
    );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId,
      baseRevision: 0,
      operations: [
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "Receiver" },
        {
          op: "pasteWordInlineContent",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          ...graph,
        },
        { op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" },
      ],
    }),
  );
  const word = session.snapshot().metadata.word!;
  expect(word.notes ?? []).toHaveLength(0);
  expect(word.noteReferences ?? []).toHaveLength(0);
  expect(word.stories.filter((s) => s.kind === "endnote")).toHaveLength(0);
  expect(word.fields).toHaveLength(0);
  expect(word.tables).toHaveLength(0);
});

for (const decision of ["accept", "reject"] as const)
  test(
    "moving a selected note retains only one owned body after native grouped review: " + decision,
    () => {
      const { session, paragraphId, footId } = wordNoteContentFixture(true),
        graph = wordNoteContentGraph(session, [footId]),
        documentId = session.snapshot().manifest.documentId;
      let current = footId,
        from = 0;
      while (current !== paragraphId) {
        const node = session.getNode(current)!,
          parent = session.getNode(node.parentNodeId!)!;
        from += parent.childNodeIds
          .slice(0, parent.childNodeIds.indexOf(current))
          .reduce((sum, id) => sum + session.getNode(id)!.text.length, 0);
        current = parent.nodeId;
      }
      let root = session.getNode(paragraphId)!;
      while (root.parentNodeId) root = session.getNode(root.parentNodeId)!;
      const apply = (operations: DocumentOperation[]) =>
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId,
            baseRevision: session.snapshot().manifest.revision,
            operations,
          }),
        );
      apply([{ op: "setWordReview", nodeId: root.nodeId, trackChanges: true, author: "Mover" }]);
      const before = session.snapshot();
      apply([
        {
          op: "cutWordContent",
          nodeId: paragraphId,
          from,
          to: from + 1,
          expectedText: "\ufffc",
          cutId: "note-move",
        },
      ]);
      apply([
        {
          op: "pasteWordInlineContent",
          nodeId: paragraphId,
          from: 0,
          to: 0,
          expectedText: "",
          ...graph,
          cutId: "note-move",
        },
      ]);
      expect(session.snapshot().metadata.word!.moves).toHaveLength(2);
      const reopened = createDocumentSession(
          openDocument(
            importDocx(exportDocx(session).bytes, { documentId, unsupportedContent: "reject" }),
            { kind: "document" },
          ),
        ),
        revision = reopened.snapshot().metadata.word!.revisions.find((r) => r.kind === "moveTo")!;
      reopened.commit(
        reopened.prepare({
          protocolVersion: 1,
          documentId,
          baseRevision: reopened.snapshot().manifest.revision,
          operations: [
            {
              op: "reviewWordRevision",
              nodeId: revision.nodeId,
              revisionId: revision.revisionId ?? revision.nodeId,
              decision,
            },
          ],
        }),
      );
      const word = reopened.snapshot().metadata.word!;
      expect(word.notes!.filter((n) => (n.type ?? "normal") === "normal")).toHaveLength(4);
      expect(word.fields).toHaveLength(2);
      expect(word.drawings).toHaveLength(2);
      expect(
        word.notes!.filter((n) =>
          reopened
            .getNode(word.stories.find((s) => s.storyId === n.storyId)!.rootNodeId)!
            .text.includes("FOOT-BODY"),
        ),
      ).toHaveLength(1);
      session.undo();
      session.undo();
      expect(session.snapshot().html).toBe(before.html);
    },
  );

test("legacy note counter owners do not hide their nested field on native output", () => {
  const documentId = "legacy-note-counter",
    session = createDocumentSession(
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
    { op: "insertWordNote", nodeId: "paragraph-1", offset: 0, kind: "footnote", text: "LEGACY" },
  ]);
  const word = session.snapshot().metadata.word!,
    story = word.stories.find((s) => s.storyId === word.notes![0]!.storyId)!,
    p = session.getNode(story.rootNodeId)!.childNodeIds[0]!;
  apply([{ op: "insertWordField", nodeId: p, offset: 0, instruction: " PAGE ", cachedText: "7" }]);
  const field = session.getNode(session.snapshot().metadata.word!.fields[0]!.nodeId)!,
    counter = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
      .map((m) => session.getNode(m[1]!)!)
      .find((n) => n.attributes["data-word-note-number"] === "true")!;
  expect(field.parentNodeId).toBe(p);
  const files = new Map(session.serialize());
  files.set(
    session.snapshot().manifest.entrypoint,
    new TextEncoder().encode(
      session
        .snapshot()
        .html.replace(field.html, "")
        .replace(counter.html, counter.html.replace("\ufffc", field.html + "\ufffc")),
    ),
  );
  const legacy = openDocument(files, { kind: "document" }),
    reopened = openDocument(
      importDocx(exportDocx(legacy).bytes, { documentId, unsupportedContent: "reject" }),
      { kind: "document" },
    );
  expect(reopened.snapshot().metadata.word!.fields).toHaveLength(1);
  expect(reopened.snapshot().metadata.word!.fields[0]!.instruction.trim()).toBe("PAGE");
});
