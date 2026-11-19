import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { gridBoundaryFixture } from "./grid-boundary-fixture.js";
import type { GridBoundaryKind } from "./grid-boundary-fixture.js";
import { strFromU8, unzipSync } from "fflate";
const setup = (kind: GridBoundaryKind, bytes = gridBoundaryFixture(kind)) => {
  const session = createDocumentSession(
    openDocument(
      importDocx(bytes, { documentId: "grid-boundaries", unsupportedContent: "reject" }),
      { kind: "document" },
    ),
  );
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "grid-boundaries",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  const table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    root = session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
  return { session, apply, table, root };
};
for (const kind of ["merged", "vertical", "before", "after"] as const)
  for (const decision of ["accept", "reject"] as const)
    test(
      "structural grid insertion wholly in " +
        kind +
        " regions supports " +
        decision +
        " through strict native reopens",
      () => {
        const { session, apply, table, root } = setup(kind),
          before = session.snapshot(),
          text = session.getNode(table)!.text,
          at = kind === "after" ? 2 : 1;
        apply([
          { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
          { op: "editWordTable", nodeId: table, change: { kind: "insertColumns", at, count: 2 } },
        ]);
        expect(session.snapshot().metadata.word!.tables[0]!.grid).toHaveLength(5);
        expect(
          session.snapshot().metadata.word!.revisions.some((item) => item.kind === "cellInsert"),
        ).toBe(false);
        expect(session.getNode(table)!.text).toBe(text);
        let bytes = exportDocx(session).bytes;
        for (let pass = 0; pass < 2; pass++) {
          const native = setup(kind, bytes);
          expect(native.session.snapshot().metadata.word!.tables[0]!.grid).toHaveLength(5);
          bytes = exportDocx(native.session).bytes;
        }
        const native = setup(kind, bytes),
          revision = native.session
            .snapshot()
            .metadata.word!.revisions.find((item) => item.property === "tableGrid")!;
        native.apply([
          {
            op: "reviewWordRevision",
            nodeId: revision.nodeId,
            revisionId: revision.revisionId!,
            decision,
          },
        ]);
        expect(native.session.snapshot().metadata.word!.tables[0]!.grid).toHaveLength(
          decision === "reject" ? 3 : 5,
        );
        if (decision === "reject")
          expect(native.session.snapshot().metadata.word!.tables[0]!.grid).toEqual([40, 60, 116]);
        expect(native.session.getNode(native.table)!.text).toBe(text);
        expect(native.session.snapshot().metadata.word!.revisions).toEqual([]);
        let reviewed = exportDocx(native.session).bytes;
        for (let pass = 0; pass < 2; pass++) {
          const model = openDocument(
            importDocx(reviewed, { documentId: "reviewed", unsupportedContent: "reject" }),
            { kind: "document" },
          );
          expect(model.snapshot().metadata.word!.tables[0]!.grid).toHaveLength(
            decision === "reject" ? 3 : 5,
          );
          reviewed = exportDocx(model).bytes;
        }
        if (decision === "reject" && (kind === "before" || kind === "after"))
          expect(strFromU8(unzipSync(reviewed)["word/document.xml"]!)).toContain('w:w="25%"');
        apply([{ op: "reviewWordRevisions", nodeId: root, decision }]);
        expect(session.snapshot().metadata.word!.tables[0]!.grid).toHaveLength(
          decision === "reject" ? 3 : 5,
        );
        session.undo();
        expect(session.snapshot().metadata.word!.tables[0]!.grid).toHaveLength(5);
        session.undo();
        expect(session.snapshot().html).toBe(before.html);
        expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
      },
    );
test("successive structural insertions retain the earliest inverse and stale candidates cannot apply", () => {
  const { session, apply, table, root } = setup("merged");
  apply([{ op: "setWordReview", nodeId: root, trackChanges: true, author: "A" }]);
  const stale = session.prepare({
    protocolVersion: 1,
    documentId: "grid-boundaries",
    baseRevision: session.snapshot().manifest.revision,
    operations: [
      { op: "editWordTable", nodeId: table, change: { kind: "insertColumns", at: 1, count: 1 } },
    ],
  });
  apply([
    { op: "editWordTable", nodeId: table, change: { kind: "insertColumns", at: 1, count: 1 } },
    { op: "editWordTable", nodeId: table, change: { kind: "insertColumns", at: 2, count: 1 } },
  ]);
  expect(() => session.commit(stale)).toThrowError(
    expect.objectContaining({ code: "RevisionConflict" }),
  );
  const native = setup("merged", exportDocx(session).bytes);
  native.apply([{ op: "reviewWordRevisions", nodeId: native.root, decision: "reject" }]);
  expect(native.session.snapshot().metadata.word!.tables[0]!.grid).toEqual([40, 60, 116]);
});

for (const kind of ["merged", "vertical", "before", "after"] as const)
  test(
    "own structural insertion can be cancelled within " + kind + " without losing earliest history",
    () => {
      const { session, apply, table, root } = setup(kind),
        at = kind === "after" ? 2 : 1;
      apply([
        { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
        { op: "editWordTable", nodeId: table, change: { kind: "insertColumns", at, count: 2 } },
      ]);
      apply([
        { op: "editWordTable", nodeId: table, change: { kind: "deleteColumns", at, count: 1 } },
      ]);
      expect(session.snapshot().metadata.word!.tables[0]!.grid).toHaveLength(4);
      const native = setup(kind, exportDocx(session).bytes);
      native.apply([{ op: "reviewWordRevisions", nodeId: native.root, decision: "reject" }]);
      expect(native.session.snapshot().metadata.word!.tables[0]!.grid).toEqual([40, 60, 116]);
      apply([
        { op: "editWordTable", nodeId: table, change: { kind: "deleteColumns", at, count: 1 } },
      ]);
      expect(session.snapshot().metadata.word!.tables[0]!.grid).toEqual([40, 60, 116]);
      expect(session.snapshot().metadata.word!.revisions).toEqual([]);
    },
  );
