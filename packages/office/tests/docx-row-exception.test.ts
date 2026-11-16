import { expect, test } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  createDocumentSession,
  openDocument,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
  wordRowExceptionProperties,
} from "@doffi/core";
import type { ContentDocumentModel, DocumentOperation, SourceStyleNode } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { rowExceptionFixture } from "./row-exception-fixture.js";
function values(model: ContentDocumentModel) {
  const css = new SourceStyleResolver(
      model.snapshot().styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    project = (id: string): SourceStyleNode => {
      const node = model.getNode(id)!;
      return {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
    },
    table = model.snapshot().metadata.word!.tables[0]!;
  return model.getNode(table.nodeId)!.childNodeIds.map((id) => {
    const value = css.get(project(id));
    return {
      direct: value["--word-row-cell-spacing"],
      exception: value["--word-row-exception-cell-spacing"],
    };
  });
}
function setup() {
  const session = createDocumentSession(
      openDocument(
        importDocx(rowExceptionFixture(), {
          documentId: "row-exception",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "row-exception",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  return { session, apply };
}
test("table-level row exceptions retain their native declaration independently of direct row properties", () => {
  let bytes = rowExceptionFixture();
  for (let pass = 0; pass < 3; pass++) {
    const model = openDocument(
      importDocx(bytes, { documentId: "row-ex-native", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(model.snapshot().metadata.word!.nativeContent?.fragments ?? []).toEqual([]);
    expect(values(model)).toEqual([
      { direct: "0pt", exception: "4pt" },
      { direct: "12pt", exception: "4pt" },
      { direct: undefined, exception: "4pt" },
    ]);
    bytes = exportDocx(model).bytes;
    expect(strFromU8(unzipSync(bytes)["word/document.xml"]!).match(/<w:tblPrEx>/g)).toHaveLength(3);
  }
});
test("resetting direct row spacing exposes its exception and keeps native reject, editing and undo independent", () => {
  const { session, apply } = setup(),
    table = session.snapshot().metadata.word!.tables[0]!,
    root = session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!,
    before = session.snapshot();
  apply([
    { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
    { op: "formatWordTable", nodeId: table.nodeId, rowIndex: 0, rowCellSpacing: null },
  ]);
  expect(values(session)[0]).toEqual({ direct: undefined, exception: "4pt" });
  let bytes = exportDocx(session).bytes;
  for (const decision of ["accept", "reject"] as const) {
    const reviewer = createDocumentSession(
        openDocument(
          importDocx(bytes, { documentId: "row-ex-review", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
      ),
      main = reviewer.snapshot().metadata.word!.stories[0]!;
    reviewer.commit(
      reviewer.prepare({
        protocolVersion: 1,
        documentId: "row-ex-review",
        baseRevision: 0,
        operations: [
          {
            op: "reviewWordRevisions",
            nodeId: reviewer.getNode(main.rootNodeId)!.parentNodeId!,
            decision,
          },
        ],
      }),
    );
    expect(values(reviewer)[0]).toEqual({
      direct: decision === "reject" ? "0pt" : undefined,
      exception: "4pt",
    });
    openDocument(
      importDocx(exportDocx(reviewer).bytes, {
        documentId: "row-ex-reviewed",
        unsupportedContent: "reject",
      }),
      { kind: "document" },
    );
  }
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
});
test("row exceptions follow surviving row identities during row insertion and column edits", () => {
  const { session, apply } = setup(),
    table = session.snapshot().metadata.word!.tables[0]!,
    rows = session.getNode(table.nodeId)!.childNodeIds,
    before = session.snapshot();
  apply([
    { op: "editWordTable", nodeId: table.nodeId, change: { kind: "insertRows", at: 0, count: 1 } },
    {
      op: "editWordTable",
      nodeId: table.nodeId,
      change: { kind: "insertColumns", at: 0, count: 1 },
    },
  ]);
  for (const id of rows)
    expect(session.getNode(id)!.attributes.class).toContain("word-row-exception");
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const model = openDocument(
      importDocx(bytes, { documentId: "row-ex-shifted", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(values(model).filter((row) => row.exception === "4pt")).toHaveLength(3);
    bytes = exportDocx(model).bytes;
  }
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
});

test("mixed native row exceptions keep their unmapped properties and the editable table through two reopens", () => {
  const files = unzipSync(rowExceptionFixture());
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replaceAll(
      "<w:tblPrEx>",
      '<w:tblPrEx><w:jc w:val="right"/>',
    ),
  );
  let bytes: Uint8Array = zipSync(files);
  for (let pass = 0; pass < 2; pass++) {
    const model = openDocument(importDocx(bytes, { documentId: "row-ex-mixed" }), {
      kind: "document",
    });
    expect(model.snapshot().metadata.word!.tables).toHaveLength(1);
    expect(values(model).filter((item) => item.exception === "4pt")).toHaveLength(3);
    bytes = exportDocx(model).bytes;
    expect(
      strFromU8(unzipSync(bytes)["word/document.xml"]!).match(/<w:jc[^>]*w:val="right"/g),
    ).toHaveLength(3);
  }
});

function rowAppearance(model: ContentDocumentModel, index = 0) {
  const resolver = new SourceStyleResolver(
    model.snapshot().styles.map((s) => s.css),
    DEFAULT_DOCUMENT_READ_LIMITS,
  );
  const project = (id: string): SourceStyleNode => {
    const node = model.getNode(id)!;
    return {
      tag: node.tagName,
      attrs: node.attributes,
      parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
    };
  };
  const table = model.snapshot().metadata.word!.tables[0]!;
  return wordRowExceptionProperties(
    resolver.get(project(model.getNode(table.nodeId)!.childNodeIds[index]!)),
  );
}
test("complete row exceptions are editable, isolated and native through reviewed reopens", () => {
  const session = createDocumentSession(
    openDocument(
      importDocx(rowExceptionFixture(3, true), {
        documentId: "full-row-ex",
        unsupportedContent: "reject",
      }),
      { kind: "document" },
    ),
  );
  const table = session.snapshot().metadata.word!.tables[0]!,
    root = session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
  const before = session.snapshot(),
    original = rowAppearance(session);
  expect(original.width).toBe("160pt");
  expect(original["background-color"]).toBe("#ffeedd");
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "full-row-ex",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  const stale = session.prepare({
    protocolVersion: 1,
    documentId: "full-row-ex",
    baseRevision: 0,
    operations: [
      {
        op: "formatWordTable",
        nodeId: table.nodeId,
        rowIndex: 0,
        rowExceptionProperties: { width: "120pt" },
      },
    ],
  });
  apply([
    { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
    {
      op: "formatWordTable",
      nodeId: table.nodeId,
      rowIndex: 0,
      rowExceptionProperties: { width: "120pt", "background-color": "#112233", padding: "7pt" },
    },
  ]);
  expect(() => session.commit(stale)).toThrowError(
    expect.objectContaining({ code: "RevisionConflict" }),
  );
  expect(rowAppearance(session, 1)).toEqual(original);
  expect(rowAppearance(session).width).toBe("120pt");
  const bytes = exportDocx(session).bytes;
  expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).toContain("<w:tblPrExChange");
  for (const decision of ["accept", "reject"] as const) {
    const reviewer = createDocumentSession(
      openDocument(importDocx(bytes, { documentId: "row-review", unsupportedContent: "reject" }), {
        kind: "document",
      }),
    );
    reviewer.commit(
      reviewer.prepare({
        protocolVersion: 1,
        documentId: "row-review",
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
    expect(rowAppearance(reviewer).width).toBe(decision === "accept" ? "120pt" : "160pt");
    let output = exportDocx(reviewer).bytes;
    for (let pass = 0; pass < 2; pass++) {
      const reopened = openDocument(
        importDocx(output, { documentId: "row-reviewed", unsupportedContent: "reject" }),
        { kind: "document" },
      );
      expect(rowAppearance(reopened).width).toBe(decision === "accept" ? "120pt" : "160pt");
      output = exportDocx(reopened).bytes;
    }
  }
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
  expect(() =>
    apply([
      {
        op: "formatWordTable",
        nodeId: table.nodeId,
        rowIndex: 999,
        rowExceptionProperties: { width: "1pt" },
      },
    ]),
  ).toThrow();
  expect(session.snapshot().html).toBe(before.html);
});

test("alignment exceptions retain inactive indentation without overriding centered placement", () => {
  const parts = unzipSync(rowExceptionFixture(3, true));
  parts["word/document.xml"] = strToU8(
    strFromU8(parts["word/document.xml"]!).replaceAll(
      '<w:jc w:val="center"/>',
      '<w:jc w:val="center"/><w:tblInd w:w="300" w:type="dxa"/>',
    ),
  );
  let bytes: Uint8Array = zipSync(parts);
  for (let pass = 0; pass < 3; pass++) {
    const model = openDocument(
      importDocx(bytes, { documentId: "row-indent", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(rowAppearance(model).indent).toBe("15pt");
    expect(rowAppearance(model)["margin-left"]).toBe("auto");
    bytes = exportDocx(model).bytes;
    expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).toContain(
      '<w:tblInd w:type="dxa" w:w="300"/>',
    );
  }
});
