import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { importDocx, exportDocx } from "@doffi/office";
import { unzipSync, strFromU8 } from "fflate";
import { rowGapFixture } from "./row-gap-fixture.js";

function current(alignment = false) {
  const session = createDocumentSession(
      openDocument(
        importDocx(rowGapFixture(alignment), {
          documentId: "gaps",
          unsupportedContent: alignment ? "preserve" : "reject",
        }),
        { kind: "document" },
      ),
    ),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "gaps",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  return { session, apply, table: session.snapshot().metadata.word!.tables[0]!.nodeId };
}
function reopen(session: ReturnType<typeof current>["session"]): void {
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(xml).toContain("w:gridBefore");
    expect(xml).toContain("w:gridAfter");
    expect(xml).toContain("w:wBefore");
    const model = openDocument(
      importDocx(bytes, { documentId: "gaps-native", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(model.snapshot().html).toContain('data-word-grid-gap="before"');
    expect(model.snapshot().html).toContain('data-word-grid-gap="after"');
    bytes = exportDocx(model).bytes;
  }
}
test("omitted row edges retain widths and editable text across two strict native reopens", () => {
  const { session, apply, table } = current(),
    nodes = [table];
  for (let index = 0; index < nodes.length; index++)
    nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
  const paragraph = nodes.find(
    (id) => session.getNode(id)!.tagName === "p" && session.getNode(id)!.text === "X",
  )!;
  expect(nodes.filter((id) => session.getNode(id)!.attributes["data-word-grid-gap"])).toHaveLength(
    2,
  );
  apply([
    { op: "replaceText", nodeId: paragraph, from: 0, to: 1, expectedText: "X", text: "EDIT" },
  ]);
  reopen(session);
  expect(strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!)).toContain(
    'w:w="25%"',
  );
  session.undo();
  expect(session.getNode(paragraph)!.text).toBe("X");
});
for (const change of [
  { kind: "insertRows", at: 1, count: 1 },
  { kind: "insertColumns", at: 0, count: 1 },
  { kind: "insertColumns", at: 3, count: 1 },
  { kind: "deleteColumns", at: 1, count: 1 },
  { kind: "splitCell", row: 0, column: 1, rows: 2, columns: 2 },
] as const)
  test("row omissions remain edge regions after " + JSON.stringify(change), () => {
    const { session, apply, table } = current(),
      before = session.snapshot();
    apply([{ op: "editWordTable", nodeId: table, change }]);
    reopen(session);
    session.undo();
    expect(session.snapshot().html).toBe(before.html);
    expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
  });
test("omission regions cannot acquire source text or be merged as editable cells", () => {
  const { session, apply, table } = current(),
    nodes = [table];
  for (let index = 0; index < nodes.length; index++)
    nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
  const gap = nodes.find((id) => session.getNode(id)!.attributes["data-word-grid-gap"])!,
    before = session.snapshot();
  expect(() =>
    apply([
      {
        op: "editWordTable",
        nodeId: table,
        change: { kind: "mergeCells", row: 1, column: 0, rows: 1, columns: 2 },
      },
    ]),
  ).toThrow();
  expect(() => apply([{ op: "replaceFragment", nodeId: gap, html: "<p>bad</p>" }])).toThrow();
  expect(session.snapshot()).toEqual(before);
});
test("resizing omitted column widths records native row preferences and rejection restores the original widths", () => {
  const { session, apply, table } = current(),
    root = session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
  apply([
    { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
    { op: "editWordTable", nodeId: table, change: { kind: "resize", grid: [60, 80, 76] } },
  ]);
  const xml = strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!);
  expect(xml).toMatch(/<w:wBefore[^>]*w:w="1200"/);
  expect(xml).toMatch(/<w:wAfter[^>]*w:w="1520"/);
  expect(xml).toContain("trPrChange");
  const reopened = createDocumentSession(
      openDocument(
        importDocx(exportDocx(session).bytes, {
          documentId: "gap-review",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    nativeRoot = reopened.getNode(
      reopened.snapshot().metadata.word!.stories[0]!.rootNodeId,
    )!.parentNodeId!;
  reopened.commit(
    reopened.prepare({
      protocolVersion: 1,
      documentId: "gap-review",
      baseRevision: 0,
      operations: [{ op: "reviewWordRevisions", nodeId: nativeRoot, decision: "reject" }],
    }),
  );
  const output = exportDocx(reopened).bytes,
    native = strFromU8(unzipSync(output)["word/document.xml"]!);
  expect(native).toMatch(/<w:wBefore[^>]*w:w="1440"/);
  expect(native).toContain('w:w="25%"');
  expect(native).not.toContain("PrChange");
  expect(reopened.snapshot().metadata.word!.tables[0]!.grid).toEqual([72, 72, 72]);
  expect(() =>
    importDocx(output, { documentId: "gap-rejected", unsupportedContent: "reject" }),
  ).not.toThrow();
});

test("preserved native row alignment remains before row history after a supported height edit", () => {
  const { session, apply, table } = current(true);
  apply([
    {
      op: "setWordReview",
      nodeId: session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!
        .parentNodeId!,
      trackChanges: true,
      author: "A",
    },
    { op: "editWordTable", nodeId: table, change: { kind: "resize", rowIndex: 1, rowHeight: 24 } },
  ]);
  const xml = strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!),
    row = xml.match(/<w:trPr>[^]*?<w:gridBefore[^]*?<\/w:trPr>/)![0];
  expect(row.indexOf("<w:jc")).toBeLessThan(row.indexOf("<w:trPrChange"));
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const model = openDocument(importDocx(bytes, { documentId: "native-row-alignment" }), {
      kind: "document",
    });
    expect(model.snapshot().metadata.word!.revisions).toHaveLength(1);
    bytes = exportDocx(model).bytes;
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(xml.match(/<w:trPrChange\b[^]*?<\/w:trPrChange>/)![0]).toContain("w:jc");
  }
  const reviewed = createDocumentSession(
      openDocument(importDocx(bytes, { documentId: "aligned-review" }), { kind: "document" }),
    ),
    root = reviewed.getNode(
      reviewed.snapshot().metadata.word!.stories[0]!.rootNodeId,
    )!.parentNodeId!;
  reviewed.commit(
    reviewed.prepare({
      protocolVersion: 1,
      documentId: "aligned-review",
      baseRevision: 0,
      operations: [{ op: "reviewWordRevisions", nodeId: root, decision: "reject" }],
    }),
  );
  const rejected = strFromU8(unzipSync(exportDocx(reviewed).bytes)["word/document.xml"]!);
  expect(rejected).toContain("w:jc");
  expect(rejected).not.toContain("trPrChange");
});
