import { expect, test } from "vitest";
import {
  createDocumentSession,
  createWordDocument,
  createAssetDescriptor,
  openDocument,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
  WordStyleSchema,
} from "@doffi/core";
import type { DocumentOperation, ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { tableStylesFixture } from "./table-style-fixture.js";
function fixture() {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "reset" }), { kind: "document" }),
  );
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "reset",
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
      text: "前禰󠄀😀后",
    },
    {
      op: "setWordStyle",
      nodeId: "document-root",
      style: { styleId: "A", kind: "paragraph", cssClass: "p-a", basedOn: "Normal" },
      runProperties: { color: "#aa0000", "font-size": "14pt" },
    },
    {
      op: "setWordStyle",
      nodeId: "document-root",
      style: { styleId: "B", kind: "paragraph", cssClass: "p-b", basedOn: "Normal" },
      runProperties: { color: "#0000aa", "font-size": "18pt" },
    },
    {
      op: "setWordStyle",
      nodeId: "document-root",
      style: { styleId: "C", kind: "character", cssClass: "c" },
      runProperties: { color: "#00aa00", "font-style": "italic" },
    },
    { op: "setWordParagraph", nodeId: "paragraph-1", styleId: "A" },
    {
      op: "applyWordCharacterStyle",
      nodeId: "paragraph-1",
      from: 0,
      to: 7,
      expectedText: "前禰󠄀😀后",
      styleId: "C",
    },
    {
      op: "formatWordText",
      nodeId: "paragraph-1",
      from: 0,
      to: 7,
      expectedText: "前禰󠄀😀后",
      properties: { "font-size": "32pt" },
    },
  ]);
  return { session, apply };
}
function runs(
  model: ContentDocumentModel,
): { text: string; css: Readonly<Record<string, string>> }[] {
  const resolver = new SourceStyleResolver(
      model.snapshot().styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    result: { text: string; css: Readonly<Record<string, string>> }[] = [];
  const visit = (id: string, parent?: SourceStyleNode): void => {
    const node = model.getNode(id)!,
      projected = { tag: node.tagName, attrs: node.attributes, parent };
    if (!node.childNodeIds.length && node.text)
      result.push({ text: node.text, css: resolver.get(projected) });
    for (const child of node.childNodeIds) visit(child, projected);
  };
  visit(model.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId);
  return result;
}
test("cleared text follows later paragraph style changes while later direct edits retain only their own overrides", () => {
  const { session, apply } = fixture();
  apply([{ op: "resetWordText", nodeId: "paragraph-1", from: 1, to: 6, expectedText: "禰󠄀😀" }]);
  expect(runs(session).find((run) => run.text === "禰󠄀😀")!.css).toMatchObject({
    color: "#aa0000",
    "font-size": "14pt",
    "font-style": "normal",
  });
  apply([{ op: "setWordParagraph", nodeId: "paragraph-1", styleId: "B" }]);
  expect(runs(session).find((run) => run.text === "禰󠄀😀")!.css).toMatchObject({
    color: "#0000aa",
    "font-size": "18pt",
  });
  apply([
    {
      op: "formatWordText",
      nodeId: "paragraph-1",
      from: 1,
      to: 6,
      expectedText: "禰󠄀😀",
      properties: { color: "#771177" },
    },
    {
      op: "setWordStyle",
      nodeId: "document-root",
      style: WordStyleSchema.parse(
        session.snapshot().metadata.word!.styles.find((style) => style.styleId === "B")!,
      ),
      runProperties: { color: "#aa5500", "font-size": "20pt" },
    },
  ]);
  expect(runs(session).find((run) => run.text === "禰󠄀😀")!.css).toMatchObject({
    color: "#771177",
    "font-size": "20pt",
    "font-style": "normal",
  });
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const reopened = openDocument(
      importDocx(bytes, { documentId: "native", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    const selected = runs(reopened).filter((run) => /禰|😀/.test(run.text));
    expect(selected.map((run) => run.text).join("")).toBe("禰󠄀😀");
    for (const run of selected)
      expect(run.css).toMatchObject({
        color: "#771177",
        "font-size": "20pt",
        "font-style": "normal",
      });
    bytes = exportDocx(reopened).bytes;
  }
  session.undo();
  expect(runs(session).find((run) => run.text === "禰󠄀😀")!.css).toMatchObject({
    color: "#0000aa",
    "font-size": "18pt",
  });
});
test("tracked clearing preserves the earliest font and character style, records the latest author and supports native rejection", () => {
  const { session, apply } = fixture();
  apply([
    { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "甲" },
    { op: "resetWordText", nodeId: "paragraph-1", from: 1, to: 6, expectedText: "禰󠄀😀" },
  ]);
  expect(session.snapshot().metadata.word!.revisions).toHaveLength(1);
  expect(session.snapshot().metadata.word!.revisions[0]!.before!.characterStyleId).toBe("C");
  apply([
    { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "乙" },
    {
      op: "formatWordText",
      nodeId: "paragraph-1",
      from: 1,
      to: 6,
      expectedText: "禰󠄀😀",
      properties: { "font-weight": "700" },
    },
    {
      op: "applyWordCharacterStyle",
      nodeId: "paragraph-1",
      from: 1,
      to: 6,
      expectedText: "禰󠄀😀",
      styleId: "C",
    },
    { op: "resetWordText", nodeId: "paragraph-1", from: 1, to: 6, expectedText: "禰󠄀😀" },
  ]);
  const revision = session.snapshot().metadata.word!.revisions[0]!;
  expect(revision.author).toBe("乙");
  expect(revision.before!.characterStyleId).toBe("C");
  const bytes = exportDocx(session).bytes,
    xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
  expect(xml).toContain("rPrChange");
  expect(xml).toContain("w:rStyle");
  const reopened = createDocumentSession(
      openDocument(importDocx(bytes, { documentId: "review", unsupportedContent: "reject" }), {
        kind: "document",
      }),
    ),
    history = reopened.snapshot().metadata.word!.revisions[0]!;
  for (const run of runs(reopened).filter((run) => /禰|😀/.test(run.text)))
    expect(run.css).toMatchObject({ "font-size": "14pt", "font-style": "normal" });
  const nodes = [
    reopened.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
  ];
  for (let index = 0; index < nodes.length; index++)
    nodes.push(...reopened.getNode(nodes[index]!)!.childNodeIds);
  expect(
    nodes
      .filter(
        (id) =>
          reopened.getNode(id)!.attributes["data-word-run"] &&
          reopened.getNode(id)!.text === "禰󠄀😀",
      )
      .map((id) => reopened.getNode(id)!.attributes["data-word-style"]),
  ).toEqual([undefined]);
  reopened.commit(
    reopened.prepare({
      protocolVersion: 1,
      documentId: "review",
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
  for (const run of runs(reopened).filter((run) => /禰|😀/.test(run.text)))
    expect(run.css).toMatchObject({
      "font-size": "32pt",
      "font-style": "italic",
      color: "#00aa00",
    });
  const rejected = openDocument(
    importDocx(exportDocx(reopened).bytes, {
      documentId: "rejected",
      unsupportedContent: "reject",
    }),
    { kind: "document" },
  );
  for (const run of runs(rejected).filter((run) => /禰|😀/.test(run.text)))
    expect(run.css).toMatchObject({
      "font-size": "32pt",
      "font-style": "italic",
      color: "#00aa00",
    });
  apply([
    {
      op: "reviewWordRevision",
      nodeId: revision.nodeId,
      revisionId: revision.revisionId!,
      decision: "reject",
    },
  ]);
  expect(runs(session).find((run) => run.text === "禰󠄀😀")!.css).toMatchObject({
    "font-size": "32pt",
    "font-style": "italic",
    color: "#00aa00",
  });
  const before = session.snapshot();
  expect(() =>
    apply([{ op: "resetWordText", nodeId: "paragraph-1", from: 2, to: 6, expectedText: "󠄀😀" }]),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});

for (const previousSize of [32, 24])
  test(
    "old whole native run properties cannot restore cleared known formatting while opaque properties and original history survive " +
      previousSize,
    () => {
      const { session: initial } = fixture(),
        native = unzipSync(exportDocx(initial).bytes),
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
      native["word/document.xml"] = strToU8(
        strFromU8(native["word/document.xml"]!).replace(
          "<w:r><w:rPr>",
          '<w:r><w:rPr><w:effect w:val="shimmer"/>',
        ),
      );
      const xml = strFromU8(native["word/document.xml"]!),
        model = openDocument(importDocx(zipSync(native), { documentId: "old-run" }), {
          kind: "document",
        }),
        snapshot = model.snapshot(),
        files = new Map(model.serialize()),
        metadata = JSON.parse(JSON.stringify(snapshot.metadata));
      const fragment = metadata.word.nativeContent.fragments.find(
          (item: { propertyName?: string }) => item.propertyName === namespace + " effect",
        ),
        path = fragment.assetPath,
        collection = JSON.parse(strFromU8(files.get(path)!));
      collection.fragments[fragment.fragmentKey] = xml
        .match(/<w:r><w:rPr>([\s\S]*?)<\/w:rPr>/)![0]
        .slice(5)
        .replace("<w:rPr>", '<w:rPr xmlns:w="' + namespace + '">');
      fragment.propertyPath = [];
      fragment.propertyName = namespace + " rPr";
      delete fragment.propertyIndices;
      delete fragment.propertyIndex;
      const bytes = strToU8(JSON.stringify(collection)),
        manifest = JSON.parse(JSON.stringify(snapshot.manifest));
      manifest.assets = manifest.assets.map((asset: { path: string; mimeType: string }) =>
        asset.path === path ? createAssetDescriptor(path, asset.mimeType, bytes) : asset,
      );
      files.set(path, bytes);
      files.set("manifest.json", strToU8(JSON.stringify(manifest)));
      files.set(snapshot.manifest.metadata, strToU8(JSON.stringify(metadata)));
      const session = createDocumentSession(openDocument(files, { kind: "document" })),
        paragraph = session.getNode(session.snapshot().metadata.sections![0]!.rootNodeId)!
          .childNodeIds[0]!,
        root = session.getNode(
          session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!
            .rootNodeId,
        )!.parentNodeId!;
      if (previousSize !== 32)
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: "old-run",
            baseRevision: 0,
            operations: [
              {
                op: "formatWordText",
                nodeId: paragraph,
                from: 0,
                to: 7,
                expectedText: "前禰󠄀😀后",
                properties: { "font-size": previousSize + "pt" },
              },
            ],
          }),
        );
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "old-run",
          baseRevision: session.snapshot().manifest.revision,
          operations: [
            { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
            { op: "resetWordText", nodeId: paragraph, from: 0, to: 7, expectedText: "前禰󠄀😀后" },
          ],
        }),
      );
      const output = exportDocx(session).bytes;
      expect(strFromU8(unzipSync(output)["word/document.xml"]!)).toContain('w:val="shimmer"');
      const reopened = createDocumentSession(
        openDocument(importDocx(output, { documentId: "old-run-review" }), { kind: "document" }),
      );
      for (const run of runs(reopened))
        expect(run.css).toMatchObject({
          color: "#aa0000",
          "font-size": "14pt",
          "font-style": "normal",
        });
      const reviewRoot = reopened.getNode(
        reopened.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!
          .rootNodeId,
      )!.parentNodeId!;
      reopened.commit(
        reopened.prepare({
          protocolVersion: 1,
          documentId: "old-run-review",
          baseRevision: 0,
          operations: [
            { op: "reviewWordRevisions", nodeId: reviewRoot, author: "A", decision: "reject" },
          ],
        }),
      );
      for (const run of runs(reopened))
        expect(run.css).toMatchObject({
          color: "#00aa00",
          "font-size": previousSize + "pt",
          "font-style": "italic",
        });
      const rejectedBytes = exportDocx(reopened).bytes;
      expect(strFromU8(unzipSync(rejectedBytes)["word/document.xml"]!)).toContain('w:val="shimmer"');
      const rejected = openDocument(importDocx(rejectedBytes, { documentId: "old-run-rejected" }), {
        kind: "document",
      });
      for (const run of runs(rejected))
        expect(run.css).toMatchObject({
          color: "#00aa00",
          "font-size": previousSize + "pt",
          "font-style": "italic",
        });
    },
  );

test("a cleared unsupported underline stays in original native history and returns after rejecting and exporting again", () => {
  const { session: initial } = fixture(),
    native = unzipSync(exportDocx(initial).bytes);
  native["word/document.xml"] = strToU8(
    strFromU8(native["word/document.xml"]!).replace(
      "<w:r><w:rPr>",
      '<w:r><w:rPr><w:u w:val="dashedHeavy"/>',
    ),
  );
  const session = createDocumentSession(
      openDocument(importDocx(zipSync(native), { documentId: "opaque-underline" }), {
        kind: "document",
      }),
    ),
    paragraph = session.getNode(session.snapshot().metadata.sections![0]!.rootNodeId)!
      .childNodeIds[0]!,
    root = session.getNode(
      session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
    )!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "opaque-underline",
      baseRevision: 0,
      operations: [
        { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
        { op: "resetWordText", nodeId: paragraph, from: 0, to: 7, expectedText: "前禰󠄀😀后" },
      ],
    }),
  );
  const output = exportDocx(session).bytes,
    xml = strFromU8(unzipSync(output)["word/document.xml"]!);
  expect(xml).toMatch(/<w:rPrChange\b[\s\S]*dashedHeavy[\s\S]*<\/w:rPrChange>/);
  const reopened = createDocumentSession(
      openDocument(importDocx(output, { documentId: "underline-review" }), { kind: "document" }),
    ),
    reviewRoot = reopened.getNode(
      reopened.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
    )!.parentNodeId!;
  reopened.commit(
    reopened.prepare({
      protocolVersion: 1,
      documentId: "underline-review",
      baseRevision: 0,
      operations: [
        { op: "reviewWordRevisions", nodeId: reviewRoot, author: "A", decision: "reject" },
      ],
    }),
  );
  const rejected = strFromU8(unzipSync(exportDocx(reopened).bytes)["word/document.xml"]!);
  expect(rejected).toContain("dashedHeavy");
  expect(rejected).not.toContain("rPrChange");
});

test("cleared native cell runs follow conditional table styles and their new row position", () => {
  const session = createDocumentSession(
    openDocument(
      importDocx(tableStylesFixture(), { documentId: "reset-table", unsupportedContent: "reject" }),
      { kind: "document" },
    ),
  );
  const nodes = [
    session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
  ];
  for (let index = 0; index < nodes.length; index++)
    nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
  const paragraph = nodes.find(
      (id) => session.getNode(id)!.tagName === "p" && session.getNode(id)!.text === "R0C1",
    )!,
    table = session.snapshot().metadata.word!.tables[0]!.nodeId;
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "reset-table",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  apply([{ op: "resetWordText", nodeId: paragraph, from: 0, to: 4, expectedText: "R0C1" }]);
  expect(runs(session).find((run) => run.text === "R0C1")!.css).toMatchObject({
    color: "#cc0000",
    "text-decoration-line": "underline",
  });
  expect(
    ["bold", "700"].includes(runs(session).find((run) => run.text === "R0C1")!.css["font-weight"]!),
  ).toBe(true);
  apply([{ op: "formatWordTable", nodeId: table, styleId: "Blue" }]);
  expect(runs(session).find((run) => run.text === "R0C1")!.css.color).toBe("#0000cc");
  apply([{ op: "editWordTable", nodeId: table, change: { kind: "insertRows", at: 0, count: 1 } }]);
  expect(runs(session).find((run) => run.text === "R0C1")!.css).toMatchObject({
    color: "#0000cc",
    "text-decoration-line": "none",
  });
  expect(
    ["bold", "700"].includes(runs(session).find((run) => run.text === "R0C1")!.css["font-weight"]!),
  ).toBe(false);
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const reopened = openDocument(
      importDocx(bytes, { documentId: "table-native", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(runs(reopened).find((run) => run.text === "R0C1")!.css).toMatchObject({
      color: "#0000cc",
      "text-decoration-line": "none",
    });
    expect(
      ["bold", "700"].includes(
        runs(reopened).find((run) => run.text === "R0C1")!.css["font-weight"]!,
      ),
    ).toBe(false);
    bytes = exportDocx(reopened).bytes;
  }
  session.undo();
  expect(
    ["bold", "700"].includes(runs(session).find((run) => run.text === "R0C1")!.css["font-weight"]!),
  ).toBe(true);
});
