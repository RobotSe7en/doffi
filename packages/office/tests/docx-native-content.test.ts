import { body, mutateDocx } from "./regression-fixtures.js";
import { expect, test } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import {
  createAssetDescriptor,
  createDocumentSession,
  createWordDocument,
  openDocument,
} from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import {
  wordNativeContentFixture,
  nativeContentGraph,
  STORE_ID,
} from "./word-native-content-fixture.js";

test("native data binding indexes only its selected datastore dependency and preserves its properties through native saves", () => {
  let model = openDocument(
    importDocx(wordNativeContentFixture(true), { documentId: "bound-source" }),
    { kind: "document" },
  );
  for (let pass = 0; pass < 3; pass++) {
    const native = model.snapshot().metadata.word!.nativeContent!,
      binding = native.fragments.find((f) => f.dependencyParts?.length)!;
    expect(binding, JSON.stringify(native.fragments)).toBeDefined();
    expect(binding.dependencyParts).toHaveLength(1);
    const part = native.parts.find((p) => p.sourcePart === binding.dependencyParts![0])!;
    expect(part.customXml?.storeItemId).toBe(STORE_ID);
    expect(part.customXml?.propertiesPart).toBeDefined();
    const output = exportDocx(model),
      files = unzipSync(output.bytes);
    expect(strFromU8(files["word/document.xml"]!)).toContain('storeItemID="' + STORE_ID + '"');
    model = openDocument(importDocx(output.bytes, { documentId: "bound-source" }), {
      kind: "document",
    });
  }
});
for (const decision of ["accept", "reject"] as const)
  test(
    "native tracked blocks remain reviewable after repeated native saves and " + decision,
    () => {
      const documentId = "native-reopen-review",
        graph = nativeContentGraph(),
        session = createDocumentSession(
          openDocument(createWordDocument({ documentId }), { kind: "document" }),
        );
      const apply = (operations: DocumentOperation[]) =>
        session.commit(
          session.prepare(
            {
              protocolVersion: 1,
              documentId,
              baseRevision: session.snapshot().manifest.revision,
              operations,
            },
            { assets: graph.assets },
          ),
        );
      apply([
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "Receiver" },
        {
          op: "pasteWordBlockContent",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          html: graph.html,
          semantics: graph.semantics,
        },
      ]);
      let output = exportDocx(session);
      expect(output.report.unexportedNativeRevisionIds).toBeUndefined();
      for (let pass = 0; pass < 3; pass++) {
        const reopened = createDocumentSession(
            openDocument(importDocx(output.bytes, { documentId }), { kind: "document" }),
          ),
          revision = reopened
            .snapshot()
            .metadata.word!.revisions.find((r) => r.kind === "nativeInsert")!;
        expect(revision).toBeDefined();
        expect(revision.author).toBe("Receiver");
        const native = strFromU8(unzipSync(output.bytes)["word/document.xml"]!);
        expect(native).toContain("SELECTED-NATIVE");
        expect(native).toContain("SELECTED-CELL");
        expect(native).not.toContain("PRIVATE-NATIVE-OUTSIDE");
        const unchanged = exportDocx(reopened);
        reopened.commit(
          reopened.prepare({
            protocolVersion: 1,
            documentId,
            baseRevision: 0,
            operations: [
              {
                op: "reviewWordRevisions",
                nodeId: (() => {
                  let n = reopened.getNode(revision.nodeId)!;
                  while (n.parentNodeId) n = reopened.getNode(n.parentNodeId)!;
                  return n.nodeId;
                })(),
                decision,
              },
            ],
          }),
        );
        const reviewed = exportDocx(reopened),
          reviewedXml = strFromU8(unzipSync(reviewed.bytes)["word/document.xml"]!);
        expect(reviewedXml.includes("SELECTED-NATIVE")).toBe(decision === "accept");
        expect(reviewedXml).not.toMatch(/<w:ins[\s>]/);
        expect(
          openDocument(importDocx(reviewed.bytes, { documentId }), { kind: "document" }).snapshot()
            .metadata.word!.revisions,
        ).toHaveLength(0);
        reopened.undo();
        expect(
          reopened.snapshot().metadata.word!.revisions.some((r) => r.kind === "nativeInsert"),
        ).toBe(true);
        output = unchanged;
      }
    },
  );

test("an opaque insertion without native run or row structure preserves bytes and explicitly reports its unrepresentable native revision", () => {
  const graph = nativeContentGraph(),
    documentId = "native-empty-insertion",
    session = createDocumentSession(
      openDocument(createWordDocument({ documentId }), { kind: "document" }),
    ),
    resource = graph.assets[0]!,
    collection = JSON.parse(new TextDecoder().decode(resource.bytes));
  collection.fragments[graph.semantics.nativeContent!.fragments[0]!.fragmentKey!] =
    '<w:customXml xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:element="empty" w:uri="urn:doffi:native-copy"/>';
  const bytes = new TextEncoder().encode(JSON.stringify(collection));
  session.commit(
    session.prepare(
      {
        protocolVersion: 1,
        documentId,
        baseRevision: 0,
        operations: [
          { op: "setWordReview", nodeId: "document-root", trackChanges: true },
          {
            op: "pasteWordBlockContent",
            nodeId: "paragraph-1",
            from: 0,
            to: 0,
            expectedText: "",
            html: graph.html,
            semantics: graph.semantics,
          },
        ],
      },
      {
        assets: [
          {
            descriptor: createAssetDescriptor(resource.descriptor.path, "application/json", bytes),
            bytes,
          },
        ],
      },
    ),
  );
  const revision = session
      .snapshot()
      .metadata.word!.revisions.find((r) => r.kind === "nativeInsert")!,
    exported = exportDocx(session);
  expect(exported.report.unexportedNativeRevisionIds).toEqual([revision.revisionId]);
  expect(strFromU8(unzipSync(exported.bytes)["word/document.xml"]!)).toContain('element="empty"');
  expect(
    openDocument(session.serialize(), { kind: "document" })
      .snapshot()
      .metadata.word!.revisions.some((r) => r.kind === "nativeInsert"),
  ).toBe(true);
});

test("a partly inserted opaque block never acquires a guessed whole-object inverse", () => {
  const input = mutateDocx("01-text-numbering.docx", {
      "word/document.xml": (source) =>
        body(
          source,
          '<w:customXml w:element="mixed" w:uri="urn:doffi:native-copy"><x:untracked xmlns:x="urn:doffi:native-extra">KEEP-UNTRACKED</x:untracked><w:p><w:pPr><w:rPr><w:ins w:id="8" w:author="Original"/></w:rPr></w:pPr><w:ins w:id="8" w:author="Original"><w:r><w:t>NEW-TEXT</w:t></w:r></w:ins></w:p></w:customXml>',
        ),
    }),
    model = openDocument(importDocx(input, { documentId: "native-mixed-insertion" }), {
      kind: "document",
    });
  expect(model.snapshot().metadata.word!.revisions.some((r) => r.kind === "nativeInsert")).toBe(
    false,
  );
  const output = strFromU8(unzipSync(exportDocx(model).bytes)["word/document.xml"]!);
  expect(output).toContain("KEEP-UNTRACKED");
  expect(output).toContain("NEW-TEXT");
  expect(output).toContain("ins");
});

test("opaque paragraph and content insertions must share native ownership before whole-object review is exposed", () => {
  const input = mutateDocx("01-text-numbering.docx", {
      "word/document.xml": (source) =>
        body(
          source,
          '<w:customXml w:element="mixed" w:uri="urn:doffi:native-copy"><w:p><w:pPr><w:rPr><w:ins w:id="8" w:author="Paragraph"/></w:rPr></w:pPr><w:ins w:id="9" w:author="Content"><w:r><w:t>KEEP-CONTENT-OWNER</w:t></w:r></w:ins></w:p></w:customXml>',
        ),
    }),
    model = openDocument(importDocx(input, { documentId: "native-different-insertions" }), {
      kind: "document",
    });
  expect(model.snapshot().metadata.word!.revisions.some((r) => r.kind === "nativeInsert")).toBe(
    false,
  );
  expect(strFromU8(unzipSync(exportDocx(model).bytes)["word/document.xml"]!)).toContain(
    "KEEP-CONTENT-OWNER",
  );
});
