import { expect, test } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { SaxesParser } from "saxes";
import {
  createDocumentSession,
  openDocument,
  createAssetDescriptor,
  WordNativeFragmentCollectionSchema,
} from "@doffi/core";
import type { WordContentSemantics, DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { nativeCommentFixture } from "./native-comment-fixture.js";

function opaqueGraph(
  placement: Parameters<typeof nativeCommentFixture>[0],
  inline: boolean,
  outsideReference = false,
) {
  let input = nativeCommentFixture(placement, inline);
  if (outsideReference) {
    const files = unzipSync(input),
      reference = '<w:r><w:commentReference w:id="7"/></w:r>';
    files["word/document.xml"] = strToU8(
      strFromU8(files["word/document.xml"]!)
        .replace(reference, "")
        .replace("</w:customXml>", "</w:customXml>" + reference),
    );
    input = zipSync(files);
  }
  const source = openDocument(importDocx(input, { documentId: "opaque-copy-source" }), {
      kind: "document",
    }),
    snapshot = source.snapshot(),
    word = snapshot.metadata.word!,
    native = word.nativeContent!,
    atom = native.fragments.find((f) => f.kind !== "property")!;
  const collection = WordNativeFragmentCollectionSchema.parse(
      JSON.parse(new TextDecoder().decode(source.serialize().get(atom.assetPath)!)),
    ),
    bytes = new TextEncoder().encode(
      JSON.stringify({
        ...collection,
        fragments: { [atom.fragmentKey!]: collection.fragments[atom.fragmentKey!] },
      }),
    ),
    descriptor = createAssetDescriptor("assets/selected-native.json", "application/json", bytes);
  const annotation = JSON.parse(JSON.stringify(snapshot.annotations[0]!));
  // Only the selected native atom crosses the clipboard. Source endpoints in
  // adjacent paragraphs are intentionally clipped to that exact object.
  annotation.anchor = { kind: "node", nodeId: atom.nodeId! };
  const semantics: WordContentSemantics = {
    nativeContent: {
      mainPart: native.mainPart,
      fragments: [{ ...JSON.parse(JSON.stringify(atom)), assetPath: descriptor.path }],
      parts: [],
      relationships: [],
      identities: native.identities?.filter((i) => i.nodeId === atom.nodeId).map((i) => ({ ...i })),
    },
    annotations: [annotation],
    comments: JSON.parse(JSON.stringify(word.comments)),
    stories: word.stories
      .filter((s) => s.kind === "comment")
      .map((s) => ({ ...s, kind: "comment", html: source.getNode(s.rootNodeId)!.html })),
  };
  return { html: source.getNode(atom.nodeId!)!.html, semantics, assets: [{ descriptor, bytes }] };
}
function controls(bytes: Uint8Array) {
  const controls: { id: string; kind: string; native: boolean }[] = [],
    stack: string[] = [],
    parser = new SaxesParser({ xmlns: true });
  parser.on("opentag", (tag) => {
    stack.push(tag.local);
    if (["commentRangeStart", "commentRangeEnd", "commentReference"].includes(tag.local))
      controls.push({
        kind: tag.local,
        id: Object.values(tag.attributes).find((a) => a.local === "id")!.value,
        native: stack.includes("customXml"),
      });
  });
  parser.on("closetag", () => stack.pop());
  parser.write(strFromU8(unzipSync(bytes)["word/document.xml"]!)).close();
  return controls;
}
for (const inline of [false, true])
  for (const placement of ["inside", "start", "end", "point"] as const)
    test(
      "native clipboard comments retain contextual controls through source ID collisions " +
        inline +
        " " +
        placement,
      () => {
        const graph = opaqueGraph(placement, inline),
          documentId = "opaque-copy-target",
          session = createDocumentSession(
            openDocument(importDocx(nativeCommentFixture("inside", false), { documentId }), {
              kind: "document",
            }),
          ),
          target = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
            .map((m) => session.getNode(m[1]!)!)
            .find((n) => n.tagName === "p" && n.text === "AFTER")!.nodeId,
          before = session.snapshot();
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
        for (let pass = 0; pass < 2; pass++)
          apply([
            {
              op: inline ? "pasteWordInlineContent" : "pasteWordBlockContent",
              nodeId: target,
              from: 0,
              to: 0,
              expectedText: "",
              html: graph.html,
              semantics: graph.semantics,
            },
          ]);
        expect(session.snapshot().annotations).toHaveLength(3);
        let bytes = exportDocx(session).bytes;
        for (let pass = 0; pass < 3; pass++) {
          const markers = controls(bytes),
            references = markers.filter((c) => c.kind === "commentReference");
          expect(references).toHaveLength(3);
          expect(references.every((c) => c.native)).toBe(true);
          expect(new Set(references.map((c) => c.id)).size).toBe(3);
          for (const reference of references) {
            expect(
              markers.filter((c) => c.id === reference.id && c.kind === "commentRangeStart").length,
            ).toBeLessThanOrEqual(1);
            expect(
              markers.filter((c) => c.id === reference.id && c.kind === "commentRangeEnd").length,
            ).toBeLessThanOrEqual(1);
          }
          const reopened = openDocument(importDocx(bytes, { documentId }), { kind: "document" });
          expect(reopened.snapshot().annotations).toHaveLength(3);
          expect(
            new Set(
              reopened.snapshot().annotations.map((a) => reopened.getNode(a.bodyNodeId!)!.text),
            ),
          ).toEqual(new Set(["REVIEW"]));
          bytes = exportDocx(reopened).bytes;
        }
        session.undo();
        session.undo();
        expect(session.snapshot().html).toBe(before.html);
        expect(session.snapshot().annotations).toHaveLength(1);
      },
    );

test("removing copied opaque endpoints releases only their selected native bytes and undo restores them", () => {
  const graph = opaqueGraph("inside", true),
    documentId = "opaque-resource-target",
    session = createDocumentSession(
      openDocument(importDocx(nativeCommentFixture("inside", false), { documentId }), {
        kind: "document",
      }),
    ),
    before = session.snapshot();
  const target = [...before.html.matchAll(/data-node-id="([^"]+)"/g)]
    .map((m) => session.getNode(m[1]!)!)
    .find((n) => n.tagName === "p" && n.text === "AFTER")!.nodeId;
  session.commit(
    session.prepare(
      {
        protocolVersion: 1,
        documentId,
        baseRevision: 0,
        operations: [
          {
            op: "pasteWordInlineContent",
            nodeId: target,
            from: 0,
            to: 0,
            expectedText: "",
            html: graph.html,
            semantics: graph.semantics,
          },
        ],
      },
      { assets: graph.assets },
    ),
  );
  const old = new Set(before.metadata.word!.nativeContent!.fragments.map((f) => f.nodeId)),
    copied = session
      .snapshot()
      .metadata.word!.nativeContent!.fragments.find(
        (f) => f.kind !== "property" && !old.has(f.nodeId),
      )!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId,
      baseRevision: session.snapshot().manifest.revision,
      operations: [{ op: "removeNode", nodeId: copied.nodeId! }],
    }),
  );
  expect(session.serialize().has(graph.assets[0]!.descriptor.path)).toBe(false);
  expect(
    session.snapshot().metadata.word!.nativeContent!.fragments.filter((f) => f.kind !== "property"),
  ).toHaveLength(1);
  expect(
    session.getAnnotations().find((a) => a.annotationId !== before.annotations[0]!.annotationId)!
      .status,
  ).toBe("needs_reanchor");
  session.undo();
  expect(session.serialize().get(graph.assets[0]!.descriptor.path)).toEqual(graph.assets[0]!.bytes);
});

test("tracked opaque blocks reject atomically until their full native structural transfer is available", () => {
  const graph = opaqueGraph("inside", false),
    documentId = "opaque-block-reject",
    session = createDocumentSession(
      openDocument(importDocx(nativeCommentFixture("inside", false), { documentId }), {
        kind: "document",
      }),
    );
  const target = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
    .map((m) => session.getNode(m[1]!)!)
    .find((n) => n.tagName === "p" && n.text === "AFTER")!.nodeId;
  let root = session.getNode(target)!;
  while (root.parentNodeId) root = session.getNode(root.parentNodeId)!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId,
      baseRevision: 0,
      operations: [{ op: "setWordReview", nodeId: root.nodeId, trackChanges: true }],
    }),
  );
  const before = session.snapshot();
  expect(() =>
    session.prepare(
      {
        protocolVersion: 1,
        documentId,
        baseRevision: before.manifest.revision,
        operations: [
          {
            op: "pasteWordBlockContent",
            nodeId: target,
            from: 0,
            to: 0,
            expectedText: "",
            html: graph.html,
            semantics: graph.semantics,
          },
        ],
      },
      { assets: graph.assets },
    ),
  ).toThrow(/追踪粘贴/);
  expect(session.snapshot()).toEqual(before);
});

for (const textAnchor of [false, true])
  test(
    "native clipboard retains a selected internal range when its reference is outside the selection " +
      textAnchor,
    () => {
      const graph = opaqueGraph("inside", true, true),
        documentId = "native-outside-reference",
        session = createDocumentSession(
          openDocument(importDocx(nativeCommentFixture("inside", false), { documentId }), {
            kind: "document",
          }),
        );
      const target = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
        .map((m) => session.getNode(m[1]!)!)
        .find((n) => n.tagName === "p" && n.text === "AFTER")!.nodeId;
      if (textAnchor) {
        graph.html = '<span data-node-id="selected-container">' + graph.html + "</span>";
        graph.semantics.annotations![0]!.anchor = {
          kind: "text",
          nodeId: "selected-container",
          from: 0,
          to: 1,
          quote: "\ufffc",
        };
      }
      session.commit(
        session.prepare(
          {
            protocolVersion: 1,
            documentId,
            baseRevision: 0,
            operations: [
              {
                op: "pasteWordInlineContent",
                nodeId: target,
                from: 0,
                to: 0,
                expectedText: "",
                html: graph.html,
                semantics: graph.semantics,
              },
            ],
          },
          { assets: graph.assets },
        ),
      );
      const bytes = exportDocx(session).bytes,
        markers = controls(bytes),
        nativeStarts = markers.filter((c) => c.native && c.kind === "commentRangeStart"),
        nativeEnds = markers.filter((c) => c.native && c.kind === "commentRangeEnd");
      expect(nativeStarts).toHaveLength(2);
      expect(nativeEnds).toHaveLength(2);
      expect(markers.filter((c) => c.kind === "commentReference")).toHaveLength(2);
      expect(
        openDocument(importDocx(bytes, { documentId }), { kind: "document" }).snapshot()
          .annotations,
      ).toHaveLength(2);
    },
  );
