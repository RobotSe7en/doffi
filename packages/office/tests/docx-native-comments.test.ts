import { expect, test } from "vitest";
import { SaxesParser } from "saxes";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { importDocx, exportDocx } from "@doffi/office";
import { nativeCommentFixture } from "./native-comment-fixture.js";
function controls(bytes: Uint8Array) {
  const items: { kind: string; id: string; native: boolean }[] = [],
    stack: string[] = [],
    parser = new SaxesParser({ xmlns: true });
  parser.on("opentag", (tag) => {
    stack.push(tag.local);
    if (["commentRangeStart", "commentRangeEnd", "commentReference"].includes(tag.local))
      items.push({
        kind: tag.local,
        id: Object.values(tag.attributes).find((attr) => attr.local === "id")!.value,
        native: stack.includes("customXml"),
      });
  });
  parser.on("closetag", () => stack.pop());
  parser.write(strFromU8(unzipSync(bytes)["word/document.xml"]!)).close();
  return items;
}
for (const inline of [false, true])
  for (const placement of ["inside", "start", "end", "point"] as const)
    test(`native ${inline ? "inline" : "block"} comment ${placement} imports and retains its original endpoints across current edits and reopens`, () => {
      const session = createDocumentSession(
          openDocument(
            importDocx(nativeCommentFixture(placement, inline), { documentId: "native-comment" }),
            { kind: "document" },
          ),
        ),
        snapshot = session.snapshot(),
        annotation = snapshot.annotations[0]!,
        atom = snapshot.metadata.word!.nativeContent!.fragments.find(
          (item) => item.kind !== "property",
        )!.nodeId!,
        binding = snapshot.metadata.word!.comments![0]!;
      expect(binding.referenceNodeId).toBe(atom);
      expect(session.getNode(atom)!.text).toBe("\ufffc");
      expect(session.getNode(annotation.bodyNodeId!)!.text).toBe("REVIEW");
      const root = snapshot.metadata.word!.stories[0]!.rootNodeId,
        nodes = [root];
      for (let i = 0; i < nodes.length; i++)
        nodes.push(...session.getNode(nodes[i]!)!.childNodeIds);
      const run = nodes.find(
        (id) =>
          session.getNode(id)!.text === "AFTER" && session.getNode(id)!.childNodeIds.length === 0,
      )!;
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "native-comment",
          baseRevision: 0,
          operations: [
            { op: "replaceText", nodeId: run, from: 5, to: 5, expectedText: "", text: "X" },
            {
              op: "updateAnnotation",
              annotationId: annotation.annotationId,
              changes: { status: "resolved" },
            },
          ],
        }),
      );
      let bytes = exportDocx(session).bytes;
      for (let pass = 0; pass < 3; pass++) {
        const markers = controls(bytes);
        expect(markers.filter((item) => item.kind === "commentRangeStart")).toHaveLength(
          placement === "point" ? 0 : 1,
        );
        expect(markers.filter((item) => item.kind === "commentRangeEnd")).toHaveLength(
          placement === "point" ? 0 : 1,
        );
        expect(markers.filter((item) => item.kind === "commentReference")).toHaveLength(1);
        expect(new Set(markers.map((item) => item.id)).size).toBe(1);
        expect(markers.find((item) => item.kind === "commentReference")!.native).toBe(true);
        if (placement !== "point") {
          expect(markers.find((item) => item.kind === "commentRangeStart")!.native).toBe(
            placement !== "end",
          );
          expect(markers.find((item) => item.kind === "commentRangeEnd")!.native).toBe(
            placement !== "start",
          );
        }
        const model = openDocument(importDocx(bytes, { documentId: "native-reopened" }), {
          kind: "document",
        });
        expect(model.snapshot().annotations).toHaveLength(1);
        expect(model.snapshot().annotations[0]!.status).toBe("resolved");
        expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).not.toContain(
          "urn:doffi:transient:native",
        );
        bytes = exportDocx(model).bytes;
      }
      session.undo();
      expect(session.getNode(run)!.text).toBe("AFTER");
      session.redo();
      expect(session.getNode(run)!.text).toBe("AFTERX");
    });
test("removing or reanchoring an opaque comment removes its old native controls atomically", () => {
  for (const action of ["remove", "reanchor"] as const) {
    const session = createDocumentSession(
        openDocument(importDocx(nativeCommentFixture(), { documentId: "native-comment-edit" }), {
          kind: "document",
        }),
      ),
      annotation = session.snapshot().annotations[0]!,
      root = session.snapshot().metadata.word!.stories[0]!.rootNodeId,
      nodes = [root];
    for (let i = 0; i < nodes.length; i++) nodes.push(...session.getNode(nodes[i]!)!.childNodeIds);
    const paragraph = nodes.find(
        (id) => session.getNode(id)!.tagName === "p" && session.getNode(id)!.text === "AFTER",
      )!,
      operation: DocumentOperation =
        action === "remove"
          ? { op: "removeAnnotation", annotationId: annotation.annotationId }
          : {
              op: "updateAnnotation",
              annotationId: annotation.annotationId,
              changes: { anchor: { kind: "text", nodeId: paragraph, from: 0, to: 2, quote: "AF" } },
            };
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "native-comment-edit",
        baseRevision: 0,
        operations: [operation],
      }),
    );
    const bytes = exportDocx(session).bytes,
      markers = controls(bytes);
    expect(markers.filter((item) => item.native)).toHaveLength(0);
    expect(markers).toHaveLength(action === "remove" ? 0 : 3);
    const model = openDocument(importDocx(bytes, { documentId: "native-edit-reopened" }), {
      kind: "document",
    });
    expect(model.snapshot().annotations).toHaveLength(action === "remove" ? 0 : 1);
    session.undo();
    expect(controls(exportDocx(session).bytes).filter((item) => item.native)).toHaveLength(3);
    session.redo();
    expect(controls(exportDocx(session).bytes).filter((item) => item.native)).toHaveLength(0);
  }
});
test("new comments over an opaque block emit valid outside anchors without reviving its deleted original thread", () => {
  const session = createDocumentSession(
      openDocument(importDocx(nativeCommentFixture(), { documentId: "native-comment-new" }), {
        kind: "document",
      }),
    ),
    old = session.snapshot().annotations[0]!,
    atom = session
      .snapshot()
      .metadata.word!.nativeContent!.fragments.find((item) => item.kind !== "property")!.nodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "native-comment-new",
      baseRevision: 0,
      operations: [
        { op: "removeAnnotation", annotationId: old.annotationId },
        {
          op: "addAnnotation",
          annotation: {
            annotationId: "new-comment",
            documentId: "native-comment-new",
            baseRevision: 0,
            status: "open",
            anchor: { kind: "node", nodeId: atom },
            instruction: "NEW",
          },
        },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 3; pass++) {
    const markers = controls(bytes);
    expect(markers).toHaveLength(3);
    expect(markers.some((item) => item.native)).toBe(false);
    const model = openDocument(importDocx(bytes, { documentId: "native-new-reopened" }), {
      kind: "document",
    });
    expect(model.snapshot().annotations).toHaveLength(1);
    expect(model.getNode(model.snapshot().annotations[0]!.bodyNodeId!)!.text).toBe("NEW");
    bytes = exportDocx(model).bytes;
  }
});

test("an unmodeled child inside a run consumes one source offset before later comment text", () => {
  const files = unzipSync(nativeCommentFixture());
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace(
      /<w:body>[\s\S]*<\/w:body>/,
      '<w:body><w:p><w:commentRangeStart w:id="7"/><w:r><w:t>BEFORE</w:t><w:dayLong/><w:t>AFTER</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p></w:body>',
    ),
  );
  let model = openDocument(importDocx(zipSync(files), { documentId: "native-comment-run" }), {
    kind: "document",
  });
  for (let pass = 0; pass < 3; pass++) {
    const annotation = model.snapshot().annotations[0]!,
      ranges =
        annotation.anchor.kind === "text-set"
          ? annotation.anchor.ranges
          : annotation.anchor.kind === "text"
            ? [annotation.anchor]
            : [];
    expect(ranges.map((range) => range.quote).join("")).toBe("BEFORE\ufffcAFTER");
    const bytes = exportDocx(model).bytes;
    model = openDocument(importDocx(bytes, { documentId: "native-run-reopened" }), {
      kind: "document",
    });
  }
});
