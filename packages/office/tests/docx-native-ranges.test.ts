import { expect, test } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { importDocx, exportDocx } from "@doffi/office";
import {
  nativeRangeFixture as fixture,
  nativeRangeFamilies as families,
} from "./native-range-fixture.js";
import { nativeRangeControl, visitNativeRanges } from "../src/docx-native-ranges.js";
import { xml } from "../src/ooxml/xml.js";
import { DEFAULT_DOCUMENT_READ_LIMITS } from "@doffi/core";

function controls(bytes: Uint8Array) {
  const source = strFromU8(unzipSync(bytes)["word/document.xml"]!),
    root = xml(strToU8(source), "word/document.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 }),
    result: ReturnType<typeof nativeRangeControl>[] = [];
  visitNativeRanges(root, (node) => {
    const c = nativeRangeControl(node);
    if (c) result.push(c);
  });
  return result;
}
for (const family of families)
  for (const inside of ["start", "end", "both"] as const)
    test(`${family} pairs retain source ownership across native ${inside}, text edits, deletion and undo`, () => {
      const documentId = "native-ranges",
        session = createDocumentSession(
          openDocument(importDocx(fixture(family, inside), { documentId }), { kind: "document" }),
        ),
        apply = (operations: DocumentOperation[]) =>
          session.commit(
            session.prepare({
              protocolVersion: 1,
              documentId,
              baseRevision: session.snapshot().manifest.revision,
              operations,
            }),
          ),
        native = session.snapshot().metadata.word!.nativeContent!,
        object = native.fragments.find((f) => f.kind === "block")!,
        nodes = [session.snapshot().metadata.word!.stories[0]!.rootNodeId];
      for (let i = 0; i < nodes.length; i++)
        nodes.push(...session.getNode(nodes[i]!)!.childNodeIds);
      const paragraph = nodes.find(
        (id) =>
          session.getNode(id)!.tagName === "p" && /BEFORE|EDITABLE/.test(session.getNode(id)!.text),
      )!;
      expect(session.getNode(paragraph)!.text).toBe(inside === "both" ? "EDITABLE" : "BEFORE");
      expect(native.fragments.flatMap((f) => f.rangeBoundaries ?? [])).toHaveLength(2);
      apply([
        {
          op: "replaceText",
          nodeId: paragraph,
          from: 0,
          to: 1,
          expectedText: inside === "both" ? "E" : "B",
          text: "X",
        },
      ]);
      let bytes = exportDocx(session).bytes;
      for (let pass = 0; pass < 2; pass++) {
        const pair = controls(bytes);
        expect(pair).toHaveLength(2);
        expect(pair[0]!.sourceId).toBe(pair[1]!.sourceId);
        expect(pair.map((p) => p!.side)).toEqual(["start", "end"]);
        bytes = exportDocx(
          openDocument(importDocx(bytes, { documentId: "reopen" }), { kind: "document" }),
        ).bytes;
      }
      const before = session.serialize(),
        candidate = session.prepare({
          protocolVersion: 1,
          documentId,
          baseRevision: session.snapshot().manifest.revision,
          operations: [{ op: "removeNode", nodeId: object.nodeId! }],
        });
      apply([
        { op: "replaceText", nodeId: paragraph, from: 1, to: 1, expectedText: "", text: "Y" },
      ]);
      expect(() => session.commit(candidate)).toThrow();
      session.undo();
      apply([{ op: "removeNode", nodeId: object.nodeId! }]);
      expect(controls(exportDocx(session).bytes)).toHaveLength(0);
      expect(
        session.snapshot().metadata.word!.nativeContent?.fragments.some((f) => f.rangeBoundary),
      ).toBeFalsy();
      session.undo();
      for (const [path, bytes] of before)
        if (path !== "manifest.json") expect(session.serialize().get(path)).toEqual(bytes);
      expect(controls(exportDocx(session).bytes)).toHaveLength(2);
      session.redo();
      const reopened = openDocument(
        importDocx(exportDocx(session).bytes, { documentId: "deleted" }),
        { kind: "document" },
      );
      expect(controls(exportDocx(reopened).bytes)).toHaveLength(0);
    });

test("unpaired or repeated native boundaries fail instead of silently losing controls", () => {
  for (const mode of ["missing", "duplicate"] as const) {
    const files = unzipSync(fixture("bookmark"));
    files["word/document.xml"] = strToU8(
      strFromU8(files["word/document.xml"]!).replace(
        '<w:bookmarkEnd w:id="9"/>',
        mode === "missing" ? "" : '<w:bookmarkEnd w:id="9"/><w:bookmarkEnd w:id="9"/>',
      ),
    );
    expect(() => importDocx(zipSync(files), { documentId: "malformed-range" })).toThrow(
      /range|bookmark/i,
    );
  }
});

test("splitting editable paragraphs retains zero-width native endpoints and their ordering", () => {
  const documentId = "split-ranges",
    session = createDocumentSession(
      openDocument(importDocx(fixture("bookmark"), { documentId }), { kind: "document" }),
    ),
    boundary = session
      .snapshot()
      .metadata.word!.nativeContent!.fragments.find((f) => f.rangeBoundary)!,
    paragraph = session.getNode(boundary.nodeId!)!.parentNodeId!,
    before = session.snapshot();
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId,
      baseRevision: 0,
      operations: [
        {
          op: "pasteWordBlockContent",
          nodeId: paragraph,
          from: 3,
          to: 3,
          expectedText: "",
          html: "<p>INSERTED</p>",
        },
      ],
    }),
  );
  expect(controls(exportDocx(session).bytes).map((c) => c!.side)).toEqual(["start", "end"]);
  expect(session.getNode(boundary.nodeId!)!.text).toBe("");
  const model = openDocument(
    importDocx(exportDocx(session).bytes, { documentId: "split-reopen" }),
    { kind: "document" },
  );
  expect(model.snapshot().html).toContain("INSERTED");
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
});

test("moved content spanning a native boundary stays opaque with its original author", () => {
  const files = unzipSync(fixture("moveFromRange"));
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace(
      "<w:r><w:t>BEFORE</w:t></w:r>",
      '<w:moveFrom w:id="2" w:author="Mover" w:date="2026-11-16T00:00:00Z"><w:r><w:t>BEFORE</w:t></w:r></w:moveFrom>',
    ),
  );
  const model = openDocument(importDocx(zipSync(files), { documentId: "native-move" }), {
      kind: "document",
    }),
    bytes = exportDocx(model).bytes;
  expect(controls(bytes)).toHaveLength(2);
  expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).toContain('w:author="Mover"');
  expect(() => importDocx(bytes, { documentId: "move-reopen" })).not.toThrow();
});
