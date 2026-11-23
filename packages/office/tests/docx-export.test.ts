import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync, readdirSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, DocumentError } from "@doffi/core";
import { exportDocx, importDocx as importNativeDocx } from "@doffi/office";
// These historical rejection cases exercise the explicitly retained strict policy.
const importDocx: typeof importNativeDocx = (bytes, options) =>
  importNativeDocx(bytes, { ...options, unsupportedContent: "reject" });
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
const directory = new URL("../../../examples/docx/", import.meta.url);
const fixture = (name: string) => new Uint8Array(readFileSync(new URL(name, directory)));
const model = (name: string) =>
  openDocument(importNativeDocx(fixture(name), { documentId: "export-test" }));
const mainText = (document: ReturnType<typeof openDocument>) =>
  document
    .snapshot()
    .metadata.word!.stories.filter((story) => story.kind === "main")
    .map((story) => document.getNode(story.rootNodeId)!.text)
    .join("");
for (const producer of ["", "producer/"])
  for (const name of readdirSync(new URL(producer || ".", directory)).filter((name) =>
    name.endsWith(".docx"),
  )) {
    test("native export and reimport retain source semantics: " + producer + name, () => {
      const source = model(producer + name),
        snapshot = source.snapshot(),
        before = source.serialize();
      const result = exportDocx(source),
        native = unzipSync(result.bytes);
      expect(source.serialize()).toEqual(before);
      expect(Object.keys(native)).not.toContain("manifest.json");
      expect(strFromU8(native["word/document.xml"]!)).not.toContain("altChunk");
      expect(strFromU8(native["[Content_Types].xml"]!)).toContain('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"');
      expect(strFromU8(native["word/_rels/document.xml.rels"]!)).toContain('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"');
      const reopened = openDocument(importNativeDocx(result.bytes, { documentId: "reopened" }));
      expect(mainText(reopened)).toBe(mainText(source));
      const actual = reopened.snapshot();
      expect(actual.metadata.sections).toHaveLength(snapshot.metadata.sections!.length);
      expect(actual.metadata.locale).toBe(snapshot.metadata.locale);
      const word = actual.metadata.word!,
        expected = snapshot.metadata.word!;
      expect(
        word.fields.map((field) => ({
          instruction: field.instruction,
          locked: field.locked,
          dirty: field.dirty,
          complex: Boolean(field.endNodeId),
        })),
      ).toEqual(
        expected.fields.map((field) => ({
          instruction: field.instruction,
          locked: field.locked,
          dirty: field.dirty,
          complex: Boolean(field.endNodeId),
        })),
      );
      expect(word.revisions.map(({ kind, author, date }) => ({ kind, author, date }))).toEqual(
        expected.revisions.map(({ kind, author, date }) => ({ kind, author, date })),
      );
      expect(word.hyperlinks.map(({ target, bookmark }) => ({ target, bookmark }))).toEqual(
        expected.hyperlinks.map(({ target, bookmark }) => ({ target, bookmark })),
      );
      expect(word.tables.map(({ grid }) => grid)).toEqual(expected.tables.map(({ grid }) => grid));
      expect(actual.annotations).toHaveLength(snapshot.annotations.length);
      expect(result.report.unexportedAnnotations).toEqual([]);
      expect(exportDocx(source).bytes).toEqual(result.bytes);
    });
  }
test("native comments retain resolved state and precise ranges after an edit", () => {
  const session = createDocumentSession(model("08-comments-revisions.docx"));
  const annotation = session.getAnnotations()[0]!;
  const candidate = session.prepare({
    protocolVersion: 1,
    documentId: "export-test",
    baseRevision: 0,
    operations: [
      {
        op: "updateAnnotation",
        annotationId: annotation.annotationId,
        changes: { status: "resolved" },
      },
    ],
  });
  session.commit(candidate);
  const exported = exportDocx(session),
    parts = unzipSync(exported.bytes);
  expect(strFromU8(parts["word/commentsExtended.xml"]!)).toContain('w15:done="1"');
  const reopened = openDocument(importDocx(exported.bytes, { documentId: "resolved" }));
  expect(reopened.snapshot().annotations[0]!.status).toBe("resolved");
  expect(exported.report.revision).toBe(1);
  session.dispose();
});
test("export is bounded and leaves source unchanged on failure", () => {
  const source = model("01-text-numbering.docx"),
    before = source.serialize();
  expect(() => exportDocx(source, { limits: { maxNodes: 20 } })).toThrow(DocumentError);
  expect(() => exportDocx(source, { maxZipBytes: 20 })).toThrow(DocumentError);
  expect(source.serialize()).toEqual(before);
});

import {
  crossSectionRangesFixture,
  editedChineseFontFixture,
  nestedFieldFixture,
  styleInheritanceFixture,
  indentedTabsFixture,
} from "./regression-fixtures.js";
import { readDocumentZip } from "@doffi/office";
for (const [name, input] of [
  ["cross-section ranges", () => crossSectionRangesFixture(false)],
  ["cross-section instruction", () => crossSectionRangesFixture(true)],
  ["nested complex field", () => nestedFieldFixture(false)],
  ["nested simple field", () => nestedFieldFixture(true)],
  ["style inheritance", styleInheritanceFixture],
  ["indented tabs", indentedTabsFixture],
] as const)
  test("exports declared reading regression: " + name, () => {
    const source = openDocument(importDocx(input(), { documentId: "regression" })),
      exported = exportDocx(source);
    const reopened = openDocument(importDocx(exported.bytes, { documentId: "roundtrip" }));
    expect(mainText(reopened)).toBe(mainText(source));
    expect(
      reopened.snapshot().metadata.word!.fields.map(({ instruction, instructionFields }) => ({
        instruction,
        offsets: instructionFields?.map(({ offset }) => offset),
      })),
    ).toEqual(
      source.snapshot().metadata.word!.fields.map(({ instruction, instructionFields }) => ({
        instruction,
        offsets: instructionFields?.map(({ offset }) => offset),
      })),
    );
    expect(reopened.snapshot().annotations).toHaveLength(source.snapshot().annotations.length);
    expect(exported.report.unexportedAnnotations).toEqual([]);
  });
test("exports the current edited Chinese font roles and complete IVS sequence", () => {
  const source = openDocument(readDocumentZip(editedChineseFontFixture())),
    exported = exportDocx(source);
  const reopened = openDocument(importDocx(exported.bytes, { documentId: "fonts" }));
  expect(mainText(reopened)).toBe(mainText(source));
  const native = strFromU8(unzipSync(exported.bytes)["word/document.xml"]!);
  expect(native).toContain('w:eastAsia="FandolFang"');
  expect(native).toContain("𠀀\u{e0100}");
});

test.each([
  [
    "invalid boolean",
    (xml: string) => xml.replace('w15:done="0"', 'w15:done="2"'),
    "InvalidDocument",
  ],
  [
    "unknown paragraph",
    (xml: string) => xml.replace('w15:paraId="00000001"', 'w15:paraId="0000FFFF"'),
    "InvalidDocument",
  ],
  [
    "self reply parent",
    (xml: string) => xml.replace("<w15:commentEx ", '<w15:commentEx w15:paraIdParent="00000001" '),
    "InvalidDocument",
  ],
  [
    "unknown extension",
    (xml: string) => xml.replace("<w15:commentEx ", '<w15:commentEx w15:future="1" '),
    "UnsupportedFeature",
  ],
] as const)(
  "resolved-comment reader rejects ambiguous or unsupported extension: %s",
  (_name, mutate, code) => {
    const exported = exportDocx(model("08-comments-revisions.docx")),
      parts = unzipSync(exported.bytes);
    parts["word/commentsExtended.xml"] = strToU8(
      mutate(strFromU8(parts["word/commentsExtended.xml"]!)),
    );
    expect(() => importDocx(zipSync(parts), { documentId: "invalid" })).toThrowError(
      expect.objectContaining({ code }),
    );
  },
);
