import { expect, test } from "vitest";
import { posix } from "node:path";
import {
  createDocumentSession,
  openDocument,
  createAssetDescriptor,
  parseContract,
} from "@doffi/core";
import type { DocumentOperation, DocumentPackage } from "@doffi/core";
import { importDocx, exportDocx } from "@doffi/office";
import {
  legacyNativeContent,
  wordNativeContentFixture,
  nativePreviewContentFixture,
} from "../../office/tests/word-native-content-fixture.js";
import { body, mutateDocx, styles } from "../../office/tests/regression-fixtures.js";
const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
function setup(files: DocumentPackage) {
  const session = createDocumentSession(openDocument(files, { kind: "document" })),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: session.snapshot().manifest.documentId,
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  return { session, apply };
}
function sameFiles(expected: DocumentPackage, actual: DocumentPackage) {
  expect([...actual.keys()].sort()).toEqual([...expected.keys()].sort());
  for (const [path, bytes] of expected)
    if (path !== "manifest.json") expect(actual.get(path)).toEqual(bytes);
}
test("legacy XML collections release deleted private entries while undo restores every original byte", () => {
  const { session, apply } = setup(
      legacyNativeContent(
        importDocx(wordNativeContentFixture(), { documentId: "legacy-collection" }),
      ),
    ),
    before = session.serialize(),
    native = session.snapshot().metadata.word!.nativeContent!,
    [first, second] = native.fragments.filter((f) => f.kind === "block"),
    old = first!.assetPath;
  expect(first!.assetPath).toBe(second!.assetPath);
  apply([{ op: "removeNode", nodeId: first!.nodeId! }]);
  const remaining = session
    .snapshot()
    .metadata.word!.nativeContent!.fragments.find((f) => f.kind === "block")!;
  expect(remaining.assetPath).not.toBe(old);
  expect(session.serialize().has(old)).toBe(false);
  expect(decode(session.serialize().get(remaining.assetPath)!)).not.toContain("SELECTED-NATIVE");
  expect(decode(session.serialize().get(remaining.assetPath)!)).toContain("PRIVATE-NATIVE-OUTSIDE");
  expect(() =>
    openDocument(importDocx(exportDocx(session).bytes, { documentId: "legacy-reopen" }), {
      kind: "document",
    }),
  ).not.toThrow();
  apply([{ op: "removeNode", nodeId: second!.nodeId! }]);
  expect(session.snapshot().metadata.word!.nativeContent?.fragments ?? []).toHaveLength(0);
  for (const bytes of session.serialize().values())
    expect(decode(bytes)).not.toContain("PRIVATE-NATIVE-OUTSIDE");
  session.undo();
  session.undo();
  sameFiles(before, session.serialize());
  session.redo();
  expect(session.serialize().has(old)).toBe(false);
});
function sharedFixture() {
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/document.xml": (s) =>
      body(
        s,
        ["A", "B"]
          .map(
            (label) =>
              `<w:customXml w:element="${label}" w:uri="urn:test" r:id="shared"><w:p><w:r><w:t>PRIVATE-${label}</w:t></w:r></w:p></w:customXml>`,
          )
          .join("") + "<w:p><w:r><w:t>KEEP</w:t></w:r></w:p>",
      ),
    "word/shared.xml": () => "<data>SHARED-CONTENT</data>",
    "word/child.xml": () => "<data>SHARED-CHILD</data>",
    "word/global.xml": () => "<data>INDEPENDENT-GLOBAL</data>",
    "word/_rels/shared.xml.rels": () =>
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="child" Type="urn:test:child" Target="child.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": (s) =>
      s.replace(
        "</Relationships>",
        '<Relationship Id="shared" Type="urn:test:shared" Target="shared.xml"/><Relationship Id="global" Type="urn:test:global" Target="global.xml"/></Relationships>',
      ),
    "[Content_Types].xml": (s) =>
      s.replace(
        "</Types>",
        ["shared", "child", "global"]
          .map((p) => `<Override PartName="/word/${p}.xml" ContentType="application/xml"/>`)
          .join("") + "</Types>",
      ),
  });
}
for (const legacy of [false, true])
  test(`shared native dependency ownership is released after its final object, retaining independent roots (${legacy})`, () => {
    const imported = importDocx(sharedFixture(), { documentId: "shared-legacy" }),
      { session, apply } = setup(legacy ? legacyNativeContent(imported) : imported),
      before = session.serialize(),
      native = session.snapshot().metadata.word!.nativeContent!,
      atoms = native.fragments.filter((f) => f.kind === "block"),
      shared = native.parts.find((p) => p.sourcePart === "word/shared.xml")!,
      child = native.parts.find((p) => p.sourcePart === "word/child.xml")!,
      global = native.parts.find((p) => p.sourcePart === "word/global.xml")!;
    apply([{ op: "removeNode", nodeId: atoms[0]!.nodeId! }]);
    expect(session.serialize().has(shared.assetPath)).toBe(true);
    expect(session.serialize().has(child.assetPath)).toBe(true);
    expect(() =>
      importDocx(exportDocx(session).bytes, { documentId: "shared-reopen" }),
    ).not.toThrow();
    const candidate = session.prepare({
      protocolVersion: 1,
      documentId: "shared-legacy",
      baseRevision: session.snapshot().manifest.revision,
      operations: [{ op: "removeNode", nodeId: atoms[1]!.nodeId! }],
    });
    apply([{ op: "setStyle", nodeId: atoms[1]!.nodeId!, property: "color", value: "#112233" }]);
    expect(() => session.commit(candidate)).toThrow();
    session.undo();
    apply([{ op: "removeNode", nodeId: atoms[1]!.nodeId! }]);
    expect(session.serialize().has(shared.assetPath)).toBe(false);
    expect(session.serialize().has(child.assetPath)).toBe(false);
    expect(session.serialize().has(global.assetPath)).toBe(true);
    session.undo();
    session.undo();
    sameFiles(before, session.serialize());
  });

test("legacy opaque UUID bindings release only the removed control's datastore closure", () => {
  const files = importDocx(wordNativeContentFixture(true), { documentId: "legacy-binding" }),
    model = openDocument(files, { kind: "document" }),
    bound = model
      .snapshot()
      .metadata.word!.nativeContent!.fragments.find((f) =>
        f.dependencyParts?.includes("customXml/item1.xml"),
      )!,
    { session, apply } = setup(legacyNativeContent(files)),
    before = session.serialize(),
    parts = session.snapshot().metadata.word!.nativeContent!.parts,
    selected = parts.filter((p) => /(?:item1|itemProps1)\.xml$/.test(p.sourcePart)),
    unbound = parts.find((p) => p.sourcePart === "customXml/item2.xml")!;
  expect(selected).toHaveLength(2);
  apply([{ op: "removeNode", nodeId: bound.nodeId! }]);
  for (const part of selected) expect(session.serialize().has(part.assetPath)).toBe(false);
  expect(session.serialize().has(unbound.assetPath)).toBe(true);
  const reopened = openDocument(
    importDocx(exportDocx(session).bytes, { documentId: "binding-reopen" }),
    { kind: "document" },
  );
  expect(
    reopened
      .snapshot()
      .metadata.word!.nativeContent?.parts.some((p) =>
        decode(reopened.serialize().get(p.assetPath)!).includes("SELECTED-BOUND-DATA"),
      ),
  ).toBeFalsy();
  session.undo();
  sameFiles(before, session.serialize());
});

test("native image bytes shared with a surviving source image are retained until that image is removed", () => {
  const { session, apply } = setup(
      legacyNativeContent(
        importDocx(nativePreviewContentFixture(), { documentId: "shared-source-image" }),
      ),
    ),
    native = session.snapshot().metadata.word!.nativeContent!,
    picture = native.parts.find((p) => p.contentType === "image/png")!,
    atom = native.fragments.find((f) => f.preview)!,
    root = session.snapshot().metadata.word!.stories[0]!.rootNodeId,
    reference = posix.relative(
      posix.dirname(session.snapshot().manifest.entrypoint),
      picture.assetPath,
    );
  apply([
    {
      op: "insertFragment",
      nodeId: root,
      position: "end",
      html: `<p><img src="${reference}" alt="shared"/></p>`,
    },
  ]);
  const ids = [root];
  for (let i = 0; i < ids.length; i++) ids.push(...session.getNode(ids[i]!)!.childNodeIds);
  const image = ids.find((id) => session.getNode(id)!.tagName === "img")!;
  apply([{ op: "removeNode", nodeId: atom.nodeId! }]);
  expect(session.serialize().has(picture.assetPath)).toBe(true);
  apply([{ op: "removeNode", nodeId: image }]);
  expect(session.serialize().has(picture.assetPath)).toBe(false);
});

test("legacy collection parsing stays bounded and rejects malformed XML atomically", () => {
  const original = legacyNativeContent(
      importDocx(wordNativeContentFixture(), { documentId: "legacy-invalid" }),
    ),
    model = openDocument(original, { kind: "document" }),
    atom = model
      .snapshot()
      .metadata.word!.nativeContent!.fragments.find((f) => f.kind === "block")!;
  for (const mode of ["doctype", "processing", "malformed", "deep", "nodes"] as const) {
    const files = new Map(original),
      manifest = parseContract("document-manifest", model.snapshot().manifest),
      source = decode(files.get(atom.assetPath)!),
      content =
        mode === "doctype"
          ? '<!DOCTYPE native [<!ENTITY private "value">]>' + source
          : mode === "processing"
            ? "<?native instruction?>" + source
            : mode === "nodes"
              ? source.replace("SELECTED-NATIVE", "<n/>".repeat(1000))
              : mode === "malformed"
                ? source.replace("</native:fragments>", "")
                : source.replace("SELECTED-NATIVE", "<n>".repeat(80) + "TEXT" + "</n>".repeat(80)),
      bytes = new TextEncoder().encode(content);
    files.set(atom.assetPath, bytes);
    manifest.assets = manifest.assets.map((a) =>
      a.path === atom.assetPath ? createAssetDescriptor(a.path, a.mimeType, bytes) : a,
    );
    files.set("manifest.json", new TextEncoder().encode(JSON.stringify(manifest)));
    const session = createDocumentSession(openDocument(files, { kind: "document" }), {
        limits: { maxDepth: 64, ...(mode === "nodes" ? { maxNodes: 200 } : {}) },
      }),
      before = session.snapshot();
    expect(() =>
      session.prepare({
        protocolVersion: 1,
        documentId: "legacy-invalid",
        baseRevision: 0,
        operations: [{ op: "removeNode", nodeId: atom.nodeId! }],
      }),
    ).toThrow();
    expect(session.snapshot()).toEqual(before);
  }
});

test("legacy XML slicing preserves inherited namespaces, UTF-16 input and supplementary text", () => {
  const imported = legacyNativeContent(
      importDocx(wordNativeContentFixture(), { documentId: "legacy-encoding" }),
    ),
    model = openDocument(imported, { kind: "document" }),
    [first, second] = model
      .snapshot()
      .metadata.word!.nativeContent!.fragments.filter((f) => f.kind === "block"),
    files = new Map(imported),
    source =
      '<?xml version="1.0" encoding="UTF-16"?>\r\n' +
      decode(files.get(first!.assetPath)!)
        .replace("<native:fragments ", '<native:fragments xmlns:extra="urn:test:inherited" ')
        .replaceAll("<w:customXml ", '<w:customXml extra:value="extra:sample" ')
        .replace("SELECTED-NATIVE", "REMOVED-𠀀\u{e0100}\r\n")
        .replace("PRIVATE-NATIVE-OUTSIDE", "RETAINED-中文𠀀\u{e0100}"),
    bytes = new Uint8Array(
      Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(source, "utf16le")]),
    ),
    manifest = parseContract("document-manifest", model.snapshot().manifest);
  files.set(first!.assetPath, bytes);
  manifest.assets = manifest.assets.map((a) =>
    a.path === first!.assetPath ? createAssetDescriptor(a.path, a.mimeType, bytes) : a,
  );
  files.set("manifest.json", new TextEncoder().encode(JSON.stringify(manifest)));
  const { session, apply } = setup(files),
    before = session.serialize(),
    retainedItem = source.slice(
      source.indexOf('<native:item id="' + second!.fragmentKey + '"'),
      source.indexOf(
        "</native:item>",
        source.indexOf('<native:item id="' + second!.fragmentKey + '"'),
      ) + "</native:item>".length,
    );
  apply([{ op: "removeNode", nodeId: first!.nodeId! }]);
  const remaining = session
      .snapshot()
      .metadata.word!.nativeContent!.fragments.find((f) => f.kind === "block")!,
    content = decode(session.serialize().get(remaining.assetPath)!);
  expect(content).toContain(retainedItem);
  expect(content).toContain('encoding="UTF-8"');
  expect(content).not.toContain("REMOVED-");
  const reopened = openDocument(
    importDocx(exportDocx(session).bytes, { documentId: "encoding-reopen" }),
    { kind: "document" },
  );
  const native = reopened.snapshot().metadata.word!.nativeContent!;
  expect(decode(reopened.serialize().get(native.fragments[0]!.assetPath)!)).toContain(
    "RETAINED-中文𠀀\u{e0100}",
  );
  session.undo();
  sameFiles(before, session.serialize());
});

test("a remaining unkeyed native owner retains the entire shared XML resource", () => {
  const imported = legacyNativeContent(
      importDocx(wordNativeContentFixture(), { documentId: "legacy-whole-owner" }),
    ),
    model = openDocument(imported, { kind: "document" }),
    files = new Map(imported),
    metadata = parseContract("document-metadata", model.snapshot().metadata);
  if (!("word" in metadata) || !metadata.word) throw new Error("Expected Word fixture.");
  const atoms = metadata.word.nativeContent!.fragments.filter((f) => f.kind === "block"),
    path = atoms[0]!.assetPath;
  delete atoms[1]!.fragmentKey;
  files.set(model.snapshot().manifest.metadata, new TextEncoder().encode(JSON.stringify(metadata)));
  const { session, apply } = setup(files);
  apply([{ op: "removeNode", nodeId: atoms[0]!.nodeId! }]);
  expect(
    session
      .snapshot()
      .metadata.word!.nativeContent!.fragments.find((f) => f.nodeId === atoms[1]!.nodeId)!
      .assetPath,
  ).toBe(path);
  expect(session.serialize().get(path)).toEqual(files.get(path));
});
