import { NativeWordIdentities } from "../src/docx-native-word-identities.js";
import { xml } from "../src/ooxml/xml.js";
import { DEFAULT_DOCUMENT_READ_LIMITS } from "@doffi/core";
import { expect, test } from "vitest";
import { SaxesParser } from "saxes";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { nativeWordIdFixture } from "./native-word-id-fixture.js";
function identities(bytes: Uint8Array) {
  const result = {
      bookmarks: [] as { name: string; id: string }[],
      ends: [] as string[],
      permissions: [] as string[],
      permissionEnds: [] as string[],
      revisions: [] as string[],
    },
    parser = new SaxesParser({ xmlns: true });
  parser.on("opentag", (tag) => {
    const attrs = Object.fromEntries(
      Object.values(tag.attributes).map((attr) => [attr.local, attr.value]),
    );
    if (tag.local === "bookmarkStart") result.bookmarks.push({ name: attrs.name!, id: attrs.id! });
    if (tag.local === "bookmarkEnd") result.ends.push(attrs.id!);
    if (tag.local === "permStart") result.permissions.push(attrs.id!);
    if (tag.local === "permEnd") result.permissionEnds.push(attrs.id!);
    if (tag.local === "ins") result.revisions.push(attrs.id!);
  });
  parser.write(strFromU8(unzipSync(bytes)["word/document.xml"]!)).close();
  return result;
}
test("opaque native bookmarks and revisions cannot collide with current generated IDs while paired boundaries remain linked", () => {
  const session = createDocumentSession(
      openDocument(importDocx(nativeWordIdFixture(), { documentId: "native-word-ids" }), {
        kind: "document",
      }),
    ),
    nodes = [session.snapshot().metadata.word!.stories[0]!.rootNodeId];
  for (let i = 0; i < nodes.length; i++) nodes.push(...session.getNode(nodes[i]!)!.childNodeIds);
  const paragraph = nodes.find(
      (id) => session.getNode(id)!.tagName === "p" && session.getNode(id)!.text === "VISIBLE",
    )!,
    root = session.getNode(nodes[0]!)!.parentNodeId!,
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "native-word-ids",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    {
      op: "insertWordBookmark",
      nodeId: paragraph,
      from: 0,
      to: 1,
      expectedText: "V",
      name: "Added",
    },
    { op: "setWordReview", nodeId: root, trackChanges: true, author: "Current" },
    { op: "trackWordText", nodeId: paragraph, from: 7, to: 7, expectedText: "", text: "X" },
  ]);
  const before = session.snapshot();
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 3; pass++) {
    const ids = identities(bytes);
    expect(ids.bookmarks.map((item) => item.name).sort()).toEqual(["Added", "Hidden", "Visible"]);
    expect(new Set(ids.bookmarks.map((item) => item.id)).size).toBe(3);
    expect([...ids.ends].sort()).toEqual(ids.bookmarks.map((item) => item.id).sort());
    expect(ids.permissions).toEqual(ids.permissionEnds);
    expect(new Set(ids.revisions).size).toBe(2);
    const model = openDocument(importDocx(bytes, { documentId: "native-word-reopened" }), {
      kind: "document",
    });
    expect(model.snapshot().metadata.word!.bookmarks).toHaveLength(2);
    expect(model.snapshot().metadata.word!.revisions).toHaveLength(1);
    bytes = exportDocx(model).bytes;
  }
  session.undo();
  expect(session.getNode(paragraph)!.text).toBe("VISIBLE");
  session.redo();
  expect(session.snapshot().html).toBe(before.html);
  expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
});

test("numeric range aliases share their integer identity while permission IDs retain string pairing", () => {
  const files = unzipSync(nativeWordIdFixture());
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!)
      .replace('<w:bookmarkEnd w:id="0"/>', '<w:bookmarkEnd w:id="+000"/>')
      .replaceAll('w:permStart w:id="0"', 'w:permStart w:id="permission-alpha"')
      .replaceAll('w:permEnd w:id="0"', 'w:permEnd w:id="permission-alpha"'),
  );
  const model = openDocument(importDocx(zipSync(files), { documentId: "native-id-lexical" }), {
      kind: "document",
    }),
    bytes = exportDocx(model).bytes,
    ids = identities(bytes);
  expect(new Set(ids.bookmarks.map((item) => item.id)).size).toBe(2);
  expect([...ids.ends].sort()).toEqual(ids.bookmarks.map((item) => item.id).sort());
  expect(ids.permissions).toEqual(ids.permissionEnds);
  openDocument(importDocx(bytes, { documentId: "native-id-lexical-reopened" }), {
    kind: "document",
  });
});

test("Word-looking identities inside an unrelated native data namespace remain opaque data", () => {
  const encode = (value: string) => new TextEncoder().encode(value),
    generated = xml(
      encode(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:bookmarkStart w:id="0" w:name="Real"/><w:bookmarkEnd w:id="0"/></w:p></w:body></w:document>',
      ),
      "generated.xml",
      DEFAULT_DOCUMENT_READ_LIMITS,
      { nodes: 0 },
    ),
    payload = xml(
      encode(
        '<x:payload xmlns:x="urn:vendor" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bookmarkStart w:id="0" w:name="Cached"/><w:bookmarkEnd w:id="0"/></x:payload>',
      ),
      "native.xml",
      DEFAULT_DOCUMENT_READ_LIMITS,
      { nodes: 0 },
    ),
    allocator = new NativeWordIdentities([generated], () => ({ part: "native.xml" }));
  generated.content.push(payload);
  allocator.rewrite([generated]);
  for (const child of payload.content)
    if (typeof child !== "string")
      expect(
        child.attrs.get("http://schemas.openxmlformats.org/wordprocessingml/2006/main id"),
      ).toBe("0");
});
