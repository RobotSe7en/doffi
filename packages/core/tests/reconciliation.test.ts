import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { expect, test } from "vitest";
import { createDocumentSession, DocumentError } from "@doffi/core";
import type { DocumentContentOperation, DocumentSession, DocumentSessionOptions, ReconciledTextEdit } from "@doffi/core";
import { repository } from "./fixtures.js";

function session(text = "甲甲甲", options?: DocumentSessionOptions): DocumentSession {
  const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path, new Uint8Array(readFileSync(join(repository, "examples/report", path)))]));
  files.set("content.html", new TextEncoder().encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body><article data-node-id="root" data-office-kind="document"><p data-node-id="p"><span data-node-id="s">'
    + text + '</span></p><p data-node-id="q">后</p></article></body></html>'));
  files.set("annotations.json", new TextEncoder().encode("[]"));
  return createDocumentSession(openDocument(files), options);
}
function prepare(s: DocumentSession, operation: DocumentContentOperation) {
  return s.prepare({ protocolVersion: 1, documentId: "report-demo", baseRevision: s.snapshot().manifest.revision, operations: [operation] });
}
function reconcile(html: string, textEdits: ReconciledTextEdit[] = [], nodeId = "root"): DocumentContentOperation {
  return { op: "reconcileFragment", nodeId, html, textEdits };
}
function annotate(s: DocumentSession, nodeId = "p", from = 1, to = 2): void {
  s.commit(s.prepare({ protocolVersion: 1, documentId: "report-demo", baseRevision: s.snapshot().manifest.revision, operations: [{ op: "addAnnotation",
    annotation: { annotationId: "note", documentId: "report-demo", baseRevision: s.snapshot().manifest.revision, status: "open", instruction: "edit",
      anchor: { kind: "text", nodeId, from, to, quote: s.getNode(nodeId)!.text.slice(from, to) } } }] }));
}
function fails(action: () => unknown, code = "InvalidOperation"): void {
  try { action(); } catch (error) { expect(error).toBeInstanceOf(DocumentError); expect((error as DocumentError).code).toBe(code); return; }
  throw new Error("Expected rejection");
}

test("reconciliation retains IDs, allocates new IDs, maps an exact repeated-text insertion and restores history", () => {
  const s = session(); annotate(s); const original = s.snapshot();
  const candidate = prepare(s, reconcile('<p data-node-id="p"><span data-node-id="s">甲甲甲甲</span></p><p data-node-id="q">后</p><p>新段</p>',
    ["p", "s"].map((nodeId) => ({ nodeId, from: 1, to: 1, expectedText: "", text: "甲" }))));
  expect(s.snapshot()).toEqual(original);
  s.commit(candidate);
  expect(s.getNode("q")!.html).toBe('<p data-node-id="q">后</p>');
  expect(s.getNode("root")!.childNodeIds).toHaveLength(3);
  expect(s.getNode("root")!.childNodeIds[2]).toMatch(/^edit-/);
  expect(s.getAnnotations()[0]!.anchor).toEqual({ kind: "text", nodeId: "p", from: 2, to: 3, quote: "甲" });
  expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().html).toBe(original.html); expect(s.getAnnotations()[0]!.anchor).toEqual(original.annotations[0]!.anchor);
  s.redo(); expect(s.getNode("s")!.text).toBe("甲甲甲甲");
});

test("formatting with unchanged text retains anchors, including an explicit no-op proof", () => {
  const s = session(); annotate(s);
  s.commit(prepare(s, reconcile('<span data-node-id="s" style="font-weight:700">甲甲甲</span>',
    [{ nodeId: "p", from: 0, to: 3, expectedText: "甲甲甲", text: "甲甲甲" }], "p")));
  expect(s.getAnnotations()[0]).toMatchObject({ status: "open", anchor: { from: 1, to: 2, quote: "甲" } });
});

test("missing mapping evidence explicitly invalidates text anchors instead of guessing offsets", () => {
  const s = session(); annotate(s);
  s.commit(prepare(s, reconcile('<span data-node-id="s">前甲甲甲</span>', [], "p")));
  expect(s.getAnnotations()[0]).toMatchObject({ status: "needs_reanchor", anchor: { from: 1, to: 2, quote: "甲" } });
});

test("reconciliation cannot reuse foreign, guessed, target-root or duplicate node identities", () => {
  const s = session(), before = s.serialize();
  for (const html of ['<span data-node-id="q">foreign</span>', '<span data-node-id="guessed">new</span>',
    '<span data-node-id="p">target</span>', '<span data-node-id="s">one</span><span data-node-id="s">two</span>'])
    fails(() => prepare(s, reconcile(html, [], "p")));
  expect(s.serialize()).toEqual(before);
});

test("mapping evidence must match source, target scope, final text and UTF-16 boundaries", () => {
  const s = session("甲😀乙"), before = s.serialize();
  const html = '<span data-node-id="s">甲新乙</span>';
  for (const edit of [
    { nodeId: "p", from: 1, to: 3, expectedText: "wrong", text: "新" },
    { nodeId: "q", from: 0, to: 1, expectedText: "后", text: "新" },
    { nodeId: "p", from: 1, to: 2, expectedText: "\ud83d", text: "新" },
    { nodeId: "p", from: 1, to: 3, expectedText: "😀", text: "different" },
  ]) fails(() => prepare(s, reconcile(html, [edit], "p")));
  expect(s.serialize()).toEqual(before);
});

test("sequential evidence validates against preceding edits and one malformed proof rejects the whole batch", () => {
  const s = session(), before = s.serialize();
  const proof: ReconciledTextEdit[] = [
    { nodeId: "p", from: 1, to: 1, expectedText: "", text: "新" },
    { nodeId: "p", from: 2, to: 3, expectedText: "甲", text: "乙" },
  ];
  fails(() => prepare(s, reconcile('<span data-node-id="s">甲新乙甲</span>', [...proof, { nodeId: "s", from: 0, to: 1, expectedText: "wrong", text: "X" }], "p")));
  expect(s.serialize()).toEqual(before);
  s.commit(prepare(s, reconcile('<span data-node-id="s">甲新乙甲</span>', proof, "p")));
  expect(s.getNode("p")!.text).toBe("甲新乙甲");
});

test("mapping proofs share the operation budget and unsafe HTML remains rejected", () => {
  const s = session("甲甲甲", { sessionLimits: { maxOperations: 1 } });
  fails(() => prepare(s, reconcile('<span data-node-id="s">新甲甲甲</span>', [{ nodeId: "p", from: 0, to: 0, expectedText: "", text: "新" }], "p")), "ResourceLimitExceeded");
  for (const html of ['<span onclick="bad()">x</span>', '<img src="https://example.com/private">', '<span style="background-image:url(https://example.com/x)">x</span>']) {
    try { prepare(s, reconcile(html, [], "p")); throw new Error("Expected rejection"); }
    catch (error) { expect(error).toBeInstanceOf(DocumentError); }
  }
  expect(s.snapshot().manifest.revision).toBe(0);
});

test("unchanged reconciliation is a no-op and leaves the redo branch intact", () => {
  const s = session(); s.commit(prepare(s, { op: "setStyle", nodeId: "p", property: "color", value: "red" })); s.undo();
  const history = s.history();
  expect(s.commit(prepare(s, reconcile(s.getNode("p")!.html + s.getNode("q")!.html))).changed).toBe(false);
  expect(s.history()).toEqual(history); expect(s.redo().changed).toBe(true);
});

test("reconciliation preserves unchanged source font scopes instead of reclassifying their text", () => {
  const s = session("é");
  s.commit(prepare(s, reconcile('<span data-node-id="s" data-word-run="true" data-word-script="ascii" style="font-family:CustomFace">é</span>', [], "p")));
  expect(s.getNode("s")!.childNodeIds).toEqual([]);
  const original = s.getNode("s")!.html;
  s.commit(prepare(s, reconcile(s.getNode("p")!.html + '<p data-node-id="q">更新</p>')));
  expect(s.getNode("s")!.html).toBe(original);
});

test("editing an existing font scope classifies inserted text while preserving original characters", () => {
  const s = session("é");
  s.commit(prepare(s, reconcile('<span data-node-id="s" data-word-run="true" data-word-script="ascii" style="font-family:CustomFace;--word-font-ascii:CustomFace;--word-font-hAnsi:OtherFace">é</span>', [], "p")));
  const source = s.getNode("s")!.html;
  s.commit(prepare(s, reconcile(source.replace(">é<", ">éA<"), ["p", "s"].map((nodeId) => ({ nodeId, from: 1, to: 1, expectedText: "", text: "A" })), "p")));
  expect(s.getNode("s")!.childNodeIds).toEqual([]);
  s.commit(prepare(s, reconcile(s.getNode("s")!.html.replace(">éA<", ">éA中<"), ["p", "s"].map((nodeId) => ({ nodeId, from: 2, to: 2, expectedText: "", text: "中" })), "p")));
  expect(s.getNode("s")!.text).toBe("éA中");
  expect(s.getNode("s")!.html).not.toContain('data-word-script="hAnsi"');
  expect(s.getNode("s")!.html).toContain('data-word-script="eastAsia"');
});

test.each([["漢", "eastAsia", "\u{e0100}"], ["e", "ascii", "\u0301"]])("incremental marks inherit the base font scope: %s", (base, script, mark) => {
  const s = session(base);
  s.commit(prepare(s, reconcile('<span data-node-id="s" data-word-run="true" data-word-script="' + script + '" style="font-family:BaseFace;--word-font-hAnsi:OtherFace">' + base + '</span>', [], "p")));
  s.commit(prepare(s, reconcile(s.getNode("s")!.html.replace('>' + base + '<', '>' + base + mark + '<'),
    ["p", "s"].map((nodeId) => ({ nodeId, from: base.length, to: base.length, expectedText: "", text: mark })), "p")));
  expect(s.getNode("s")!.childNodeIds).toEqual([]);
  expect(s.getNode("s")!.text).toBe(base + mark);
  expect(s.getNode("s")!.attributes.style).toContain("font-family:BaseFace");
});

test("one composition's adjacent inserted base and selector keep a shared new font scope", () => {
  const s = session("");
  s.commit(prepare(s, reconcile('<span data-node-id="s" data-word-run="true" data-word-script="ascii"></span>', [], "p")));
  s.commit(prepare(s, reconcile('<span data-node-id="s" data-word-run="true" data-word-script="ascii">漢\u{e0100}</span>',
    ["p", "s"].flatMap((nodeId) => [
      { nodeId, from: 0, to: 0, expectedText: "", text: "漢" },
      { nodeId, from: 1, to: 1, expectedText: "", text: "\u{e0100}" },
    ]), "p")));
  const children = s.getNode("s")!.childNodeIds;
  expect(children).toHaveLength(1);
  expect(s.getNode(children[0]!)!.attributes["data-word-script"]).toBe("eastAsia");
  expect(s.getNode(children[0]!)!.text).toBe("漢\u{e0100}");
});
