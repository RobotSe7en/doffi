import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, test } from "vitest";
import { createAssetDescriptor, createDocumentSession, DocumentError } from "@doffi/core";
import type { Anchor, Annotation, DocumentEditingSession, DocumentOperation, DocumentPackage } from "@doffi/core";
import { repository } from "./fixtures.js";

const encode = (value: string) => new TextEncoder().encode(value);
function files(content: string, anchors: Anchor[], status: Annotation["status"] = "open"): DocumentPackage {
  const result = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(
    (path) => [path, new Uint8Array(readFileSync(join(repository, "examples/report", path)))],
  ));
  result.set("content.html", encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body>'
    + '<article data-node-id="root" data-office-kind="document">' + content + '</article></body></html>'));
  result.set("annotations.json", encode(JSON.stringify(anchors.map((anchor, index) => ({
    annotationId: "a" + index, documentId: "report-demo", baseRevision: 0, status, instruction: "检查", anchor,
  })))));
  return result;
}
const part = (nodeId: string, from: number, quote: string) => ({ nodeId, from, to: from + quote.length, quote });
const range = (nodeId: string, from: number, quote: string): Extract<Anchor, { kind: "text" }> => ({ kind: "text", nodeId, from, to: from + quote.length, quote });
const replace = (nodeId: string, from: number, expectedText: string, text: string): DocumentOperation => ({ op: "replaceText", nodeId, from, to: from + expectedText.length, expectedText, text });
function prepare(s: DocumentEditingSession, operations: DocumentOperation[]) {
  return s.prepare({ protocolVersion: 1, documentId: "report-demo", baseRevision: s.snapshot().manifest.revision, operations });
}
function session(content: string, anchors: Anchor[], status?: Annotation["status"]) {
  return createDocumentSession(openDocument(files(content, anchors, status)));
}
function expectHistory(s: DocumentEditingSession, operations: DocumentOperation[], expected: Record<string, unknown>) {
  const original = s.snapshot();
  const candidate = prepare(s, operations);
  expect(s.snapshot()).toEqual(original);
  expect(candidate.preview.annotations[0]).toMatchObject(expected);
  s.commit(candidate);
  expect(s.getAnnotations()[0]).toMatchObject(expected);
  expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  s.undo();
  expect(s.getAnnotations()[0]).toEqual({ ...original.annotations[0], baseRevision: 2 });
  expect(s.snapshot().html).toBe(original.html);
  s.redo();
  expect(s.getAnnotations()[0]).toMatchObject(expected);
  expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
}

describe("precise annotation mapping", () => {
  test.each([
    [0, "", "前", 3, "丙丁戊", "open"],
    [2, "", "前", 3, "丙丁戊", "open"],
    [5, "", "后", 2, "丙丁戊", "open"],
    [3, "", "中", 2, "丙中丁戊", "open"],
    [0, "甲乙", "前", 1, "丙丁戊", "open"],
    [5, "己庚", "后", 2, "丙丁戊", "open"],
    [3, "丁", "修改", 2, "丙修改戊", "open"],
    [3, "丁", "", 2, "丙戊", "open"],
    [1, "乙丙", "改", 1, "改丁戊", "open"],
    [4, "戊己", "改", 2, "丙丁改", "open"],
    [2, "丙丁戊", "改", 2, "丙丁戊", "needs_reanchor"],
    [2, "丙丁戊", "", 2, "丙丁戊", "needs_reanchor"],
    [1, "乙丙丁戊己", "改", 2, "丙丁戊", "needs_reanchor"],
  ] as const)("maps replacement at %s: %s → %s", (from, old, text, start, quote, status) => {
    const s = session('<p data-node-id="p">甲乙丙丁戊己庚</p>', [range("p", 2, "丙丁戊")]);
    expectHistory(s, [replace("p", from, old, text)], { status, anchor: range("p", start, quote) });
    expect(s.getAnnotations()[0]!.baseRevision).toBe(status === "needs_reanchor" ? 0 : 3);
  });

  test("preserves resolved state and maps sequential edits to ancestors and nested runs", () => {
    const s = session('<p data-node-id="p">甲<strong data-node-id="b">乙丙丁</strong>戊</p>', [range("p", 2, "丙丁"), range("b", 1, "丙丁")], "resolved");
    const candidate = prepare(s, [replace("b", 0, "", "前"), replace("p", 4, "丁", "新丁")]);
    expect(candidate.preview.annotations.map((a) => a.anchor)).toEqual([range("p", 3, "丙新丁"), range("b", 2, "丙新丁")]);
    expect(candidate.preview.annotations.every((a) => a.status === "resolved")).toBe(true);
    expect(candidate.changedAnnotationIds).toEqual(["a0", "a1"]);
    const receipt = s.commit(candidate);
    const undo = s.undo();
    expect(undo.mapping.textEdits.filter((item) => item.nodeId === "p")).toEqual([
      { nodeId: "p", from: 4, to: 6, insertedLength: 1 }, { nodeId: "p", from: 1, to: 2, insertedLength: 0 },
    ]);
    expect(s.getAnnotations().map((a) => a.anchor)).toEqual([range("p", 2, "丙丁"), range("b", 1, "丙丁")]);
    const redone = s.redo();
    expect(redone.mapping.textEdits).toEqual(receipt.mapping.textEdits);
    expect([...redone.mapping.nodes]).toEqual([...receipt.mapping.nodes]);
  });

  test("maps a range across multiple inline styles while preserving each node's coordinates", () => {
    const s = session('<p data-node-id="p"><strong data-node-id="b">甲乙</strong><em data-node-id="e">丙丁戊</em></p>', [range("p", 1, "乙丙丁"), range("e", 1, "丁戊")]);
    s.commit(prepare(s, [replace("p", 1, "乙丙", "新") ]));
    expect(s.getAnnotations().map((a) => a.anchor)).toEqual([range("p", 1, "新丁"), range("e", 0, "丁戊")]);
  });

  test.each([0, 1, 2, 3])("maps point anchor %s with following-text affinity", (point) => {
    const s = session('<p data-node-id="p">甲乙丙</p>', [range("p", point, "")]);
    s.commit(prepare(s, [replace("p", 1, "", "新")]));
    expect(s.getAnnotations()[0]!.anchor).toEqual(range("p", point < 1 ? point : point + 1, ""));
  });

  test("invalidates deleted interior points and preserves their source revision", () => {
    const s = session('<p data-node-id="p">甲乙丙丁</p>', [range("p", 2, "")]);
    expectHistory(s, [replace("p", 1, "乙丙", "新")], { status: "needs_reanchor", baseRevision: 0, anchor: range("p", 2, "") });
  });

  test("maps UTF-16 emoji, supplementary ideographs, variation sequences and line breaks", () => {
    const s = session('<p data-node-id="p">甲😀𠀀\u{e0100}<br>乙</p>', [range("p", 1, "😀𠀀\u{e0100}\n")]);
    expectHistory(s, [replace("p", 0, "", "前😀"), replace("p", 4, "😀", "中文\n")], { status: "open", anchor: range("p", 4, "中文\n𠀀\u{e0100}\n") });
  });

  test("structural image insertion shifts offsets by one U+FFFC and preserves existing object text", () => {
    const input = new Map(files('<p data-node-id="p">甲<span data-node-id="b">乙</span><br>丙</p>', [range("p", 1, "乙\n丙")]));
    const bytes = new Uint8Array(readFileSync(join(repository, "examples/fidelity/assets/sample.png")));
    const manifest = JSON.parse(new TextDecoder().decode(input.get("manifest.json")!));
    manifest.assets = [createAssetDescriptor("assets/a.png", "image/png", bytes)];
    input.set("manifest.json", encode(JSON.stringify(manifest)));
    input.set("assets/a.png", bytes);
    const s = createDocumentSession(openDocument(input));
    expectHistory(s, [{ op: "insertFragment", nodeId: "b", position: "before", html: '<img src="assets/a.png">' }], {
      status: "open", anchor: range("p", 2, "乙\n丙"),
    });
    expect(s.getNode("p")!.text).toBe("甲\ufffc乙\n丙");
  });

  test("maps selections across table cells without allowing content edits to cross cell boundaries", () => {
    const anchor: Anchor = { kind: "text-set", ranges: [part("a", 0, "甲乙"), part("b", 0, "丙丁")] };
    const s = session('<table data-node-id="t"><tr data-node-id="tr"><td data-node-id="a">甲乙</td><td data-node-id="b">丙丁</td></tr></table>', [anchor]);
    expectHistory(s, [replace("a", 0, "", "前"), replace("b", 1, "丁", "新丁")], {
      status: "open", anchor: { kind: "text-set", ranges: [part("a", 1, "甲乙"), part("b", 0, "丙新丁")] },
    });
  });

  test.each(["before", "after", "start", "end"] as const)("maps structural insertions at %s", (position) => {
    const s = session('<p data-node-id="p">甲<span data-node-id="b">乙丙</span>丁</p>', [range("p", 1, "乙丙")]);
    const nodeId = position === "before" || position === "after" ? "b" : "p";
    expectHistory(s, [{ op: "insertFragment", nodeId, position, html: "<span>新</span>" }], {
      status: "open", anchor: range("p", position === "before" || position === "start" ? 2 : 1, "乙丙"),
    });
  });

  test("maps normalized text around structural deletions and replacements, including empty nodes", () => {
    const s = session('<p data-node-id="p"><span data-node-id="a">甲</span><span data-node-id="empty"></span><span data-node-id="b">乙丙</span>丁</p>', [range("p", 1, "乙丙")]);
    s.commit(prepare(s, [{ op: "insertFragment", nodeId: "empty", position: "end", html: "<br>" },
      { op: "replaceFragment", nodeId: "a", html: "<strong>前文</strong>" }, { op: "removeNode", nodeId: "empty" }]));
    expect(s.getAnnotations()[0]).toMatchObject({ status: "open", anchor: range("p", 2, "乙丙") });
    expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  });

  test("invalidates wholesale fragment replacement even when the text is identical", () => {
    const s = session('<p data-node-id="p">甲乙丙</p>', [range("p", 1, "乙"), { kind: "node", nodeId: "p" }]);
    expectHistory(s, [{ op: "replaceFragment", nodeId: "p", html: "<strong>甲乙丙</strong>" }], { status: "needs_reanchor", anchor: range("p", 1, "乙") });
    expect(s.getAnnotations()[1]).toMatchObject({ status: "open", anchor: { kind: "node", nodeId: "p" } });
  });

  test.each(["replace", "remove-insert"])("does not infer ancestor text identity from matching text after structural %s", (kind) => {
    const s = session('<p data-node-id="p">甲<span data-node-id="b">乙</span>丙</p>', [range("p", 1, "乙"), range("p", 2, "丙")]);
    const operations: DocumentOperation[] = kind === "replace"
      ? [{ op: "replaceFragment", nodeId: "b", html: "<strong>乙</strong>" }]
      : [{ op: "insertFragment", nodeId: "b", position: "before", html: "<em>乙</em>" }, { op: "removeNode", nodeId: "b" }];
    expectHistory(s, operations, { status: "needs_reanchor", anchor: range("p", 1, "乙") });
    expect(s.getAnnotations()[1]).toMatchObject({ status: "open", anchor: range("p", 2, "丙") });
  });

  test("maps cross-node ranges and preserves their complete old anchor if any range is lost", () => {
    const anchor: Anchor = { kind: "text-set", ranges: [part("a", 1, "乙丙"), part("b", 0, "丁戊")] };
    const s = session('<p data-node-id="a">甲乙丙</p><p data-node-id="b">丁戊己</p>', [anchor]);
    const candidate = prepare(s, [replace("a", 0, "", "前"), replace("b", 1, "戊", "新戊")]);
    expect(candidate.preview.annotations[0]!.anchor).toEqual({ kind: "text-set", ranges: [part("a", 2, "乙丙"), part("b", 0, "丁新戊")] });
    s.discard(candidate);
    expectHistory(s, [replace("a", 0, "", "前"), { op: "removeNode", nodeId: "b" }], { status: "needs_reanchor", anchor, baseRevision: 0 });
  });

  test("maps node identities and invalidates a complete node set when any member is deleted", () => {
    const anchor: Anchor = { kind: "node-set", nodeIds: ["a", "b"] };
    const s = session('<p data-node-id="a">甲</p><p data-node-id="b">乙</p>', [anchor, { kind: "node", nodeId: "b" }]);
    s.commit(prepare(s, [replace("a", 0, "甲", "新甲")]));
    expect(s.getAnnotations()[0]).toMatchObject({ status: "open", baseRevision: 1, anchor });
    s.commit(prepare(s, [{ op: "removeNode", nodeId: "b" }]));
    expect(s.getAnnotations().every((a) => a.status === "needs_reanchor" && a.baseRevision === 1)).toBe(true);
    expect(s.getAnnotations()[0]!.anchor).toEqual(anchor);
    expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  });

  test("keeps no-net-change edits from moving or invalidating annotations and preserves redo", () => {
    const s = session('<p data-node-id="p">甲乙丙</p>', [range("p", 1, "乙")]);
    s.commit(prepare(s, [replace("p", 1, "乙", "新乙")]));
    s.undo();
    const before = s.snapshot();
    const history = s.history();
    const candidate = prepare(s, [replace("p", 0, "甲乙丙", "暂存"), replace("p", 0, "暂存", "甲乙丙")]);
    expect(candidate.preview).toEqual(before);
    expect(s.commit(candidate).changed).toBe(false);
    expect(s.history()).toEqual(history);
    expect(s.redo().changed).toBe(true);
  });

  test("text-set overlap caused by a partial replacement invalidates the entire old selection", () => {
    const anchor: Anchor = { kind: "text-set", ranges: [part("p", 1, "乙丙"), part("p", 5, "己庚")] };
    const s = session('<p data-node-id="p">甲乙丙丁戊己庚辛</p>', [anchor]);
    expectHistory(s, [replace("p", 2, "丙丁戊己", "新")], { status: "needs_reanchor", anchor });
  });

  test("restored text retains its anchor even when the batch also changes styling", () => {
    const s = session('<p data-node-id="p">甲乙丙</p>', [range("p", 1, "乙")]);
    s.commit(prepare(s, [replace("p", 0, "甲乙丙", "暂存"), replace("p", 0, "暂存", "甲乙丙"),
      { op: "setStyle", nodeId: "p", property: "color", value: "red" }]));
    expect(s.getAnnotations()[0]).toMatchObject({ status: "open", baseRevision: 1, anchor: range("p", 1, "乙") });
  });

  test("reanchoring cannot guess identities allocated by a preceding fragment operation", () => {
    const s = session('<p data-node-id="p">甲</p>', [range("p", 0, "甲")]);
    expect(() => prepare(s, [{ op: "insertFragment", nodeId: "p", position: "after", html: "<p>乙</p>" },
      { op: "updateAnnotation", annotationId: "a0", changes: { anchor: range("edit-1", 0, "乙") } }])).toThrow(DocumentError);
    expect(s.snapshot().manifest.revision).toBe(0);
  });

  test("activating stale recovery data cannot silently replace its quote during unrelated edits", () => {
    const s = session('<p data-node-id="p">新文</p><p data-node-id="q">旁文</p>', [range("p", 0, "旧文")], "needs_reanchor");
    const original = s.snapshot();
    expect(() => prepare(s, [{ op: "updateAnnotation", annotationId: "a0", changes: { status: "open" } },
      replace("q", 0, "旁文", "新旁文")])).toThrow(DocumentError);
    expect(s.snapshot()).toEqual(original);
  });

  test("changing status after a recovery anchor cannot bypass base-node identity checks", () => {
    const s = session('<p data-node-id="p">甲</p>', [range("p", 0, "甲")]);
    expect(() => prepare(s, [{ op: "insertFragment", nodeId: "p", position: "after", html: "<p>乙</p>" },
      { op: "updateAnnotation", annotationId: "a0", changes: { anchor: range("edit-1", 0, "乙"), status: "needs_reanchor" } },
      { op: "updateAnnotation", annotationId: "a0", changes: { status: "open" } }])).toThrow(DocumentError);
    expect(s.snapshot().manifest.revision).toBe(0);
  });

  test("explicit anchors use final candidate coordinates, while instruction/status edits still map", () => {
    const s = session('<p data-node-id="p">甲乙丙</p>', [range("p", 1, "乙"), range("p", 1, "乙")]);
    s.commit(prepare(s, [
      { op: "updateAnnotation", annotationId: "a0", changes: { anchor: range("p", 2, "乙"), status: "resolved" } },
      { op: "updateAnnotation", annotationId: "a1", changes: { instruction: "更新说明", status: "resolved" } },
      replace("p", 0, "", "前"),
    ]));
    expect(s.getAnnotations().map((a) => a.anchor)).toEqual([range("p", 2, "乙"), range("p", 2, "乙")]);
    expect(s.getAnnotations().every((a) => a.status === "resolved")).toBe(true);
  });

  test("explicit reanchoring and annotation addition validate final text, fail atomically and reject stale candidates", () => {
    const s = session('<p data-node-id="p">甲乙丙</p>', [range("p", 1, "乙")]);
    const original = s.snapshot();
    expect(() => prepare(s, [replace("p", 0, "甲", "新甲"), { op: "updateAnnotation", annotationId: "a0", changes: { anchor: range("p", 0, "错误") } }])).toThrow(DocumentError);
    expect(s.snapshot()).toEqual(original);
    expect(s.history().undoSteps).toBe(0);
    const stale = prepare(s, [replace("p", 0, "", "前")]);
    s.commit(prepare(s, [replace("p", 0, "甲乙丙", "新文"), { op: "updateAnnotation", annotationId: "a0", changes: { anchor: range("p", 0, "新文"), status: "open" } },
      { op: "addAnnotation", annotation: { annotationId: "new", documentId: "report-demo", baseRevision: 0, status: "open", instruction: "新增", anchor: range("p", 1, "文") } }]));
    expect(s.getAnnotations().every((a) => a.status === "open" && a.baseRevision === 1)).toBe(true);
    expect(() => s.commit(stale)).toThrow(expect.objectContaining({ code: "RevisionConflict" }));
    expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  });
});

const region = (nodeIds: string[] = []): Extract<Anchor, { kind: "region" }> => ({ kind: "region", surfaceId: "root", nodeIds, layoutKey: "flow-1", rect: { x: 12, y: 24, width: 50, height: 20 } });
describe("region validity and explicit relocation", () => {
  test.each([{ hits: [] }, { hits: ["p"] }])("retains recovery geometry for hits $hits after any source layout change", ({ hits }) => {
    const anchor = region(hits);
    const s = session('<p data-node-id="p">甲</p><p data-node-id="q">乙</p>', [anchor]);
    expectHistory(s, [{ op: "setStyle", nodeId: "q", property: "font-size", value: "24pt" }], { status: "needs_reanchor", anchor, baseRevision: 0 });
  });

  test("annotation-only changes keep region validity; explicit renderer relocation uses the same transaction and history", () => {
    const s = session('<p data-node-id="p">甲</p>', [region()]);
    s.commit(prepare(s, [{ op: "updateAnnotation", annotationId: "a0", changes: { instruction: "新说明" } }]));
    expect(s.getAnnotations()[0]).toMatchObject({ status: "open", baseRevision: 1, anchor: region() });
    s.commit(prepare(s, [replace("p", 0, "", "前")]));
    const relocated = { ...region(["p"]), layoutKey: "flow-2", pageId: "page-2", rect: { x: 10, y: 30, width: 60, height: 22 } };
    s.commit(prepare(s, [{ op: "updateAnnotation", annotationId: "a0", changes: { anchor: relocated, status: "open" } }]));
    expect(s.getAnnotations()[0]).toMatchObject({ status: "open", baseRevision: 3, anchor: relocated });
    s.undo();
    expect(s.getAnnotations()[0]).toMatchObject({ status: "needs_reanchor", baseRevision: 1, anchor: region() });
    s.redo();
    expect(s.getAnnotations()[0]).toMatchObject({ status: "open", baseRevision: 5, anchor: relocated });
    expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  });

  test.each(["open", "resolved"] as const)("recovering a region to %s requires an explicit measured anchor", (status) => {
    const s = session('<p data-node-id="p">甲</p>', [region()]);
    s.commit(prepare(s, [replace("p", 0, "", "前")]));
    const original = s.snapshot();
    const history = s.history();
    expect(() => prepare(s, [{ op: "updateAnnotation", annotationId: "a0", changes: { status } }])).toThrow(DocumentError);
    expect(s.snapshot()).toEqual(original);
    expect(s.history()).toEqual(history);
    const anchor = { ...region(["p"]), layoutKey: "flow-2" };
    s.commit(prepare(s, [{ op: "updateAnnotation", annotationId: "a0", changes: { anchor } },
      { op: "updateAnnotation", annotationId: "a0", changes: { status } }]));
    expect(s.getAnnotations()[0]).toMatchObject({ status, baseRevision: 2, anchor });
    s.undo();
    expect(s.getAnnotations()[0]).toEqual(original.annotations[0]);
    s.redo();
    expect(s.getAnnotations()[0]).toMatchObject({ status, baseRevision: 4, anchor });
  });

  test("matching text does not allow status-only recovery after a structural replacement", () => {
    const s = session('<p data-node-id="p">甲乙</p>', [range("p", 0, "甲乙")]);
    s.commit(prepare(s, [{ op: "replaceFragment", nodeId: "p", html: "<strong>甲乙</strong>" }]));
    const original = s.snapshot();
    expect(() => prepare(s, [{ op: "updateAnnotation", annotationId: "a0", changes: { status: "open" } },
      { op: "setStyle", nodeId: "p", property: "color", value: "red" }])).toThrow(DocumentError);
    expect(s.snapshot()).toEqual(original);
  });

  test("renderer environment changes can invalidate a region explicitly without changing document HTML", () => {
    const s = session('<p data-node-id="p">甲</p>', [region()]);
    const html = s.snapshot().html;
    s.commit(prepare(s, [{ op: "updateAnnotation", annotationId: "a0", changes: { status: "needs_reanchor" } }]));
    expect(s.snapshot().html).toBe(html);
    expect(s.getAnnotations()[0]).toMatchObject({ status: "needs_reanchor", baseRevision: 0, anchor: region() });
  });
});

describe("anchor source validation", () => {
  test.each([
    range("p", 2, "\ude00"), range("p", 3, "𠀀"),
    { kind: "text-set", ranges: [part("q", 0, "乙"), part("p", 0, "甲")] },
    { kind: "text-set", ranges: [part("p", 0, "甲😀"), part("p", 1, "😀")] },
    { kind: "text-set", ranges: [part("p", 0, "甲"), part("b", 0, "甲")] },
    { kind: "node-set", nodeIds: ["p", "p"] },
    { ...region(), surfaceId: "missing" },
    { ...region(), surfaceId: "p" },
    { ...region(["p", "p"]) },
    { ...region(["q"]), surfaceId: "container" },
  ] satisfies Anchor[])("rejects invalid active source anchor %j and preserves it as recovery data", (anchor) => {
    const content = '<div data-node-id="container"><p data-node-id="p"><span data-node-id="b">甲😀𠀀\u{e0100}</span></p></div><p data-node-id="q">乙</p>';
    expect(() => openDocument(files(content, [anchor]))).toThrow(expect.objectContaining({ code: "InvalidDocument" }));
    const recovery = openDocument(files(content, [anchor], "needs_reanchor"));
    expect(openDocument(recovery.serialize()).snapshot().annotations[0]!.anchor).toEqual(anchor);
  });

  test("accepts ordered empty-paragraph ranges sharing a normalized offset, and rejects reversed or duplicate points", () => {
    const content = '<p data-node-id="p"></p><p data-node-id="q"></p><p data-node-id="r">乙</p>';
    const ranges = [part("p", 0, ""), part("q", 0, ""), part("r", 0, "乙")];
    const document = openDocument(files(content, [{ kind: "text-set", ranges }]));
    expect(openDocument(document.serialize()).snapshot()).toEqual(document.snapshot());
    for (const invalid of [[...ranges].reverse(), [ranges[0]!, ranges[0]!]]) {
      expect(() => openDocument(files(content, [{ kind: "text-set", ranges: invalid }]))).toThrow(DocumentError);
    }
    const s = createDocumentSession(document);
    s.commit(prepare(s, [replace("r", 0, "", "前")]));
    expect(s.getAnnotations()[0]).toMatchObject({ status: "open", anchor: { kind: "text-set", ranges: [ranges[0], ranges[1], part("r", 1, "乙")] } });
  });

  test("ancestor-node text anchors use the content domains actually covered by their ranges", () => {
    const content = '<div data-node-id="main" data-office-story="main"><p data-node-id="p">甲乙</p></div>'
      + '<div data-node-id="comment" data-office-story="comment"><p data-node-id="q">丙丁</p></div>';
    for (const anchor of [range("root", 1, "乙丙"), { kind: "text-set", ranges: [part("root", 0, "甲"), part("root", 2, "丙")] } as Anchor]) {
      expect(() => openDocument(files(content, [anchor]))).toThrow(DocumentError);
    }
    const document = openDocument(files(content, [range("root", 0, "甲乙"), range("root", 2, "丙丁")]));
    expect(document.snapshot().annotations.every((a) => a.status === "open")).toBe(true);
    expect(openDocument(document.serialize()).snapshot()).toEqual(document.snapshot());
  });

  test("ancestor point anchors follow actual source stories, including the anchored node's end", () => {
    const content = '<div data-node-id="main" data-office-story="main"><p data-node-id="p">甲乙丙</p></div>'
      + '<div data-node-id="comment" data-office-story="comment"><p data-node-id="q">丁</p></div>';
    const anchors: Anchor[] = [
      { kind: "text-set", ranges: [part("root", 1, ""), part("root", 2, "丙")] },
      { kind: "text-set", ranges: [part("p", 0, "甲"), part("p", 3, "")] },
      { kind: "text-set", ranges: [part("q", 0, "丁"), part("root", 4, "")] },
    ];
    const document = openDocument(files(content, anchors));
    expect(openDocument(document.serialize()).snapshot()).toEqual(document.snapshot());
    expect(() => openDocument(files(content, [{ kind: "text-set", ranges: [part("p", 2, "丙"), part("root", 3, "")] }]))).toThrow(DocumentError);
  });

  test("rejects cross-story text sets", () => {
    const content = '<div data-node-id="main" data-office-story="main"><p data-node-id="p">甲</p></div>'
      + '<div data-node-id="comment" data-office-story="comment"><p data-node-id="q">乙</p></div>';
    expect(() => openDocument(files(content, [{ kind: "text-set", ranges: [part("p", 0, "甲"), part("q", 0, "乙")] }]))).toThrow(DocumentError);
  });
});
