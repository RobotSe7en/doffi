import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, test } from "vitest";
import { captureContentModel, createAssetDescriptor, createDocumentSession, createWordDocument, DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, parseContract } from "@doffi/core";
import type { CandidateChange, DocumentEditingSession, DocumentErrorCode, DocumentOperation, DocumentSessionOptions } from "@doffi/core";
import { annotations, proposal, repository } from "./fixtures.js";

const encode = (text: string) => new TextEncoder().encode(text);
function files(content?: string): Map<string, Uint8Array> {
  const result = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(
    (path) => [path, new Uint8Array(readFileSync(join(repository, "examples/report", path)))],
  ));
  if (content !== undefined) {
    result.set("content.html", encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body>'
      + '<article data-node-id="root" data-office-kind="document">' + content + '</article></body></html>'));
    result.set("annotations.json", encode("[]"));
  }
  return result;
}
function session(content?: string, options?: DocumentSessionOptions): DocumentEditingSession {
  return createDocumentSession(openDocument(files(content)), options);
}
function prepare(s: DocumentEditingSession, operations: DocumentOperation[]): CandidateChange {
  return s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations });
}
function text(nodeId: string, expectedText: string, replacement: string, from = 0): DocumentOperation {
  return { op: "replaceText", nodeId, from, to: from + expectedText.length, expectedText, text: replacement };
}
function failure(action: () => unknown, code: DocumentErrorCode = "InvalidOperation"): DocumentError {
  try { action(); } catch (error) {
    expect(error).toBeInstanceOf(DocumentError);
    expect((error as DocumentError).code).toBe(code);
    expect(parseContract("document-error", (error as DocumentError).toJSON())).toEqual((error as DocumentError).toJSON());
    return error as DocumentError;
  }
  throw new Error("Expected failure");
}

describe("isolated document transactions and history", () => {
  test("local text revisions retain the validated Word structure and exact ancestor offsets", () => {
    const input = new Map(createWordDocument({ documentId: "local-runs" }));
    const html = new TextDecoder().decode(input.get("content.html"));
    const source = html.replace(/(<p\b[^>]*>)[\s\S]*?(<\/p>)/,
      '$1<span data-node-id="left">甲</span><span data-node-id="right">乙</span>$2');
    input.set("content.html", encode(source));
    const s = createDocumentSession(openDocument(input));
    const original = s.snapshot(), first = prepare(s, [text("left", "", "中文", 1)]);
    expect(first.preview.metadata).toBe(original.metadata);
    s.commit(first);
    const second = prepare(s, [text("right", "乙", "<&")]);
    const root = original.metadata.word!.stories.find(story => story.kind === "main")!.rootNodeId;
    expect(second.mapping.textEdits).toContainEqual({ nodeId: root, from: 3, to: 4, insertedLength: 2 });
    s.commit(second);
    expect(openDocument(s.serialize()).getNode(root)?.text).toBe("甲中文<&");
    s.undo();
    expect(s.getNode(root)?.text).toBe("甲中文乙");
    s.undo();
    expect(s.snapshot().html).toBe(original.html);
    s.redo();
    expect(s.getNode(root)?.text).toBe("甲中文乙");
    const quoted = '<span data-node-id="left">甲</span>';
    input.set("content.html", encode(source.replace(quoted, "<!--" + quoted + "-->" + quoted)));
    const withComment = createDocumentSession(openDocument(input));
    withComment.commit(prepare(withComment, [text("left", "甲", "丙")]));
    expect(openDocument(withComment.serialize()).getNode("left")?.text).toBe("丙");
    expect(withComment.snapshot().html).toContain("<!--" + quoted + "-->");
    withComment.dispose();
    s.dispose();
  });
  test("many edits in a large file retain immutable text patches and reopen the latest source", () => {
    const input = new Map(createWordDocument({ documentId: "bounded-local-history" }));
    const original = new TextDecoder().decode(input.get("content.html"));
    input.set("content.html", encode(original.replace("</section>", "<!--" + "layout ".repeat(30_000) + "--></section>")
      .replace("</p>", "初</p>")));
    const model = openDocument(input);
    const initialBytes = [...model.serialize().values()].reduce((size, value) => size + value.length, 0);
    const s = createDocumentSession(model, { sessionLimits: { maxHistoryBytes: initialBytes + 64_000 } });
    for (let index = 0; index < 80; index++)
      s.commit(prepare(s, [text("paragraph-1", "", "字", index + 1)]));
    expect(s.history().undoSteps).toBe(80);
    expect(s.history().retainedBytes).toBeLessThan(initialBytes + 64_000);
    expect(openDocument(s.serialize()).getNode("paragraph-1")?.text).toBe("初" + "字".repeat(80));
    for (let index = 0; index < 80; index++) s.undo();
    expect(s.snapshot().html).toBe(model.snapshot().html);
    s.dispose();
  });
  test("captures private revisions and reuses unchanged CSS without relaxing package limits", () => {
    const input = files('<p data-node-id="p">甲</p>');
    input.set("styles.css", encode("p{color:blue}"));
    const s = createDocumentSession(openDocument(input), { limits: { maxCssNodes: 10 } });
    const previous = captureContentModel(s);
    expect(previous.snapshot()).toBe(s.snapshot());
    s.commit(prepare(s, [text("p", "甲", "乙")]));
    expect(previous.getNode("p")!.text).toBe("甲");
    expect(s.snapshot().styles[0]).toBe(previous.snapshot().styles[0]);
    failure(() => prepare(s, [{ op: "setStyle", nodeId: "p", property: "color", value: "red" }]), "ResourceLimitExceeded");
    s.undo();
    expect(s.snapshot().styles[0]).toBe(previous.snapshot().styles[0]);
    expect(s.getNode("p")!.text).toBe("甲");

    const unsafe = new Map(input);
    const source = new TextDecoder().decode(unsafe.get("content.html"));
    unsafe.set("content.html", encode(source.replace("</body>", "<script>0</script></body>")));
    expect(() => captureContentModel({ ...previous, serialize: () => unsafe })).toThrowError(DocumentError);

    const max = DEFAULT_DOCUMENT_READ_LIMITS.maxFileBytes;
    input.set("extra.bin", new Uint8Array(max + 1));
    const wide = openDocument(input, { limits: { maxFileBytes: max + 1 } });
    failure(() => captureContentModel(wide), "ResourceLimitExceeded");
    s.dispose();
  });
  test("preserves explicit carriage returns through text transactions, conflicts, history and reopening", () => {
    const s = session('<p data-node-id="p" title="甲&#x000D;乙">甲&#13;乙</p>');
    expect(s.getNode("p")!.text).toBe("甲\r乙");
    expect(s.getNode("p")!.attributes.title).toBe("甲\r乙");
    const first = prepare(s, [text("p", "甲", "丙😀")]);
    const stale = prepare(s, [text("p", "甲\r乙", "不能应用")]);
    s.commit(first);
    expect(s.getNode("p")!.text).toBe("丙😀\r乙");
    failure(() => s.commit(stale), "RevisionConflict");
    expect(openDocument(s.serialize()).getNode("p")!.text).toBe("丙😀\r乙");
    s.undo(); expect(s.getNode("p")!.text).toBe("甲\r乙");
    s.redo(); expect(s.getNode("p")!.text).toBe("丙😀\r乙");
    for (const reference of ["&#1;", "&#128;", "&#xD", "&#0;"]) expect(() => session('<p data-node-id="p">' + reference + '</p>')).toThrowError(DocumentError);
  });
  test("runs the report's content operations, preserves the source, and saves/reopens the result", () => {
    const source = openDocument(files());
    const original = source.snapshot();
    const s = createDocumentSession(source);
    const candidate = prepare(s, proposal.operations);
    expect(s.snapshot()).toEqual(original);
    expect(s.history()).toMatchObject({ undoSteps: 0, redoSteps: 0 });
    expect(candidate.preview.manifest.revision).toBe(0);
    expect(candidate.preview.html).toContain("第四季度");
    expect(candidate.changes).toEqual(expect.arrayContaining([
      { nodeId: "n-summary", kind: "modified" }, { nodeId: "n-title", kind: "modified" },
    ]));
    expect(Reflect.set(candidate.preview.manifest, "revision", 99)).toBe(false);
    expect("set" in candidate.mapping.nodes).toBe(false);
    const receipt = s.commit(candidate);
    expect(receipt).toMatchObject({ changed: true, revision: 1 });
    expect(s.getNode("n-title")!.attributes.style).toBe("color:#1f4b99");
    expect(s.snapshot().annotations[0]).toMatchObject({ status: "needs_reanchor", anchor: annotations[0]!.anchor });
    expect(source.snapshot()).toEqual(original);
    expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
    expect(s.undo().revision).toBe(2);
    expect(s.snapshot().html).toBe(original.html);
    expect(s.getAnnotations()[0]).toMatchObject({ status: "open", baseRevision: 2 });
    expect(s.redo().revision).toBe(3);
    expect(s.snapshot().html).toBe(candidate.preview.html);
    expect(s.history()).toMatchObject({ undoSteps: 1, redoSteps: 0 });
  });

  test("executes ranges against preceding candidate edits, and a later failure rolls back everything", () => {
    const s = session('<p data-node-id="p">甲乙丙</p>');
    const original = s.serialize();
    failure(() => prepare(s, [text("p", "甲", "AB"), text("p", "甲", "X")]));
    expect(s.serialize()).toEqual(original);
    const candidate = prepare(s, [text("p", "甲", "AB"), text("p", "乙", "丁", 2)]);
    s.commit(candidate);
    expect(s.getNode("p")!.text).toBe("AB丁丙");
    expect(candidate.mapping.textEdits.filter((item) => item.nodeId === "p")).toEqual([
      { nodeId: "p", from: 0, to: 1, insertedLength: 2 }, { nodeId: "p", from: 2, to: 3, insertedLength: 1 },
    ]);
  });

  test("no-op and reverted batches do not advance revisions, emit changes or clear redo", () => {
    const s = session('<p data-node-id="p" style="color:blue">原文</p>');
    const events: string[] = [];
    s.subscribe((event) => events.push(event.type));
    s.commit(prepare(s, [text("p", "原文", "新文")]));
    s.undo();
    const history = s.history();
    const revision = s.snapshot().manifest.revision;
    events.length = 0;
    const candidate = prepare(s, [text("p", "原文", "暂存"), text("p", "暂存", "原文"),
      { op: "setStyle", nodeId: "p", property: "color", value: "blue" }]);
    expect(s.commit(candidate)).toMatchObject({ changed: false, revision });
    expect(s.history()).toEqual(history);
    expect(events).toEqual([]);
    failure(() => s.commit(candidate));
    expect(s.redo().changed).toBe(true);
  });

  test("rejects stale, forged, repeated, discarded and cross-session candidates", () => {
    const s = session();
    const another = session();
    const first = prepare(s, proposal.operations);
    failure(() => another.commit(first));
    failure(() => s.commit({ ...first }));
    const second = prepare(s, proposal.operations);
    s.commit(first);
    failure(() => s.commit(first));
    failure(() => s.commit(second), "RevisionConflict");
    s.discard(second);
    failure(() => s.commit(second));
    failure(() => s.prepare({ protocolVersion: 1, documentId: "report-demo", baseRevision: 0, operations: [] }), "RevisionConflict");
  });

  test("new editing after undo drops the redo branch; snapshots and saved bytes are isolated", () => {
    const s = session('<p data-node-id="p">初稿</p>');
    const old = s.snapshot();
    s.commit(prepare(s, [text("p", "初稿", "二稿")]));
    s.undo();
    s.commit(prepare(s, [text("p", "初稿", "三稿")]));
    expect(s.history()).toMatchObject({ undoSteps: 1, redoSteps: 0 });
    expect(s.redo().changed).toBe(false);
    const saved = s.serialize();
    saved.get("content.html")!.fill(0);
    expect(old.html).toContain("初稿");
    expect(s.getNode("p")!.text).toBe("三稿");
  });

  test("preserves formatting outside the range and inherits the starting run's formatting", () => {
    const s = session('<p data-node-id="p"><strong data-node-id="b">甲乙</strong><em data-node-id="i">丙丁</em></p>');
    s.commit(prepare(s, [text("p", "乙丙", "新", 1)]));
    expect(s.getNode("b")!.text).toBe("甲新");
    expect(s.getNode("i")!.text).toBe("丁");
    expect(s.getNode("p")!.text).toBe("甲新丁");
    expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  });

  test("preserves UTF-16, combining/variation sequences and line breaks; rejects split surrogate ranges", () => {
    const s = session('<p data-node-id="p">甲😀𠀀\u{e0100}e\u0301<br>乙</p>');
    failure(() => prepare(s, [{ op: "replaceText", nodeId: "p", from: 2, to: 3, expectedText: "\ude00", text: "X" }]));
    failure(() => prepare(s, [text("p", "甲", "\ud800")]));
    failure(() => prepare(s, [text("p", "𠀀", "字", 3)]));
    s.commit(prepare(s, [text("p", "😀", "中文\nABC", 1)]));
    expect(s.getNode("p")!.text).toBe("甲中文\nABC𠀀\u{e0100}e\u0301\n乙");
    expect(openDocument(s.serialize()).getNode("p")!.text).toBe(s.getNode("p")!.text);
  });

  test("rejects text replacement across paragraphs and cells, root removal and damaged section references", () => {
    const s = session('<p data-node-id="p">甲</p><p data-node-id="q">乙</p>');
    failure(() => prepare(s, [text("root", "甲乙", "新文")]));
    failure(() => prepare(s, [{ op: "removeNode", nodeId: "root" }]));
    const table = session('<table data-node-id="t"><tr data-node-id="r"><td data-node-id="a">甲</td><td data-node-id="b">乙</td></tr></table>');
    failure(() => prepare(table, [text("t", "甲乙", "新文")]));
  });

  test.each([
    ['a<br>', 2, "a\nX"],
    ['<span data-node-id="run">a<br></span>', 2, "a\nX"],
    ['<br>', 1, "\nX"],
    ['a<br><br>', 3, "a\n\nX"],
    ['a<br>', 1, "aX\n"],
  ] as const)("inserts at the normalized break boundary: %s at %s", (content, offset, expected) => {
    const s = session('<p data-node-id="p">' + content + '</p>');
    const before = s.snapshot();
    const candidate = prepare(s, [text("p", "", "X", offset)]);
    expect(s.snapshot()).toEqual(before);
    expect(candidate.mapping.textEdits.filter((item) => item.nodeId === "p")).toEqual([
      { nodeId: "p", from: offset, to: offset, insertedLength: 1 },
    ]);
    const stale = prepare(s, [text("p", "", "Y", offset)]);
    s.commit(candidate);
    expect(s.getNode("p")!.text).toBe(expected);
    expect(openDocument(s.serialize()).getNode("p")!.text).toBe(expected);
    failure(() => s.commit(stale), "RevisionConflict");
    s.undo();
    expect(s.snapshot().html).toBe(before.html);
    s.redo();
    expect(s.getNode("p")!.text).toBe(expected);
    s.dispose();
  });

  test("inserts contextual fragments, assigns identities, replaces children and removes nodes atomically", () => {
    const s = session('<p data-node-id="p">甲</p><table data-node-id="t"><tbody><tr data-node-id="r"><td data-node-id="c">单元格</td></tr></tbody></table>');
    const candidate = prepare(s, [
      { op: "insertFragment", nodeId: "p", position: "after", html: '<p class="new">新段落<strong>粗体</strong></p>' },
      { op: "insertFragment", nodeId: "r", position: "after", html: "<tr><td>新行</td></tr>" },
      { op: "replaceFragment", nodeId: "p", html: "<em>替换</em>" },
      { op: "removeNode", nodeId: "r" },
    ]);
    s.commit(candidate);
    expect(s.getNode("root")!.text).toBe("替换新段落粗体新行");
    expect(s.getNode("r")).toBeUndefined();
    expect(candidate.mapping.nodes.get("r")).toBeNull();
    expect(s.getNode("p")!.childNodeIds[0]).toMatch(/^edit-/);
    s.undo();
    expect(s.getNode("r")!.text).toBe("单元格");
  });

  test.each([
    '<p data-node-id="chosen">非法身份</p>', '<script>alert(1)</script>', '<p onclick="bad()">事件</p>',
    '<img src="https://example.com/private.png">', '<p style="background-color:url(secret)">非法</p>',
  ])("rejects unsafe or externally identified fragments: %s", (html) => {
    const s = session();
    const before = s.serialize();
    expect(() => prepare(s, [{ op: "insertFragment", nodeId: "n-summary", position: "after", html }])).toThrow(DocumentError);
    expect(s.serialize()).toEqual(before);
  });

  test("rejects removed targets and guessing identities allocated within the same batch", () => {
    const s = session('<p data-node-id="p">原文</p>');
    failure(() => prepare(s, [{ op: "removeNode", nodeId: "p" }, text("p", "原文", "新文")]));
    failure(() => prepare(s, [{ op: "insertFragment", nodeId: "p", position: "after", html: "<p>新文</p>" }, text("edit-1", "新文", "猜测")]));
    // Unsafe intermediate HTML is rejected even when a later operation would remove it.
    failure(() => prepare(s, [{ op: "replaceFragment", nodeId: "p", html: '<span onclick="bad()">中间内容</span>' },
      { op: "replaceFragment", nodeId: "p", html: "安全结果" }]));
  });

  test("style changes preserve stylesheet bytes, unrelated declarations, variables and validate a single declaration", () => {
    const s = session('<p data-node-id="p" style="font-size:12pt;color:red">原文</p>');
    const styles = s.snapshot().styles;
    s.commit(prepare(s, [{ op: "setStyle", nodeId: "p", property: "color", value: "#123456" }]));
    expect(s.getNode("p")!.attributes.style).toBe("font-size:12pt;color:#123456");
    expect(s.snapshot().styles).toEqual(styles);
    failure(() => prepare(s, [{ op: "setStyle", nodeId: "p", property: "color", value: "red;font-size:99pt" }]));
    failure(() => prepare(s, [{ op: "setStyle", nodeId: "p", property: "color;bad", value: "red" }]));
  });

  test("places explicit longhands after shorthand declarations and roundtrips their order", () => {
    const s = session('<p data-node-id="p" style="margin-left:10pt;margin:20pt">原文</p>');
    s.commit(prepare(s, [{ op: "setStyle", nodeId: "p", property: "margin-left", value: "5pt" }]));
    expect(s.getNode("p")!.attributes.style).toBe("margin:20pt;margin-left:5pt");
    expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  });

  test("guards revision overflow without changing state or consuming the candidate", () => {
    const input = files();
    const value = JSON.parse(new TextDecoder().decode(input.get("manifest.json")!));
    value.revision = Number.MAX_SAFE_INTEGER;
    input.set("manifest.json", encode(JSON.stringify(value)));
    const s = createDocumentSession(openDocument(input));
    const original = s.serialize();
    const candidate = prepare(s, proposal.operations);
    failure(() => s.commit(candidate), "ResourceLimitExceeded");
    expect(s.serialize()).toEqual(original);
    expect(s.commit(prepare(s, [])).changed).toBe(false);
    s.discard(candidate);
  });

  test("replaces only declared image assets and preserves source dimensions and bytes", () => {
    const input = files('<p data-node-id="p"><img data-node-id="image" src="assets/a.png" width="10" height="20"></p>');
    const bytes = new Uint8Array(readFileSync(join(repository, "examples/fidelity/assets/sample.png")));
    const manifest = JSON.parse(new TextDecoder().decode(input.get("manifest.json")!));
    manifest.assets = ["assets/a.png", "assets/b.png"].map((path) => createAssetDescriptor(path, "image/png", bytes));
    input.set("manifest.json", encode(JSON.stringify(manifest)));
    input.set("assets/a.png", bytes); input.set("assets/b.png", bytes);
    const s = createDocumentSession(openDocument(input));
    failure(() => prepare(s, [{ op: "replaceAsset", nodeId: "image", assetPath: "assets/unknown.png" }]), "AssetUnavailable");
    s.commit(prepare(s, [{ op: "replaceAsset", nodeId: "image", assetPath: "assets/b.png" }]));
    expect(s.getNode("image")!.attributes).toMatchObject({ src: "assets/b.png", width: "10", height: "20" });
    expect(s.serialize().get("assets/a.png")).toEqual(bytes);
    expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  });

  test("annotation operations use the same history and reject mismatched quotes or identities", () => {
    const s = session();
    s.commit(prepare(s, [{ op: "updateAnnotation", annotationId: "a-heading", changes: { status: "resolved" } },
      { op: "removeAnnotation", annotationId: "a-summary" }]));
    expect(s.getAnnotations()).toHaveLength(1);
    expect(s.getAnnotations()[0]!.status).toBe("resolved");
    s.undo();
    expect(s.getAnnotations()).toHaveLength(2);
    failure(() => prepare(s, [{ op: "addAnnotation", annotation: { ...annotations[0]!, baseRevision: 2 } }]));
    s.commit(prepare(s, [{ op: "addAnnotation", annotation: { ...annotations[0]!, annotationId: "new", baseRevision: 2 } }]));
    expect(s.getAnnotations()).toHaveLength(3);
  });

  test("budgets reject excessive input, pending candidates and history without partial state changes", () => {
    const limited = session(undefined, { sessionLimits: { maxOperations: 1, maxCandidates: 1, maxHistorySteps: 1 } });
    failure(() => prepare(limited, proposal.operations), "ResourceLimitExceeded");
    const first = prepare(limited, [proposal.operations[0]!]);
    failure(() => prepare(limited, []), "ResourceLimitExceeded");
    limited.discard(first);
    limited.commit(prepare(limited, [proposal.operations[0]!]));
    const next = prepare(limited, [proposal.operations[1]!]);
    const before = limited.serialize();
    failure(() => limited.commit(next), "ResourceLimitExceeded");
    expect(limited.serialize()).toEqual(before);
    limited.clearHistory();
    limited.commit(next);
    expect(limited.snapshot().manifest.revision).toBe(2);
    failure(() => prepare(session(undefined, { sessionLimits: { maxTransactionBytes: 16 } }), []), "ResourceLimitExceeded");
    failure(() => prepare(session(undefined, { sessionLimits: { maxCandidateBytes: 1 } }), []), "ResourceLimitExceeded");
    const history = session(undefined, { sessionLimits: { maxHistoryBytes: 1 } });
    failure(() => history.commit(prepare(history, [proposal.operations[0]!])), "ResourceLimitExceeded");
    expect(history.history().undoSteps).toBe(0);
    const nodes = session('<p data-node-id="p"><span data-node-id="run" data-word-run="true">ABC</span></p>', { limits: { maxNodes: 30 } });
    failure(() => prepare(nodes, [text("run", "ABC", "中A".repeat(100))]), "ResourceLimitExceeded");
    expect(nodes.getNode("run")!.text).toBe("ABC");
  });

  test("events follow committed state; callback errors cannot undo success, and disposal releases the session", () => {
    const s = session();
    const events: string[] = [];
    const unsubscribe = s.subscribe((event) => {
      expect(s.snapshot().manifest.revision).toBe(event.revision);
      events.push(event.type);
      if (event.type === "documentChanged") failure(() => s.undo());
      throw new Error("consumer error");
    });
    s.commit(prepare(s, proposal.operations));
    expect(events).toEqual(["documentChanged", "annotationChanged"]);
    unsubscribe();
    const candidate = prepare(s, []);
    s.dispose(); s.dispose();
    failure(() => s.snapshot()); failure(() => s.commit(candidate)); failure(() => s.serialize());
  });
});
