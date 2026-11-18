import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { join } from "node:path";
import { describe, expect, expectTypeOf, test } from "vitest";
import { createAssetDescriptor, DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, parseContract } from "@doffi/core";
import type { DocumentErrorCode, ContentDocumentModel as DocumentModel, DocumentPackage, DocumentReadLimits } from "@doffi/core";
import { annotations, manifest, metadata, repository } from "./fixtures.js";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
function files(): Map<string, Uint8Array> {
  return new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(
    (path) => [path, new Uint8Array(readFileSync(join(repository, "examples/report", path)))],
  ));
}
function json(input: Map<string, Uint8Array>, path: string, value: unknown): void {
  input.set(path, encoder.encode(JSON.stringify(value)));
}
function body(content: string): Map<string, Uint8Array> {
  const input = files();
  input.set("content.html", encoder.encode('<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">'
    + '<title>测试</title><link rel="stylesheet" href="styles.css"></head><body>'
    + '<article data-node-id="root" data-office-kind="document">' + content + '</article></body></html>'));
  json(input, "annotations.json", []);
  return input;
}
function failure(input: DocumentPackage, code: DocumentErrorCode = "InvalidDocument", limits?: Partial<DocumentReadLimits>): DocumentError {
  try {
    openDocument(input, limits ? { limits } : {});
  } catch (error) {
    if (!(error instanceof DocumentError)) throw error;
    expect(error.code).toBe(code);
    expect(parseContract("document-error", error.toJSON())).toEqual(error.toJSON());
    return error;
  }
  throw new Error("Expected opening the document to fail.");
}

describe("document package reading and serialization", () => {
  test("opens the report and preserves document identity, revisions, styles and annotations on reopen", () => {
    const input = files();
    const before = new Map([...input].map(([path, bytes]) => [path, [...bytes]]));
    const document = openDocument(input);
    expectTypeOf(document).toEqualTypeOf<DocumentModel>();
    expect(document.snapshot()).toMatchObject({ manifest, metadata, annotations });
    expect(document.getNode("n-summary")).toMatchObject({
      nodeId: "n-summary", tagName: "p", parentNodeId: "root", text: "项目计划于第三季度完成。",
    });
    expect(document.getNode("root")!.childNodeIds).toEqual(["n-title", "n-summary", "n-deliverables"]);
    expect(document.getNode("n-deliverables")!.childNodeIds).toEqual(["n-header-row", "n-body-row"]);
    expect(document.getNode("missing")).toBeUndefined();
    const saved = document.serialize();
    const reopened = openDocument(saved);
    expect(reopened.snapshot()).toEqual(document.snapshot());
    expect(reopened.getNode("n-deliverables")).toEqual(document.getNode("n-deliverables"));
    expect(reopened.serialize()).toEqual(saved);
    expect(new Map([...input].map(([path, bytes]) => [path, [...bytes]]))).toEqual(before);
  });

  test("isolates the model from caller-owned bytes, snapshots, node views and serialized files", () => {
    const input = files();
    const document = openDocument(input);
    const snapshot = document.snapshot();
    const node = document.getNode("n-summary")!;
    input.get("content.html")!.fill(0);
    input.clear();
    expect(Reflect.set(snapshot.manifest, "revision", 99)).toBe(false);
    expect(Reflect.set(snapshot.metadata.pageSetup!.margins, "top", 0)).toBe(false);
    expect(Reflect.set(snapshot.annotations[0]!, "instruction", "changed")).toBe(false);
    expect(Reflect.set(node.attributes, "style", "color:red")).toBe(false);
    expect(Reflect.set(node.childNodeIds, "0", "changed")).toBe(false);
    const first = document.serialize() as Map<string, Uint8Array>;
    first.get("content.html")!.fill(0);
    first.clear();
    expect(document.snapshot().manifest.revision).toBe(0);
    expect(document.getNode("n-summary")!.text).toBe("项目计划于第三季度完成。");
    expect(openDocument(document.serialize()).snapshot()).toEqual(snapshot);
  });

  test("keeps separate opened documents and unreferenced binary files isolated", () => {
    const input = files();
    input.set("sources/original.bin", new Uint8Array([0, 255, 7]));
    const first = openDocument(input);
    const second = openDocument(input);
    expect(first.snapshot()).not.toBe(second.snapshot());
    const saved = first.serialize();
    expect([...saved.get("sources/original.bin")!]).toEqual([0, 255, 7]);
    saved.get("sources/original.bin")!.fill(1);
    expect([...second.serialize().get("sources/original.bin")!]).toEqual([0, 255, 7]);
  });

  test.each([0, 2, 99])("rejects format version %i without changing bytes", (formatVersion) => {
    const input = files();
    json(input, "manifest.json", { ...manifest, formatVersion });
    const before = [...input.get("manifest.json")!];
    failure(input, "UnsupportedVersion");
    expect([...input.get("manifest.json")!]).toEqual(before);
  });

  test.each(["presentation", "workbook"])("rejects unsupported kinds or metadata belonging to another kind: %s", (kind) => {
    const input = files();
    json(input, "manifest.json", { ...manifest, kind });
    // A workbook now requires a chunks manifest and workbook metadata.
    failure(input, "InvalidDocument");
  });

  test.each(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"])("rejects missing %s", (path) => {
    const input = files();
    input.delete(path);
    failure(input);
  });

  test("rejects malformed JSON and UTF-8 without exposing raw input", () => {
    const input = files();
    input.set("metadata.json", encoder.encode('{"private text":'));
    expect(JSON.stringify(failure(input))).not.toContain("private text");
    input.set("metadata.json", new Uint8Array([0xc0, 0xaf]));
    failure(input);
  });

  test.each(["../outside", "a\\b", "/absolute", "a\u2028/../outside"])("rejects unsafe package filename %j", (path) => {
    const input = files();
    input.set(path, new Uint8Array());
    failure(input);
  });

  test("rejects conflicting file roles and repeated resource declarations", () => {
    const input = files();
    for (const changed of [
      { ...manifest, metadata: "manifest.json" },
      { ...manifest, styles: ["styles.css", "styles.css"] },
      { ...manifest, assets: [{ path: "content.html", mimeType: "image/png", sha256: "a".repeat(64) }] },
    ]) {
      json(input, "manifest.json", changed);
      failure(input);
    }
  });

  test.each([false, true])("rejects file ancestors across all package paths in either order: %s", (reverse) => {
    const input = files();
    input.set("content.html/metadata.json", input.get("metadata.json")!);
    json(input, "manifest.json", { ...manifest, metadata: "content.html/metadata.json" });
    failure(reverse ? new Map([...input].reverse()) : input);

    const unreferenced = files();
    unreferenced.set("sources/original", new Uint8Array());
    unreferenced.set("sources/original-copy", new Uint8Array());
    unreferenced.set("sources/original/metadata.json", new Uint8Array());
    failure(reverse ? new Map([...unreferenced].reverse()) : unreferenced);
  });

  test("preserves sibling prefixes and deeply nested unreferenced package files", () => {
    const input = files();
    for (const path of ["sources/a", "sources/a-copy", "sources/ab/item", "a/".repeat(32_000) + "b"]) {
      input.set(path, new Uint8Array([1]));
    }
    const document = openDocument(input);
    expect(openDocument(document.serialize()).snapshot()).toEqual(document.snapshot());
    for (const [path, bytes] of input) {
      if (path.startsWith("sources/") || path.length > 64_000) expect(document.serialize().get(path)).toEqual(bytes);
    }
  });

  test.each([
    [{ sectionId: "s", rootNodeId: "missing" }],
    [{ sectionId: "s", rootNodeId: "n-summary" }],
    [{ sectionId: "s", rootNodeId: "root" }, { sectionId: "s", rootNodeId: "root" }],
  ])("rejects invalid metadata section references %j", (...sections) => {
    const input = files();
    json(input, "metadata.json", { ...metadata, sections });
    failure(input);
  });

  test("checks annotation ownership, identity and future revisions", () => {
    const input = files();
    for (const changed of [
      [annotations[0], annotations[0]],
      [{ ...annotations[0], documentId: "another-document" }],
      [{ ...annotations[0], baseRevision: 1 }],
    ]) {
      json(input, "annotations.json", changed);
      failure(input);
    }
  });

  test("preserves unresolved anchors without inventing replacements or changing revisions", () => {
    const input = files();
    const unresolved = [{ ...annotations[0], status: "needs_reanchor", anchor: { kind: "node", nodeId: "deleted" } }];
    json(input, "annotations.json", unresolved);
    const document = openDocument(input);
    expect(document.snapshot().annotations).toEqual(unresolved);
    expect(document.snapshot().manifest.revision).toBe(0);
  });

  test.each(["open", "resolved"])("rejects missing anchor nodes for %s annotations", (status) => {
    for (const anchor of [
      { kind: "text", nodeId: "deleted", from: 0, to: 0, quote: "" },
      { kind: "text-set", ranges: [{ nodeId: "n-summary", from: 0, to: 0, quote: "" }, { nodeId: "deleted", from: 0, to: 0, quote: "" }] },
      { kind: "node", nodeId: "deleted" },
      { kind: "node-set", nodeIds: ["n-summary", "deleted"] },
      { kind: "region", surfaceId: "root", rect: { x: 0, y: 0, width: 1, height: 1 }, nodeIds: ["deleted"], layoutKey: "flow" },
    ]) {
      const input = files();
      json(input, "annotations.json", [{ ...annotations[0], status, anchor }]);
      failure(input);
    }
  });

  test("accepts existing anchor nodes and empty region node sets", () => {
    const input = files();
    const savedAnnotations = [
      { ...annotations[0], annotationId: "text-set", anchor: { kind: "text-set", ranges: [{ nodeId: "n-summary", from: 0, to: 2, quote: "项目" }] } },
      { ...annotations[0], annotationId: "node-set", anchor: { kind: "node-set", nodeIds: ["n-title", "n-summary"] } },
      { ...annotations[0], annotationId: "region", anchor: { kind: "region", surfaceId: "root", rect: { x: 0, y: 0, width: 1, height: 1 }, nodeIds: [], layoutKey: "flow" } },
    ];
    json(input, "annotations.json", savedAnnotations);
    const document = openDocument(input);
    expect(openDocument(document.serialize()).snapshot().annotations).toEqual(savedAnnotations);
  });
});

describe("HTML model and resource references", () => {
  test("preserves inline structure, combining characters, emoji, line breaks, images and table spans", () => {
    const input = body('<p data-node-id="p" class="text" data-custom="保留" aria-label="正文" style="color: #123456">'
      + '中e\u0301😀<strong data-node-id="s">粗体</strong><br><img data-node-id="i" src="assets/%E5%9B%BE%20%E7%89%87.png" width="1" height="2" alt="图"></p>'
      + '<ol data-node-id="list" start="3" reversed><li data-node-id="li" value="4">列表</li></ol>'
      + '<table data-node-id="t"><tr data-node-id="tr"><td data-node-id="td" colspan="2" rowspan="2">单元格</td></tr></table>');
    const image = new Uint8Array(readFileSync(join(repository, "examples/fidelity/assets/sample.png")));
    json(input, "manifest.json", { ...manifest, assets: [createAssetDescriptor("assets/图 片.png", "image/png", image)] });
    input.set("assets/图 片.png", image);
    const document = openDocument(input);
    expect(document.getNode("p")!.text).toBe("中e\u0301😀粗体\n\ufffc");
    expect(document.getNode("p")!.text.length).toBe(9);
    expect(document.getNode("p")!.childNodeIds).toEqual(["s", "i"]);
    expect(document.getNode("p")!.attributes).toMatchObject({ "data-custom": "保留", "aria-label": "正文", style: "color:#123456" });
    expect(document.getNode("td")!.attributes).toMatchObject({ colspan: "2", rowspan: "2" });
    expect(document.snapshot().html).toContain("<tbody>");
    expect(openDocument(document.serialize()).snapshot()).toEqual(document.snapshot());
    expect(document.serialize().get("assets/图 片.png")).toEqual(image);
  });

  test("treats opaque node identities as map keys", () => {
    const document = openDocument(body('<p data-node-id="__proto__">one</p><p data-node-id="constructor">two</p>'));
    expect(document.getNode("__proto__")!.text).toBe("one");
    expect(document.getNode("constructor")!.text).toBe("two");
  });

  test.each([
    '<p>missing identity</p>',
    '<p data-node-id="root">duplicate identity</p>',
    '<p data-node-id="">empty identity</p>',
    '<p data-node-id="a" data-node-id="b">duplicate attribute</p>',
    '<p data-node-id="a" onclick="private text">event handler</p>',
    '<img data-node-id="a" src="javascript:alert(1)">',
    '<img data-node-id="a" src="jav&#x61;script:alert(1)">',
    '<img data-node-id="a" src="https://example.test/image.png">',
    '<img data-node-id="a" src="//example.test/image.png">',
    '<img data-node-id="a" src="data:image/png;base64,AAAA">',
    '<img data-node-id="a" src="../outside.png">',
    '<img data-node-id="a" src="%2e%2e/outside.png">',
    '<img data-node-id="a" src="assets%2fimage.png">',
    '<img data-node-id="a" src="assets/%0aimage.png">',
    '<img data-node-id="a" src="assets/%invalid.png">',
    '<img data-node-id="a" src="assets/missing.png">',
  ])("rejects invalid HTML or references: %s", (content) => { failure(body(content)); });

  test.each([
    '<script>private text</script>', '<iframe src="https://example.test"></iframe>',
    '<frame>', '<frame src="https://example.test" onload="private text">',
    '<frameset><frame src="https://example.test"></frameset>',
    '<svg><script>private text</script></svg>', '<object></object>', '<style>p{color:red}</style>',
    '<p data-node-id="a" contenteditable="true">editor state</p>',
    '<img data-node-id="a" src="assets/a.png" srcset="https://example.test/a.png 2x">',
  ])("rejects unsupported structure rather than deleting it: %s", (content) => {
    const input = body(content);
    const before = [...input.get("content.html")!];
    expect(JSON.stringify(failure(input, "UnsupportedFeature"))).not.toContain("private text");
    expect([...input.get("content.html")!]).toEqual(before);
  });

  test.each(["td", "th", "tr", "thead", "tbody", "tfoot"])("rejects a misplaced %s that parse5 would silently discard", (tag) => {
    const input = body('<' + tag + ' data-node-id="discarded" style="color:red">private text</' + tag + '>');
    const before = [...input.get("content.html")!];
    expect(JSON.stringify(failure(input))).not.toContain("private text");
    expect([...input.get("content.html")!]).toEqual(before);
  });

  test("preserves optional end tags and tag-like text in title and comments", () => {
    const input = body('<p data-node-id="one">one<p data-node-id="two">two<!--<frame onload="text">-->');
    const source = decoder.decode(input.get("content.html")!)
      .replace("<title>测试</title>", '<title><frame src="https://example.test"> &lt;script&gt;</title>');
    input.set("content.html", encoder.encode(source));
    const document = openDocument(input);
    expect(document.getNode("one")!.text).toBe("one");
    expect(document.getNode("two")!.text).toBe("two");
    expect(document.snapshot().html).toContain('<title>&lt;frame src="https://example.test"&gt; &lt;script&gt;</title>');
    expect(openDocument(document.serialize()).snapshot()).toEqual(document.snapshot());
  });

  test("rejects missing declared image bytes", () => {
    const input = body('<img data-node-id="image" src="assets/image.png">');
    json(input, "manifest.json", { ...manifest, assets: [{ path: "assets/image.png", mimeType: "image/png", sha256: "a".repeat(64) }] });
    failure(input, "AssetUnavailable");
  });

  test("resolves encoded references relative to a nested HTML entrypoint", () => {
    const input = files();
    const original = decoder.decode(input.get("content.html")!);
    input.delete("content.html");
    input.set("正文/index.html", encoder.encode(original.replace('href="styles.css"', 'href="../styles.css"')));
    json(input, "manifest.json", { ...manifest, entrypoint: "正文/index.html" });
    const document = openDocument(input);
    expect(openDocument(document.serialize()).snapshot()).toEqual(document.snapshot());
  });

  test("requires stylesheet links to agree with the manifest and its order", () => {
    const input = files();
    const original = decoder.decode(input.get("content.html")!);
    input.set("content.html", encoder.encode(original.replace('href="styles.css"', 'href="https://example.test/styles.css"')));
    failure(input);
    input.set("other.css", encoder.encode("p { color: red }"));
    json(input, "manifest.json", { ...manifest, styles: ["other.css", "styles.css"] });
    input.set("content.html", encoder.encode(original.replace('<link rel="stylesheet" href="styles.css">',
      '<link rel="stylesheet" href="styles.css"><link rel="stylesheet" href="other.css">')));
    failure(input);
  });
});

describe("CSS support and validation", () => {
  test("preserves selector semantics, variable references, declaration order and inline overrides", () => {
    const input = body('<p data-node-id="p" class="x" style="color: blue !important; margin: 2pt">测试</p>');
    input.set("styles.css", encoder.encode(':root{--正文色:#123456;--spacing:2pt} article .x[data-node-id="p"], p {'
      + 'color:red;color:var(--正文色, blue);margin:calc(2pt + var(--spacing));font-family:"中文字体",sans-serif}'));
    const document = openDocument(input);
    expect(document.snapshot().styles[0]!.css).toContain("color:red;color:var(--正文色,blue)");
    expect(document.getNode("p")!.attributes.style).toBe("color:blue!important;margin:2pt");
    expect(openDocument(document.serialize()).snapshot()).toEqual(document.snapshot());
  });

  test.each([
    "p { color: definitely-not-a-color }", "p { width: red }", "p { color: }", "p { broken }",
    "p { color: var(ordinary) }", "p { span {color:red} }",
    "p { --asset: url(https://example.test/a); color: var(--asset) }",
    "p { color: url(https://example.test/a) }", "p { color: u\\72l(https://example.test/a) }",
  ])("rejects malformed or unsafe CSS: %s", (css) => {
    const input = files();
    input.set("styles.css", encoder.encode(css));
    failure(input);
  });

  test.each([
    "@import 'https://example.test/style.css';", "@font-face {font-family:x;src:url(x)}",
    "@media print {p{color:red}}", "p { position: fixed }", "p { color: expression(alert(1)) }",
    "p:hover { color: red }", "p::before { color: red }", "article > p {color:red}",
    "#p {color:red}", "[href] {color:red}",
  ])("reports unsupported CSS instead of silently dropping it: %s", (css) => {
    const input = files();
    input.set("styles.css", encoder.encode(css));
    failure(input, "UnsupportedFeature");
  });

  test("applies the same style policy to HTML inline declarations", () => {
    failure(body('<p data-node-id="p" style="color: url(https://example.test/a)">unsafe</p>'));
    failure(body('<p data-node-id="p" style="position:fixed">unsupported</p>'), "UnsupportedFeature");
  });
});

describe("bounded document input", () => {
  test.each([
    ["annotations.json", "JSON.stringify(Array.from({ length: 100_000 }, () => ({})))", "InvalidDocument"],
    ["styles.css", "'p{color:red}'.repeat(200_000)", "ResourceLimitExceeded"],
  ])("rejects oversized structure in %s within a 128 MiB heap", (path, source, code) => {
    const output = execFileSync(process.execPath, ["--max-old-space-size=128", "--input-type=module", "-e", `
      import { readFileSync } from "node:fs";
      import { DocumentError, openDocument } from "./packages/core/dist/index.js";
      const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"]
        .map(path => [path, readFileSync("examples/report/" + path)]));
      files.set(${JSON.stringify(path)}, new TextEncoder().encode(${source}));
      try { openDocument(files); throw new Error("Expected a bounded failure."); }
      catch (error) {
        if (!(error instanceof DocumentError) || error.code !== ${JSON.stringify(code)}) throw error;
        console.log(error.code);
      }
    `], { cwd: repository, encoding: "utf8", timeout: 15_000 });
    expect(output.trim()).toBe(code);
  });

  test.each([
    { maxFiles: 4 }, { maxPackageBytes: 10 }, { maxFileBytes: 10 },
    { maxNodes: 10 }, { maxDepth: 3 }, { maxCssNodes: 3 },
  ])("reports the configured limit exceeded by %j", (limits) => {
    expect(failure(files(), "ResourceLimitExceeded", limits).message).toContain(Object.keys(limits)[0]);
  });

  test("counts CSS budgets across files and inline styles", () => {
    const input = body('<p data-node-id="p" style="color:red">正文</p>');
    input.set("styles.css", encoder.encode("p{color:blue}"));
    // Each input alone is below ten AST nodes; the combined document is above it.
    expect(failure(input, "ResourceLimitExceeded", { maxCssNodes: 10 }).message)
      .toBe("Document input exceeds maxCssNodes (12 > 10) in CSS.");
  });

  test("shares identical Word inline declarations while budgeting every use", () => {
    const input = body(Array.from({ length: 20 }, (_, index) =>
      `<p data-node-id="inline-${index}" style="color: red">text</p>`).join(""));
    input.set("styles.css", encoder.encode(""));
    const limits = { maxCssNodes: 24 };
    const opened = openDocument(input, { limits });
    expect(openDocument(opened.serialize(), { limits }).snapshot()).toEqual(opened.snapshot());
    expect(failure(input, "ResourceLimitExceeded", { maxCssNodes: 23 }).message)
      .toBe("Document input exceeds maxCssNodes (24 > 23) in inline CSS.");
  });

  test.each([
    ["p{color:red}", 9],
    ['p[data-x="v"] p{color:red}', 14],
    ["p{--x: }", 9],
  ] as const)("enforces exact CSS node and depth budgets including implicit nodes: %s", (css, maxCssNodes) => {
    const input = body("");
    input.set("styles.css", encoder.encode(css));
    const limits = { maxCssNodes, maxDepth: 6 };
    const document = openDocument(input, { limits });
    expect(openDocument(document.serialize(), { limits }).snapshot()).toEqual(document.snapshot());
    failure(input, "ResourceLimitExceeded", { ...limits, maxCssNodes: maxCssNodes - 1 });
    failure(input, "ResourceLimitExceeded", { ...limits, maxDepth: 5 });
  });

  test.each([
    ["stylesheet", "p{color:red}".repeat(1000) + "p{broken}"],
    ["declarationList", "color:red;".repeat(1000) + "broken"],
  ])("stops CSS tree construction before parsing the tail in %s", (context, css) => {
    const input = context === "stylesheet" ? body("") : body('<p data-node-id="p" style="' + css + '">text</p>');
    input.set("styles.css", encoder.encode(context === "stylesheet" ? css : ""));
    failure(input, "ResourceLimitExceeded", { maxCssNodes: 64 });
    // Failed parser state must not affect a subsequent document.
    expect(() => openDocument(files())).not.toThrow();
  });

  test("stops deeply nested CSS functions before overflowing the parser stack", () => {
    const input = body("");
    input.set("styles.css", encoder.encode("p{--x:" + "var(--x,".repeat(30_000) + "red" + ")".repeat(30_000) + "}"));
    failure(input, "ResourceLimitExceeded");
    expect(() => openDocument(files())).not.toThrow();
  });

  test("stops deeply nested HTML before parsing the remaining source", () => {
    const input = body(Array.from({ length: 30_000 }, (_, index) => '<div data-node-id="n' + index + '">').join("")
      + "text" + "</div>".repeat(30_000) + '<br title="one" title="malformed tail">');
    input.set("styles.css", encoder.encode(""));
    // A complete parse would reach the duplicate attribute and report InvalidDocument.
    failure(input, "ResourceLimitExceeded");
  });

  test.each(["<br>", "<!--comment-->", "text<br>"])("stops excess HTML nodes before parsing the tail: %s", (content) => {
    const input = body(content.repeat(2_000) + '<br title="one" title="malformed tail">');
    input.set("styles.css", encoder.encode(""));
    failure(input, "ResourceLimitExceeded", { maxNodes: 64 });
  });

  test.each([
    ["", 10, 4],
    ["one two", 11, 4],
    ["<!--one--><!--two-->", 12, 4],
    ["<br>", 11, 4],
    ['<table data-node-id="t"><tr data-node-id="r"><td data-node-id="c">text</td></tr></table>', 15, 8],
  ] as const)("counts exact HTML node and depth limits including implicit nodes: %s", (content, maxNodes, maxDepth) => {
    const input = body(content);
    input.set("styles.css", encoder.encode(""));
    const limits = { maxNodes, maxDepth };
    const document = openDocument(input, { limits });
    expect(openDocument(document.serialize(), { limits }).snapshot()).toEqual(document.snapshot());
    failure(input, "ResourceLimitExceeded", { ...limits, maxNodes: maxNodes - 1 });
    failure(input, "ResourceLimitExceeded", { ...limits, maxDepth: maxDepth - 1 });
  });

  test("rejects normalized output that would exceed the same byte limits on reopen", () => {
    const input = body("");
    input.set("styles.css", encoder.encode(""));
    json(input, "manifest.json", manifest);
    json(input, "metadata.json", metadata);
    const maxFileBytes = Math.max(...[...input.values()].map((bytes) => bytes.byteLength));
    failure(input, "ResourceLimitExceeded", { maxFileBytes });
  });

  test.each([0, -1, 1.5, Number.POSITIVE_INFINITY])("rejects invalid limits %s", (maxNodes) => {
    failure(files(), "InvalidDocument", { maxNodes });
  });

  test("exports immutable defaults and allows a report at its exact file-count limit", () => {
    expect(Object.isFrozen(DEFAULT_DOCUMENT_READ_LIMITS)).toBe(true);
    expect(openDocument(files(), { limits: { maxFiles: 5 } }).snapshot().manifest.documentId).toBe("report-demo");
  });
});
