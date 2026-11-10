import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, SourceStyleResolver, DEFAULT_DOCUMENT_READ_LIMITS } from "../src/index.js";
import type { SourceStyleNode } from "../src/index.js";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookSession } from "../src/index.js";

const owned: WorkbookSession[] = [];
afterEach(() => owned.splice(0).forEach(session => session.dispose()));
const commit = (session: WorkbookSession, operations: DocumentOperation[]) => session.commit(session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: session.snapshot().manifest.revision, operations }));
function fixture() {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: "drawing-text" }), { kind: "workbook" })); owned.push(session);
  commit(session, ["text", "geometry"].map(kind => ({ op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "absolute", xPt: 20, yPt: 30, widthPt: 180, heightPt: 90 },
    object: kind === "text" ? { kind: "text", html: '<p style="text-align:right">A<strong>禰󠄀😀</strong>&#13;<br>Z</p><p>second<em>段</em></p>', text: { wrap: true } } : { kind: "geometry", preset: "ellipse", html: "<p>另一对象</p>" } })));
  const object = session.snapshot().metadata.workbook.objects![0]!, other = session.snapshot().metadata.workbook.objects![1]!;
  return { session, nodeId: object.nodeId, paragraphId: session.getNode(object.nodeId)!.childNodeIds[0]!, second: session.getNode(object.nodeId)!.childNodeIds[1]!, other };
}

test("partial decoration clearing preserves text, native font roles, identities and exact surrounding paint through history and reopen", () => {
  const { session: s, nodeId } = fixture();
  commit(s, [{ op: "setWorkbookObjectText", nodeId, html: '<p style="text-decoration:underline line-through double red;--office-font-ascii:&quot;Liberation Sans&quot;;--office-font-east-asia:&quot;Noto Sans CJK SC&quot;;font-size:18pt"><strong>A禰󠄀😀Z</strong></p>' }]); s.clearHistory();
  const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, strong = s.getNode(paragraphId)!.childNodeIds[0]!, before = s.snapshot();
  function appearance(): ReturnType<SourceStyleResolver["get"]>[] {
    const resolver = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS);
    const read = (id: string, parent?: SourceStyleNode): ReturnType<SourceStyleResolver["get"]>[] => {
      const item = s.getNode(id)!, node = { tag: item.tagName, attrs: item.attributes, parent };
      return [resolver.get(node), ...item.childNodeIds.flatMap(id => read(id, node))];
    };
    return read(nodeId).slice(-2);
  }
  const operation: DocumentOperation = { op: "formatWorkbookObjectText", nodeId, paragraphId, from: 1, to: 6, expectedText: "禰󠄀😀", properties: { "text-decoration-line": "none", "--office-font-decoration-reset": "underline" } };
  commit(s, [operation]); expect(s.getNode(strong)!.text).toBe("A禰󠄀😀Z"); expect(s.history().undoSteps).toBe(1);
  expect(appearance().map(style => style["text-decoration-line"])).toEqual(["underline line-through", "line-through"]);
  // The unselected text remains in the original strong node; only selected
  // leaves receive the source override. Its native role declarations inherit.
  const wrapper = s.getNode(s.getNode(strong)!.childNodeIds[0]!)!;
  expect(wrapper.text).toBe("禰󠄀😀"); expect(wrapper.attributes.style).toContain("--office-font-decoration-reset:underline");
  expect(appearance()[1]).toMatchObject({ "--office-font-ascii": '"Liberation Sans"', "--office-font-east-asia": '"Noto Sans CJK SC"', "font-size": "18pt", "font-weight": "bold" });
  expect(commit(s, [operation]).changed).toBe(false);
  const reopened = openDocument(s.serialize(), { kind: "workbook" }); expect(reopened.snapshot().content!.html).toBe(s.snapshot().content!.html);
  s.undo(); expect(s.snapshot().content!.html).toBe(before.content!.html); s.redo(); expect(s.getNode(strong)!.text).toBe("A禰󠄀😀Z");
  const current = s.serialize(); expect(() => commit(s, [{ ...operation, properties: { "--office-font-decoration-reset": "all underline" } }])).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" })); expect(s.serialize()).toEqual(current);
});

test("explicit important none records a full clear and paragraph resets are persisted through the same source boundary", () => {
  const { session: s, nodeId } = fixture(); commit(s, [{ op: "setWorkbookObjectText", nodeId, html: '<p style="text-decoration:underline line-through"><strong>AB</strong></p>' }]); s.clearHistory();
  const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, strong = s.getNode(paragraphId)!.childNodeIds[0]!;
  commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "A", properties: { "text-decoration-line": "none!important" } }]);
  const wrapper = s.getNode(s.getNode(strong)!.childNodeIds[0]!)!;
  expect(wrapper.attributes.style).toContain("--office-font-decoration-reset:all");
  const parent: SourceStyleNode = { tag: "p", attrs: s.getNode(paragraphId)!.attributes, parent: undefined }, child: SourceStyleNode = { tag: "strong", attrs: s.getNode(strong)!.attributes, parent };
  const resolver = new SourceStyleResolver([], DEFAULT_DOCUMENT_READ_LIMITS);
  expect(resolver.get(child)["text-decoration-line"]).toBe("underline line-through"); expect(resolver.get({ tag: "span", attrs: wrapper.attributes, parent: child })["text-decoration-line"]).toBe("none");
  s.undo(); commit(s, [{ op: "setWorkbookObjectParagraphStyle", nodeId, paragraphId, properties: { "text-decoration": "none" } }]); expect(s.getNode(paragraphId)!.attributes.style).toContain("--office-font-decoration-reset:all");
});

test("partial font formatting preserves unrequested fonts when new spans match existing source CSS", () => {
  const { session: initial, nodeId } = fixture();
  commit(initial, [{ op: "setWorkbookObjectText", nodeId, html: '<p style="font-family:&quot;Liberation Sans&quot;;font-size:18pt;--office-font-ascii:&quot;Liberation Sans&quot;;--office-font-east-asia:&quot;Noto Sans CJK SC&quot;"><strong>Glyph禰󠄀😀</strong></p>' }]);
  const files = new Map(initial.serialize()); files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + 'span{font-family:"Liberation Serif"!important;font-size:40pt!important;font-weight:normal!important;font-style:italic!important;--office-font-ascii:"Liberation Serif"!important;--office-font-east-asia:"Liberation Serif"!important}'));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); owned.push(s);
  const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, strong = s.getNode(paragraphId)!.childNodeIds[0]!;
  const project = (id: string): SourceStyleNode => { const node = s.getNode(id)!; return { tag: node.tagName, attrs: node.attributes, parent: node.parentNodeId ? project(node.parentNodeId) : undefined }; };
  const before = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS).get(project(strong));
  commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId, from: 1, to: 4, expectedText: "lyp", properties: { color: "#a02030" } }]);
  const wrapper = s.getNode(strong)!.childNodeIds[0]!, after = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS).get(project(wrapper));
  expect(after.color).toBe("#a02030");
  for (const key of ["font-family", "font-size", "font-weight", "font-style", "--office-font-ascii", "--office-font-east-asia"]) expect(after[key], key).toBe(before[key]);
  expect(s.getNode(wrapper)!.text).toBe("lyp"); expect(s.getNode(strong)!.text).toBe("Glyph禰󠄀😀"); expect(s.history().undoSteps).toBe(1);
});

test("font no-op checks use winning declarations rather than an overridden inline value", () => {
  const { session: initial, nodeId } = fixture(); commit(initial, [{ op: "setWorkbookObjectText", nodeId, html: '<p><span style="font-size:18pt">AB</span></p>' }]);
  const files = new Map(initial.serialize()); files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + 'span{font-size:40pt!important}'));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); owned.push(s);
  const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, original = s.getNode(paragraphId)!.childNodeIds[0]!, untouched = s.getNode(original)!.attributes;
  const operation: DocumentOperation = { op: "formatWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "A", properties: { "font-size": "18pt" } };
  commit(s, [operation]); const wrapper = s.getNode(s.getNode(original)!.childNodeIds[0]!)!;
  const parent: SourceStyleNode = { tag: "span", attrs: untouched, parent: undefined }, resolver = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS);
  expect(resolver.get(parent)["font-size"]).toBe("40pt"); expect(resolver.get({ tag: "span", attrs: wrapper.attributes, parent })["font-size"]).toBe("18pt");
  expect(s.getNode(original)!.attributes).toEqual(untouched); expect(commit(s, [operation]).changed).toBe(false); expect(s.history().undoSteps).toBe(1);
});

test("requested font variables retain the original source scope rather than a new span's defaults", () => {
  const { session: initial, nodeId } = fixture(); commit(initial, [{ op: "setWorkbookObjectText", nodeId, html: '<p style="--paint:#123456"><strong>AB</strong></p>' }]);
  const files = new Map(initial.serialize()); files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + 'span{--paint:#fedcba!important}'));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); owned.push(s); const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, strong = s.getNode(paragraphId)!.childNodeIds[0]!;
  commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "A", properties: { color: "var(--paint)" } }]);
  const wrapper = s.getNode(s.getNode(strong)!.childNodeIds[0]!)!, paragraph: SourceStyleNode = { tag: "p", attrs: s.getNode(paragraphId)!.attributes, parent: undefined }, owner: SourceStyleNode = { tag: "strong", attrs: s.getNode(strong)!.attributes, parent: paragraph };
  const css = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS).get({ tag: "span", attrs: wrapper.attributes, parent: owner });
  expect(css.color).toBe("#123456"); expect(wrapper.attributes.style).toContain("--paint:inherit!important");
});

test("one hundred edits of the same font range keep bounded source structure and one hundred undo steps", () => {
  const { session: s, nodeId } = fixture(); commit(s, [{ op: "setWorkbookObjectText", nodeId, html: '<p><strong>ABCDEF</strong></p>' }]); s.clearHistory();
  const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, strong = s.getNode(paragraphId)!.childNodeIds[0]!, before = s.snapshot().content!.html; let wrapper: string | undefined;
  for (let index = 0; index < 100; index++) {
    commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId, from: 1, to: 5, expectedText: "BCDE", properties: { color: index % 2 ? "#102030" : "#203040" } }]);
    wrapper ??= s.getNode(strong)!.childNodeIds[0]!; expect(s.getNode(strong)!.childNodeIds).toEqual([wrapper]); expect(s.getNode(wrapper)!.childNodeIds).toEqual([]);
  }
  expect(s.history().undoSteps).toBe(100); expect(s.getNode(strong)!.text).toBe("ABCDEF"); for (let index = 0; index < 100; index++) s.undo(); expect(s.snapshot().content!.html).toBe(before);
}, 20_000);

test("updating a partial range inside a protected font span retains the surrounding source formatting", () => {
  const { session: s, nodeId } = fixture(); commit(s, [{ op: "setWorkbookObjectText", nodeId, html: '<p><strong>ABCDEF</strong></p>' }]); s.clearHistory();
  const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, strong = s.getNode(paragraphId)!.childNodeIds[0]!;
  commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId, from: 1, to: 5, expectedText: "BCDE", properties: { color: "#102030" } }]);
  const outer = s.getNode(strong)!.childNodeIds[0]!, original = s.getNode(outer)!.attributes;
  commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId, from: 2, to: 3, expectedText: "C", properties: { color: "#506070" } }]);
  expect(s.getNode(outer)!.attributes).toEqual(original); expect(s.getNode(outer)!.text).toBe("BCDE"); expect(s.getNode(s.getNode(outer)!.childNodeIds[0]!)!.text).toBe("C"); expect(s.history().undoSteps).toBe(2);
  s.undo(); expect(s.getNode(outer)!.childNodeIds).toEqual([]); expect(s.getNode(outer)!.attributes).toEqual(original);
});

test.each([["font-family", "--office-font-ascii", '"Liberation Sans"'], ["color", "--office-text-color", "#123456"]] as const)("native %s variables do not become self-referential during partial formatting", (property, variable, expected) => {
  const { session: s, nodeId } = fixture(); commit(s, [{ op: "setWorkbookObjectText", nodeId, html: '<p style="font-family:&quot;Noto Sans CJK SC&quot;;--office-font-ascii:&quot;Liberation Sans&quot;;--office-text-color:#123456"><strong>ABC</strong></p>' }]); s.clearHistory();
  const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, strong = s.getNode(paragraphId)!.childNodeIds[0]!;
  commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId, from: 1, to: 2, expectedText: "B", properties: { [property]: "var(" + variable + ")" } }]);
  const wrapper = s.getNode(s.getNode(strong)!.childNodeIds[0]!)!, paragraph: SourceStyleNode = { tag: "p", attrs: s.getNode(paragraphId)!.attributes, parent: undefined }, owner: SourceStyleNode = { tag: "strong", attrs: s.getNode(strong)!.attributes, parent: paragraph };
  const css = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS).get({ tag: "span", attrs: wrapper.attributes, parent: owner });
  expect(css[property]).toBe(expected); expect(css[variable]).toBe(expected);
});

test("a missing native font variable uses its requested fallback ahead of the old theme family", () => {
  const { session: s, nodeId } = fixture(); commit(s, [{ op: "setWorkbookObjectText", nodeId, html: '<p style="font-family:&quot;Noto Sans CJK SC&quot;;--office-minor-ascii:&quot;Liberation Serif&quot;"><strong style="--office-font-ascii:&quot;Liberation Sans&quot;">AB</strong><em>CD</em></p>' }]);
  const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, original = s.getNode(paragraphId)!.childNodeIds;
  commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId, from: 0, to: 4, expectedText: "ABCD", properties: { "font-family": 'var(--office-font-ascii,"Liberation Sans")' } }]);
  for (const [index, id] of original.entries()) {
    const item = s.getNode(id)!, wrapper = s.getNode(item.childNodeIds[0]!)!, paragraph: SourceStyleNode = { tag: "p", attrs: s.getNode(paragraphId)!.attributes, parent: undefined }, owner: SourceStyleNode = { tag: item.tagName, attrs: item.attributes, parent: paragraph };
    const css = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS).get({ tag: "span", attrs: wrapper.attributes, parent: owner });
    expect(css["font-family"]).toBe('"Liberation Sans"'); expect(css["--office-font-ascii"]).toBe(index === 0 ? '"Liberation Sans"' : undefined); expect(css["--office-minor-ascii"]).toBeUndefined();
  }
});

test("unrepresentable native self-reference expressions reject earlier text changes atomically", () => {
  const { session: s, nodeId, paragraphId } = fixture(), original = s.serialize(), before = s.getNode(paragraphId)!.text;
  const invalid: Record<string, string>[] = [{ "--office-text-color": "var(--office-text-color)" }, { color: "var(--missing,var(--office-text-color))" }, { "font-family": "var(--missing,var(--office-font-ascii))" }];
  for (const properties of invalid) {
    expect(() => commit(s, [{ op: "replaceWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: before.slice(0, 1), text: "X" }, { op: "formatWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "X", properties }])).toThrowError(expect.objectContaining({ code: expect.stringMatching(/InvalidOperation|UnsupportedFeature/) }));
    expect(s.serialize()).toEqual(original);
  }
});

test("paragraph text and partial fonts retain native source runs, CRLF, UTF-16 and one atomic history step", () => {
  const { session: s, nodeId, paragraphId, other } = fixture(), before = s.serialize(), paragraph = s.getNode(paragraphId)!, strong = paragraph.childNodeIds.find(id => s.getNode(id)!.tagName === "strong")!, unrelated = s.getNode(other.nodeId)!.html;
  const font = { "--office-font-ascii": '"Liberation Sans"', "--office-font-east-asia": '"Noto Sans CJK SC"', color: "#c02030", "font-size": "16pt" };
  const pending = s.prepare({ protocolVersion: 1, documentId: "drawing-text", baseRevision: 1, operations: [
    { op: "formatWorkbookObjectText", nodeId, paragraphId, from: 1, to: 6, expectedText: "禰󠄀😀", properties: font },
    { op: "replaceWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "A", text: "中文" },
  ] });
  expect(s.serialize()).toEqual(before); s.commit(pending); expect(s.history().undoSteps).toBe(2);
  expect(s.getNode(paragraphId)!.text).toBe("中文禰󠄀😀\r\nZ"); expect(s.getNode(strong)!.text).toBe("禰󠄀😀"); expect(s.getNode(strong)!.html).toContain("--office-text-source:direct");
  expect(s.getNode(paragraphId)!.html).toContain("&#13;<br"); expect(s.getNode(other.nodeId)!.html).toBe(unrelated); expect(s.snapshot().manifest.chunks).toEqual([]);
  const current = s.serialize(), history = s.history();
  expect(commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId, from: 2, to: 7, expectedText: "禰󠄀😀", properties: font }]).changed).toBe(false); expect(s.serialize()).toEqual(current); expect(s.history()).toEqual(history);
  expect(openDocument(current, { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().content).toEqual(openDocument(before, { kind: "workbook" }).snapshot().content); s.redo(); expect(s.snapshot().content).toEqual(openDocument(current, { kind: "workbook" }).snapshot().content);
});

test("paragraph insertion, alignment, body properties and deletion preserve old paragraph identities and the last typing paragraph", () => {
  const { session: s, nodeId, paragraphId, second } = fixture(), before = s.snapshot();
  commit(s, [{ op: "insertWorkbookObjectParagraph", nodeId, beforeParagraphId: second, html: '<p><em>插入禰󠄀😀</em><br>新行</p>' },
    { op: "setWorkbookObjectParagraphStyle", nodeId, paragraphId, properties: { "text-align": "center", "line-height": "1.5" } },
    { op: "updateWorkbookObject", nodeId, changes: { text: { wrap: false, autoFit: "normal", fontScale: .8 } } }]);
  const paragraphs = s.getNode(nodeId)!.childNodeIds; expect(paragraphs[0]).toBe(paragraphId); expect(paragraphs[2]).toBe(second); expect(s.getNode(paragraphs[1]!)!.text).toBe("插入禰󠄀😀\n新行");
  expect(s.getNode(paragraphId)!.attributes.style).toContain("text-align:center"); expect(s.snapshot().metadata.workbook.objects![0]).toMatchObject({ text: { wrap: false, fontScale: .8 } });
  commit(s, [{ op: "removeWorkbookObjectParagraph", nodeId, paragraphId }, { op: "removeWorkbookObjectParagraph", nodeId, paragraphId: second }]);
  const last = s.getNode(nodeId)!.childNodeIds[0]!;
  commit(s, [{ op: "removeWorkbookObjectParagraph", nodeId, paragraphId: last }]); expect(s.getNode(nodeId)!.childNodeIds).toEqual([last]); expect(s.getNode(last)!.text).toBe("");
  s.undo(); s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
});

test("font formatting retains already matching source spans while changing different parent styles independently", () => {
  const { session: s, nodeId } = fixture();
  commit(s, [{ op: "setWorkbookObjectText", nodeId, html: '<p><span style="font-size:18pt;color:red">A</span><span style="font-size:9pt;color:blue">B</span><span style="font-size:18pt;color:red">C</span></p>' }]);
  const paragraphId = s.getNode(nodeId)!.childNodeIds[0]!, spans = s.getNode(paragraphId)!.childNodeIds, a = s.getNode(spans[0]!)!.html, c = s.getNode(spans[2]!)!.html;
  const operation: DocumentOperation = { op: "formatWorkbookObjectText", nodeId, paragraphId, from: 0, to: 3, expectedText: "ABC", properties: { "font-size": "18pt" } };
  commit(s, [operation]); expect(s.getNode(spans[0]!)!.html).toBe(a); expect(s.getNode(spans[2]!)!.html).toBe(c);
  expect(s.getNode(spans[1]!)!.attributes.style).toContain("color:blue"); expect(s.getNode(spans[1]!)!.childNodeIds).toHaveLength(1);
  expect(s.getNode(s.getNode(spans[1]!)!.childNodeIds[0]!)!.attributes.style).toContain("font-size:18pt!important");
  const before = s.serialize(); expect(commit(s, [operation]).changed).toBe(false); expect(s.serialize()).toEqual(before);
  commit(s, [{ op: "setWorkbookObjectText", nodeId, html: '<p><span style="text-decoration-line:underline;text-decoration:none">ABC</span></p>' }]);
  const decoratedParagraph = s.getNode(nodeId)!.childNodeIds[0]!, parent = s.getNode(decoratedParagraph)!.childNodeIds[0]!;
  commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId: decoratedParagraph, from: 0, to: 3, expectedText: "ABC", properties: { "text-decoration-line": "underline" } }]);
  expect(s.getNode(parent)!.childNodeIds).toHaveLength(1); expect(s.getNode(s.getNode(parent)!.childNodeIds[0]!)!.attributes.style).toContain("text-decoration-line:underline!important");
});

test("font ranges and paragraph operations reject foreign identities, broken IVS, unsafe styles and earlier removals atomically", () => {
  const { session: s, nodeId, paragraphId, second, other } = fixture(), before = s.serialize(), history = s.history();
  const bad: DocumentOperation[] = [
    { op: "replaceWorkbookObjectText", nodeId, paragraphId: s.getNode(other.nodeId)!.childNodeIds[0]!, from: 0, to: 0, expectedText: "", text: "越界" },
    { op: "formatWorkbookObjectText", nodeId, paragraphId, from: 2, to: 6, expectedText: s.getNode(paragraphId)!.text.slice(2, 6), properties: { color: "red" } },
    { op: "formatWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "wrong", properties: {} },
    { op: "formatWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "A", properties: { width: "900pt" } },
    { op: "formatWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "A", properties: { color: "url(https://private.invalid/)" } },
    { op: "setWorkbookObjectParagraphStyle", nodeId, paragraphId, properties: { position: "absolute" } },
    { op: "insertWorkbookObjectParagraph", nodeId, html: '<p data-node-id="foreign">外部身份</p>' },
    { op: "insertWorkbookObjectParagraph", nodeId, html: "<p>一</p><p>二</p>" },
    { op: "insertWorkbookObjectParagraph", nodeId, html: '<p onclick="alert(1)">脚本</p>' },
    { op: "insertWorkbookObjectParagraph", nodeId, html: "<p><ins>超出工作簿正文子集</ins></p>" },
    { op: "replaceWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "A", text: "\r" },
  ];
  for (const op of bad) { expect(() => commit(s, [{ op: "setWorkbookObjectParagraphStyle", nodeId, paragraphId: second, properties: { "text-align": "left" } }, op])).toThrow(); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history); }
  expect(() => commit(s, [{ op: "removeWorkbookObjectParagraph", nodeId, paragraphId }, { op: "replaceWorkbookObjectText", nodeId, paragraphId, from: 0, to: 1, expectedText: "A", text: "x" }])).toThrow(); expect(s.serialize()).toEqual(before);
});

test("native object hyperlinks bound exact text edits and explicit paragraph/body replacements clean only their own removed links", () => {
  const initial = fixture(), files = new Map(initial.session.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  const strong = initial.session.getNode(initial.paragraphId)!.childNodeIds.find(id => initial.session.getNode(id)!.tagName === "strong")!, otherParagraph = initial.session.getNode(initial.other.nodeId)!.childNodeIds[0]!;
  metadata.workbook.objectHyperlinks = [{ nodeId: strong, url: "https://example.com/first" }, { nodeId: otherParagraph, url: "https://example.com/other" }, { nodeId: initial.nodeId, url: "https://example.com/object" }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); const s = createDocumentSession(openDocument(files, { kind: "workbook" })); owned.push(s);
  const before = s.serialize();
  expect(() => commit(s, [{ op: "replaceWorkbookObjectText", nodeId: initial.nodeId, paragraphId: initial.paragraphId, from: 0, to: 6, expectedText: "A禰󠄀😀", text: "不跨链接" }])).toThrow(); expect(s.serialize()).toEqual(before);
  commit(s, [{ op: "replaceWorkbookObjectText", nodeId: initial.nodeId, paragraphId: initial.paragraphId, from: 1, to: 6, expectedText: "禰󠄀😀", text: "链接内" }]); expect(s.getNode(strong)!.text).toBe("链接内"); expect(s.snapshot().metadata.workbook.objectHyperlinks).toEqual(metadata.workbook.objectHyperlinks);
  commit(s, [{ op: "setWorkbookObjectText", nodeId: initial.nodeId, html: "<p>显式替换全文</p>" }]); expect(s.snapshot().metadata.workbook.objectHyperlinks).toEqual(metadata.workbook.objectHyperlinks!.slice(1));
  s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
  commit(s, [{ op: "removeWorkbookObjectParagraph", nodeId: initial.nodeId, paragraphId: initial.paragraphId }]); expect(s.snapshot().metadata.workbook.objectHyperlinks).toEqual(metadata.workbook.objectHyperlinks!.slice(1));
});
