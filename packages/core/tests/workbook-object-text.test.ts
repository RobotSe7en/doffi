import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "../src/index.js";
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
  expect(s.getNode(s.getNode(spans[1]!)!.childNodeIds[0]!)!.attributes.style).toBe("font-size:18pt");
  const before = s.serialize(); expect(commit(s, [operation]).changed).toBe(false); expect(s.serialize()).toEqual(before);
  commit(s, [{ op: "setWorkbookObjectText", nodeId, html: '<p><span style="text-decoration-line:underline;text-decoration:none">ABC</span></p>' }]);
  const decoratedParagraph = s.getNode(nodeId)!.childNodeIds[0]!, parent = s.getNode(decoratedParagraph)!.childNodeIds[0]!;
  commit(s, [{ op: "formatWorkbookObjectText", nodeId, paragraphId: decoratedParagraph, from: 0, to: 3, expectedText: "ABC", properties: { "text-decoration-line": "underline" } }]);
  expect(s.getNode(parent)!.childNodeIds).toHaveLength(1); expect(s.getNode(s.getNode(parent)!.childNodeIds[0]!)!.attributes.style).toBe("text-decoration-line:underline");
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
