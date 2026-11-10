import { readFileSync } from "node:fs";
import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, SourceStyleResolver, DEFAULT_DOCUMENT_READ_LIMITS } from "../src/index.js";
import type { DocumentOperation, TransactionPrepareOptions, WorkbookChartDataInput, WorkbookDocumentMetadata, WorkbookDrawingAnchor, WorkbookSession } from "../src/index.js";
import { assetDigest } from "../src/document/assets.js";

const owned: WorkbookSession[] = [], sheetId = "sheet-1";
afterEach(() => owned.splice(0).forEach(session => session.dispose()));
const anchor: WorkbookDrawingAnchor = { kind: "two-cell", from: { row: 1, column: 1, offsetXPt: 2, offsetYPt: 3 }, to: { row: 8, column: 5, offsetXPt: 4, offsetYPt: 5 }, moveWithCells: true, sizeWithCells: true };
function blank() { const s = createDocumentSession(openDocument(createWorkbook({ documentId: "drawing-edit" }), { kind: "workbook" })); owned.push(s); return s; }
function prepare(s: WorkbookSession, operations: DocumentOperation[], options?: TransactionPrepareOptions) { const m = s.snapshot().manifest; return s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations }, options); }
const commit = (s: WorkbookSession, operations: DocumentOperation[], options?: TransactionPrepareOptions) => s.commit(prepare(s, operations, options));
const drawings = (s: WorkbookSession) => s.snapshot().metadata.workbook.objects!;
const chart = (s: WorkbookSession) => s.snapshot().metadata.workbook.charts![0]!;
const values: WorkbookChartDataInput = { categories: { kind: "literal", valueType: "text", count: 2, points: [{ index: 0, value: "甲" }, { index: 1, value: "乙" }] },
  series: [{ values: { expression: "=Sheet1!$A$1:$A$2", valueType: "number", cache: { state: "valid", count: 2, points: [{ index: 0, value: 999 }, { index: 1, value: 999 }] } } }] };
const shape = (): DocumentOperation => ({ op: "insertWorkbookObject", sheetId, anchor, object: { kind: "geometry", preset: "roundRect", html: '<p>中文<strong>禰󠄀😀</strong>&#13;<br>尾</p>', text: { wrap: true } }, properties: { "--office-fill-color": "#112233" }, name: "Shape" });

test("shape insertion is isolated, preserves rich UTF-16 source and allocates all object/paragraph identities", () => {
  const s = blank(), before = s.serialize(), pending = prepare(s, [shape()]); expect(s.serialize()).toEqual(before);
  expect(pending.preview.metadata.workbook.objects).toHaveLength(1); expect(pending.affectedSheetIds).toEqual([sheetId]); s.commit(pending);
  const object = drawings(s)[0]!; expect(object).toMatchObject({ kind: "geometry", preset: "roundRect", anchor, locked: true, printable: true });
  expect(s.getNode(object.nodeId)!.text).toBe("中文禰󠄀😀\r\n尾"); expect(s.getNode(object.nodeId)!.html).toContain("data-node-id");
  expect(s.snapshot().manifest.chunks).toEqual([]); expect(s.snapshot().metadata.workbook.sheets[0]!.rowDimensions).toEqual([]);
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().manifest.content).toBeUndefined(); expect(s.snapshot().metadata.workbook.objects).toBeUndefined();
  s.redo(); expect(drawings(s)[0]!.nodeId).toBe(object.nodeId);
});

test("typed anchor, shape, rotation and style changes share one history step and text equal to source preserves identities and redo", () => {
  const s = blank(); commit(s, [shape()]); const object = drawings(s)[0]!, before = s.snapshot(), html = s.getNode(object.nodeId)!.html;
  commit(s, [{ op: "updateWorkbookObject", nodeId: object.nodeId, changes: { anchor: { kind: "absolute", xPt: 30, yPt: 40, widthPt: 160, heightPt: 80 }, rotation: 45, flipH: true, preset: "ellipse", name: null, locks: { noMove: true } } },
    { op: "setWorkbookObjectStyle", nodeId: object.nodeId, properties: { "--office-fill-color": "red", "font-size": "16pt" } }]);
  expect(drawings(s)[0]).toMatchObject({ preset: "ellipse", anchor: { kind: "absolute", xPt: 30 }, locks: { noMove: true } }); expect(drawings(s)[0]).not.toHaveProperty("name");
  expect(s.getNode(object.nodeId)!.attributes).toMatchObject({ "data-office-rotation": "45", "data-office-flip-h": "true" });
  s.undo(); expect(s.getNode(object.nodeId)!.html).toBe(html); expect(s.snapshot().metadata).toEqual(before.metadata);
  const history = s.history(); expect(commit(s, [{ op: "setWorkbookObjectText", nodeId: object.nodeId, html: '<p>中文<strong>禰󠄀😀</strong>&#13;<br>尾</p>' }]).changed).toBe(false);
  expect(s.getNode(object.nodeId)!.html).toBe(html); expect(s.history()).toEqual(history);
  commit(s, [{ op: "setWorkbookObjectText", nodeId: object.nodeId, html: '<p><em>新文字禰󠄀</em></p>', text: { wrap: false, autoFit: "normal", fontScale: .8 } }]);
  expect(s.getNode(object.nodeId)!.text).toBe("新文字禰󠄀"); s.undo(); expect(s.getNode(object.nodeId)!.html).toBe(html);
});

test("authorized image bytes remain isolated, image replacement and crop preserve source identities and invalid/unused assets roll back", () => {
  const s = blank(), bytes = new Uint8Array(readFileSync(new URL("../../../examples/presentation/assets/sample.png", import.meta.url))), asset = { path: "assets/authorized.png", mimeType: "image/png", sha256: assetDigest(bytes) };
  commit(s, [{ op: "insertWorkbookObject", sheetId, anchor: { kind: "one-cell", from: anchor.from, widthPt: 120, heightPt: 80 }, object: { kind: "image", assetPath: asset.path, alt: "图片禰󠄀" } }], { assets: [{ descriptor: asset, bytes }] });
  const id = drawings(s)[0]!.nodeId, image = s.getNode(id)!.childNodeIds[0]!, before = s.serialize();
  expect(() => commit(s, [{ op: "setWorkbookImageCrop", nodeId: id, crop: { left: .7, right: .5, top: 0, bottom: 0 } }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before);
  commit(s, [{ op: "replaceWorkbookImage", nodeId: id, assetPath: asset.path, alt: "新说明" }, { op: "setWorkbookImageCrop", nodeId: id, crop: { left: .1, right: .2, top: 0, bottom: .3 } }]);
  expect(s.getNode(id)!.childNodeIds).toEqual([image]); expect(s.getNode(image)!.attributes.alt).toBe("新说明"); expect(s.getNode(id)!.attributes["data-office-crop-right"]).toBe("0.2");
  const current = s.serialize(); expect(() => commit(s, [], { assets: [{ descriptor: { ...asset, path: "assets/unused.png" }, bytes }] })).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(current);
  s.undo(); expect(s.getNode(image)!.attributes.alt).toBe("图片禰󠄀"); s.undo(); expect(s.snapshot().manifest.assets).toEqual([]);
});

test("chart vectors retain sole literal data and references discard supplied caches without creating worksheet cells", () => {
  const s = blank(); commit(s, [{ op: "insertWorkbookChart", sheetId, anchor, kind: "column", data: values, title: '<p>销售<strong>图</strong></p>' }]);
  const c = chart(s), object = drawings(s)[0]!; expect(object).toMatchObject({ kind: "chart", chartId: c.chartId });
  expect(c.series[0]!.values).toEqual({ expression: "=Sheet1!$A$1:$A$2", valueType: "number" }); expect(c.categories).toEqual(values.categories);
  expect(s.getNode(c.titleNodeId!)!.text).toBe("销售图"); expect(s.snapshot().manifest.chunks).toEqual([]);
  commit(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 7 } } }] },
    { op: "setWorkbookChartData", nodeId: c.nodeId, data: { series: [{ seriesId: c.series[0]!.seriesId, values: { kind: "literal", valueType: "number", count: 2, points: [{ index: 0, value: -3 }, { index: 1, value: 9 }] } }] } },
    { op: "setWorkbookChart", nodeId: c.nodeId, changes: { kind: "pie", showLabels: true, labelContents: ["value", "percent"] } }]);
  expect(chart(s).series[0]!.seriesId).toBe(c.series[0]!.seriesId); expect(chart(s).series[0]!.values).toMatchObject({ kind: "literal", points: [{ value: -3 }, { value: 9 }] });
  expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 7 }); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(chart(s)).toEqual(c); expect(s.getCell(sheetId, 0, 0)).toBeUndefined();
});

test("chart titles switch atomically between source text and a current reference and pie axes are rejected", () => {
  const s = blank(); commit(s, [{ op: "insertWorkbookChart", sheetId, anchor, kind: "line", data: values, title: "<p>原题</p>" }]);
  const original = chart(s), before = s.getNode(original.titleNodeId!)!.html;
  commit(s, [{ op: "setWorkbookChartTitle", nodeId: original.nodeId, target: "title", source: { reference: { expression: "=Sheet1!$B$1", valueType: "text", cache: { state: "valid", count: 1, points: [{ index: 0, value: "假缓存" }] } } } },
    { op: "setWorkbookChartTitle", nodeId: original.nodeId, target: "categoryAxis", source: { html: '<p>横轴禰󠄀</p>', text: { wrap: true } } }]);
  expect(chart(s).titleNodeId).toBeUndefined(); expect(s.getNode(original.titleNodeId!)).toBeUndefined(); expect(chart(s).titleReference).toEqual({ expression: "=Sheet1!$B$1", valueType: "text" });
  expect(s.getNode(chart(s).categoryAxisTitle!.nodeId!)!.text).toBe("横轴禰󠄀");
  const current = s.serialize(); expect(() => commit(s, [{ op: "setWorkbookChart", nodeId: original.nodeId, changes: { kind: "pie" } }])).toThrow(); expect(s.serialize()).toEqual(current);
  s.undo(); expect(chart(s)).toEqual(original); expect(s.getNode(original.titleNodeId!)!.html).toBe(before);
});

test("unchanged title references preserve only their own historical caches and no-op history across all three titles", () => {
  const source = blank(); commit(source, [{ op: "insertWorkbookChart", sheetId, anchor, kind: "column", data: values }]);
  const id = chart(source).nodeId;
  for (const target of ["title", "categoryAxis", "valueAxis"] as const) commit(source, [{ op: "setWorkbookChartTitle", nodeId: id, target, source: { reference: { expression: "=Sheet1!$B$1", valueType: "number" } } }]);
  const files = new Map(source.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata, c = metadata.workbook.charts![0]!;
  c.labelContents = [];
  for (const [index, reference] of [c.titleReference!, c.categoryAxisTitle!.reference!, c.valueAxisTitle!.reference!].entries()) reference.cache = { state: "imported", count: 1, points: [{ index: 0, value: 40 + index }], numberFormat: "0%" };
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); owned.push(s); const before = s.serialize();
  const operations: DocumentOperation[] = ["title", "categoryAxis", "valueAxis"].map(target => ({ op: "setWorkbookChartTitle", nodeId: id, target: target as "title", source: { reference: { expression: "=Sheet1!$B$1", valueType: "number", cache: { state: "valid", count: 1, points: [{ index: 0, value: -999 }] } } } }));
  expect(commit(s, operations).changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  expect(commit(s, [{ op: "setWorkbookChart", nodeId: id, changes: { showLabels: false } }]).changed).toBe(false); expect(s.serialize()).toEqual(before);
  commit(s, [{ op: "setWorkbookChartTitle", nodeId: id, target: "valueAxis", source: { reference: { expression: "=Sheet1!$C$1", valueType: "number", cache: { state: "valid", count: 1, points: [{ index: 0, value: -999 }] } } } }]);
  expect(chart(s).valueAxisTitle!.reference).toEqual({ expression: "=Sheet1!$C$1", valueType: "number" }); expect(chart(s).titleReference).toEqual(c.titleReference);
  s.undo(); expect(s.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
});

test("type-specific chart options can be explicitly cleared during one atomic kind conversion and restored by undo", () => {
  const s = blank(); commit(s, [{ op: "insertWorkbookChart", sheetId, anchor, kind: "column", data: values }]); const id = chart(s).nodeId;
  commit(s, [{ op: "setWorkbookChart", nodeId: id, changes: { gapWidth: 175, overlap: 25 } }, { op: "setWorkbookChartTitle", nodeId: id, target: "categoryAxis", source: { html: "<p>分类禰󠄀</p>" } }]);
  const before = s.snapshot();
  commit(s, [{ op: "setWorkbookChartTitle", nodeId: id, target: "categoryAxis", source: null }, { op: "setWorkbookChart", nodeId: id, changes: { kind: "pie", firstSliceAngle: 30 }, clear: ["gapWidth", "overlap"] }]);
  expect(chart(s)).toMatchObject({ kind: "pie", firstSliceAngle: 30 }); expect(chart(s)).not.toHaveProperty("gapWidth"); expect(chart(s)).not.toHaveProperty("categoryAxisTitle");
  commit(s, [{ op: "setWorkbookChart", nodeId: id, changes: { kind: "line" }, clear: ["firstSliceAngle"] }]); expect(chart(s)).not.toHaveProperty("firstSliceAngle");
  s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
});

test("duplication allocates objects, chart titles and series while removing native identities; ordering and deletion preserve unrelated source", () => {
  const s = blank(); commit(s, [shape(), { op: "insertWorkbookChart", sheetId, anchor, kind: "line", data: values, title: "<p>图</p>" }]);
  const original = chart(s), ids = drawings(s).map(object => object.nodeId), before = s.snapshot();
  commit(s, [{ op: "setWorkbookChartTitle", nodeId: original.nodeId, target: "valueAxis", source: { html: "<p>值</p>" } }]);
  commit(s, [{ op: "duplicateWorkbookObjects", nodeIds: ids, offset: { xPt: 12, yPt: 15 } }]);
  const objects = drawings(s), clone = s.snapshot().metadata.workbook.charts![1]!;
  expect(new Set(objects.map(object => object.nodeId)).size).toBe(4); expect(clone.series[0]!.seriesId).not.toBe(original.series[0]!.seriesId);
  expect(clone.titleNodeId).not.toBe(original.titleNodeId); expect(s.getNode(clone.titleNodeId!)!.text).toBe("图"); expect(s.getNode(clone.valueAxisTitle!.nodeId!)!.text).toBe("值");
  expect(objects[3]!.anchor).toMatchObject({ from: { offsetXPt: 14, offsetYPt: 18 }, to: { offsetXPt: 16, offsetYPt: 20 } });
  commit(s, [{ op: "reorderWorkbookObjects", nodeIds: [objects[3]!.nodeId], position: "back" }]); expect(drawings(s)[0]!.nodeId).toBe(objects[3]!.nodeId);
  commit(s, [{ op: "removeWorkbookObject", nodeId: clone.nodeId }]); expect(s.getNode(clone.titleNodeId!)).toBeUndefined(); expect(s.snapshot().metadata.workbook.charts).toHaveLength(1);
  s.undo(); expect(s.getNode(clone.titleNodeId!)).toBeDefined(); s.undo(); s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("unsafe or foreign drawing edits, duplicate series, and invalid geometry fail the whole batch without history", () => {
  const s = blank(); commit(s, [shape(), { op: "insertWorkbookChart", sheetId, anchor, kind: "line", data: values }]);
  const id = drawings(s)[0]!.nodeId, c = chart(s), before = s.serialize(), history = s.history();
  const bad: DocumentOperation[] = [
    { op: "setWorkbookObjectText", nodeId: id, html: '<p onclick="alert(1)">坏</p>' },
    { op: "setWorkbookObjectText", nodeId: id, html: '<p data-node-id="foreign">身份</p>' },
    { op: "setWorkbookObjectStyle", nodeId: id, properties: { width: "20pt" } },
    { op: "setWorkbookObjectStyle", nodeId: id, properties: { "--office-fill-color": "url(https://private.invalid/)" } },
    { op: "duplicateWorkbookObjects", nodeIds: [id, id], offset: { xPt: 0, yPt: 0 } },
    { op: "setWorkbookObjectText", nodeId: s.getNode(id)!.childNodeIds[0]!, html: "<p>越界</p>" },
    { op: "updateWorkbookObject", nodeId: id, changes: { anchor: { kind: "absolute", xPt: 0, yPt: 0, widthPt: -1, heightPt: 1 } } },
    { op: "setWorkbookChartData", nodeId: c.nodeId, data: { series: [{ seriesId: "foreign", values: values.series[0]!.values }] } },
    { op: "setWorkbookChartData", nodeId: c.nodeId, data: { series: [0, 1].map(() => ({ seriesId: c.series[0]!.seriesId, values: values.series[0]!.values })) } },
  ];
  for (const operation of bad) { expect(() => commit(s, [{ op: "updateWorkbookObject", nodeId: id, changes: { rotation: 13 } }, operation])).toThrow(); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history); }
});

test("duplicated drawings retain source-ID typography and child hyperlinks in their private candidate and one undo", () => {
  const original = blank(); commit(original, [shape()]); const root = drawings(original)[0]!.nodeId, p = original.getNode(root)!.childNodeIds[0]!, run = original.getNode(p)!.childNodeIds.find(id => original.getNode(id)!.tagName === "strong")!;
  const files = new Map(original.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.objectHyperlinks = [{ nodeId: run, url: "https://example.com/retained" }]; files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + '[data-node-id="' + root + '"] [data-node-id="' + run + '"]{color:#123456;font-size:20pt}'));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); owned.push(s); const before = s.snapshot(), pending = prepare(s, [{ op: "duplicateWorkbookObjects", nodeIds: [root], offset: { xPt: 0, yPt: 0 } }]);
  expect(s.snapshot()).toEqual(before); s.commit(pending); const added = drawings(s)[1]!, nextP = s.getNode(added.nodeId)!.childNodeIds[0]!, nextRun = s.getNode(nextP)!.childNodeIds.find(id => s.getNode(id)!.tagName === "strong")!;
  const resolver = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS);
  const style = (id: string): Record<string, string> => {
    const chain = []; let current: string | null = id; while (current) { const node: NonNullable<ReturnType<WorkbookSession["getNode"]>> = s.getNode(current)!; chain.push(node); current = node.parentNodeId; }
    let parent: Parameters<SourceStyleResolver["get"]>[0] | undefined;
    for (const node of chain.reverse()) parent = { tag: node.tagName, attrs: node.attributes, parent };
    return resolver.get(parent!);
  };
  expect(style(nextRun)).toEqual(style(run)); expect(style(nextRun)).toMatchObject({ color: "#123456", "font-size": "20pt" });
  expect(s.snapshot().metadata.workbook.objectHyperlinks).toEqual([{ nodeId: run, url: "https://example.com/retained" }, { nodeId: nextRun, url: "https://example.com/retained" }]);
  expect(s.getNode(root)!.html).toBe(original.getNode(root)!.html); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().styles).toEqual(before.styles); expect(s.snapshot().metadata).toEqual(before.metadata);
  const failedBefore = s.serialize(); expect(() => commit(s, [{ op: "duplicateWorkbookObjects", nodeIds: [root], offset: { xPt: 0, yPt: 0 } }, { op: "setWorkbookObjectText", nodeId: root, html: '<p onclick="bad()">bad</p>' }])).toThrow(); expect(s.serialize()).toEqual(failedBefore);
});
