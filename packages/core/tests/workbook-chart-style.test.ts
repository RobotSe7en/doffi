import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, SourceStyleResolver } from "../src/index.js";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookSession } from "../src/index.js";
import { DEFAULT_DOCUMENT_READ_LIMITS } from "../src/index.js";

const owned: WorkbookSession[] = [];
afterEach(() => owned.splice(0).forEach(session => session.dispose()));
const commit = (s: WorkbookSession, operations: DocumentOperation[]) => s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations }));
const chart = (s: WorkbookSession) => s.snapshot().metadata.workbook.charts![0]!;
function fixture() {
  const initial = createDocumentSession(openDocument(createWorkbook({ documentId: "chart-style" }), { kind: "workbook" })); owned.push(initial);
  commit(initial, [{ op: "insertWorkbookChart", sheetId: "sheet-1", kind: "line", anchor: { kind: "absolute", xPt: 10, yPt: 10, widthPt: 300, heightPt: 200 }, data: { series: [{ values: { kind: "literal", valueType: "number", count: 3, points: [{ index: 0, value: 2 }, { index: 2, value: -4 }] } }] } }]);
  commit(initial, [{ op: "duplicateWorkbookObjects", nodeIds: [chart(initial).nodeId], offset: { xPt: 310, yPt: 0 } }]);
  const files = new Map(initial.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  for (const c of metadata.workbook.charts!) c.seriesStyles = [{ seriesId: c.series[0]!.seriesId, cssClass: "shared", marker: "square", markerSize: 5 }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const path = initial.snapshot().manifest.styles[0]!;
  files.set(path, new TextEncoder().encode(new TextDecoder().decode(files.get(path)!) + 'span.shared{--office-line-color:blue!important;--office-line-width:2pt}.shared.shared{--office-line-dash:dash;--office-fill-kind:solid;--office-fill-color:var(--accent)}:root{--accent:#123456}'));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); owned.push(s); return s;
}
function own(s: WorkbookSession, name: string) { return new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS).specified({ tag: "span", attrs: { class: name }, parent: undefined }); }

test("chart styles isolate shared contextual source, preserve markers and sparse literal vectors, and undo/reopen atomically", () => {
  const s = fixture(), before = s.serialize(), c = chart(s), other = s.snapshot().metadata.workbook.charts![1]!, seriesId = c.series[0]!.seriesId;
  commit(s, [{ op: "setWorkbookChartStyle", nodeId: c.nodeId, target: "series", seriesId, properties: { "--office-line-color": "#c02030" }, marker: "circle", markerSize: 12 },
    { op: "setWorkbookChartStyle", nodeId: c.nodeId, target: "point", seriesId, index: 2, properties: { "--office-marker-fill-color": "#20a040", "--office-marker-fill-kind": "solid" }, marker: "diamond" },
    { op: "setWorkbookChartStyle", nodeId: c.nodeId, target: "pointLabel", seriesId, index: 2, properties: { color: "#703090", "font-size": "16pt" } }]);
  const current = chart(s), style = current.seriesStyles![0]!, point = current.pointStyles![0]!;
  expect(s.history().undoSteps).toBe(1); expect(style.cssClass).not.toBe("shared"); expect(style).toMatchObject({ marker: "circle", markerSize: 12 });
  expect(own(s, style.cssClass)).toMatchObject({ "--office-line-color": "#c02030", "--office-line-source": "direct", "--office-line-width": "2pt", "--office-line-dash": "dash", "--office-fill-color": "var(--accent)" });
  expect(own(s, "shared")["--office-line-color"]).toBe("blue"); expect(s.snapshot().metadata.workbook.charts![1]).toEqual(other); expect(current.series).toEqual(c.series);
  expect(point).toMatchObject({ index: 2, marker: "diamond" }); expect(own(s, point.labelCssClass!)).toMatchObject({ color: "#703090", "font-size": "16pt", "--office-text-source": "direct" });
  expect(s.snapshot().manifest.chunks).toEqual([]); const saved = s.serialize(); expect(openDocument(saved, { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata); expect(s.snapshot().styles).toEqual(openDocument(before, { kind: "workbook" }).snapshot().styles); s.redo(); expect(new Map([...s.serialize()].filter(([path]) => path !== "manifest.json"))).toEqual(new Map([...saved].filter(([path]) => path !== "manifest.json")));
});

test("all chart part slots support font/paint changes, marker-only edits and precise no-op history", () => {
  const s = fixture(), c = chart(s), seriesId = c.series[0]!.seriesId;
  const operations: DocumentOperation[] = ["plot", "legend", "text", "label", "categoryAxis", "valueAxis", "grid", "categoryGrid"].map(target => ({ op: "setWorkbookChartStyle", nodeId: c.nodeId, target: target as "plot", properties: { "--office-line-color": "#112233" } }));
  operations.push({ op: "setWorkbookChartStyle", nodeId: c.nodeId, target: "seriesLabel", seriesId, properties: { color: "red", "text-decoration": "underline double" } }, { op: "setWorkbookChartStyle", nodeId: c.nodeId, target: "point", seriesId, index: 1, properties: {}, marker: "circle", markerSize: 9 });
  commit(s, operations); const before = s.serialize(), history = s.history(); expect(chart(s).pointStyles![0]).toMatchObject({ index: 1, marker: "circle", markerSize: 9 });
  expect(commit(s, operations).changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  commit(s, [{ op: "setWorkbookChartStyle", nodeId: c.nodeId, target: "point", seriesId, index: 1, properties: {}, marker: null, markerSize: null }]); expect(chart(s).pointStyles![0]).not.toHaveProperty("marker"); expect(chart(s).pointStyles![0]).not.toHaveProperty("markerSize");
  s.undo(); expect(new Map([...s.serialize()].filter(([path]) => path !== "manifest.json"))).toEqual(new Map([...before].filter(([path]) => path !== "manifest.json")));
});

test("chart style identity/marker/property errors roll back the complete batch and stale candidates fail", () => {
  const s = fixture(), c = chart(s), before = s.serialize(), seriesId = c.series[0]!.seriesId;
  const operation: Extract<DocumentOperation, { op: "setWorkbookChartStyle" }> = { op: "setWorkbookChartStyle", nodeId: c.nodeId, target: "series", seriesId, properties: { color: "red" } };
  for (const changes of [{ seriesId: "foreign" }, { target: "plot" as const }, { index: 0 }, { target: "point" as const, index: -1 }, { target: "seriesLabel" as const, marker: "circle" as const }, { markerSize: 73 }, { properties: { width: "100pt" } }, { properties: { color: "url(https://invalid.example/)" } }]) {
    expect(() => commit(s, [{ op: "setWorkbookChart", nodeId: c.nodeId, changes: { legend: "left" } }, { ...operation, ...changes }])).toThrow(); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  }
  const stale = s.prepare({ protocolVersion: 1, documentId: "chart-style", baseRevision: 2, operations: [operation] });
  commit(s, [{ op: "setWorkbookChart", nodeId: c.nodeId, changes: { legend: "left" } }]); expect(() => s.commit(stale)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  const saved = s.serialize(); expect(() => commit(s, [{ op: "setWorkbookChartData", nodeId: c.nodeId, data: { series: [{ values: { kind: "literal", valueType: "number", count: 0, points: [] } }] } }, operation])).toThrow(); expect(s.serialize()).toEqual(saved);
});

test("chart variable decoration shorthands stay intact, retain no-op history and replace later longhand overrides", () => {
  const s = fixture(), c = chart(s), operation: DocumentOperation = { op: "setWorkbookChartStyle", nodeId: c.nodeId, target: "label", properties: { "text-decoration": "var(--decor)" } };
  commit(s, [operation]); expect(s.snapshot().styles.map(style => style.css).join("\n")).toContain("text-decoration:var(--decor)!important");
  const saved = s.serialize(); expect(commit(s, [operation]).changed).toBe(false); expect(s.serialize()).toEqual(saved);
  commit(s, [{ ...operation, properties: { "text-decoration-line": "none" } }]); expect(commit(s, [operation]).changed).toBe(true); expect(commit(s, [operation]).changed).toBe(false);
});
