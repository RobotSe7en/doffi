import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument, WorkbookChartDataInputSchema } from "@doffi/core";
import type { DocumentOperation, WorkbookLayoutMeasurer, WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { createWorkbookPreviewSource, DEFAULT_WORKBOOK_CHART_READ_LIMITS } from "@doffi/render";
// Explicit point metrics verify source/history here. Browser acceptance owns
// actual font-driven geometry; imported fixed-response objects still require
// an authorized measurer when their cell inputs change.
const metrics: WorkbookLayoutMeasurer = request => ({ sourceDigest: request.sourceDigest, fontEnvironment: 'native-source-test-points',
  sheets: request.sheetIds.map(sheetId => ({ sheetId, rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
const commit = (s: WorkbookSession, operations: DocumentOperation[]) => s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations }, { measureWorkbookLayout: metrics }));

test.each(["examples/xlsx/07-three-charts.xlsx", "examples/xlsx/producer/07-three-charts.xlsx"])("native %s keeps source caches and unrelated charts while current data, titles, axes, copied identities and history change", async path => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync(path), { documentId: "drawing-native" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), chart = before.metadata.workbook.charts![0]!, object = before.metadata.workbook.objects!.find(object => object.nodeId === chart.nodeId)!, owner = before.metadata.workbook.sheets.find(sheet => sheet.sheetId === object.sheetId)!;
    expect(commit(s, [{ op: "setWorkbookChartData", nodeId: chart.nodeId, data: WorkbookChartDataInputSchema.parse({ categories: chart.categories, series: chart.series.map(({ nativeIndex: _nativeIndex, ...item }) => item) }) }]).changed).toBe(false);
    const others = before.metadata.workbook.objects!.filter(object => object.nodeId !== chart.nodeId).map(object => [object.nodeId, s.getNode(object.nodeId)!.html] as const), series = chart.series[0]!.seriesId;
    const reference = "='" + owner.name.replace(/'/g, "''") + "'!$B$2:$B$3";
    commit(s, [{ op: "setWorkbookCells", sheetId: object.sheetId, cells: [{ row: 1, column: 1, source: { kind: "formula", expression: "=40+2" } }] },
      { op: "setWorkbookChartData", nodeId: chart.nodeId, data: WorkbookChartDataInputSchema.parse({ ...(chart.categories ? { categories: chart.categories } : {}), series: [{ seriesId: series, values: { expression: reference, valueType: "number", cache: { state: "valid", count: 2, points: [{ index: 0, value: -999 }] } } }] }) },
      { op: "setWorkbookChartTitle", nodeId: chart.nodeId, target: "title", source: { html: "<p>新标题<strong>中文禰󠄀😀</strong></p>" } },
      { op: "setWorkbookChart", nodeId: chart.nodeId, changes: { minimum: 0, maximum: 100, majorUnit: 20, legend: "right", showLabels: true } }]);
    const calculation = s.calculateSync(), source = createWorkbookPreviewSource(s, calculation);
    const response = await source.charts!.read({ ...DEFAULT_WORKBOOK_CHART_READ_LIMITS, documentId: "drawing-native", revision: s.snapshot().manifest.revision, chartIds: [chart.chartId] }, new AbortController().signal);
    const vector = response.charts[0]!.references.find(reference => reference.role === "series-values" && reference.seriesId === series)!.data;
    expect(vector.kind).toBe("cells"); if (vector.kind !== "cells") throw Error("Expected source cells"); expect(vector.areas[0]!.cells[0]!.value).toEqual({ type: "number", value: 42 });
    const current = s.snapshot().metadata.workbook.charts![0]!; expect(current.series[0]!.values).toEqual({ expression: reference, valueType: "number" });
    expect(current.categories).toEqual(chart.categories); expect(current.series[0]!.nativeIndex).toBe(chart.series[0]!.nativeIndex);
    expect(s.getNode(current.titleNodeId!)!.text).toBe("新标题中文禰󠄀😀"); for (const [id, html] of others) expect(s.getNode(id)!.html).toBe(html);
    commit(s, [{ op: "duplicateWorkbookObjects", nodeIds: [chart.nodeId], offset: { xPt: 8, yPt: 10 } }]);
    const clone = s.snapshot().metadata.workbook.charts!.at(-1)!; expect(clone.chartId).not.toBe(chart.chartId); expect(clone.series[0]!.seriesId).not.toBe(series);
    expect(s.snapshot().metadata.workbook.objects!.find(object => object.nodeId === clone.nodeId)).not.toHaveProperty("nativeId");
    expect(openDocument(await readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().content).toEqual(before.content);
  } finally { s.dispose(); }
});

test.each(["examples/xlsx/08-images-shapes.xlsx", "examples/xlsx/producer/08-images-shapes.xlsx"])("native %s retains image resources, original crop/rotation, shape runs and untouched source through independent object edits", async path => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync(path), { documentId: "drawing-native" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), picture = before.metadata.workbook.objects!.find(object => object.kind === "image")!, shape = before.metadata.workbook.objects!.find(object => object.kind === "geometry")!, assetBytes = new Map(before.manifest.assets.map(asset => [asset.path, s.serialize().get(asset.path)!]));
    const original = new Map(before.metadata.workbook.objects!.map(object => [object.nodeId, s.getNode(object.nodeId)!.html]));
    commit(s, [{ op: "setWorkbookImageCrop", nodeId: picture.nodeId, crop: { left: .1, top: .2, right: .1, bottom: .2 } },
      { op: "updateWorkbookObject", nodeId: shape.nodeId, changes: { anchor: { kind: "absolute", xPt: 70, yPt: 80, widthPt: 160, heightPt: 100 }, rotation: 30, flipV: true } },
      { op: "setWorkbookObjectText", nodeId: shape.nodeId, html: '<p><strong>新形状文字禰󠄀😀</strong></p>' },
      { op: "setWorkbookObjectStyle", nodeId: shape.nodeId, properties: { "--office-fill-color": "#234567", "--office-line-width": "2pt" } }]);
    expect(s.getNode(picture.nodeId)!.attributes).toMatchObject({ "data-office-crop-left": "0.1", "data-office-crop-bottom": "0.2" });
    expect(s.getNode(shape.nodeId)!.text).toBe("新形状文字禰󠄀😀"); for (const [id, html] of original) if (id !== picture.nodeId && id !== shape.nodeId) expect(s.getNode(id)!.html).toBe(html);
    for (const [path, bytes] of assetBytes) expect(s.serialize().get(path)).toEqual(bytes); expect(s.snapshot().manifest.assets).toEqual(before.manifest.assets); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    expect(openDocument(await readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    commit(s, [{ op: "removeWorkbookObject", nodeId: picture.nodeId }]); expect(s.getNode(picture.nodeId)).toBeUndefined(); s.undo(); s.undo();
    expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); for (const [id, html] of original) expect(s.getNode(id)!.html).toBe(html);
  } finally { s.dispose(); }
});
