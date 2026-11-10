import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, WorkbookAxisDimensions, createWorkbookScalarEvaluator } from "../src/index.js";
import type { DeepReadonly, DocumentOperation, WorkbookChartReference, WorkbookDrawingAnchor, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookModel, WorkbookOperation, WorkbookSession } from "../src/index.js";
import { WorkbookReferences } from "../src/workbook/reference.js";

const sessions: WorkbookSession[] = [], sourceId = "sheet-1";
afterEach(() => sessions.splice(0).forEach(s => s.dispose()));
const point = (row: number, column: number) => ({ row, column });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: point(row, column), to: point(lastRow, lastColumn) });
const marker = (row: number, column: number) => ({ row, column, offsetXPt: 1.25, offsetYPt: 2.5 });
const anchors: WorkbookDrawingAnchor[] = [
  { kind: "absolute", xPt: 81.25, yPt: 32.5, widthPt: 82.5, heightPt: 47.5 },
  { kind: "one-cell", from: marker(2, 2), widthPt: 82.5, heightPt: 47.5 },
  ...[false, true].flatMap(moveWithCells => [false, true].map(sizeWithCells => ({ kind: "two-cell" as const, from: marker(2, 2), to: { ...marker(5, 4), offsetXPt: 3.75, offsetYPt: 5 }, moveWithCells, sizeWithCells }))),
];
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
/** Explicit document-point metrics test the transaction geometry kernel;
 * browser/native producers provide separate font and OOXML evidence. */
function metrics(request: WorkbookLayoutRequest): WorkbookLayoutMeasurement {
  return { sourceDigest: request.sourceDigest, fontEnvironment: "explicit-drawing-transport-metrics", sheets: request.sheetIds.map(sheetId => {
    const sheet = request.model.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!;
    return { sheetId, rows: { defaultSizePt: sheet.defaultRowHeightPt, overrides: sheet.rowDimensions.map(run => ({ from: run.from, to: run.to, sizePt: run.hidden ? 0 : run.heightPt ?? sheet.defaultRowHeightPt })) },
      columns: { defaultSizePt: sheet.name === "Other" ? 70 : 40, overrides: [] } };
  }) };
}
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations), { measureWorkbookLayout: metrics })); }
function example(anchor: WorkbookDrawingAnchor) {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "drawing-transport" }), { kind: "workbook" })); sessions.push(s);
  apply(s, [{ op: "insertWorkbookSheet", name: "Other" }, { op: "insertWorkbookObject", sheetId: sourceId, anchor, name: "Native shape", object: { kind: "geometry", preset: "rect", html: "<p>禰󠄀😀<strong>预算</strong></p>" } }]);
  const targetId = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
  apply(s, [{ op: "updateWorkbookSheet", sheetId: targetId, changes: { defaultRowHeightPt: 30 } }, { op: "setWorkbookCells", sheetId: sourceId,
    cells: [{ row: 2, column: 2, source: { kind: "value", value: { type: "number", value: 3 } } }] }]);
  s.clearHistory(); return { s, targetId };
}
function rectangle(model: WorkbookModel, index = 0) {
  const object = model.snapshot().metadata.workbook.objects![index]!, anchor = object.anchor;
  const measured = metrics({ model, sourceDigest: "assertion", sheetIds: [object.sheetId] }).sheets[0]!;
  const rows = new WorkbookAxisDimensions(1_048_576, measured.rows.defaultSizePt, measured.rows.overrides), columns = new WorkbookAxisDimensions(16_384, measured.columns.defaultSizePt, measured.columns.overrides);
  if (anchor.kind === "absolute") return { x: anchor.xPt, y: anchor.yPt, width: anchor.widthPt, height: anchor.heightPt };
  const x = columns.position(anchor.from.column) + anchor.from.offsetXPt, y = rows.position(anchor.from.row) + anchor.from.offsetYPt;
  return { x, y, width: anchor.kind === "one-cell" ? anchor.widthPt : columns.position(anchor.to.column) + anchor.to.offsetXPt - x,
    height: anchor.kind === "one-cell" ? anchor.heightPt : rows.position(anchor.to.row) + anchor.to.offsetYPt - y };
}
const move = (targetId: string, back = false): DocumentOperation => ({ op: "moveWorkbookRange", sheetId: back ? targetId : sourceId, range: back ? range(259, 7) : range(2, 2), targetSheetId: back ? sourceId : targetId, target: back ? point(2, 2) : point(259, 7) });
test.each(anchors)("cut ownership and physical geometry respect $kind/$moveWithCells/$sizeWithCells", anchor => {
  const { s, targetId } = example(anchor), before = s.snapshot(), bounds = rectangle(s), object = before.metadata.workbook.objects![0]!, html = s.getNode(object.nodeId)!.html;
  const candidate = s.prepare(transaction(s, [move(targetId)]), { measureWorkbookLayout: metrics }); expect(s.snapshot()).toEqual(before); s.commit(candidate);
  const follows = anchor.kind === "one-cell" || anchor.kind === "two-cell" && anchor.moveWithCells, current = rectangle(s);
  expect(s.snapshot().metadata.workbook.objects![0]!.sheetId).toBe(follows ? targetId : sourceId);
  expect(current.x).toBeCloseTo(bounds.x + (follows ? 410 : 0), 9); expect(current.y).toBeCloseTo(bounds.y + (follows ? 7740 : 0), 9);
  expect(current.width).toBeCloseTo(bounds.width, 9); expect(current.height).toBeCloseTo(bounds.height, 9);
  expect(s.getNode(object.nodeId)!.html).toBe(html); expect(s.snapshot().metadata.workbook.objects![0]).toMatchObject({ nodeId: object.nodeId, name: object.name, kind: object.kind });
  expect(s.getCell(targetId, 259, 7)?.value).toEqual({ type: "number", value: 3 }); expect(s.history().undoSteps).toBe(1);
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  const bytes = s.serialize(); expect(apply(s, [move(targetId), move(targetId, true)]).changed).toBe(false); expect(s.serialize()).toEqual(bytes);
});
test("sorting follows record identity across unequal row heights and preserves layer and rich-text identities", () => {
  const { s } = example(anchors[5]!); apply(s, [{ op: "updateWorkbookSheet", sheetId: sourceId, changes: { rowDimensions: [{ from: 3, to: 3, heightPt: 45 }] } },
    { op: "insertWorkbookObject", sheetId: sourceId, anchor: anchors[0]!, object: { kind: "text", html: "<p>前景</p>" } }]); s.clearHistory();
  const before = s.snapshot(), bounds = rectangle(s), operation: DocumentOperation = { op: "reorderWorkbookRows", sheetId: sourceId, range: range(2, 0, 4, 2), rows: [4, 3, 2] };
  apply(s, [operation]); const current = rectangle(s); expect(current.y).toBeCloseTo(bounds.y + 60, 9); expect(current.height).toBeCloseTo(bounds.height, 9);
  expect(s.snapshot().metadata.workbook.objects!.map(o => o.nodeId)).toEqual(before.metadata.workbook.objects!.map(o => o.nodeId)); expect(s.snapshot().metadata.workbook.objects![1]).toEqual(before.metadata.workbook.objects![1]);
  s.undo(); const bytes = s.serialize(); expect(apply(s, [operation, operation]).changed).toBe(false); expect(s.serialize()).toEqual(bytes);
});
test("moving the anchor cell alone transports its drawing; covering only the extent or overwriting below another drawing does not", () => {
  const { s, targetId } = example(anchors[5]!), before = s.snapshot().metadata.workbook.objects![0];
  apply(s, [{ op: "moveWorkbookRange", sheetId: sourceId, range: range(5, 4), targetSheetId: targetId, target: point(10, 5) }]); expect(s.snapshot().metadata.workbook.objects![0]).toEqual(before);
  expect(() => s.prepare(transaction(s, [move(targetId)]))).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  apply(s, [move(targetId)]); const owner = s.snapshot().metadata.workbook.objects![0]!;
  apply(s, [{ op: "moveWorkbookRange", sheetId: sourceId, range: range(20, 1), targetSheetId: targetId, target: point(259, 7), overwrittenContent: "clear" }]); expect(s.snapshot().metadata.workbook.objects![0]).toEqual(owner);
});
test("native-limit transport retains physical extents and restores original markers in a composed inverse", () => {
  const { s, targetId } = example(anchors[5]!), before = s.serialize(), bounds = rectangle(s), target = point(1_048_575, 16_383);
  const forward: DocumentOperation = { op: "moveWorkbookRange", sheetId: sourceId, range: range(2, 2), targetSheetId: targetId, target },
    back: DocumentOperation = { op: "moveWorkbookRange", sheetId: targetId, range: { from: target, to: target }, targetSheetId: sourceId, target: point(2, 2) };
  apply(s, [forward]); const current = rectangle(s), anchor = s.snapshot().metadata.workbook.objects![0]!.anchor;
  expect(current.width).toBeCloseTo(bounds.width, 7); expect(current.height).toBeCloseTo(bounds.height, 7);
  expect(anchor).toMatchObject({ from: target, to: target }); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); const restored = s.serialize(); expect(apply(s, [forward, back]).changed).toBe(false); expect(s.serialize()).toEqual(restored);
  expect(s.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
});
function chartExample() {
  const { s: original, targetId } = example(anchors[5]!);
  apply(original, [{ op: "insertWorkbookChart", sheetId: sourceId, anchor: anchors[1]!, kind: "column", data: {
    categories: { expression: "=A2:A4", valueType: "text" }, series: [
      { name: { kind: "reference", reference: { expression: "=B1", valueType: "text" } }, values: { expression: "=B2:B4", valueType: "number" } },
      { values: { kind: "literal", valueType: "number", count: 3, points: [{ index: 0, value: -3 }, { index: 2, value: 9 }] } },
    ] } }, { op: "setWorkbookCells", sheetId: sourceId, cells: [1, 2, 3].flatMap(row => [
      { row, column: 0, source: { kind: "value" as const, value: { type: "text" as const, value: ["甲", "乙", "丙"][row - 1]! } } },
      { row, column: 1, source: { kind: "value" as const, value: { type: "number" as const, value: row + 1 } } },
    ]) }]);
  const files = new Map(original.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!));
  metadata.workbook.charts[0].series[0].values.cache = { state: "imported", count: 3, points: [{ index: 0, value: 999 }], numberFormat: "0.00" };
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); sessions.push(s); return { s, targetId };
}
test("cross-sheet chart transport retains its input bindings, literals, identities and historical cache fields", () => {
  const { s, targetId } = chartExample(), before = s.snapshot().metadata.workbook, chart = before.charts![0]!, literal = chart.series[1]!.values;
  apply(s, [move(targetId)]); const current = s.snapshot().metadata.workbook.charts![0]!;
  expect(s.snapshot().metadata.workbook.objects!.find(o => o.nodeId === chart.nodeId)!.sheetId).toBe(targetId);
  expect(current.categories).toMatchObject({ expression: "=A2:A4", originSheetId: sourceId }); expect(current.series[0]!.values).toMatchObject({ expression: "=B2:B4", originSheetId: sourceId, cache: { state: "stale", count: 3, points: [{ index: 0, value: 999 }], numberFormat: "0.00" } });
  expect(current.series[0]!.name).toMatchObject({ reference: { expression: "=B1", originSheetId: sourceId } }); expect(current.series[1]!.values).toEqual(literal);
  expect(current.series.map(series => series.seriesId)).toEqual(chart.series.map(series => series.seriesId));
  const evaluate = createWorkbookScalarEvaluator(s); expect(evaluate.evaluate("=SUM(" + (current.series[0]!.values as DeepReadonly<WorkbookChartReference>).expression.slice(1) + ")", { sheetId: sourceId, ...point(0, 0) })).toMatchObject({ value: { type: "number", value: 9 } }); evaluate.dispose();
  s.undo(); const bytes = s.serialize(); expect(apply(s, [move(targetId), move(targetId, true)]).changed).toBe(false); expect(s.serialize()).toEqual(bytes);
});
test("cancelling after the final drawing projection leaves source, rich text and history untouched", async () => {
  const { s, targetId } = example(anchors[5]!), before = s.serialize(), signal = { aborted: false }; let calls = 0;
  await expect(s.prepareAsync(transaction(s, [move(targetId)]), { measureWorkbookLayout: request => { const result = metrics(request); if (++calls === 2) signal.aborted = true; return result; } }, { signal, yield: async () => {} })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(calls).toBe(2); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
test.each([1, 2, 3])("partial cut of chart vector element %s preserves source order and reads current moved data", row => {
  const { s, targetId } = chartExample(), before = s.snapshot().metadata.workbook.charts![0]!;
  apply(s, [{ op: "moveWorkbookRange", sheetId: sourceId, range: range(row, 1), targetSheetId: targetId, target: point(19, 7) }]);
  const chart = s.snapshot().metadata.workbook.charts![0]!, source = chart.series[0]!.values as DeepReadonly<WorkbookChartReference>;
  const resolved = new WorkbookReferences(s.snapshot().metadata.workbook).resolve(source.expression, { sheetId: sourceId, ...point(0, 0) });
  if (resolved.kind !== "ranges") throw Error("Expected current chart cells");
  const values = resolved.areas.flatMap(area => Array.from({ length: area.to.row - area.from.row + 1 }, (_, i) => s.getCell(area.sheetId, area.from.row + i, area.from.column)?.value));
  expect(values).toEqual([2, 3, 4].map(value => ({ type: "number", value }))); expect(chart.categories).toEqual(before.categories);
  expect(source.cache).toMatchObject({ state: "stale", points: [{ index: 0, value: 999 }] });
  apply(s, [{ op: "setWorkbookCells", sheetId: targetId, cells: [{ row: 19, column: 7, source: { kind: "value", value: { type: "number", value: 42 } } }] }]);
  const evaluator = createWorkbookScalarEvaluator(s); expect(evaluator.evaluate("=SUM(" + source.expression.slice(1) + ")", { sheetId: sourceId, ...point(0, 0) })).toMatchObject({ value: { type: "number", value: 9 - (row + 1) + 42 } }); evaluator.dispose();
});
test("chart transport preserves local names and distinguishes global names from a destination local shadow", () => {
  const { s, targetId } = chartExample(), chart = s.snapshot().metadata.workbook.charts![0]!;
  apply(s, [{ op: "setWorkbookDefinedName", name: { nameId: "local-source", name: "Amounts", sheetId: sourceId, origin: point(0, 0), expression: "=$B$2:$B$4" } },
    { op: "setWorkbookDefinedName", name: { nameId: "local-target", name: "Amounts", sheetId: targetId, origin: point(0, 0), expression: "=$K$2:$K$4" } },
    { op: "setWorkbookDefinedName", name: { nameId: "global", name: "GlobalAmounts", origin: point(0, 0), expression: "=Sheet1!$B$2:$B$4" } },
    { op: "setWorkbookDefinedName", name: { nameId: "global-shadow", name: "GlobalAmounts", sheetId: targetId, origin: point(0, 0), expression: "=$L$2:$L$4" } },
    { op: "setWorkbookChartData", nodeId: chart.nodeId, data: { series: [{ seriesId: chart.series[0]!.seriesId, values: { expression: "=Amounts", valueType: "number" } },
      { seriesId: chart.series[1]!.seriesId, values: { expression: "=GlobalAmounts", valueType: "number" } }] } }]); s.clearHistory();
  apply(s, [move(targetId)]); const current = s.snapshot().metadata.workbook.charts![0]!;
  expect(current.series[0]!.values).toMatchObject({ expression: "=Amounts", originSheetId: sourceId }); expect(current.series[1]!.values).toMatchObject({ expression: "=GlobalAmounts", originSheetId: sourceId });
  const references = new WorkbookReferences(s.snapshot().metadata.workbook);
  for (const series of current.series) expect(references.resolve((series.values as DeepReadonly<WorkbookChartReference>).expression, { sheetId: sourceId, ...point(0, 0) })).toMatchObject({ kind: "ranges", areas: [{ sheetId: sourceId, from: point(1, 1), to: point(3, 1) }] });
});
test("a scoped cell cut cannot indirectly move a drawing without its frozen object permission", () => {
  const { s, targetId } = example(anchors[5]!), m = s.snapshot().manifest;
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "request", documentId: m.documentId, baseRevision: m.revision, status: "open", instruction: "移动记录与绘图", anchor: { kind: "cells", sheetId: sourceId, ranges: [range(2, 2)] } } }]); s.clearHistory();
  const ranges = [{ sheetId: sourceId, range: range(2, 2) }, { sheetId: targetId, range: range(259, 7) }], before = s.serialize();
  const request = (objectIds: string[]) => s.buildInteractionRequest({ kind: "workbook", annotationIds: ["request"], scope: { kind: "ranges", ranges, objectIds } });
  const proposal = (input: ReturnType<typeof request>) => ({ protocolVersion: 1 as const, documentId: input.documentId, baseRevision: input.baseRevision, requestId: input.requestId, proposalId: "move-drawing", annotationIds: ["request"], summary: "移动记录与绘图", operations: [move(targetId) as WorkbookOperation] });
  const denied = request([]); expect(() => s.prepareProposal(denied, proposal(denied), { measureWorkbookLayout: metrics })).toThrowError(expect.objectContaining({ code: "InvalidProposal" })); expect(s.serialize()).toEqual(before);
  const allowed = request([s.snapshot().metadata.workbook.objects![0]!.nodeId]), candidate = s.prepareProposal(allowed, proposal(allowed), { measureWorkbookLayout: metrics });
  expect(s.serialize()).toEqual(before); expect(candidate.preview.metadata.workbook.objects![0]!.sheetId).toBe(targetId); s.commit(candidate); expect(s.history().undoSteps).toBe(1);
  s.undo(); expect(s.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
});
