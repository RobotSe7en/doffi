import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookSession } from "@doffi/core";
import { createWorkbookPreviewSource, DEFAULT_WORKBOOK_CHART_READ_LIMITS } from "@doffi/render";
import type { WorkbookLayoutSourceRequest } from "@doffi/render";
import { connectWorkbookSession, serveWorkbookSession } from "@doffi/editor";

const sessions: WorkbookSession[] = [];
afterEach(() => sessions.splice(0).forEach(session => session.dispose()));
const metrics = (request: WorkbookLayoutRequest): WorkbookLayoutMeasurement => ({ sourceDigest: request.sourceDigest, fontEnvironment: "chart-origin-explicit-point-metrics", sheets: request.sheetIds.map(sheetId => ({ sheetId,
  rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
function apply(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations }, { measureWorkbookLayout: metrics })); }
function example(expression: string) {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "chart-origin" }), { kind: "workbook" })); sessions.push(s); apply(s, [{ op: "insertWorkbookSheet", name: "Other" }]);
  const [source, other] = s.snapshot().metadata.workbook.sheets;
  apply(s, [source!, other!].map((sheet, index) => ({ op: "setWorkbookCells", sheetId: sheet.sheetId, cells: [
    { row: 0, column: 0, source: { kind: "value", value: { type: "text", value: index ? "Other category" : "来源禰󠄀😀" } } },
    { row: 0, column: 1, source: { kind: "value", value: { type: "number", value: index ? 99 : 10 } } },
  ] })));
  apply(s, [{ op: "setWorkbookDefinedName", name: { nameId: "global", name: "CallerValue", origin: { row: 0, column: 0 }, expression } },
    { op: "setWorkbookDefinedName", name: { nameId: "nested", name: "Nested", origin: { row: 0, column: 0 }, expression: "=CallerValue" } },
    { op: "insertWorkbookChart", sheetId: source!.sheetId, anchor: { kind: "one-cell", from: { row: 2, column: 0, offsetXPt: 0, offsetYPt: 0 }, widthPt: 100, heightPt: 50 }, kind: "column",
      data: { categories: { expression: "=A1", valueType: "text" }, series: [{ values: { expression: "=Nested", valueType: "number" } }] } }]); s.clearHistory();
  const move: DocumentOperation = { op: "moveWorkbookRange", sheetId: source!.sheetId, range: { from: { row: 2, column: 0 }, to: { row: 2, column: 0 } }, targetSheetId: other!.sheetId, target: { row: 19, column: 7 } };
  const back: DocumentOperation = { op: "moveWorkbookRange", sheetId: other!.sheetId, range: { from: { row: 19, column: 7 }, to: { row: 19, column: 7 } }, targetSheetId: source!.sheetId, target: { row: 2, column: 0 } };
  return { s, source: source!, other: other!, move, back };
}
async function read(s: WorkbookSession) {
  const source = createWorkbookPreviewSource(s), m = s.snapshot().manifest;
  return source.charts!.read({ ...DEFAULT_WORKBOOK_CHART_READ_LIMITS, documentId: m.documentId, revision: m.revision, chartIds: s.snapshot().metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
}
test.each(["=!$B$1", "=$B$1", "=!B1", "=B1"])("moving a chart retains its expression worksheet for nested global recipe %s and current public data", async expression => {
  const { s, source, other, move, back } = example(expression), before = await read(s), bytes = s.serialize();
  apply(s, [move]); const after = await read(s);
  expect(after.charts[0]!.sheetId).toBe(other.sheetId); expect(after.charts[0]!.references).toEqual(before.charts[0]!.references);
  expect(s.snapshot().metadata.workbook.charts![0]!.series[0]!.values).toMatchObject({ expression: "=Nested", originSheetId: source.sheetId });
  s.undo(); const restored = s.serialize(); expect(apply(s, [move, back]).changed).toBe(false); expect(s.serialize()).toEqual(restored);
  expect(s.snapshot().metadata).toEqual(openDocument(bytes, { kind: "workbook" }).snapshot().metadata);
});
test("a moved chart reads current original inputs after unrelated target edits and source renaming", async () => {
  const { s, source, other, move } = example("=!$B$1"); apply(s, [move]);
  apply(s, [{ op: "setWorkbookCells", sheetId: source.sheetId, cells: [{ row: 0, column: 1, source: { kind: "value", value: { type: "number", value: 42 } } }] },
    { op: "setWorkbookCells", sheetId: other.sheetId, cells: [{ row: 0, column: 1, source: { kind: "value", value: { type: "number", value: 123 } } }] },
    { op: "updateWorkbookSheet", sheetId: source.sheetId, changes: { name: "来源 新名" } }]);
  const data = (await read(s)).charts[0]!.references.find(reference => reference.role === "series-values")!.data;
  expect(data).toMatchObject({ kind: "cells", areas: [{ range: { sheetId: source.sheetId }, cells: [{ value: { type: "number", value: 42 } }] }] });
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
});
test("deleting an expression's source worksheet invalidates that context without deleting its independently owned chart", async () => {
  const { s, source, other, move } = example("=!$B$1"); apply(s, [move]); const before = s.snapshot();
  apply(s, [{ op: "removeWorkbookSheet", sheetId: source.sheetId }]); const current = s.snapshot().metadata.workbook;
  expect(current.objects![0]!.sheetId).toBe(other.sheetId); expect(current.charts![0]!.series[0]!.values).toEqual({ expression: "=#REF!", valueType: "number" });
  expect((await read(s)).charts[0]!.references.find(reference => reference.role === "series-values")!.data).toEqual({ kind: "error", code: "#REF!" });
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
});
test.each([false, true])("whole-sheet copy maps an owning origin and retains an independent origin (moved=%s)", async moved => {
  const { s, source, other, move } = example("=!$B$1");
  if (moved) apply(s, [move]); else {
    const chart = s.snapshot().metadata.workbook.charts![0]!;
    apply(s, [{ op: "setWorkbookChartData", nodeId: chart.nodeId, data: { series: [{ seriesId: chart.series[0]!.seriesId, values: { expression: "=CallerValue", valueType: "number", originSheetId: source.sheetId } }] } }]);
  }
  apply(s, [{ op: "copyWorkbookSheet", sheetId: moved ? other.sheetId : source.sheetId, name: "Copied" }]);
  const snapshot = s.snapshot(), clone = snapshot.metadata.workbook.sheets.at(-1)!, copiedChart = snapshot.metadata.workbook.charts!.at(-1)!;
  expect(copiedChart.series[0]!.values).toMatchObject({ expression: moved ? "=Nested" : "=CallerValue", originSheetId: moved ? source.sheetId : clone.sheetId });
  apply(s, [{ op: "setWorkbookCells", sheetId: clone.sheetId, cells: [{ row: 0, column: 1, source: { kind: "value", value: { type: "number", value: 77 } } }] }]);
  const actual = (await read(s)).charts.find(chart => chart.chartId === copiedChart.chartId)!.references.find(reference => reference.role === "series-values")!.data;
  expect(actual).toMatchObject({ kind: "cells", areas: [{ range: { sheetId: moved ? source.sheetId : clone.sheetId }, cells: [{ value: { type: "number", value: moved ? 10 : 77 } }] }] });
});
test("frozen object context discloses the bound source and rejects changing to undisclosed target inputs", () => {
  const { s, source, other, move } = example("=!$B$1"); apply(s, [move]); const chart = s.snapshot().metadata.workbook.charts![0]!, m = s.snapshot().manifest;
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "request", documentId: m.documentId, baseRevision: m.revision, status: "open", instruction: "复核图表输入", anchor: { kind: "cells", sheetId: other.sheetId, ranges: [{ from: { row: 19, column: 7 }, to: { row: 19, column: 7 } }] } } }]); const before = s.serialize();
  const request = (expanded = false) => s.buildInteractionRequest({ kind: "workbook", annotationIds: ["request"], scope: { kind: "ranges", ranges: [{ sheetId: other.sheetId, range: { from: { row: 19, column: 7 }, to: { row: 19, column: 7 } } }], objectIds: [chart.nodeId] },
    ...(expanded ? { readRanges: [{ sheetId: other.sheetId, range: { from: { row: 0, column: 1 }, to: { row: 0, column: 1 } } }] } : {}) });
  const input = request(); expect(input.context.cells).toEqual(expect.arrayContaining([expect.objectContaining({ sheetId: source.sheetId, row: 0, column: 1, access: "read-only" })]));
  expect(input.context.cells.some(cell => cell.sheetId === other.sheetId && cell.row === 0 && cell.column === 1)).toBe(false);
  const proposal = (input: ReturnType<typeof request>) => ({ protocolVersion: 1 as const, requestId: input.requestId, documentId: input.documentId, baseRevision: input.baseRevision, proposalId: "rebind", annotationIds: ["request"], summary: "选择新的图表输入", operations: [{ op: "setWorkbookChartData" as const, nodeId: chart.nodeId, data: { series: [{ seriesId: chart.series[0]!.seriesId, values: { expression: "=Nested", valueType: "number" as const, originSheetId: other.sheetId } }] } }] });
  expect(() => s.prepareProposal(input, proposal(input))).toThrowError(expect.objectContaining({ code: "InvalidProposal" })); expect(s.serialize()).toEqual(before);
  const expanded = request(true), candidate = s.prepareProposal(expanded, proposal(expanded)); expect(candidate.preview.metadata.workbook.charts![0]!.series[0]!.values).toMatchObject({ originSheetId: other.sheetId }); s.discard(candidate); expect(s.serialize()).toEqual(before);
});
test("missing expression origins and origins on literal vectors are rejected atomically", () => {
  const { s } = example("=!$B$1"), chart = s.snapshot().metadata.workbook.charts![0]!, before = s.serialize();
  expect(() => apply(s, [{ op: "setWorkbookChartData", nodeId: chart.nodeId, data: { series: [{ seriesId: chart.series[0]!.seriesId, values: { expression: "=Nested", valueType: "number", originSheetId: "missing" } }] } }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  const files = new Map(before), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)); metadata.workbook.charts[0].series[0].values = { kind: "literal", count: 0, points: [], valueType: "number", originSheetId: "sheet-1" };
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); expect(() => openDocument(files, { kind: "workbook" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" })); expect(s.serialize()).toEqual(before);
});
test("explicitly selecting the current expression worksheet preserves source spellings, caches and history when the binding is unchanged", () => {
  const { s, source, other, move } = example("=!$B$1"), chart = s.snapshot().metadata.workbook.charts![0]!, before = s.serialize();
  expect(apply(s, [{ op: "setWorkbookChartData", nodeId: chart.nodeId, data: { categories: { expression: "=A1", valueType: "text", originSheetId: source.sheetId },
    series: [{ seriesId: chart.series[0]!.seriesId, values: { expression: "=Nested", valueType: "number", originSheetId: source.sheetId } }] } }]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0); apply(s, [move]);
  expect(apply(s, [{ op: "setWorkbookChartData", nodeId: chart.nodeId, data: { categories: { expression: "=A1", valueType: "text", originSheetId: source.sheetId },
    series: [{ seriesId: chart.series[0]!.seriesId, values: { expression: "=Nested", valueType: "number", originSheetId: other.sheetId } }] } }]).changed).toBe(true);
});
test("move and data replacement in one candidate compare against the current chart origin and their inverse restores exact source", async () => {
  const { s, source, move, back } = example("=!$B$1"), before = s.serialize(), initial = await read(s), chart = s.snapshot().metadata.workbook.charts![0]!;
  const preserve: DocumentOperation = { op: "setWorkbookChartData", nodeId: chart.nodeId, data: { categories: { expression: "=A1", valueType: "text", originSheetId: source.sheetId },
    series: [{ seriesId: chart.series[0]!.seriesId, values: { expression: "=Nested", valueType: "number", originSheetId: source.sheetId } }] } };
  apply(s, [move, preserve]); expect((await read(s)).charts[0]!.references).toEqual(initial.charts[0]!.references);
  s.undo(); const restored = s.serialize(); expect(apply(s, [move, preserve, back]).changed).toBe(false); expect(s.serialize()).toEqual(restored);
  expect(s.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
});
test("sole Worker transports explicit chart origins through private candidate measurements and the public chart endpoint", async () => {
  const { s, source, other, move } = example("=!$B$1"), before = s.snapshot(), initial = await read(s);
  const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const m = before.manifest, command = { kind: "transaction" as const, transaction: { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations: [move] } };
    const measureLayout = (request: WorkbookLayoutSourceRequest): WorkbookLayoutMeasurement => ({ sourceDigest: request.sourceDigest, fontEnvironment: "worker-chart-origin-points", sheets: request.sheetIds.map(sheetId => ({ sheetId,
      rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
    const discarded = await client.prepare(command, undefined, { measureLayout }); await client.discard(discarded); expect(s.snapshot()).toEqual(before);
    await client.commit(await client.prepare(command, undefined, { measureLayout })); const view = await client.view();
    try {
      const manifest = view.cells.snapshot.manifest, chart = view.cells.snapshot.metadata.workbook.charts![0]!;
      expect(chart.series[0]!.values).toMatchObject({ expression: "=Nested", originSheetId: source.sheetId });
      const actual = await view.charts!.read({ ...DEFAULT_WORKBOOK_CHART_READ_LIMITS, documentId: manifest.documentId, revision: manifest.revision, chartIds: [chart.chartId] }, new AbortController().signal);
      expect(actual.charts[0]!.sheetId).toBe(other.sheetId); expect(actual.charts[0]!.references).toEqual(initial.charts[0]!.references);
    } finally { view.dispose(); }
    await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.history().undoSteps).toBe(0);
  } finally { client.dispose(); server.dispose(); }
});
