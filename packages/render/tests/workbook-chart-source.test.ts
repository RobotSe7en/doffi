import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { describe, expect, it } from "vitest";
import { createWorkbook, createWorkbookCalculator, createWorkbookNumberFormatter, openDocument } from "@doffi/core";
import type { WorkbookChart, WorkbookDocumentMetadata, WorkbookManifest, WorkbookModel } from "@doffi/core";
import { readXlsxRecords } from "../../office/src/xlsx/read.js";
import { importXlsx } from "../../office/src/xlsx-import.js";
import { projectXlsxRecords } from "../../office/tests/helpers/xlsx-record-model.js";
import { DEFAULT_CHART_READ_LIMITS, modelChartSource } from "../src/workbook/chart-source.js";
import type { WorkbookChartReadRequest, WorkbookChartReadResponse } from "../src/workbook/chart-source.js";

import { bindWorkbookChartData } from "../src/workbook/chart-values.js";

const encode = (text: string) => new TextEncoder().encode(text), decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
const reference = (expression: string) => ({ expression, valueType: "number" as const, cache: { state: "imported" as const, count: 1, points: [{ index: 0, value: 999 }] } });
function seeded(edit: (chart: WorkbookChart, metadata: WorkbookDocumentMetadata) => void = () => {}, unavailableFormula = '=UNSUPPORTED()'): WorkbookModel {
  const files = new Map(createWorkbook({ documentId: "chart-data", sheetName: "Data" }));
  const manifest = JSON.parse(decode(files.get("manifest.json")!)) as WorkbookManifest;
  const metadata = JSON.parse(decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets.push({ ...metadata.workbook.sheets[0]!, sheetId: "second", name: "第二表" });
  manifest.content = "content.html";
  metadata.workbook.objects = [{ kind: "chart", nodeId: "chart-node", chartId: "chart", sheetId: "sheet-1", locked: true, printable: true, anchor: { kind: "absolute", xPt: 0, yPt: 0, widthPt: 400, heightPt: 250 } }];
  const chart: WorkbookChart = { chartId: "chart", nodeId: "chart-node", kind: "column", legend: "bottom", showLabels: false,
    categories: { expression: "=A1:A2", valueType: "text" }, series: [{ seriesId: "values", values: reference("=B1:B6") }] };
  metadata.workbook.charts = [chart]; edit(chart, metadata);
  files.set("content.html", encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body><section data-node-id="content" data-office-kind="workbook-content"><div data-node-id="chart-node" data-office-kind="chart"></div></section></body></html>'));
  const path = "sheets/sheet-1/0-0.html", bytes = encode('<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="0" data-column="0"><tbody>'
    + '<tr data-row="0"><td data-column="0" data-value-type="text">中文😀</td><td data-column="1" data-value-type="number" data-value="12">旧显示</td></tr>'
    + '<tr data-row="1"><td data-column="0" data-value-type="boolean" data-value="true">TRUE</td><td data-column="1" data-value-type="number" data-value="999" data-formula="=20+3" data-cache-state="imported">999</td></tr>'
    + '<tr data-row="2"><td data-column="1" data-value-type="number" data-value="999" data-formula="=30+4" data-cache-state="valid">999</td></tr>'
    + '<tr data-row="3"><td data-column="1" data-value-type="number" data-value="999" data-formula="=40+5" data-cache-state="stale">999</td></tr>'
    + '<tr data-row="4"><td data-column="1" data-value-type="number" data-value="999" data-formula="' + unavailableFormula + '" data-cache-state="valid">999</td></tr>'
    + '<tr data-row="5"><td data-column="1" data-value-type="error" data-value="#N/A">#N/A</td></tr>'
    + '</tbody></table></body></html>');
  files.set(path, bytes); manifest.chunks.push({ sheetId: "sheet-1", row: 0, column: 0, path, sha256: createHash("sha256").update(bytes).digest("hex") });
  files.set("manifest.json", encode(JSON.stringify(manifest))); files.set("metadata.json", encode(JSON.stringify(metadata)));
  return openDocument(files, { kind: "workbook" });
}
const request = (model: WorkbookModel, overrides: Partial<WorkbookChartReadRequest> = {}): WorkbookChartReadRequest => ({ ...DEFAULT_CHART_READ_LIMITS,
  documentId: model.snapshot().manifest.documentId, revision: model.snapshot().manifest.revision, chartIds: ["chart"], ...overrides });
const read = (model: WorkbookModel, overrides: Partial<WorkbookChartReadRequest> = {}, signal = new AbortController().signal) => modelChartSource(model).read(request(model, overrides), signal);
function cells(response: WorkbookChartReadResponse, role = "series-values") {
  const result = response.charts[0]!.references.find(item => item.role === role)!.data;
  if (result.kind !== "cells") throw new Error("Expected source cells");
  return result.areas.flatMap(area => area.cells);
}

describe("bounded current-source workbook chart reads", () => {
  it("reads current formulas with the completed batch's explicit time context after calculator disposal", async () => {
    const model = seeded(() => {}, "=TODAY()"), before = model.serialize(), calculator = createWorkbookCalculator(model);
    const calculation = calculator.calculateSync({ context: { now: { unixMilliseconds: 0, utcOffsetMinutes: 0 } } });
    const provider = modelChartSource(model, calculation); calculator.dispose();
    const response = await provider.read(request(model), new AbortController().signal);
    expect(cells(response)[4]).toMatchObject({ state: "valid", value: { type: "number", value: 25569 } });
    expect(model.serialize()).toEqual(before);
  });
  it("returns immutable sparse literals without resolving formulas or reading worksheet cells", async () => {
    const model = seeded(chart => {
      chart.categories = { kind: "literal", valueType: "text", count: 4, points: [{ index: 0, value: "季度😀" }] };
      chart.series[0]!.values = { kind: "literal", valueType: "number", count: 4, numberFormat: "0.00", points: [{ index: 0, value: -3 }, { index: 2, value: null }, { index: 3, value: "#N/A", numberFormat: "0%" }] };
    });
    const provider = { ...model, resolveReference() { throw new Error("No literal has a reference"); }, readWindow() { throw new Error("No literal owns cells"); } };
    const response = await read(provider, { maxPoints: 8 });
    expect(response.charts[0]!.references.map(ref => ref.data)).toEqual([model.snapshot().metadata.workbook.charts![0]!.categories, model.snapshot().metadata.workbook.charts![0]!.series[0]!.values]);
    expect(structuredClone(response)).toEqual(response); expect(Object.isFrozen(response.charts[0]!.references[0]!.data)).toBe(true);
    await expect(read(provider, { maxPoints: 7 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    await expect(read(provider, { maxCodeUnits: 24 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  });
  it("preflights literal counts before reading any referenced cells and yields during literal text admission", async () => {
    let queried = 0;
    const model = seeded(chart => { chart.series[0]!.values = { kind: "literal", valueType: "number", count: 1024, points: Array.from({ length: 1024 }, (_, index) => ({ index, value: index })) }; });
    const provider = { ...model, readWindow(window: Parameters<WorkbookModel["readWindow"]>[0]) { queried++; return model.readWindow(window); } };
    await expect(read(provider, { maxPoints: 1025 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" }); expect(queried).toBe(0);
    const controller = new AbortController(), pending = read(provider, {}, controller.signal); controller.abort();
    await expect(pending).rejects.toMatchObject({ code: "OperationCancelled" }); expect(queried).toBe(0);
    const response = await read(provider); expect(cells(response, "categories")).toHaveLength(2); expect(response.charts[0]!.references[1]!.data.kind).toBe("literal");
  });
  it("keeps typed values, formula states and sparse blanks distinct from imported chart caches", async () => {
    const model = seeded(), source = model.serialize(), response = await read(model);
    expect(cells(response).map(cell => cell.state)).toEqual(["constant", "valid", "valid", "valid", "unavailable", "constant"]);
    expect(cells(response)[4]).toMatchObject({ issue: { reason: "unsupported-function", functionName: "UNSUPPORTED" }, value: { type: "blank" } });
    expect(model.readWindow({ sheetId: 'sheet-1', row: 1, column: 1, rows: 3, columns: 1 }).map(cell => [cell.value, cell.formula?.cacheState]))
      .toEqual(['imported', 'valid', 'stale'].map(state => [{ type: 'number', value: 999 }, state]));
    expect(cells(response).map(cell => cell.value)).toEqual([{ type: "number", value: 12 }, { type: "number", value: 23 }, { type: "number", value: 34 }, { type: "number", value: 45 }, { type: "blank" }, { type: "error", value: "#N/A" }]);
    expect(cells(response, "categories").map(cell => cell.value)).toEqual([{ type: "text", value: "中文😀" }, { type: "boolean", value: true }]);
    expect(model.snapshot().metadata.workbook.charts![0]!.series[0]!.values).toMatchObject({ cache: { points: [{ index: 0, value: 999 }] } });
    expect(model.serialize()).toEqual(source); expect(Object.isFrozen(response.charts[0]?.references)).toBe(true); expect(Object.isFrozen(cells(response)[0]?.value)).toBe(true);
    expect(structuredClone(response)).toEqual(response);
  });
  it("retains exact region order, repeated cells, cross-sheet blanks and rectangular row order", async () => {
    const model = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference("=(B3:B4,第二表!XFD1048576,A1:B2,B1)"); });
    const response = await read(model), data = response.charts[0]!.references[0]!.data;
    expect(data.kind === "cells" && data.areas.map(area => area.range.sheetId)).toEqual(["sheet-1", "second", "sheet-1", "sheet-1"]);
    expect(cells(response).map(cell => [cell.row, cell.column])).toEqual([[2, 1], [3, 1], [1048575, 16383], [0, 0], [0, 1], [1, 0], [1, 1], [0, 1]]);
    expect(cells(response)[2]).toMatchObject({ state: "constant", value: { type: "blank" } });
    await expect(read(model, { maxPoints: 7 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  });
  it('bounds cumulative formula work across slots and reuses repeated current coordinates within one request', async () => {
    const single = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference('=B2'); });
    await expect(read(single, { maxFormulaSteps: 1 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    const repeated = seeded(chart => { chart.categories = reference('=B2'); chart.series[0]!.values = reference('=(B2,B2)'); });
    expect(cells(await read(repeated, { maxFormulaCells: 1 })).map(cell => cell.value)).toEqual([{ type: 'number', value: 23 }, { type: 'number', value: 23 }]);
    const distinct = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference('=B2:B3'); });
    await expect(read(distinct, { maxFormulaCells: 1 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    // A rejected read cannot poison another read's calculation or budgets.
    expect(cells(await read(distinct, { maxFormulaCells: 2 })).map(cell => cell.value)).toEqual([{ type: 'number', value: 23 }, { type: 'number', value: 34 }]);
  });
  it.each([['=B5', 'circular-reference'], ['=Loop', 'circular-name']] as const)('keeps %s unavailable despite a source cache marked valid', async (expression, reason) => {
    const model = seeded((chart, metadata) => {
      delete chart.categories; chart.series[0]!.values = reference('=B5');
      metadata.workbook.names = [{ nameId: 'loop', name: 'Loop', expression: '=Loop', origin: { row: 0, column: 0 } }];
    }, expression), before = model.serialize();
    const points = cells(await read(model));
    expect(points[0]).toMatchObject({ state: 'unavailable', value: { type: 'blank' }, issue: { reason } });
    expect(model.serialize()).toEqual(before);
  });
  it("resolves names in the owning worksheet rather than the active sheet, including title and series slots", async () => {
    const model = seeded((chart, metadata) => {
      metadata.workbook.activeSheetId = "second";
      metadata.workbook.names = [{ nameId: "local", name: "LocalValues", sheetId: "sheet-1", expression: "=$B$1:$B$2", origin: { row: 0, column: 0 } }];
      chart.series[0]!.values = reference("=LocalValues"); chart.series[0]!.name = { kind: "reference", reference: { expression: "=A1", valueType: "text" } };
      chart.titleReference = { expression: "=A1", valueType: "text" }; chart.categoryAxisTitle = { reference: { expression: "=A1", valueType: "text" } }; chart.valueAxisTitle = { reference: { expression: "=A2", valueType: "text" } };
    });
    const response = await read(model);
    expect(response.charts[0]!.references.map(ref => ref.role)).toEqual(["categories", "title", "category-axis-title", "value-axis-title", "series-name", "series-values"]);
    expect(cells(response).map(cell => cell.value)).toEqual([{ type: "number", value: 12 }, { type: "number", value: 23 }]);
    expect(cells(response, "title")[0]!.value).toEqual({ type: "text", value: "中文😀" });
    expect(response.charts[0]!.references.at(-1)).toMatchObject({ seriesId: "values" });
  });
  it("returns dynamic, circular and broken reference states without using their imported caches", async () => {
    for (const [expression, expected] of [["=INDEX(A1:B5,0,1)", { kind: "unresolved", reason: "calculation-required" }], ["=Loop", { kind: "unresolved", reason: "circular-name" }], ["=Missing!B1", { kind: "error", code: "#REF!" }], ["=MissingName", { kind: "error", code: "#NAME?" }]] as const) {
      const model = seeded((chart, metadata) => { delete chart.categories; chart.series[0]!.values = reference(expression); metadata.workbook.names = [{ nameId: "loop", name: "Loop", expression: "=Loop", origin: { row: 0, column: 0 } }]; });
      expect((await read(model)).charts[0]!.references[0]!.data).toEqual(expected);
    }
  });
  it("preserves hidden flags and inherited number formats without filtering or coercing points", async () => {
    const model = seeded((chart, metadata) => {
      delete chart.categories; chart.series[0]!.values = reference("=A1:B3");
      metadata.workbook.numberFormats.push({ formatId: "row-format", code: "0.00" }, { formatId: "column-format", code: "0%" });
      const sheet = metadata.workbook.sheets[0]!; sheet.defaultRowsHidden = true;
      sheet.rowDimensions = [{ from: 1, to: 1, hidden: false, numberFormatId: "row-format" }];
      sheet.columnDimensions = [{ from: 1, to: 1, unit: "characters", hidden: true, numberFormatId: "column-format" }];
    });
    const points = cells(await read(model));
    expect(points.map(point => point.rowHidden)).toEqual([true, true, false, false, true, true]);
    expect(points.map(point => point.columnHidden)).toEqual([false, true, false, true, false, true]);
    expect(points[1]!.numberFormatId).toBe("column-format"); expect(points[3]!.numberFormatId).toBe("row-format");
  });
  it('accumulates formula dependency text while charging a cached execution only once', async () => {
    function longDependencies(expression: string) {
      const files = new Map(seeded(chart => { delete chart.categories; chart.series[0]!.values = reference(expression); }).serialize());
      const manifest = JSON.parse(decode(files.get('manifest.json')!)) as WorkbookManifest, part = manifest.chunks[0]!;
      let html = decode(files.get(part.path)!);
      for (const [row, expression] of [[1, '=20+3'], [2, '=30+4']] as const) {
        html = html.replace(`data-formula="${expression}"`, `data-formula="=LEN(C${row + 1})"`);
        html = html.replace(new RegExp('<tr data-row="' + row + '">.*?</tr>', 's'), source => source.replace('</tr>', `<td data-column="2" data-value-type="text">${'x'.repeat(60)}</td></tr>`));
      }
      const bytes = encode(html); files.set(part.path, bytes); part.sha256 = createHash('sha256').update(bytes).digest('hex');
      files.set('manifest.json', encode(JSON.stringify(manifest))); return openDocument(files, { kind: 'workbook' });
    }
    const distinct = longDependencies('=B2:B3'), before = distinct.serialize();
    await expect(read(distinct, { maxCodeUnits: 150 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    expect(cells(await read(distinct, { maxCodeUnits: 1024 })).map(cell => cell.value)).toEqual([{ type: 'number', value: 60 }, { type: 'number', value: 60 }]);
    expect(cells(await read(longDependencies('=(B2,B2)'), { maxCodeUnits: 150 }))).toHaveLength(2);
    expect(distinct.serialize()).toEqual(before);
  });
  it("admits all references before reading any cells and rejects invalid identity or duplicate requests", async () => {
    let reads = 0;
    const source = seeded(chart => { chart.series[0]!.values = reference("=B:B"); }), model = { ...source, readWindow: (window: Parameters<WorkbookModel["readWindow"]>[0]) => { reads++; return source.readWindow(window); } };
    for (const override of [{}, { maxReferences: 1 }, { maxCharts: 0 }, { chartIds: ["chart", "chart"] }, { chartIds: ["missing"] }, { revision: 999 }, { documentId: "other" }]) await expect(read(model, override)).rejects.toBeDefined();
    expect(reads).toBe(0);
    expect((await read(model, { chartIds: [] })).charts).toEqual([]);
  });
  it("counts actual UTF-16 values and repeated occurrences against the text budget", async () => {
    const model = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference("=(A1,A1)"); });
    // chart id, owner, expression, series id, two area sheet ids, two four-unit values.
    const units = "chart".length + "sheet-1".length + "=(A1,A1)".length + "values".length + 2 * "sheet-1".length + 2 * "中文😀".length;
    expect(cells(await read(model, { maxCodeUnits: units }))).toHaveLength(2);
    await expect(read(model, { maxCodeUnits: units - 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  });
  it.each(["=A1:A1000", "=A1:ALL1", "=A1:AD20"])("yields bounded queries and cancels %s before reading the next batch", async expression => {
    const model = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference(expression); }), queried: number[] = [];
    const controller = new AbortController(), provider = modelChartSource({ ...model, readWindow(window) { queried.push(window.rows * window.columns); return model.readWindow(window); } });
    const pending = provider.read(request(model), controller.signal);
    expect(queried.length).toBeGreaterThan(0); expect(Math.max(...queried)).toBeLessThanOrEqual(256); const completed = [...queried]; controller.abort();
    await expect(pending).rejects.toMatchObject({ code: "OperationCancelled" }); expect(queried).toEqual(completed);
    await expect(provider.read(request(model), controller.signal)).rejects.toMatchObject({ code: "OperationCancelled" });
  });
  it("isolates requests before yielding and returns only the selected immutable revision", async () => {
    const model = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference("=A1:A512"); });
    const ids = ["chart"], input = { ...request(model), chartIds: ids }, pending = modelChartSource(model).read(input, new AbortController().signal);
    ids[0] = "missing"; input.documentId = "other"; input.revision = 999; input.maxPoints = 1;
    const response = await pending; expect(response.documentId).toBe("chart-data"); expect(response.revision).toBe(model.snapshot().manifest.revision); expect(cells(response)).toHaveLength(512);
  });
  it("allows cancellation during reference admission before any cells are read", async () => {
    const model = seeded(chart => { delete chart.categories; chart.series = Array.from({ length: 64 }, (_, index) => ({ seriesId: `series-${index}`, values: reference("=B1") })); });
    let bound = 0, queried = 0;
    const provider = modelChartSource({ ...model, resolveReference(expression, origin) { bound++; return model.resolveReference(expression, origin); }, readWindow(window) { queried++; return model.readWindow(window); } });
    const controller = new AbortController(), pending = provider.read(request(model), controller.signal);
    expect(bound).toBe(32); expect(queried).toBe(0); controller.abort();
    await expect(pending).rejects.toMatchObject({ code: "OperationCancelled" }); expect(bound).toBe(32); expect(queried).toBe(0);
  });
  it.each(["examples/xlsx/07-three-charts.xlsx", "examples/xlsx/producer/07-three-charts.xlsx"])("reads all three independent native chart sources from %s", async path => {
    const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync(path))), definitions = model.snapshot().metadata.workbook.charts!;
    const response = await read(model, { chartIds: definitions.map(chart => chart.chartId) });
    expect(response.charts).toHaveLength(3);
    for (const chart of response.charts) {
      const values = chart.references.filter(ref => ref.role === "series-values");
      expect(values[0]!.data.kind === "cells" && values[0]!.data.areas[0]!.cells.map(cell => cell.value)).toEqual([2, 5, 8].map(value => ({ type: "number", value })));
      expect(values.every(ref => ref.data.kind === "cells" && ref.data.areas[0]!.cells.every(cell => cell.state === "constant"))).toBe(true);
    }
  });
  it("reads mixed values and hidden-source coordinates independently of deliberately stale native chart caches", async () => {
    const bytes = execFileSync('python3', ['-c', "import importlib.util,sys; s=importlib.util.spec_from_file_location('chart_values','scripts/check-xlsx-chart-values.py'); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); sys.stdout.buffer.write(m.generate(m.inputs()))"], { maxBuffer: 4 * 1024 * 1024 });
    const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-chart-value-semantics.json', 'utf8')) as { sourceSha256: string; cases: { hidden: string; visibleOnly: boolean }[] };
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
    const { model } = projectXlsxRecords(await readXlsxRecords(bytes)), definitions = model.snapshot().metadata.workbook.charts!;
    const response = await read(model, { chartIds: definitions.map(chart => chart.chartId) }); expect(response.charts).toHaveLength(15);
    for (const [index, chart] of response.charts.entries()) {
      const points = cells({ ...response, charts: [chart] });
      expect(points.map(point => point.value)).toEqual([{ type: 'number', value: 2 }, { type: 'text', value: '7' }, { type: 'text', value: 'word' }, { type: 'boolean', value: true }, { type: 'boolean', value: false }, { type: 'blank' }, { type: 'blank' }, { type: 'text', value: '' }, { type: 'error', value: '#N/A' }, { type: 'error', value: '#DIV/0!' }, { type: 'number', value: -3 }, { type: 'number', value: 0 }]);
      expect(points[7]!.state).toBe('valid');
      expect(points.filter(point => point.rowHidden).map(point => point.row)).toEqual(evidence.cases[index]!.hidden === 'rows' ? [2, 4, 11] : []);
      expect(points.every(point => point.columnHidden)).toBe(evidence.cases[index]!.hidden === 'values-column');
      expect(definitions[index]!.series[0]!.values).toMatchObject({ cache: { points: Array.from({ length: 12 }, (_, index) => ({ index, value: 999 })) } });
    }
  });
  it('reads and reopens independent native chart formulas using current values instead of either native cache', async () => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-chart-formulas.xlsx');
    const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-chart-formulas.json', 'utf8')) as { sourceSha256: string; expected: Record<string, unknown> };
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
    let model = openDocument(await importXlsx(bytes, { documentId: 'chart-current' }), { kind: 'workbook' });
    for (let pass = 0; pass < 2; pass++) {
      const before = model.serialize(), response = await read(model, { chartIds: model.snapshot().metadata.workbook.charts!.map(chart => chart.chartId) });
      const address = (row: number, column: number) => String.fromCharCode(65 + column) + (row + 1);
      const points = response.charts[0]!.references.flatMap(reference => reference.data.kind === 'cells' ? reference.data.areas.flatMap(area => area.cells) : []);
      expect(points).toHaveLength(25);
      for (const point of points) { expect(point.state).toBe('valid'); expect(point.value).toEqual(evidence.expected[address(point.row, point.column)]); }
      const cached = model.readWindow({ sheetId: 'sheet-1', row: 1, column: 1, rows: 12, columns: 1 });
      expect(cached.every(cell => cell.value.type === 'number' && cell.value.value === 777 && cell.formula?.cacheState === 'imported')).toBe(true);
      const formatter = createWorkbookNumberFormatter({ dateSystem: '1900' });
      try {
        const [plot] = bindWorkbookChartData(model.snapshot(), response, formatter);
        expect(plot).toMatchObject({ kind: 'available', data: { categories: Array.from({ length: 12 }, (_, index) => 'Point' + index) } });
        expect(plot!.diagnostics).toEqual([expect.objectContaining({ code: 'error-point', detail: '#DIV/0!' })]);
      } finally { formatter.dispose(); }
      expect(model.serialize()).toEqual(before); model = openDocument(before, { kind: 'workbook' });
    }
  });
});


describe("workbook chart display binding", () => {
  it("uses current formats and formula states without borrowing chart caches", async () => {
    const model = seeded((chart, metadata) => {
      chart.categories = reference("=B1:B3"); chart.series[0]!.values = reference("=B1:B3");
      metadata.workbook.numberFormats.push({ formatId: "percent", code: "0.0%" });
      metadata.workbook.sheets[0]!.rowDimensions = [{ from: 1, to: 1, numberFormatId: "percent" }];
    }), source = model.serialize(), formatter = createWorkbookNumberFormatter({ dateSystem: "1900" });
    const [plot] = bindWorkbookChartData(model.snapshot(), await read(model), formatter);
    expect(plot?.kind).toBe("available"); if (plot?.kind !== "available") throw new Error("Expected available plot");
    expect(plot.data.categories).toEqual(["12", "2300.0%", "34"]); expect(plot.data.series[0]!.values).toEqual([12, 23, 34]);
    expect(plot.valueFormats[0]![1]).toEqual({ formatId: "percent", code: "0.0%" });
    expect(plot.diagnostics).toEqual([]);
    expect(Object.isFrozen(plot.data.series[0]!.values)).toBe(true); expect(model.serialize()).toEqual(source); formatter.dispose();
  });
  it("reports stale, unavailable and unresolved references instead of plotting old chart caches", async () => {
    const formatter = createWorkbookNumberFormatter({ dateSystem: "1900" });
    for (const expression of ["=B1:B6", '=OFFSET(B1,0,0)', '=#REF!']) {
      const model = seeded(chart => { chart.series[0]!.values = reference(expression); });
      const [plot] = bindWorkbookChartData(model.snapshot(), await read(model), formatter);
      expect(plot?.kind).toBe("unavailable"); expect(plot).not.toHaveProperty("data");
      const codes = plot!.diagnostics.map(d => d.code);
      if (expression === "=B1:B6") {
        expect(codes).not.toContain("stale-cell"); expect(codes).toContain("unavailable-cell");
        expect(plot!.diagnostics).toContainEqual(expect.objectContaining({ code: 'unavailable-cell', issue: { reason: 'unsupported-function', functionName: 'UNSUPPORTED' } }));
      }
      else expect(codes).toContain(expression.includes("OFFSET") ? "unresolved-reference" : "reference-error");
    }
    const model = seeded(chart => { chart.titleReference = { expression: "=A1:A2", valueType: "text" }; chart.series[0]!.values = reference("=B1"); });
    expect(bindWorkbookChartData(model.snapshot(), await read(model), formatter)[0]).toMatchObject({ kind: "unavailable", diagnostics: expect.arrayContaining([expect.objectContaining({ code: "non-scalar-label", role: "title" })]) });
    formatter.dispose();
  });
  it("retains literal formats, sparse positions and error gaps without fabricating worksheet values", async () => {
    const model = seeded(chart => {
      chart.categories = { kind: "literal", valueType: "text", count: 4, points: [{ index: 0, value: "季度😀" }] };
      chart.series[0]!.values = { kind: "literal", valueType: "number", count: 4, numberFormat: "0.00", points: [{ index: 0, value: -3 }, { index: 2, value: null }, { index: 3, value: "#N/A", numberFormat: "0%" }] };
    });
    const formatter = createWorkbookNumberFormatter({ dateSystem: "1900" }), [plot] = bindWorkbookChartData(model.snapshot(), await read(model), formatter, { maxPoints: 8 });
    if (plot?.kind !== "available") throw new Error("Expected literal plot");
    expect(plot.data.categories).toEqual(["季度😀", "", "", ""]); expect(plot.data.sourceIndexes).toEqual([0, 1, 2, 3]);
    expect(plot.data.series[0]!.values).toEqual([-3, null, null, null]); expect(plot.data.series[0]!.missing).toEqual([undefined, undefined, undefined, "gap"]);
    expect(plot.valueFormats[0]![0]).toEqual({ code: "0.00" }); expect(plot.valueFormats[0]![3]).toEqual({ code: "0%" });
    expect(Object.isFrozen(plot.valueFormats[0]![0])).toBe(true); expect(Object.isFrozen(plot.categoryNumbers)).toBe(true);
    expect(plot.valueAxisFormat).toBe(plot.valueFormats[0]![0]);
    expect(plot.diagnostics).toContainEqual(expect.objectContaining({ code: "error-point", index: 3, detail: "#N/A" }));
    const response = await read(model), wrong = { ...response, charts: response.charts.map(chart => ({ ...chart, references: chart.references.map(ref => ({ ...ref,
      data: ref.data.kind === "literal" && ref.role === "categories" ? { ...ref.data, points: [{ index: 0, value: "other workbook" }] } : ref.data })) })) };
    expect(() => bindWorkbookChartData(model.snapshot(), wrong, formatter)).toThrowError(expect.objectContaining({ code: "InvalidDocument" })); formatter.dispose();
  });
  it("bounds both incoming vectors and padded plot allocations before returning a candidate", async () => {
    const model = seeded(chart => { chart.categories = { kind: "literal", valueType: "text", count: 20, points: [] }; chart.series[0]!.values = reference("=B1"); });
    const formatter = createWorkbookNumberFormatter({ dateSystem: "1900" }), response = await read(model), snapshot = model.snapshot();
    for (const limits of [{ maxPoints: 20 }, { maxPoints: 39 }, { maxCodeUnits: 1 }, { maxReferences: 1 }]) expect(() => bindWorkbookChartData(snapshot, response, formatter, limits)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    const padded = bindWorkbookChartData(snapshot, response, formatter, { maxPoints: 40 })[0]!;
    if (padded.kind !== "available") throw new Error("Expected padded plot");
    expect(padded.data.categories).toHaveLength(20); expect(padded.data.series[0]!.values).toEqual([12, ...Array(19).fill(null)]);
    expect(() => bindWorkbookChartData(snapshot, { ...response, revision: response.revision + 1 }, formatter)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
    for (const charts of [[response.charts[0]!, response.charts[0]!], [{ ...response.charts[0]!, references: [] }], [{ ...response.charts[0]!, sheetId: "second" }]])
      expect(() => bindWorkbookChartData(snapshot, { ...response, charts }, formatter)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    const cached = { ...response, charts: response.charts.map(chart => ({ ...chart, references: chart.references.map(ref => ref.role === "series-values"
      ? { ...ref, data: { kind: "literal" as const, valueType: "number" as const, count: 1, points: [{ index: 0, value: 999 }] } } : ref) })) };
    expect(() => bindWorkbookChartData(snapshot, cached, formatter)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    formatter.dispose();
  });
  it("binds all independent mixed-type and hidden-data probes to current source positions", async () => {
    const bytes = execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('chart_values','scripts/check-xlsx-chart-values.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"]);
    const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-chart-value-semantics.json', 'utf8')) as { cases: { hidden: string; visibleOnly: boolean }[] };
    const { model } = projectXlsxRecords(await readXlsxRecords(bytes)), definitions = model.snapshot().metadata.workbook.charts!;
    const response = await read(model, { chartIds: definitions.map(chart => chart.chartId) }), formatter = createWorkbookNumberFormatter({ dateSystem: '1900' });
    const plots = bindWorkbookChartData(model.snapshot(), response, formatter);
    for (const [index, plot] of plots.entries()) {
      if (plot.kind !== 'available') throw new Error('Expected independently supplied current data');
      const sample = evidence.cases[index]!, positions = Array.from({ length: 12 }, (_, i) => i).filter(i => !sample.visibleOnly || sample.hidden !== 'rows' || ![1, 3, 10].includes(i));
      const hiddenSeries = sample.visibleOnly && sample.hidden === 'values-column';
      expect(plot.data.sourceIndexes).toEqual(hiddenSeries ? [] : positions);
      expect(plot.data.series[0]!.hidden).toBe(hiddenSeries);
      const values = [2, null, null, 1, 0, null, null, null, null, null, -3, 0];
      expect(plot.data.series[0]!.values).toEqual(hiddenSeries ? [] : positions.map(i => values[i]));
      expect(plot.data.categories).toEqual(hiddenSeries ? [] : positions.map(i => sample.visibleOnly && sample.hidden === 'categories-column' ? '' : `Point${i}`));
      expect(plot.data.series[0]!.values).not.toContain(999);
    }
    formatter.dispose();
  });
});
