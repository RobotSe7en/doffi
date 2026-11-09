import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { WorkbookLayoutMeasurement, WorkbookLayoutRequest } from "@doffi/core";
import { createWorkbookRangeCopyTransaction } from "@doffi/editor";
import { createWorkbookPreviewSource, DEFAULT_WORKBOOK_CHART_READ_LIMITS } from "@doffi/render";
import type { WorkbookChartReferenceData } from "@doffi/render";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

interface Case { name: string; sourceRange: { from: { row: number; column: number }; to: { row: number; column: number } }; target: { row: number; column: number }; otherSheet: boolean; outputFile: string; outputSha256: string; observed: { drawings: unknown[]; charts: string[][] } }
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-copy-objects.json", "utf8")) as { sources: { name: string; source: string; sourceSha256: string; cases: Case[] }[] };
const cases = evidence.sources.flatMap(source => source.cases.map(producer => ({ source, producer, label: source.name + "/" + producer.name })));
const metrics = (request: WorkbookLayoutRequest): WorkbookLayoutMeasurement => ({ sourceDigest: request.sourceDigest, fontEnvironment: "native-object-copy-explicit-points", sheets: request.sheetIds.map(sheetId => ({ sheetId,
  rows: { defaultSizePt: request.model.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!.defaultRowHeightPt, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
const dataValues = (data: WorkbookChartReferenceData) => data.kind === "cells" ? data.areas.flatMap(area => area.cells.map(cell => cell.value)) : data;
test("eight native object copies retain original bytes and actual producer vector boundaries", () => {
  for (const source of evidence.sources) expect(createHash("sha256").update(readFileSync(source.source)).digest("hex")).toBe(source.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-copy-objects.py"], { encoding: "utf8" })).toContain("Verified eight");
});
test.each(cases)("independent source $label copies current cells, fresh drawings and authorized resources in one undo step", async ({ source, producer }) => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync(source.source), { documentId: "native-object-copy" }), { kind: "workbook" }));
  try {
    if (producer.otherSheet) s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations: [{ op: "insertWorkbookSheet", name: "Other" }] }, { measureWorkbookLayout: metrics }));
    s.clearHistory(); const before = s.snapshot(), original = before.metadata.workbook.objects!, sheetId = before.metadata.workbook.sheets[0]!.sheetId, targetSheetId = producer.otherSheet ? before.metadata.workbook.sheets.at(-1)!.sheetId : sheetId;
    const read = async () => { const m = s.snapshot().manifest, input = createWorkbookPreviewSource(s); return input.charts!.read({ ...DEFAULT_WORKBOOK_CHART_READ_LIMITS, documentId: m.documentId, revision: m.revision, chartIds: s.snapshot().metadata.workbook.charts?.map(chart => chart.chartId) ?? [] }, new AbortController().signal); };
    const priorData = await read(), html = new Map(original.map(object => [object.nodeId, s.getNode(object.nodeId)!.html]));
    const tx = await createWorkbookRangeCopyTransaction(s, sheetId, producer.sourceRange, producer.target, { targetSheetId }), candidate = s.prepare(tx, { measureWorkbookLayout: metrics });
    expect(s.snapshot()).toEqual(before); s.commit(candidate); const after = s.snapshot(), added = after.metadata.workbook.objects!.filter(object => !original.some(old => old.nodeId === object.nodeId));
    const count = source.name === "charts" ? producer.name.startsWith("full-") ? 3 : 1 : producer.name === "partial-box" ? 2 : 3;
    expect(added).toHaveLength(count); expect(after.metadata.workbook.objects!.slice(0, original.length)).toEqual(original); expect(after.manifest.assets).toEqual(before.manifest.assets); expect(s.history().undoSteps).toBe(1);
    for (const object of original) expect(s.getNode(object.nodeId)!.html).toBe(html.get(object.nodeId));
    expect(added.every(object => object.sheetId === targetSheetId && object.nativeId === undefined)).toBe(true);
    const current = await read(), newCharts = after.metadata.workbook.charts?.filter(chart => !before.metadata.workbook.charts!.some(old => old.chartId === chart.chartId)) ?? [];
    for (const [index, chart] of newCharts.entries()) {
      const previous = priorData.charts[index]!, copied = current.charts.find(value => value.chartId === chart.chartId)!;
      expect(copied.references.map(reference => [reference.role, dataValues(reference.data)])).toEqual(previous.references.map(reference => [reference.role, dataValues(reference.data)]));
      expect(chart.series.every(series => series.nativeIndex === undefined)).toBe(true);
    }
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    if (source.name === "charts" && producer.name === "full-other") {
      const m = s.snapshot().manifest; s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [{ op: "setWorkbookCells", sheetId: targetSheetId, cells: [{ row: 60, column: 8, source: { kind: "value", value: { type: "number", value: 12345.678 } } }] }] }, { measureWorkbookLayout: metrics }));
      const result = (await read()).charts.find(chart => chart.chartId === newCharts[0]!.chartId)!.references.find(reference => reference.role === "series-values")!.data;
      expect(dataValues(result)).toEqual(expect.arrayContaining([{ type: "number", value: 12345.678 }])); expect(s.getCell(sheetId, 1, 1)?.value).toEqual(openDocument(await importXlsx(readFileSync(source.source), { documentId: "original-check" }), { kind: "workbook" }).getCell(sheetId, 1, 1)?.value); s.undo();
    }
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
test.each(cases)("actual native producer $label reopens all observed copied objects and accepts another source edit", async ({ producer }) => {
  const bytes = readFileSync(producer.outputFile); expect(createHash("sha256").update(bytes).digest("hex")).toBe(producer.outputSha256);
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "object-copy-producer" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), m = before.manifest; expect(before.metadata.workbook.objects).toHaveLength(producer.observed.drawings.length); expect(before.metadata.workbook.charts ?? []).toHaveLength(producer.observed.charts.length);
    s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [{ op: "updateWorkbookObject", nodeId: before.metadata.workbook.objects!.at(-1)!.nodeId, changes: { name: "独立副本再编辑禰󠄀😀" } }] }, { measureWorkbookLayout: metrics }));
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
