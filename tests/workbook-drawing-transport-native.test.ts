import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookCoordinate, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookSession } from "@doffi/core";
import { connectWorkbookSession, serveWorkbookSession } from "@doffi/editor";
import type { WorkbookLayoutSourceRequest } from "@doffi/render";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

interface NativeCase {
  name: string; outputFile: string; outputSha256: string;
  sdkOperation: { kind: "cut" | "sort"; from: [number, number]; to: [number, number]; target?: [number, number]; otherSheet?: boolean; rows?: number[] };
  observed: { drawings: { name: string | null; from?: { row: string; col: string } }[]; charts: string[][] };
}
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-drawing-transport.json", "utf8")) as { sources: { name: string; source: string; sourceSha256: string; cases: NativeCase[] }[] };
const cases = evidence.sources.flatMap(source => source.cases.map(producer => ({ source, producer, label: source.name + "/" + producer.name })));
const point = ([row, column]: [number, number]): WorkbookCoordinate => ({ row, column });
function metrics(request: WorkbookLayoutRequest): WorkbookLayoutMeasurement {
  return { sourceDigest: request.sourceDigest, fontEnvironment: "native-drawing-explicit-point-metrics", sheets: request.sheetIds.map(sheetId => ({ sheetId,
    rows: { defaultSizePt: request.model.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!.defaultRowHeightPt, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) };
}
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
test("native drawing transport retains immutable originals and the actual producer's normalized geometry and chart-copy difference", () => {
  for (const source of evidence.sources) expect(createHash("sha256").update(readFileSync(source.source)).digest("hex")).toBe(source.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-drawing-transport.py"], { encoding: "utf8" })).toContain("Verified six");
});
test.each(cases)("independent source $label transports only owned cell-bound drawings and reopens exact one-step history", async ({ source, producer }) => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync(source.source), { documentId: "native-drawing-transport" }), { kind: "workbook" }));
  try {
    const sheetId = s.snapshot().metadata.workbook.sheets[0]!.sheetId, spec = producer.sdkOperation;
    const apply = (operations: DocumentOperation[]) => s.commit(s.prepare(transaction(s, operations), { measureWorkbookLayout: metrics }));
    if (spec.otherSheet) apply([{ op: "insertWorkbookSheet", name: "Other" }]);
    const targetSheetId = spec.otherSheet ? s.snapshot().metadata.workbook.sheets.at(-1)!.sheetId : sheetId;
    if (spec.kind === "sort") apply([{ op: "setWorkbookCells", sheetId, cells: [3, 1, 2].map((value, index) => ({ row: spec.from[0] + index, column: spec.from[1], source: { kind: "value" as const, value: { type: "number" as const, value } } })) }]);
    s.clearHistory(); const before = s.snapshot(), objects = before.metadata.workbook.objects!, html = new Map(objects.map(object => [object.nodeId, s.getNode(object.nodeId)!.html]));
    const operation: DocumentOperation = spec.kind === "sort" ? { op: "reorderWorkbookRows", sheetId, range: { from: point(spec.from), to: point(spec.to) }, rows: spec.rows! }
      : { op: "moveWorkbookRange", sheetId, range: { from: point(spec.from), to: point(spec.to) }, targetSheetId, target: point(spec.target!) };
    const candidate = s.prepare(transaction(s, [operation]), { measureWorkbookLayout: metrics }); expect(s.snapshot()).toEqual(before); s.commit(candidate);
    const current = s.snapshot().metadata.workbook.objects!;
    expect(current.map(object => object.nodeId)).toEqual(objects.map(object => object.nodeId));
    for (const object of objects) {
      const after = current.find(value => value.nodeId === object.nodeId)!, anchor = object.anchor;
      const moving = anchor.kind !== "absolute" && (anchor.kind === "one-cell" || anchor.moveWithCells) && anchor.from.row >= spec.from[0] && anchor.from.row <= spec.to[0] && anchor.from.column >= spec.from[1] && anchor.from.column <= spec.to[1];
      if (moving) expect(after).toMatchObject({ sheetId: targetSheetId, anchor: { kind: anchor.kind, from: {
        row: spec.kind === "sort" ? spec.rows![anchor.from.row - spec.from[0]] : anchor.from.row + spec.target![0] - spec.from[0],
        column: spec.kind === "sort" ? anchor.from.column : anchor.from.column + spec.target![1] - spec.from[1],
      } } });
      else expect(after).toEqual(object);
      expect(s.getNode(object.nodeId)!.html).toBe(html.get(object.nodeId));
    }
    expect(s.snapshot().manifest.assets).toEqual(before.manifest.assets); expect(s.snapshot().metadata.workbook.charts?.map(chart => [chart.chartId, chart.series.map(series => series.seriesId)])).toEqual(before.metadata.workbook.charts?.map(chart => [chart.chartId, chart.series.map(series => series.seriesId)]));
    expect(s.history().undoSteps).toBe(1); expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().content).toEqual(before.content);
  } finally { s.dispose(); }
});
test.each(cases)("native producer $label retains its actual object count, current references and an editable package closure", async ({ source, producer }) => {
  const bytes = readFileSync(producer.outputFile); expect(createHash("sha256").update(bytes).digest("hex")).toBe(producer.outputSha256);
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "native-drawing-producer" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), object = before.metadata.workbook.objects![0]!;
    expect(before.metadata.workbook.objects).toHaveLength(producer.observed.drawings.length); expect(before.metadata.workbook.charts ?? []).toHaveLength(producer.observed.charts.length);
    if (source.name === "charts") expect(before.metadata.workbook.charts).toHaveLength(producer.name === "sort" ? 3 : 4);
    s.commit(s.prepare(transaction(s, [{ op: "updateWorkbookObject", nodeId: object.nodeId, changes: { name: "原生对象再编辑禰󠄀😀" } }]), { measureWorkbookLayout: metrics }));
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
test.each(evidence.sources)("sole Worker $name transports from private measured source and rejects cancellation, discard and stale commands", async source => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync(source.source), { documentId: "worker-drawing-transport" }), { kind: "workbook" }));
  s.commit(s.prepare(transaction(s, [{ op: "insertWorkbookSheet", name: "Other" }]), { measureWorkbookLayout: metrics })); s.clearHistory();
  const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), object = before.metadata.workbook.objects![0]!, anchor = object.anchor; if (anchor.kind === "absolute") throw Error("Expected independent cell-bound source");
    const targetSheetId = before.metadata.workbook.sheets.at(-1)!.sheetId, command = { kind: "transaction" as const, transaction: transaction(s, [{ op: "moveWorkbookRange", sheetId: object.sheetId,
      range: { from: { row: anchor.from.row, column: anchor.from.column }, to: { row: anchor.from.row, column: anchor.from.column } }, targetSheetId, target: { row: 19, column: 7 } }]) };
    const measureLayout = (request: WorkbookLayoutSourceRequest): WorkbookLayoutMeasurement => ({ sourceDigest: request.sourceDigest, fontEnvironment: "native-drawing-private-point-metrics", sheets: request.sheetIds.map(sheetId => ({ sheetId,
      rows: { defaultSizePt: request.source.cells.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!.defaultRowHeightPt, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
    await expect(client.prepare(command)).rejects.toMatchObject({ code: "InvalidOperation" }); expect(s.snapshot()).toEqual(before);
    const candidate = await client.prepare(command, undefined, { measureLayout }); expect(s.snapshot()).toEqual(before); await client.discard(candidate); expect(s.snapshot()).toEqual(before);
    const abort = new AbortController(); await expect(client.prepare(command, undefined, { signal: abort.signal, measureLayout: request => { abort.abort(); return measureLayout(request); } })).rejects.toMatchObject({ code: "OperationCancelled" });
    expect(s.snapshot()).toEqual(before); await client.commit(await client.prepare(command, undefined, { measureLayout })); expect(s.history().undoSteps).toBe(1);
    const view = await client.view(); expect(view.cells.snapshot.metadata.workbook.objects!.find(value => value.nodeId === object.nodeId)).toMatchObject({ sheetId: targetSheetId, anchor: { from: { row: 19, column: 7 } } }); view.dispose();
    await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    await expect(client.prepare(command, undefined, { measureLayout })).rejects.toMatchObject({ code: "RevisionConflict" });
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
