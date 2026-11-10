import { afterEach, expect, test, vi } from "vitest";
import { createDocumentSession, createWorkbook, DEFAULT_WORKBOOK_RULE_LIMITS, DocumentError, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookSession } from "@doffi/core";
import type { WorkbookLayoutSourceRequest, WorkbookPreviewSource } from "@doffi/render";
import { connectWorkbookSession, serveWorkbookSession } from "../src/index.js";
import type { WorkbookSessionPreparationExecution, WorkbookSessionTransportOptions } from "../src/index.js";

const cleanups: (() => void)[] = [], sheetId = "sheet-1";
afterEach(() => cleanups.splice(0).reverse().forEach(clean => clean()));
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
const resize: DocumentOperation = { op: "updateWorkbookSheet", sheetId, changes: { defaultRowHeightPt: 30 } };
async function pair(options: WorkbookSessionTransportOptions = {}) {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "worker-layout" }), { kind: "workbook" })); cleanups.push(() => s.dispose());
  s.commit(s.prepare(transaction(s, [{ op: "insertWorkbookObject", sheetId, anchor: { kind: "two-cell", from: { row: 2, column: 1, offsetXPt: 0, offsetYPt: 0 }, to: { row: 5, column: 3, offsetXPt: 0, offsetYPt: 0 }, moveWithCells: false, sizeWithCells: false }, object: { kind: "geometry", preset: "rect", html: "<p>禰󠄀😀</p>" } },
    { op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 5 } } }, { row: 0, column: 1, source: { kind: "formula", expression: "=A1+1" } }] }]), { measureWorkbookLayout: metrics })); s.clearHistory();
  const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s, options); cleanups.push(() => server.dispose());
  const client = await connectWorkbookSession(channel.port2, options); cleanups.push(() => client.dispose());
  return { s, client, server, command: () => ({ kind: "transaction" as const, transaction: transaction(s, [resize]) }) };
}
function metrics(request: WorkbookLayoutSourceRequest | WorkbookLayoutRequest): WorkbookLayoutMeasurement {
  const snapshot = 'model' in request ? request.model.snapshot() : request.source.cells.snapshot;
  return { sourceDigest: request.sourceDigest, fontEnvironment: "fixed-test-fonts", sheets: request.sheetIds.map(id => ({ sheetId: id,
    rows: { defaultSizePt: snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === id)!.defaultRowHeightPt, overrides: [] }, columns: { defaultSizePt: 40, overrides: [] } })) };
}
test("layout deltas bind exact successive private providers across candidate and committed revisions", async () => {
  const { s, client } = await pair(); let previous: WorkbookPreviewSource | undefined;
  const seen: { sourceRows: number[]; calculatedRows: number[] }[] = [];
  const measureLayout: NonNullable<WorkbookSessionPreparationExecution["measureLayout"]> = request => {
    const impact = request.source.impact;
    if (impact) seen.push({ sourceRows: impact.sourceRanges.map(value => value.range.from.row), calculatedRows: impact.calculatedRanges.map(value => value.range.from.row) });
    else expect(previous).toBeUndefined();
    previous = request.source; return metrics(request);
  };
  for (const value of [10, 20]) {
    const candidate = await client.prepare({ kind: "transaction", transaction: transaction(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value } } }] }]) }, undefined, { measureLayout });
    await client.commit(candidate);
  }
  expect(seen).toHaveLength(3); expect(seen.filter(value => value.sourceRows.includes(0) && value.calculatedRows.includes(0))).toHaveLength(2);
  await expect(read(previous!)).rejects.toThrow();
});
async function read(source: WorkbookPreviewSource, signal = new AbortController().signal) {
  const m = source.cells.snapshot.manifest;
  return source.cells.read({ documentId: m.documentId, revision: m.revision, maxCells: 2, maxCodeUnits: 8192, windows: [{ sheetId, row: 0, column: 0, rows: 1, columns: 2 }] }, signal);
}

test("private layout ports expose current bounded reads, retain one session and close after the measurement", async () => {
  const { s, client, command } = await pair(), before = s.serialize(), sources: WorkbookPreviewSource[] = [], contexts: string[] = [];
  const measureLayout: NonNullable<WorkbookSessionPreparationExecution["measureLayout"]> = async request => {
    expect("model" in request || "serialize" in request.source).toBe(false); expect(Object.isFrozen(request.sheetIds)).toBe(true);
    sources.push(request.source); contexts.push(request.calculationContext!.locale!);
    expect((await read(request.source, request.signal)).cells[0]!.value).toEqual({ type: "number", value: 5 });
    const m = request.source.cells.snapshot.manifest;
    const current = await request.source.rules!.read({ ...DEFAULT_WORKBOOK_RULE_LIMITS, documentId: m.documentId, revision: m.revision, sheetId, includeCurrent: true, coordinates: [{ row: 0, column: 1 }] }, request.signal!);
    expect(current.cells[0]!.current).toMatchObject({ value: { type: "number", value: 6 } }); return metrics(request);
  };
  const candidate = await client.prepare(command(), { calculationContext: { locale: "zh-CN" } }, { measureLayout });
  expect(s.serialize()).toEqual(before); expect(sources).toHaveLength(2); expect(contexts).toEqual(["zh-CN", "zh-CN"]);
  for (const source of sources) await expect(read(source)).rejects.toThrow();
  const preview = await client.view(candidate); expect(preview.cells.snapshot.metadata.workbook.objects![0]!.anchor).toMatchObject({ from: { row: 2, offsetYPt: -30 }, to: { row: 5, offsetYPt: -75 } }); preview.dispose();
  expect(await client.commit(candidate)).toMatchObject({ changed: true }); expect(s.history().undoSteps).toBe(1); await client.undo();
  expect(s.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
});

test.each(["stale", "throw", "oversize"])("%s layout replies fail atomically and leave the owning connection usable", async kind => {
  const { s, client, command } = await pair(kind === "oversize" ? { maxMessageBytes: 100_000 } : {}), before = s.serialize(); let source: WorkbookPreviewSource | undefined;
  await expect(client.prepare(command(), undefined, { measureLayout: request => {
    source = request.source;
    if (kind === "throw") throw new DocumentError("InvalidOperation", "host refused measurement");
    const result = metrics(request); return kind === "stale" ? { ...result, sourceDigest: "old-source" } : { ...result, fontEnvironment: "x".repeat(100_000) };
  } })).rejects.toMatchObject({ code: kind === "oversize" ? "ResourceLimitExceeded" : "InvalidOperation" });
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0); await expect(read(source!)).rejects.toThrow();
  const next = await client.prepare({ kind: "transaction", transaction: transaction(s, []) }); await client.discard(next); expect(client.state().revision).toBe(s.snapshot().manifest.revision);
});

test("cancelling the parent preparation aborts the main-thread measure and releases its private read endpoint", async () => {
  const { s, client, command } = await pair(), before = s.serialize(), abort = new AbortController(); let started!: () => void, source: WorkbookPreviewSource | undefined;
  const begin = new Promise<void>(resolve => { started = resolve; });
  const job = client.prepare(command(), undefined, { signal: abort.signal, measureLayout: request => {
    source = request.source; started(); return new Promise((_, reject) => request.signal!.addEventListener("abort", () => reject(new DocumentError("OperationCancelled", "cancelled measurement")), { once: true }));
  } });
  await begin; abort.abort(); await expect(job).rejects.toMatchObject({ code: "OperationCancelled" }); expect(s.serialize()).toEqual(before); await expect(read(source!)).rejects.toThrow();
  const next = await client.prepare({ kind: "transaction", transaction: transaction(s, []) }); await client.discard(next);
});

test("an external commit during private measurement rejects the old candidate without losing the host change", async () => {
  const { s, client, command } = await pair(); let started!: () => void, release!: () => void;
  const begin = new Promise<void>(resolve => { started = resolve; }), wait = new Promise<void>(resolve => { release = resolve; });
  const job = client.prepare(command(), undefined, { measureLayout: async request => { started(); await wait; return metrics(request); } });
  await begin; s.commit(s.prepare(transaction(s, [{ op: "updateWorkbookSheet", sheetId, changes: { name: "External" } }]))); const before = s.serialize();
  await vi.waitFor(() => expect(client.state().revision).toBe(s.snapshot().manifest.revision)); release();
  await expect(job).rejects.toMatchObject({ code: "RevisionConflict" }); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(1);
});
