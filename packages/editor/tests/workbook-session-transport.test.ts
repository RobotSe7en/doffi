import { afterEach, expect, test, vi } from "vitest";
import { createDocumentSession, createWorkbook, DEFAULT_WORKBOOK_RULE_LIMITS, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { connectWorkbookSession, serveWorkbookSession } from "../src/index.js";
import type { WorkbookSessionConnection, WorkbookSessionProgress, WorkbookSessionTransportOptions } from "../src/index.js";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { workbookPreviewChangeImpact } from "../../render/dist/workbook/preview-source.js";

const cleanups: (() => void)[] = [], sheetId = "sheet-1";
afterEach(() => { for (const cleanup of cleanups.splice(0).reverse()) cleanup(); });
function workbook(): WorkbookSession { const s = createDocumentSession(openDocument(createWorkbook({ documentId: "worker-session" }), { kind: "workbook" })); cleanups.push(() => s.dispose()); return s; }
const number = (row: number, column: number, value: number): DocumentOperation => ({ op: "setWorkbookCells", sheetId,
  cells: [{ row, column, source: { kind: "value", value: { type: "number", value } } }] });
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { return { protocolVersion: 1 as const, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations }; }
async function pair(s = workbook(), options: WorkbookSessionTransportOptions = {}) {
  const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s, options); cleanups.push(() => server.dispose());
  const client = await connectWorkbookSession(channel.port2, options); cleanups.push(() => client.dispose()); return { s, client, server, channel };
}
function command(s: WorkbookSession, operations: DocumentOperation[]) { return { kind: "transaction" as const, transaction: transaction(s, operations) }; }
async function read(client: WorkbookSessionConnection, row: number, column: number) {
  const source = await client.view(), manifest = source.cells.snapshot.manifest;
  try { return await source.cells.read({ documentId: manifest.documentId, revision: manifest.revision, maxCells: 1, maxCodeUnits: 4096,
    windows: [{ sheetId, row, column, rows: 1, columns: 1 }] }, new AbortController().signal); }
  finally { source.dispose(); }
}

test("a dedicated endpoint prepares privately, returns bounded views and shares core history", async () => {
  const { s, client } = await pair(), before = s.serialize(), states: number[] = []; client.subscribe(state => states.push(state.revision));
  expect(Object.isFrozen(client.state().history)).toBe(true); expect("serialize" in client || "model" in client).toBe(false);
  const candidate = await client.prepare(command(s, [number(256, 32, 12)]));
  expect("model" in candidate || "preview" in candidate).toBe(false); expect(Object.isFrozen(candidate.mapping.nodes)).toBe(true);
  expect(s.serialize()).toEqual(before); expect(client.state()).toMatchObject({ revision: 0, history: { undoSteps: 0 } });
  const preview = await client.view(candidate);
  const query = { documentId: "worker-session", revision: 0, maxCells: 1, maxCodeUnits: 4096, windows: [{ sheetId, row: 256, column: 32, rows: 1, columns: 1 }] };
  expect((await preview.cells.read(query, new AbortController().signal)).cells[0]?.value).toEqual({ type: "number", value: 12 });
  expect(await client.commit(candidate)).toMatchObject({ changed: true, revision: 1, affectedSheetIds: [sheetId] });
  expect(s.history().undoSteps).toBe(1); expect(client.state().revision).toBe(1);
  await expect(preview.cells.read(query, new AbortController().signal)).rejects.toThrow(); preview.dispose();
  expect((await read(client, 256, 32)).cells[0]?.value).toEqual({ type: "number", value: 12 });
  expect(await client.undo()).toMatchObject({ revision: 2, changed: true }); expect(s.getCell(sheetId, 256, 32)).toBeUndefined();
  expect(await client.redo()).toMatchObject({ revision: 3, changed: true }); expect(states).toEqual([1, 2, 3]);
});

test("remote no-op commits preserve redo and emit no core error", async () => {
  const { s, client } = await pair(); s.commit(s.prepare(transaction(s, [number(0, 0, 1)]))); s.undo();
  await vi.waitFor(() => expect(client.state().revision).toBe(2)); const before = s.history(), events: string[] = []; s.subscribe(event => events.push(event.type));
  const candidate = await client.prepare(command(s, [])); expect(await client.commit(candidate)).toMatchObject({ changed: false, revision: 2 });
  expect(s.history()).toEqual(before); expect(events).toEqual([]); expect((await client.redo()).changed).toBe(true);
});

test("Worker input preserves literal types, no-op revisions and failed source drafts", async () => {
  const { s, client } = await pair(), at = { row: 256, column: 32 };
  const input = (revision: number, text: string, mode?: "auto" | "text" | "number") => ({ kind: "input" as const, documentId: "worker-session", revision, sheetId, at, text, ...(mode ? { mode } : {}) });
  await client.commit(await client.prepare(input(0, "12", "text"))); expect(s.getCell(sheetId, at.row, at.column)?.value).toEqual({ type: "text", value: "12" });
  expect(await client.commit(await client.prepare(input(1, "'12")))).toMatchObject({ changed: false, revision: 1 });
  await client.commit(await client.prepare(input(1, "12", "number"))); expect(s.getCell(sheetId, at.row, at.column)?.value).toEqual({ type: "number", value: 12 });
  const before = s.serialize(); await expect(client.prepare(input(1, "old"))).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  await expect(client.prepare(input(2, "="))).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(client.state().history.undoSteps).toBe(2);
});

test("Worker text input retains rich spans, CRLF and variation sequences against its authoritative source", async () => {
  const s = workbook(), at = { row: 0, column: 0 };
  s.commit(s.prepare(transaction(s, [{ op: "setWorkbookCells", sheetId, cells: [{ ...at, source: { kind: "richText", html: '<span class="cell-default">ABC&#13;&#10;</span><span class="cell-default">中文禰󠄀</span>' } }] }])));
  const original = s.getCell(sheetId, 0, 0)!, { client } = await pair(s), before = s.serialize();
  const candidate = await client.prepare({ kind: "input", documentId: "worker-session", revision: 1, sheetId, at, text: original.text + "尾" });
  expect(s.serialize()).toEqual(before); await client.commit(candidate);
  expect(s.getCell(sheetId, 0, 0)?.html).toContain('<span class="cell-default">ABC&#13;\n</span><span class="cell-default">中文禰󠄀尾</span>');
  await client.undo(); expect(s.getCell(sheetId, 0, 0)?.html).toBe(original.html);
});

test("Worker TSV paste crosses both chunk boundaries as one typed history step", async () => {
  const { s, client } = await pair(), before = s.serialize();
  const candidate = await client.prepare({ kind: "paste", documentId: "worker-session", revision: 0, sheetId, at: { row: 255, column: 31 }, text: '"甲\t乙"\tTRUE\n=1+1\t00123' });
  expect(s.serialize()).toEqual(before); await client.commit(candidate);
  expect(s.getCell(sheetId, 255, 31)?.value).toEqual({ type: "text", value: "甲\t乙" });
  expect(s.getCell(sheetId, 255, 32)?.value).toEqual({ type: "boolean", value: true });
  expect(s.getCell(sheetId, 256, 31)?.formula?.expression).toBe("=1+1"); expect(s.getCell(sheetId, 256, 32)?.value).toEqual({ type: "text", value: "00123" });
  expect(s.history().undoSteps).toBe(1); await client.undo();
  expect(s.snapshot().manifest.chunks).toEqual([]); expect([...s.serialize()].filter(([path]) => path.startsWith("sheets/"))).toEqual([]);
  await client.redo(); expect(s.getCell(sheetId, 256, 31)?.formula?.expression).toBe("=1+1");
});

test.each(["fill", "cells"] as const)("cancelling at %s releases work without a source or history change", async phase => {
  const { s, client } = await pair(); s.commit(s.prepare(transaction(s, [number(0, 0, 1)]))); s.clearHistory();
  await vi.waitFor(() => expect(client.state().revision).toBe(1)); const before = s.serialize(), abort = new AbortController(), progress: WorkbookSessionProgress[] = [];
  await expect(client.prepare({ kind: "fill", documentId: "worker-session", revision: 1, sheetId, source: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } },
    target: { from: { row: 1, column: 0 }, to: { row: 600, column: 0 } } }, undefined, { signal: abort.signal, onProgress: value => {
      progress.push(value); if (value.phase === phase && value.completed === 256) abort.abort();
    } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  expect(progress.some(value => value.phase === phase && value.completed === 256)).toBe(true); expect(progress.every(Object.isFrozen)).toBe(true);
  // Wait for the remote cancellation message and cleanup, rather than treating
  // immediate local Promise rejection as proof that the Worker has finished.
  await vi.waitFor(async () => { const next = await client.prepare(command(s, [])); await client.discard(next); });
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("copy values uses current long-chain calculation, and source views retain original caches", async () => {
  const s = createDocumentSession(scalarWorkbook(Array.from({ length: 1200 }, (_, row) => ({ row, column: 0, value: { type: "number", value: 999 }, formula: row ? `=A${row}+1` : "=1" }))));
  cleanups.push(() => s.dispose()); const { client } = await pair(s), view = await client.view();
  expect(view.rules?.currentValues).toBe(true);
  const request = { ...DEFAULT_WORKBOOK_RULE_LIMITS, documentId: "scalar-kernel", revision: 0, sheetId, includeCurrent: true, coordinates: [{ row: 1199, column: 0 }] };
  expect((await view.rules!.read(request, new AbortController().signal)).cells[0]?.current).toMatchObject({ value: { type: "number", value: 1200 } });
  expect(s.getCell(sheetId, 1199, 0)?.value).toEqual({ type: "number", value: 999 }); view.dispose();
  const candidate = await client.prepare({ kind: "copy", documentId: "scalar-kernel", revision: 0, sheetId, source: { from: { row: 1199, column: 0 }, to: { row: 1199, column: 0 } }, target: { row: 0, column: 1 }, mode: "values" });
  await client.commit(candidate); expect(s.getCell(sheetId, 0, 1)).toMatchObject({ value: { type: "number", value: 1200 } }); expect(s.getCell(sheetId, 0, 1)?.formula).toBeUndefined();
  await client.undo(); expect(s.getCell(sheetId, 0, 1)).toBeUndefined();
});

test("find and replacement stay revision-bound in the Worker", async () => {
  const { s, client } = await pair(); s.commit(s.prepare(transaction(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 256, column: 0, source: { kind: "value", value: { type: "text", value: "甲猫😀禰󠄀猫" } } }] }])));
  await vi.waitFor(() => expect(client.state().revision).toBe(1));
  const found = await client.find({ text: "猫", field: "source" }); expect(found.matches.map(match => [match.from, match.to])).toEqual([[1, 2], [7, 8]]);
  const candidate = await client.prepare({ kind: "replace", documentId: "worker-session", revision: 1, query: { text: "猫" }, replacement: "$&" });
  await client.commit(candidate); expect(s.getCell(sheetId, 256, 0)?.text).toBe("甲$&😀禰󠄀$&"); expect(s.history().undoSteps).toBe(2);
  await expect(client.prepare({ kind: "copy", documentId: "worker-session", revision: 1, sheetId, source: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }, target: { row: 0, column: 1 } })).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
});

test("foreign and invalidated candidates cannot be applied through another endpoint", async () => {
  const s = workbook(), a = await pair(s), b = await pair(s), candidate = await a.client.prepare(command(s, [number(0, 0, 1)]));
  await expect(b.client.commit(candidate)).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  await expect(a.client.commit({ ...candidate })).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  await b.client.commit(await b.client.prepare(command(s, [number(0, 0, 2)])));
  await vi.waitFor(() => expect(a.client.state().revision).toBe(1));
  await expect(a.client.commit(candidate)).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" })); await a.client.discard(candidate);
  expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 2 });
});

test("connection disposal releases candidates and views while its host session remains usable", async () => {
  const { s, client, server } = await pair(), candidate = await client.prepare(command(s, [number(0, 0, 1)])), view = await client.view(candidate);
  client.dispose(); await vi.waitFor(() => expect(server.disposed).toBe(true));
  expect(s.snapshot().manifest.revision).toBe(0); expect(s.history().undoSteps).toBe(0); await expect(view.cells.read({ documentId: "worker-session", revision: 0, maxCells: 1, maxCodeUnits: 1024,
    windows: [{ sheetId, row: 0, column: 0, rows: 1, columns: 1 }] }, new AbortController().signal)).rejects.toThrow();
  s.commit(s.prepare(transaction(s, [number(0, 0, 7)]))); expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 7 });
});

test("retained Worker views link exact source and calculation deltas across source revisions", async () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 1 } },
    { row: 1000, column: 1, value: { type: "number", value: 999 }, formula: "=A1+1" }])); cleanups.push(() => s.dispose());
  const { client } = await pair(s), previous = await client.view(), original = await client.view();
  expect(previous.impact).toBeUndefined();
  await client.commit(await client.prepare(command(s, [number(0, 0, 2)])));
  const next = await client.view(undefined, { previous });
  expect(next.impact).toMatchObject({ fromRevision: 0, revision: 1,
    sourceRanges: [{ sheetId, range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } }],
    calculatedRanges: [{ sheetId, range: { from: { row: 1000, column: 1 }, to: { row: 1000, column: 1 } } }] });
  expect(Object.isFrozen(next.impact?.calculatedRanges[0]?.range)).toBe(true);
  expect(workbookPreviewChangeImpact(previous, next)).toBe(next.impact);
  expect(workbookPreviewChangeImpact(original, next)).toBeUndefined();
  previous.dispose(); original.dispose(); next.dispose();
  await expect(client.view(undefined, { previous })).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
});

test("Worker view deltas distinguish calculation contexts at the same source revision", async () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 999 }, formula: "=TODAY()" }])); cleanups.push(() => s.dispose());
  const a = await pair(s), b = await pair(s), previous = await a.client.view(undefined, { calculationContext: { now: { unixMilliseconds: 0, utcOffsetMinutes: 0 } } });
  const other = await a.client.view(undefined, { calculationContext: { now: { unixMilliseconds: 86400000, utcOffsetMinutes: 0 } } });
  const next = await a.client.view(undefined, { previous, calculationContext: { now: { unixMilliseconds: 86400000, utcOffsetMinutes: 0 } } });
  expect(next.impact).toMatchObject({ fromRevision: 0, revision: 0, sourceRanges: [], calculatedRanges: [{ sheetId, range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } }] });
  expect(workbookPreviewChangeImpact(previous, next)).toBe(next.impact); expect(workbookPreviewChangeImpact(other, next)).toBeUndefined();
  await expect(b.client.view(undefined, { previous })).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  await expect(a.client.view(undefined, { previous: { ...previous } })).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  previous.dispose(); other.dispose(); next.dispose();
});

test("views enforce retained limits and reclaim cancelled or explicitly released ports", async () => {
  const { client } = await pair(workbook(), { maxViews: 1 }); const a = await client.view();
  await expect(client.view()).rejects.toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); a.dispose();
  const b = await client.view(); b.dispose();
  const abort = new AbortController(); await expect(client.view(undefined, { signal: abort.signal, onProgress: () => abort.abort() })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  const c = await client.view(); c.dispose();
});

test("message budgets and unsupported values fail before sending source work", async () => {
  const { s, client } = await pair(workbook(), { maxMessageBytes: 4096 }), before = s.serialize();
  await expect(client.find({ text: "x".repeat(4096) })).rejects.toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  await expect(client.prepare({ ...command(s, []), unknown: true } as never)).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  const cyclic: unknown[] = []; cyclic.push(cyclic);
  await expect(client.prepare({ kind: "transaction", transaction: cyclic } as never)).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  await expect(client.prepare({ kind: "transaction", transaction: new Array(100_000_000) } as never)).rejects.toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("a cancelled prepared response releases its core candidate even before delivery", async () => {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "worker-session" }), { kind: "workbook" }), { sessionLimits: { maxCandidates: 1 } }); cleanups.push(() => s.dispose());
  const { client, channel } = await pair(s), send = channel.port1.postMessage.bind(channel.port1);
  let held: { message: unknown; transfer: Transferable[] } | undefined;
  Object.defineProperty(channel.port1, "postMessage", { configurable: true, value(message: { value?: { handle?: unknown } }, transfer: Transferable[] = []) {
    if (message.value?.handle) { held = { message, transfer }; return; } send(message, transfer);
  } });
  const abort = new AbortController(), work = client.prepare(command(s, [number(0, 0, 1)]), undefined, { signal: abort.signal });
  await vi.waitFor(() => expect(held).toBeDefined()); abort.abort(); await expect(work).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  await vi.waitFor(() => { const next = s.prepare(transaction(s, [])); s.discard(next); });
  send(held!.message, held!.transfer); expect(s.history().undoSteps).toBe(0); expect(s.getCell(sheetId, 0, 0)).toBeUndefined();
});

test("a cancelled transferred view response does not exhaust the server view budget", async () => {
  const { client, channel } = await pair(workbook(), { maxViews: 1 }), send = channel.port1.postMessage.bind(channel.port1);
  let held: { message: unknown; transfer: Transferable[] } | undefined;
  Object.defineProperty(channel.port1, "postMessage", { configurable: true, value(message: { port?: MessagePort }, transfer: Transferable[] = []) {
    if (message.port && !held) { held = { message, transfer }; return; } send(message, transfer);
  } });
  const abort = new AbortController(), work = client.view(undefined, { signal: abort.signal });
  await vi.waitFor(() => expect(held).toBeDefined()); abort.abort(); await expect(work).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  // Port order ensures cancel is handled before this next view request.
  const next = await client.view(); next.dispose(); send(held!.message, held!.transfer);
});

test("a lost commit response reports an unknown outcome and retains the committed host source", async () => {
  const { s, client, channel } = await pair(workbook(), { requestTimeoutMs: 1000 }), send = channel.port1.postMessage.bind(channel.port1);
  const candidate = await client.prepare(command(s, [number(0, 0, 5)]));
  Object.defineProperty(channel.port1, "postMessage", { configurable: true, value(message: { value?: { receipt?: unknown } }, transfer: Transferable[] = []) {
    if (message.value?.receipt) return; send(message, transfer);
  } });
  await expect(client.commit(candidate)).rejects.toThrowError(/outcome is unknown/);
  expect(s.snapshot().manifest.revision).toBe(1); expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 5 }); expect(s.history().undoSteps).toBe(1);
  expect(() => client.state()).toThrow();
});

test("views retain manual mode until explicit calculation and use only supplied host time", async () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 999 }, formula: "=TODAY()" }], metadata => { metadata.workbook.calculationMode = "manual"; }));
  cleanups.push(() => s.dispose()); const { client } = await pair(s), before = s.serialize(), manual = await client.view();
  expect(manual.rules?.currentValues).toBeUndefined(); manual.dispose();
  const computed = await client.view(undefined, { calculate: true, calculationContext: { now: { unixMilliseconds: 0, utcOffsetMinutes: 0 } } });
  expect(computed.rules?.currentValues).toBe(true);
  const result = await computed.rules!.read({ ...DEFAULT_WORKBOOK_RULE_LIMITS, documentId: "scalar-kernel", revision: 0, sheetId, coordinates: [{ row: 0, column: 0 }], includeCurrent: true }, new AbortController().signal);
  expect(result.cells[0]?.current).toMatchObject({ value: { type: "number", value: 25569 } }); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0); computed.dispose();
  expect((await client.find({ text: "25569", field: "display" }, { calculationContext: { now: { unixMilliseconds: 0, utcOffsetMinutes: 0 } } })).matches).toMatchObject([{ row: 0, column: 0 }]);
});

test("read transport options are validated before a session connection starts", async () => {
  const s = workbook(), ports = new MessageChannel(); cleanups.push(() => { ports.port1.close(); ports.port2.close(); });
  expect(() => serveWorkbookSession(ports.port1, s, { reads: { requestTimeoutMs: -1 } })).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  await expect(connectWorkbookSession(ports.port2, { reads: { readLimits: { unknown: 1 } } } as never)).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.history().undoSteps).toBe(0);
});

test("an oversized mutation receipt explicitly reports the already committed revision", async () => {
  const s = workbook(); s.commit(s.prepare(transaction(s, [{ op: "setWorkbookCells", sheetId, cells: Array.from({ length: 200 }, (_, row) => ({ row: row * 2, column: 0,
    source: { kind: "value", value: { type: "number", value: row } } })) }])));
  const { client } = await pair(s, { maxMessageBytes: 4096 });
  await expect(client.undo()).rejects.toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded", message: expect.stringContaining("committed at revision 2"),
    details: [expect.objectContaining({ code: "committed-result-too-large" })] }));
  expect(client.state().revision).toBe(2); expect(s.snapshot().manifest.revision).toBe(2); expect(s.getCell(sheetId, 0, 0)).toBeUndefined(); expect(s.history().redoSteps).toBe(1);
});

test("a host edit during a yielding calculation rejects the previous view", async () => {
  const s = createDocumentSession(scalarWorkbook(Array.from({ length: 1200 }, (_, row) => ({ row, column: 0, value: { type: "number", value: 999 }, formula: row ? `=A${row}+1` : "=1" }))));
  cleanups.push(() => s.dispose()); const { client } = await pair(s); let changed = false;
  await expect(client.view(undefined, { onProgress: value => { if (!changed && value.phase === "calculation") { changed = true; s.commit(s.prepare(transaction(s, [number(0, 0, 10)]))); } } })).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  expect(changed).toBe(true); expect(s.history().undoSteps).toBe(1);
  const source = await client.view(); const value = await source.rules!.read({ ...DEFAULT_WORKBOOK_RULE_LIMITS, documentId: "scalar-kernel", revision: 1, sheetId,
    includeCurrent: true, coordinates: [{ row: 1199, column: 0 }] }, new AbortController().signal);
  expect(value.cells[0]?.current).toMatchObject({ value: { type: "number", value: 1209 } }); source.dispose();
});
