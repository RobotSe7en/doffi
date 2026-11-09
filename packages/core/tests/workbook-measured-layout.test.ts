import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, WorkbookAxisDimensions } from "../src/index.js";
import type { DeepReadonly, DocumentOperation, WorkbookDrawingAnchor, WorkbookLayoutMeasurement, WorkbookLayoutMeasurer, WorkbookLayoutRequest, WorkbookModel, WorkbookSession } from "../src/index.js";

const sessions: WorkbookSession[] = [], sheetId = "sheet-1";
afterEach(() => sessions.splice(0).forEach(session => session.dispose()));
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function commit(s: WorkbookSession, operations: DocumentOperation[], measureWorkbookLayout?: WorkbookLayoutMeasurer) { return s.commit(s.prepare(transaction(s, operations), measureWorkbookLayout ? { measureWorkbookLayout } : undefined)); }
const anchor = (moveWithCells: boolean, sizeWithCells: boolean): WorkbookDrawingAnchor => ({ kind: "two-cell", from: { row: 2, column: 2, offsetXPt: 1, offsetYPt: 2 }, to: { row: 5, column: 4, offsetXPt: 3, offsetYPt: 4 }, moveWithCells, sizeWithCells });
function blank() {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "measured-layout" }), { kind: "workbook" })); sessions.push(s);
  commit(s, [false, true].flatMap(move => [false, true].map(size => ({ op: "insertWorkbookObject" as const, sheetId, anchor: anchor(move, size), name: `${move}/${size}`, object: { kind: "geometry" as const, preset: "rect" as const, html: "<p>禰󠄀😀</p>" } })))); s.clearHistory(); return s;
}
/** Deliberately supplied document-point metrics: this tests core compensation,
 * not browser font measurement or native Excel geometry acceptance. */
function metrics(request: WorkbookLayoutRequest): WorkbookLayoutMeasurement {
  return { sourceDigest: request.sourceDigest, fontEnvironment: "fixed-test-fonts", sheets: request.sheetIds.map(id => {
    const sheet = request.model.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === id)!;
    return { sheetId: id, rows: { defaultSizePt: sheet.defaultRowsHidden ? 0 : sheet.defaultRowHeightPt, overrides: sheet.rowDimensions.map(run => ({ from: run.from, to: run.to, sizePt: run.hidden ? 0 : run.heightPt ?? sheet.defaultRowHeightPt })) },
      columns: { defaultSizePt: 40, overrides: sheet.columnDimensions.map(run => ({ from: run.from, to: run.to, sizePt: run.hidden ? 0 : run.width ?? 40 })) } };
  }) };
}
function rect(anchor: DeepReadonly<WorkbookDrawingAnchor>, measurement: WorkbookLayoutMeasurement) {
  const sheet = measurement.sheets[0]!, rows = new WorkbookAxisDimensions(1_048_576, sheet.rows.defaultSizePt, sheet.rows.overrides), cols = new WorkbookAxisDimensions(16_384, sheet.columns.defaultSizePt, sheet.columns.overrides);
  if (anchor.kind === "absolute") return { x: anchor.xPt, y: anchor.yPt, width: anchor.widthPt, height: anchor.heightPt };
  const x = cols.position(anchor.from.column) + anchor.from.offsetXPt, y = rows.position(anchor.from.row) + anchor.from.offsetYPt;
  return { x, y, width: anchor.kind === "one-cell" ? anchor.widthPt : cols.position(anchor.to.column) + anchor.to.offsetXPt - x, height: anchor.kind === "one-cell" ? anchor.heightPt : rows.position(anchor.to.row) + anchor.to.offsetYPt - y };
}
function rectangles(model: WorkbookModel) {
  const measurement = metrics({ model, sourceDigest: "test", sheetIds: [sheetId] });
  return Object.fromEntries(model.snapshot().metadata.workbook.objects!.map(object => [object.name!, rect(object.anchor, measurement)]));
}

test("all response modes retain their intended physical bounds after simultaneous measured row/column changes", () => {
  const s = blank(), before = s.serialize();
  const pending = s.prepare(transaction(s, [{ op: "updateWorkbookSheet", sheetId, changes: { rowDimensions: [{ from: 0, to: 0, heightPt: 30 }, { from: 3, to: 3, heightPt: 45 }], columnDimensions: [{ from: 0, to: 0, width: 70, unit: "characters" }, { from: 3, to: 3, width: 80, unit: "characters" }] } }]), { measureWorkbookLayout: metrics });
  expect(s.serialize()).toEqual(before);
  expect(rectangles(pending.model)).toEqual({ "false/false": { x: 81, y: 32, width: 82, height: 47 }, "false/true": { x: 81, y: 32, width: 122, height: 77 },
    "true/false": { x: 111, y: 47, width: 82, height: 47 }, "true/true": { x: 111, y: 47, width: 122, height: 77 } });
  s.commit(pending); expect(s.history().undoSteps).toBe(1); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata.workbook.objects).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata.workbook.objects);
  s.redo(); expect(rectangles(s)["false/false"]).toEqual({ x: 81, y: 32, width: 82, height: 47 });
});

test.each(["row", "column"] as const)("axis insertion before and inside an object composes measured %s movement and resizing in one history step", axis => {
  const s = blank(); commit(s, [{ op: "editWorkbookAxis", sheetId, axis, action: "insert", index: 1, count: 2 }, { op: "editWorkbookAxis", sheetId, axis, action: "insert", index: 5, count: 2 }], metrics);
  const values = rectangles(s), unit = axis === "row" ? 15 : 40, position = axis === "row" ? "y" : "x", extent = axis === "row" ? "height" : "width";
  expect(values["false/false"]![position]).toBe(axis === "row" ? 32 : 81); expect(values["false/false"]![extent]).toBe(axis === "row" ? 47 : 82);
  expect(values["true/false"]![position]).toBe((axis === "row" ? 32 : 81) + 2 * unit); expect(values["true/false"]![extent]).toBe(axis === "row" ? 47 : 82);
  expect(values["false/true"]![position]).toBe(axis === "row" ? 32 : 81); expect(values["false/true"]![extent]).toBe((axis === "row" ? 47 : 82) + 2 * unit);
  expect(values["true/true"]![extent]).toBe((axis === "row" ? 47 : 82) + 2 * unit); expect(s.history().undoSteps).toBe(1);
});

test("projection preserves pending chunks when an intermediate empty chunk is reused later in the batch", () => {
  const s = blank(), measurements: WorkbookLayoutRequest[] = [];
  const set = (text: string): DocumentOperation => ({ op: "setWorkbookCells", sheetId, cells: [{ row: 256, column: 32, source: { kind: "value", value: { type: "text", value: text } } }] });
  commit(s, [set("old")]); s.clearHistory();
  commit(s, [{ op: "clearWorkbookRanges", sheetId, mode: "all", ranges: [{ from: { row: 256, column: 32 }, to: { row: 256, column: 32 } }] }, set("新禰󠄀😀")], request => { measurements.push(request); return metrics(request); });
  expect(s.getCell(sheetId, 256, 32)?.text).toBe("新禰󠄀😀"); expect(measurements).toHaveLength(4);
  expect(measurements[1]!.model.getCell(sheetId, 256, 32)).toBeUndefined(); expect(measurements[3]!.model.getCell(sheetId, 256, 32)?.text).toBe("新禰󠄀😀");
  expect(new Set(measurements.map(value => value.sourceDigest)).size).toBe(3); expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.getCell(sheetId, 256, 32)?.text).toBe("old");
});

test("explicit anchor replacement keeps the requested bounds; inverse axes preserve exact original source and redo", () => {
  const s = blank(), id = s.snapshot().metadata.workbook.objects![0]!.nodeId;
  commit(s, [{ op: "updateWorkbookObject", nodeId: id, changes: { anchor: { kind: "absolute", xPt: 9, yPt: 8, widthPt: 7, heightPt: 6 } } }], metrics);
  expect(rectangles(s)["false/false"]).toEqual({ x: 9, y: 8, width: 7, height: 6 }); s.undo(); const before = s.serialize(), history = s.history();
  const candidate = s.prepare(transaction(s, [{ op: "editWorkbookAxis", sheetId, axis: "column", action: "insert", index: 1, count: 2 }, { op: "editWorkbookAxis", sheetId, axis: "column", action: "delete", index: 1, count: 2 }]), { measureWorkbookLayout: metrics });
  expect(candidate.preview.metadata.workbook.objects).toEqual(s.snapshot().metadata.workbook.objects);
  expect(s.commit(candidate).changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
});

test.each(["stale", "fonts", "duplicate", "invalid-run", "unknown"])("invalid %s measurement rolls back all earlier operations", kind => {
  const s = blank(), before = s.serialize(), history = s.history(); let call = 0;
  const bad: WorkbookLayoutMeasurer = request => {
    const result = metrics(request); if (++call === 1) return result;
    if (kind === "stale") return { ...result, sourceDigest: "stale" };
    if (kind === "fonts") return { ...result, fontEnvironment: "changed-fonts" };
    if (kind === "duplicate") return { ...result, sheets: [...result.sheets, ...result.sheets] };
    if (kind === "invalid-run") return { ...result, sheets: [{ ...result.sheets[0]!, rows: { defaultSizePt: 15, overrides: [{ from: 4, to: 3, sizePt: 20 }] } }] };
    return { ...result, unexpected: true } as WorkbookLayoutMeasurement;
  };
  expect(() => commit(s, [{ op: "updateWorkbookSheet", sheetId, changes: { defaultRowHeightPt: 20 } }], bad)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
});

test("asynchronous measurements preserve exact sync results and reject late cancelled results", async () => {
  const a = blank(), b = blank(), operations: DocumentOperation[] = [{ op: "updateWorkbookSheet", sheetId, changes: { defaultRowHeightPt: 25 } }];
  commit(a, operations, metrics); b.commit(await b.prepareAsync(transaction(b, operations), { measureWorkbookLayout: async request => metrics(request) }));
  expect(rectangles(b)).toEqual(rectangles(a)); const before = b.serialize(); let release!: () => void, started!: () => void;
  const begin = new Promise<void>(resolve => { started = resolve; }), wait = new Promise<void>(resolve => { release = resolve; }), abort = new AbortController();
  const job = b.prepareAsync(transaction(b, operations), { measureWorkbookLayout: async request => { expect(request.signal).toBe(abort.signal); started(); await wait; return metrics(request); } }, { signal: abort.signal });
  await begin; abort.abort(); release(); await expect(job).rejects.toMatchObject({ code: "OperationCancelled" }); expect(b.serialize()).toEqual(before);
  let syncSignal: WorkbookLayoutRequest["signal"];
  expect(() => b.prepare(transaction(b, operations), { measureWorkbookLayout: async request => { syncSignal = request.signal; return metrics(request); } })).toThrowError(/asynchronous preparation/);
  expect(syncSignal?.aborted).toBe(true);
});

test("both classic native note boxes keep their own response flags and exact rich-body identities", () => {
  const original = blank(), box = { anchor: anchor(false, false), visible: true, printable: true, locked: true, lockText: false, defaultSize: false,
    disabled: false, autoFill: false, autoLine: false, autoScale: false, justifyLastLine: false };
  commit(original, [{ op: "addWorkbookComment", anchor: { kind: "cells", sheetId, ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] }, format: "classic", html: "<p>禰󠄀😀</p>", box }]);
  const files = new Map(original.serialize()), m = original.snapshot().manifest, metadata = JSON.parse(new TextDecoder().decode(files.get(m.metadata)!));
  Object.assign(metadata.workbook.comments[0], { boxSource: "vml", commentPrBox: { ...box, anchor: anchor(true, false) }, boxDifferences: ["anchor"] });
  files.set(m.metadata, new TextEncoder().encode(JSON.stringify(metadata)));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); sessions.push(s);
  const before = s.snapshot(), body = s.getNode(before.annotations[0]!.bodyNodeId!)!.html;
  commit(s, [{ op: "updateWorkbookSheet", sheetId, changes: { rowDimensions: [{ from: 0, to: 0, heightPt: 30 }, { from: 3, to: 3, heightPt: 45 }] } }], metrics);
  const comment = s.snapshot().metadata.workbook.comments![0]!; expect(comment.format).toBe("classic");
  if (comment.format !== "classic") throw Error("Missing classic comment.");
  const measured = metrics({ model: s, sourceDigest: "test", sheetIds: [sheetId] });
  expect(rect(comment.box!.anchor, measured)).toEqual({ x: 81, y: 32, width: 82, height: 47 });
  expect(rect(comment.commentPrBox!.anchor, measured)).toEqual({ x: 81, y: 47, width: 82, height: 47 });
  expect(s.getNode(before.annotations[0]!.bodyNodeId!)!.html).toBe(body); expect(s.history().undoSteps).toBe(1);
  s.undo(); expect(s.snapshot().metadata.workbook.comments).toEqual(before.metadata.workbook.comments);
});

test("hidden axes may invert temporary marker positions while fixed physical extents remain valid", () => {
  const s = blank(), id = s.snapshot().metadata.workbook.objects![0]!.nodeId;
  const value = anchor(false, false); if (value.kind !== "two-cell") throw Error("Missing cell anchor.");
  value.from.row = 2; value.from.offsetYPt = 20; value.to.row = 4; value.to.offsetYPt = 0;
  commit(s, [{ op: "updateWorkbookObject", nodeId: id, changes: { anchor: value } }]); s.clearHistory();
  commit(s, [{ op: "updateWorkbookSheet", sheetId, changes: { rowDimensions: [{ from: 2, to: 3, hidden: true }] } }], metrics);
  expect(rectangles(s)["false/false"]).toEqual({ x: 81, y: 50, width: 82, height: 10 }); expect(s.history().undoSteps).toBe(1);
});

test("measurement digests also bind frozen calculation time and locale", () => {
  const s = blank(), digests: string[] = [];
  const measure: WorkbookLayoutMeasurer = request => { digests.push(request.sourceDigest); return metrics(request); };
  for (const locale of ["zh-CN", "en-US"]) {
    const candidate = s.prepare(transaction(s, [{ op: "updateWorkbookSheet", sheetId, changes: { showGridLines: true } }]), { measureWorkbookLayout: measure, calculationContext: { locale } });
    s.discard(candidate);
  }
  expect(digests).toHaveLength(4); expect(digests[0]).not.toBe(digests[2]);
});

test("late measurement cannot overwrite an external commit", async () => {
  const s = blank(); let release!: () => void, started!: () => void;
  const begin = new Promise<void>(resolve => { started = resolve; }), wait = new Promise<void>(resolve => { release = resolve; });
  const job = s.prepareAsync(transaction(s, [{ op: "updateWorkbookSheet", sheetId, changes: { defaultRowHeightPt: 25 } }]), {
    measureWorkbookLayout: async request => { started(); await wait; return metrics(request); },
  });
  await begin; commit(s, [{ op: "updateWorkbookSheet", sheetId, changes: { name: "External" } }]); const before = s.serialize(), history = s.history();
  release(); await expect(job).rejects.toMatchObject({ code: "RevisionConflict" }); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
});

test("measured dimension runs share the document node budget", () => {
  const original = blank(), s = createDocumentSession(openDocument(original.serialize(), { kind: "workbook" }), { limits: { maxNodes: 100 } }); sessions.push(s);
  const before = s.serialize();
  expect(() => commit(s, [{ op: "updateWorkbookSheet", sheetId, changes: { defaultRowHeightPt: 25 } }], request => ({ ...metrics(request), sheets: [{ sheetId,
    rows: { defaultSizePt: 15, overrides: Array.from({ length: 101 }, (_, row) => ({ from: row, to: row, sizePt: 20 })) }, columns: { defaultSizePt: 40, overrides: [] } }] })))
    .toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); expect(s.serialize()).toEqual(before);
});
