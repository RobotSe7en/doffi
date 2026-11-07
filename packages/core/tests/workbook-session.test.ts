import { describe, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { workbookSource } from "../src/workbook/source.js";

const sheetId = "sheet-1";
function session(options: Parameters<typeof createDocumentSession>[1] = {}): WorkbookSession {
  return createDocumentSession(openDocument(createWorkbook({ documentId: "editing" }), { kind: "workbook" }), options);
}
function prepare(s: WorkbookSession, operations: DocumentOperation[]) {
  return s.prepare({ protocolVersion: 1, documentId: "editing", baseRevision: s.snapshot().manifest.revision, operations });
}
const number = (row: number, column: number, value: number): DocumentOperation => ({ op: "setWorkbookCells", sheetId,
  cells: [{ row, column, source: { kind: "value", value: { type: "number", value } } }] });

describe("workbook shared transaction and history boundary", () => {
  test("prepares isolated cross-block candidates and saves/reopens a single monotonic revision", () => {
    const s = session(), before = s.serialize();
    const candidate = prepare(s, [number(0, 0, 12), number(256, 32, 25), { op: "setWorkbookCells", sheetId, cells: [
      { row: 257, column: 33, source: { kind: "formula", expression: "=AG257+1" } },
      { row: 258, column: 33, source: { kind: "value", value: { type: "text", value: "甲😀禰󠄀\r\n乙" } } },
    ] }]);
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
    expect(candidate.model.getCell(sheetId, 256, 32)?.value).toEqual({ type: "number", value: 25 });
    expect(Object.isFrozen(candidate.preview.manifest)).toBe(true); expect("set" in candidate.mapping.nodes).toBe(false);
    expect(candidate.affectedSheetIds).toEqual([sheetId]); expect(candidate.changedRanges).toHaveLength(4);
    expect(s.commit(candidate)).toMatchObject({ changed: true, revision: 1, affectedSheetIds: [sheetId] });
    const reopened = openDocument(s.serialize(), { kind: "workbook" });
    expect(reopened.getCell(sheetId, 257, 33)?.formula).toEqual({ expression: "=AG257+1", cacheState: "stale" });
    expect(reopened.getCell(sheetId, 258, 33)?.value).toEqual({ type: "text", value: "甲😀禰󠄀\r\n乙" });
    expect(reopened.snapshot()).toEqual(s.snapshot());
    expect(s.undo().revision).toBe(2); expect(s.getCell(sheetId, 0, 0)).toBeUndefined();
    expect(s.redo().revision).toBe(3); expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 12 });
    expect(candidate.preview.manifest.revision).toBe(0);
  });
  test("rejects late failures, wrong formats, forged and stale candidates without partial writes", () => {
    const s = session(), original = s.serialize();
    expect(() => prepare(s, [number(0, 0, 1), { op: "setWorkbookCells", sheetId, cells: [{ row: 1, column: 0, source: { kind: "richText", html: '<span onclick="bad()">unsafe</span>' } }] }])).toThrow();
    expect(s.serialize()).toEqual(original);
    expect(() => prepare(s, [number(0, 0, 1), { op: "setWorkbookCells", sheetId: "foreign", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "blank" } } }] }])).toThrow();
    expect(() => prepare(s, [{ op: "replaceText", nodeId: "guess", from: 0, to: 0, text: "X", expectedText: "" }])).toThrow();
    const candidate = prepare(s, [number(0, 0, 1)]), stale = prepare(s, [number(0, 0, 2)]);
    expect(() => session().commit(candidate)).toThrow(); expect(() => s.commit({ ...candidate })).toThrow();
    s.commit(candidate); expect(() => s.commit(stale)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
    expect(() => s.commit(candidate)).toThrow(); s.discard(stale);
    expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 1 });
  });
  test("invalid intermediate rich input cannot disappear through overwrite, clear or worksheet deletion", () => {
    for (const html of ['<span onclick="bad()">unsafe</span>', '<script>unsafe</script>', "长".repeat(32768), "\ud800"]) for (const action of ["overwrite", "clear", "sheet"] as const) {
      const s = session(); s.commit(prepare(s, [{ op: "insertWorkbookSheet", name: "Retained" }])); s.clearHistory(); const before = s.serialize();
      const remove: DocumentOperation = action === "overwrite" ? number(0, 0, 7) : action === "sheet" ? { op: "removeWorkbookSheet", sheetId }
        : { op: "clearWorkbookRanges", sheetId, ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], mode: "all" };
      expect(() => prepare(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "richText", html } }] }, remove])).toThrow();
      expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0); s.dispose();
    }
  });
  test("removed rich input retains its node budget while temporary aggregate block size is checked on the final source", () => {
    const s = session({ limits: { maxNodes: 40 } }), before = s.serialize();
    expect(() => prepare(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "richText", html: "<span>x</span>".repeat(100) } }] },
      { op: "clearWorkbookRanges", sheetId, ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], mode: "all" }])).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
    const candidate = prepare(s, [{ op: "setWorkbookCells", sheetId, cells: Array.from({ length: 20 }, (_, column) => ({ row: 0, column, source: { kind: "value" as const, value: { type: "number" as const, value: column } } })) },
      { op: "clearWorkbookRanges", sheetId, ranges: [{ from: { row: 0, column: 10 }, to: { row: 0, column: 19 } }], mode: "all" }]);
    s.commit(candidate); expect(s.readWindow({ sheetId, row: 0, column: 0, rows: 1, columns: 20 })).toHaveLength(10); expect(s.history().undoSteps).toBe(1);
  });
  test("no-op and reverted batches preserve revision, redo and events", () => {
    const s = session(); s.commit(prepare(s, [number(0, 0, 1)])); s.commit(prepare(s, [number(0, 0, 2)])); s.undo();
    const history = s.history(), revision = s.snapshot().manifest.revision, events: string[] = []; s.subscribe(event => events.push(event.type));
    expect(s.commit(prepare(s, [number(0, 0, 1)]))).toMatchObject({ changed: false, revision });
    expect(s.commit(prepare(s, [number(0, 0, 9), number(0, 0, 1)]))).toMatchObject({ changed: false, revision });
    expect(s.history()).toEqual(history); expect(events).toEqual([]); expect(s.redo().changed).toBe(true);
  });
  test("contents/all clear differ and sparse far ranges never enumerate implicit cells", () => {
    const s = session(); s.commit(prepare(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 1_048_575, column: 16_383, cssClass: "cell-default", formatId: "general", source: { kind: "value", value: { type: "number", value: 10 } } }] }]));
    const ranges = [{ from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16_383 } }];
    s.commit(prepare(s, [{ op: "clearWorkbookRanges", sheetId, ranges, mode: "contents" }]));
    expect(s.getCell(sheetId, 1_048_575, 16_383)).toMatchObject({ value: { type: "blank" }, cssClass: "cell-default", formatId: "general" });
    s.commit(prepare(s, [{ op: "clearWorkbookRanges", sheetId, ranges, mode: "all" }]));
    expect(s.snapshot().manifest.chunks).toEqual([]); expect(openDocument(s.serialize(), { kind: "workbook" }).getCell(sheetId, 1_048_575, 16_383)).toBeUndefined();
    s.undo(); expect(s.getCell(sheetId, 1_048_575, 16_383)?.cssClass).toBe("cell-default");
  });
  test("cell annotations share atomic batches and return through undo with revision rebinding", () => {
    const s = session();
    s.commit(prepare(s, [number(0, 0, 3), { op: "addAnnotation", annotation: { annotationId: "a", documentId: "editing", baseRevision: 0, status: "open", instruction: "检查", anchor: { kind: "cells", sheetId, ranges: [{ from: { row: 0, column: 0 }, to: { row: 2, column: 2 } }] } } }]));
    expect(s.getAnnotations()[0]?.baseRevision).toBe(1); s.undo(); expect(s.getAnnotations()).toEqual([]); s.redo();
    expect(s.getAnnotations()[0]).toMatchObject({ baseRevision: 3, status: "open" });
    expect(() => prepare(s, [number(0, 0, 4), { op: "updateAnnotation", annotationId: "missing", changes: { status: "resolved" } }])).toThrow();
    expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 3 });
  });
  test("candidate and history budgets fail before publication and retain recoverable candidates", () => {
    const s = session({ sessionLimits: { maxCandidates: 1, maxHistorySteps: 1 } });
    const first = prepare(s, [number(0, 0, 1)]); expect(() => prepare(s, [])).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); s.commit(first);
    const next = prepare(s, [number(0, 0, 2)]), before = s.serialize();
    expect(() => s.commit(next)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); expect(s.serialize()).toEqual(before);
    s.clearHistory(); expect(s.commit(next).revision).toBe(2);
    const small = session({ sessionLimits: { maxCandidateBytes: 1 } }); expect(() => prepare(small, [number(0, 0, 1)])).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
  test("shares unchanged chunk buffers and indexes across 100 small edit history entries", () => {
    const seed = session();
    seed.commit(prepare(seed, Array.from({ length: 8 }, (_, index) => number(index * 256, 0, index))));
    const s = createDocumentSession(openDocument(seed.serialize(), { kind: "workbook" })), candidate = prepare(s, [number(0, 0, 100)]);
    const original = workbookSource(candidate.model)!, untouched = candidate.preview.manifest.chunks.find(chunk => chunk.row === 256)!.path;
    s.commit(candidate);
    for (let index = 1; index < 100; index++) {
      const next = prepare(s, [number(0, 0, 100 + index)]), source = workbookSource(next.model)!;
      expect(source.files.get(untouched)).toBe(original.files.get(untouched)); expect(source.chunks.get(untouched)).toBe(original.chunks.get(untouched)); s.commit(next);
    }
    expect(s.history().undoSteps).toBe(100); expect(s.history().retainedBytes).toBeLessThan(300_000);
    s.dispose(); expect(() => s.snapshot()).toThrow(); expect(() => s.getCell(sheetId, 0, 0)).toThrow(); s.dispose();
  });
});
