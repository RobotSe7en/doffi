import { describe, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "../src/index.js";
import type { DocumentOperation, DocumentSessionOptions, DocumentTransaction, TransactionPreparationProgress, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sheetId = "sheet-1";
const number = (row: number, column: number, value: number): DocumentOperation => ({ op: "setWorkbookCells", sheetId,
  cells: [{ row, column, source: { kind: "value", value: { type: "number", value } } }] });
function session(options: DocumentSessionOptions = {}): WorkbookSession {
  return createDocumentSession(openDocument(createWorkbook({ documentId: "async-editing" }), { kind: "workbook" }), options);
}
function transaction(s: WorkbookSession, operations: DocumentOperation[]): DocumentTransaction {
  return { protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations };
}
const batch = (): DocumentOperation => ({ op: "setWorkbookCells", sheetId, cells: Array.from({ length: 600 }, (_, row) => ({ row, column: 0,
  source: row === 599 ? { kind: "formula", expression: "=A599+1" } : { kind: "value", value: { type: "number", value: row } } })) });
const cancelled = expect.objectContaining({ code: "OperationCancelled" });

describe("asynchronous preparation on the shared workbook boundary", () => {
  test("sync and async prepare the same source, diff and single undo entry", async () => {
    const a = session(), b = session(), source = a.serialize(), progress: TransactionPreparationProgress[] = [];
    const operations: DocumentOperation[] = [batch(), number(257, 32, 11), { op: "addAnnotation", annotation: {
      annotationId: "review", documentId: "async-editing", baseRevision: 0, status: "open", instruction: "检查", anchor: { kind: "cells", sheetId,
        ranges: [{ from: { row: 256, column: 0 }, to: { row: 257, column: 1 } }] } } }];
    const sync = a.prepare(transaction(a, operations));
    let yields = 0;
    const async = await b.prepareAsync(transaction(b, operations), undefined, { yield: async () => { yields++; }, onProgress: value => { progress.push(value); } });
    expect(progress.every(Object.isFrozen)).toBe(true);
    expect(yields).toBeGreaterThan(10); expect(progress.some(value => value.phase === "cells" && value.completed === 256)).toBe(true);
    expect(progress.some(value => value.phase === "validation" && value.total === 4)).toBe(true);
    expect(async.model.serialize()).toEqual(sync.model.serialize()); expect(async.changedRanges).toEqual(sync.changedRanges);
    expect([...async.mapping.nodes]).toEqual([...sync.mapping.nodes]); expect(async.mapping.textEdits).toEqual(sync.mapping.textEdits); expect(async.mapping.workbook).toEqual(sync.mapping.workbook);
    expect(b.serialize()).toEqual(source); expect(b.history().undoSteps).toBe(0);
    const committed = b.commit(async), expected = a.commit(sync);
    expect({ ...committed, mapping: undefined }).toEqual({ ...expected, mapping: undefined }); expect(b.history().undoSteps).toBe(1);
    expect(b.calculateSync()).toMatchObject({ calculated: 1 }); expect(b.getCalculatedCell(sheetId, 599, 0)).toMatchObject({ state: "valid", value: { type: "number", value: 599 } });
    b.undo(); expect(b.getCell(sheetId, 256, 0)).toBeUndefined(); expect(b.getAnnotations()).toEqual([]); b.redo(); expect(b.getAnnotations()).toHaveLength(1);
    a.dispose(); b.dispose();
  });

  test.each(["cells", "chunks", "last-validation"] as const)("cancellation at %s releases private work and the reserved candidate slot", async phase => {
    const s = session({ sessionLimits: { maxCandidates: 1 } }), before = s.serialize(), abort = new AbortController(), events: string[] = [];
    s.subscribe(event => events.push(event.type)); let reached = false;
    await expect(s.prepareAsync(transaction(s, [batch()]), undefined, { signal: abort.signal, onProgress: progress => {
      if (phase === "cells" && progress.phase === "cells" && progress.completed === 256
        || phase === "chunks" && progress.phase === "chunks" && progress.completed === 1
        || phase === "last-validation" && progress.phase === "validation" && progress.completed === 1 && progress.total === 1) { reached = true; abort.abort(); }
    } })).rejects.toThrowError(cancelled);
    expect(reached).toBe(true); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0); expect(events).toEqual(["error"]);
    s.commit(await s.prepareAsync(transaction(s, [number(0, 0, 7)]))); expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 7 }); s.dispose();
  });

  test("sparse clear yields on actual cells and cancellation restores every original byte", async () => {
    const s = session(); s.commit(s.prepare(transaction(s, [batch()]))); s.clearHistory(); const before = s.serialize(), abort = new AbortController(); let reached = false;
    await expect(s.prepareAsync(transaction(s, [{ op: "clearWorkbookRanges", sheetId, mode: "all", ranges: [{ from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16_383 } }] }]), undefined,
      { signal: abort.signal, onProgress: progress => { if (progress.phase === "cells" && progress.completed === 256) { expect(progress.total).toBe(600); reached = true; abort.abort(); } } })).rejects.toThrowError(cancelled);
    expect(reached).toBe(true); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0); s.dispose();
  });

  test("a source commit during a yield wins, and the stale private candidate never publishes", async () => {
    const s = session(); let changed = false;
    await expect(s.prepareAsync(transaction(s, [batch()]), undefined, { yield: async () => {
      if (!changed) { changed = true; s.commit(s.prepare(transaction(s, [number(900, 3, 42)]))); }
    } })).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
    expect(s.snapshot().manifest.revision).toBe(1); expect(s.history().undoSteps).toBe(1); expect(s.getCell(sheetId, 0, 0)).toBeUndefined();
    expect(s.getCell(sheetId, 900, 3)?.value).toEqual({ type: "number", value: 42 }); s.dispose();
  });

  test("pending work reserves quota, rejects overlapping async work and releases it on cancellation", async () => {
    const s = session({ sessionLimits: { maxCandidates: 1 } }), abort = new AbortController();
    let resume!: () => void, started!: () => void;
    const ready = new Promise<void>(resolve => { started = resolve; }), suspended = new Promise<void>(resolve => { resume = resolve; });
    const first = s.prepareAsync(transaction(s, [batch()]), undefined, { signal: abort.signal, yield: async () => { started(); await suspended; } });
    await ready;
    expect(() => s.prepare(transaction(s, [number(1, 1, 1)]))).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    await expect(s.prepareAsync(transaction(s, [number(2, 2, 2)]))).rejects.toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    abort.abort(); resume(); await expect(first).rejects.toThrowError(cancelled);
    s.commit(s.prepare(transaction(s, [number(3, 3, 3)]))); expect(s.history().undoSteps).toBe(1); s.dispose();
  });

  test("captures caller transactions before callbacks and ignores progress callback errors", async () => {
    const s = session(), input = transaction(s, [number(0, 0, 12)]); let changed = false;
    const candidate = await s.prepareAsync(input, undefined, { onProgress: () => {
      if (!changed) { changed = true; input.documentId = "foreign"; input.operations.splice(0, 1, number(0, 0, 999)); }
      throw new Error("host reporter");
    } });
    s.commit(candidate); expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 12 }); s.dispose();
  });

  test("disposed work cancels without late publication or events", async () => {
    const s = session(), events: string[] = []; s.subscribe(event => events.push(event.type)); let reached = false;
    await expect(s.prepareAsync(transaction(s, [batch()]), undefined, { yield: async () => { reached = true; s.dispose(); } })).rejects.toThrowError(cancelled);
    expect(reached).toBe(true); expect(events).toEqual([]); expect(() => s.snapshot()).toThrow(); s.dispose();
  });

  test("invalid intermediate rich input remains invalid when later operations overwrite or clear it", async () => {
    const s = session(), before = s.serialize();
    for (const remove of [number(0, 0, 7), { op: "clearWorkbookRanges", sheetId, ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], mode: "all" } satisfies DocumentOperation]) {
      await expect(s.prepareAsync(transaction(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "richText", html: '<span onclick="bad()">unsafe</span>' } }] }, remove]), undefined,
        { yield: async () => {} })).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
      expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
    }
    s.dispose();
  });

  test.each([false, true])("captures validation acknowledgments before yields (accepted=%s)", async accepted => {
    const s = createDocumentSession(scalarWorkbook([], metadata => { metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "warning", kind: "whole", operator: "lessThan", formulas: ["=5"],
      origin: { row: 0, column: 0 }, ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], allowBlank: true, errorStyle: "warning", showErrorMessage: true, showInputMessage: true,
      errorTitle: "", error: "", promptTitle: "", prompt: "" }]; }));
    const acceptedRuleIds = accepted ? ["warning"] : [];
    const work = s.prepareAsync(transaction(s, [number(0, 0, 9)]), { validation: { acceptedRuleIds } }, { yield: async () => { acceptedRuleIds.splice(0, acceptedRuleIds.length, ...(accepted ? [] : ["warning"])); } });
    if (accepted) { s.commit(await work); expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 9 }); }
    else { await expect(work).rejects.toThrowError(expect.objectContaining({ details: [expect.objectContaining({ code: "validation-confirmation-required" })] })); expect(s.getCell(sheetId, 0, 0)).toBeUndefined(); }
    s.dispose();
  });

  test("captures host time and cancels during validation calculation without publishing its results", async () => {
    const make = () => createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 999 }, formula: "=TODAY()" }], metadata => {
      metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "date", kind: "whole", operator: "equal", formulas: ["=25569"], origin: { row: 0, column: 0 },
        ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], allowBlank: false, errorStyle: "stop", showErrorMessage: true, showInputMessage: true,
        errorTitle: "", error: "", promptTitle: "", prompt: "" }];
    }));
    const s = make(), now = { unixMilliseconds: 0, utcOffsetMinutes: 0 };
    // A formula rewrite changes its source identity and triggers target validation.
    const input = transaction(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=TODAY()+0" } }] }]);
    s.commit(await s.prepareAsync(input, { calculationContext: { now } }, { yield: async () => { now.unixMilliseconds = 86_400_000; } }));
    expect(s.history().undoSteps).toBe(1); expect(s.getCell(sheetId, 0, 0)?.formula?.expression).toBe("=TODAY()+0"); s.dispose();
    const c = make(), before = c.serialize(), abort = new AbortController(); let reached = false;
    await expect(c.prepareAsync({ ...input, baseRevision: 0 }, { calculationContext: { now: { unixMilliseconds: 0, utcOffsetMinutes: 0 } } }, { signal: abort.signal,
      onProgress: value => { if (value.phase === "calculation") { reached = true; abort.abort(); } } })).rejects.toThrowError(cancelled);
    expect(reached).toBe(true); expect(c.serialize()).toEqual(before); expect(c.history().undoSteps).toBe(0); c.dispose();
  });

  test("cancels after validation calculation has actually indexed formulas", async () => {
    const s = createDocumentSession(scalarWorkbook(Array.from({ length: 1200 }, (_, row) => ({ row, column: 0, value: { type: "number" as const, value: -1 }, formula: row ? `=A${row}+1` : "=1" })), metadata => {
      metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "last", kind: "whole", operator: "greaterThan", formulas: ["=0"], origin: { row: 1199, column: 0 },
        ranges: [{ from: { row: 1199, column: 0 }, to: { row: 1199, column: 0 } }], allowBlank: true, errorStyle: "stop", showErrorMessage: true, showInputMessage: true,
        errorTitle: "", error: "", promptTitle: "", prompt: "" }];
    })), before = s.serialize(), abort = new AbortController(); let completed = 0;
    await expect(s.prepareAsync(transaction(s, [number(0, 0, 2)]), undefined, { signal: abort.signal, onProgress: value => {
      if (value.phase === "calculation" && value.completed > 0) { completed = value.completed; abort.abort(); }
    }, yield: async () => {} })).rejects.toThrowError(cancelled);
    expect(completed).toBeGreaterThanOrEqual(256); expect(s.serialize()).toEqual(before); expect(s.getCalculatedCell(sheetId, 1199, 0).state).toBe("stale");
    const candidate = await s.prepareAsync(transaction(s, [number(0, 0, 2)])); s.commit(candidate);
    expect(s.history().undoSteps).toBe(1); s.calculateSync(); expect(s.getCalculatedCell(sheetId, 1199, 0)).toMatchObject({ state: "valid", value: { type: "number", value: 1201 } }); s.dispose();
  });
});
