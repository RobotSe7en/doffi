import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument } from "../src/index.js";
import type { Annotation, DocumentOperation, WorkbookDocumentMetadata, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sheetId = "sheet-1", sessions: WorkbookSession[] = [];
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const number = (value: number) => ({ type: "number" as const, value });
function session(model = scalarWorkbook([])) { const s = createDocumentSession(model); sessions.push(s); return s; }
afterEach(() => { for (const s of sessions.splice(0)) s.dispose(); });
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations))); }
function other(metadata: WorkbookDocumentMetadata) { metadata.workbook.sheets.push({ ...metadata.workbook.sheets[0]!, sheetId: "other", name: "Other", rowDimensions: [], columnDimensions: [], merges: [] }); }
const move = (source = range(0, 0), row = 3, column = 0, targetSheetId = sheetId, clear = false): DocumentOperation => ({ op: "moveWorkbookRange", sheetId, range: source, targetSheetId, target: { row, column }, overwrittenContent: clear ? "clear" : "reject" });

test("overlapping cut captures old cells, preserves rich source and moves hyperlink references exactly once", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(3) }, { row: 1, column: 0, value: number(7) },
    { row: 5, column: 0, value: number(999), formula: "=SUM(A1:A2)" }], metadata => {
    metadata.workbook.sheets[0]!.hyperlinks = [{ range: range(0, 0), location: "A2" }];
  }));
  apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 1, source: { kind: "richText", html: '<span class="cell-default">中文&#13;禰󠄀😀</span>' } }] }]); s.clearHistory();
  const before = s.snapshot(), rich = s.getCell(sheetId, 0, 1)!.html;
  apply(s, [move(range(0, 0, 1, 1), 1)]);
  expect(s.getCell(sheetId, 0, 0)).toBeUndefined(); expect(s.getCell(sheetId, 1, 0)?.value).toEqual(number(3)); expect(s.getCell(sheetId, 2, 0)?.value).toEqual(number(7));
  expect(s.getCell(sheetId, 1, 1)?.html).toBe(rich); expect(s.snapshot().metadata.workbook.sheets[0]!.hyperlinks).toEqual([{ range: range(1, 0), location: "A3" }]);
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 5, 0)).toMatchObject({ value: number(10) }); expect(s.history().undoSteps).toBe(1);
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("cross-sheet cuts retain unmoved formula inputs on their original sheet and rewrite partial outside ranges into exact fragments", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(7) }, { row: 0, column: 1, value: number(999), formula: "=A1+Z1" },
    { row: 0, column: 2, value: number(5) }, { row: 0, column: 25, value: number(3) }, { row: 2, column: 0, value: number(999), formula: "=SUM(A1:C1)" }], metadata => {
    other(metadata); metadata.workbook.names = [{ nameId: "all", name: "AllValues", expression: "=Sheet1!$A$1:$C$1", origin: { row: 0, column: 0 } }];
    metadata.workbook.sheets[0]!.hyperlinks = [{ range: range(0, 0, 0, 2), location: "Z1" }];
  })); const before = s.snapshot();
  const candidate = s.prepare(transaction(s, [move(range(0, 0, 0, 1), 3, 0, "other")])); expect(s.snapshot()).toEqual(before);
  s.commit(candidate); s.calculateSync();
  expect(s.getCell(sheetId, 0, 0)).toBeUndefined(); expect(s.getCell("other", 3, 0)?.value).toEqual(number(7));
  expect(s.getCell("other", 3, 1)).toMatchObject({ value: number(999), formula: { expression: "=(A4+'Sheet1'!Z1)", cacheState: "stale" } });
  expect(s.getCalculatedCell("other", 3, 1)).toMatchObject({ value: number(10) }); expect(s.getCalculatedCell(sheetId, 2, 0)).toMatchObject({ value: number(22) });
  const areas = s.resolveReference("=AllValues", { sheetId, row: 0, column: 0 });
  expect(areas).toEqual({ kind: "ranges", areas: [{ sheetId, ...range(0, 2) }, { sheetId: "other", ...range(3, 0, 3, 1) }] });
  expect(s.snapshot().metadata.workbook.sheets[0]!.hyperlinks).toEqual([{ range: range(0, 2), location: "Z1" }]);
  expect(s.snapshot().metadata.workbook.sheets[1]!.hyperlinks).toEqual([{ range: range(3, 0, 3, 1), location: "'Sheet1'!Z1" }]);
  const reopened = openDocument(s.serialize(), { kind: "workbook" }); expect(reopened.snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); s.redo(); expect(s.snapshot().metadata).toEqual(reopened.snapshot().metadata);
});

test("overwrites require an explicit choice, whole merges move and a partial merge leaves the entire batch unchanged", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(7) }, { row: 3, column: 0, value: { type: "text", value: "" } }], metadata => {
    metadata.workbook.sheets[0]!.merges = [range(0, 0, 0, 1)];
  })); const before = s.serialize();
  for (const op of [move(range(0, 0)), move(range(0, 0, 0, 1))]) {
    expect(() => apply(s, [op])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  }
  apply(s, [move(range(0, 0, 0, 1), 3, 0, sheetId, true)]);
  expect(s.getCell(sheetId, 3, 0)?.value).toEqual(number(7)); expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([range(3, 0, 3, 1)]);
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.serialize().size).toBe(before.size);
});

test("cutting partial rules splits exact coverage across sheets and rebases current formulas without redirecting unmoved inputs", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(3) }, { row: 1, column: 0, value: number(7) }, { row: 2, column: 0, value: number(5) },
    { row: 0, column: 25, value: number(1) }], metadata => {
    other(metadata); const sheet = metadata.workbook.sheets[0]!;
    sheet.conditionalRules = [{ ruleId: "paint", ranges: [range(0, 0, 2, 0)], origin: { row: 0, column: 0 }, kind: "expression", priority: 1, stopIfTrue: true, expression: "=A1>$Z$1", cssClass: "cell-default" }];
    sheet.dataValidations = [{ ruleId: "valid", ranges: [range(0, 0, 2, 0)], origin: { row: 0, column: 0 }, kind: "custom", expression: "=A1>$Z$1",
      allowBlank: true, showErrorMessage: true, showInputMessage: false, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" }];
    metadata.workbook.sheets[1]!.conditionalRules = [{ ruleId: "overwrite", ranges: [range(3, 2, 6, 2)], origin: { row: 3, column: 2 }, kind: "expression", priority: 2, stopIfTrue: false, expression: "=TRUE", cssClass: "cell-default" }];
  })); const before = s.snapshot();
  apply(s, [move(range(1, 0), 4, 2, "other")]);
  const [source, target] = s.snapshot().metadata.workbook.sheets;
  expect(source!.conditionalRules).toMatchObject([{ ruleId: "paint", ranges: [range(0, 0), range(2, 0)], expression: "=A1>$Z$1" }]);
  expect(source!.dataValidations).toMatchObject([{ ruleId: "valid", ranges: [range(0, 0), range(2, 0)] }]);
  expect(target!.conditionalRules).toMatchObject([{ ruleId: "overwrite", ranges: [range(3, 2), range(5, 2, 6, 2)] },
    { ranges: [range(4, 2)], origin: { row: 4, column: 2 }, expression: "=(C5>'Sheet1'!$Z$1)" }]);
  expect(target!.dataValidations).toMatchObject([{ ranges: [range(4, 2)], expression: "=(C5>'Sheet1'!$Z$1)" }]);
  expect(new Set([source!.conditionalRules![0]!.ruleId, target!.conditionalRules![1]!.ruleId]).size).toBe(2);
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  const revision = s.snapshot().manifest.revision;
  expect(() => apply(s, [{ op: "setWorkbookCells", sheetId: "other", cells: [{ row: 4, column: 2, source: { kind: "value", value: number(0) } }] }]))
    .toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.snapshot().manifest.revision).toBe(revision); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
});

test("annotations move within one sheet, invalidate unrepresentable cross-sheet fragments, and keep unrelated native replies at their real sheet", () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(7) }], other), files = new Map(model.serialize());
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  const annotations: Annotation[] = [
    { annotationId: "moved", documentId: "scalar-kernel", baseRevision: 0, status: "open", instruction: "移动正文", anchor: { kind: "cells", sheetId, ranges: [range(0, 0)] }, replies: [{ replyId: "moved-reply", instruction: "回复" }] },
    { annotationId: "other", documentId: "scalar-kernel", baseRevision: 0, status: "resolved", instruction: "保留另一表", anchor: { kind: "cells", sheetId: "other", ranges: [range(0, 0)] }, replies: [{ replyId: "other-reply", instruction: "不移动" }] },
    { annotationId: "split", documentId: "scalar-kernel", baseRevision: 0, status: "resolved", instruction: "保留跨表片段正文😀", anchor: { kind: "cells", sheetId, ranges: [range(0, 0, 0, 1)] } },
  ];
  metadata.workbook.comments = [{ format: "modern", annotationId: "moved", replies: [{ replyId: "moved-reply", nativeCell: { row: 0, column: 0 } }] },
    { format: "modern", annotationId: "other", replies: [{ replyId: "other-reply", nativeCell: { row: 0, column: 0 } }] }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); files.set("annotations.json", new TextEncoder().encode(JSON.stringify(annotations)));
  const s = session(openDocument(files, { kind: "workbook" })), before = s.snapshot();
  apply(s, [move(range(0, 0), 3, 0, "other")]);
  expect(s.getAnnotations().find(value => value.annotationId === "moved")).toMatchObject({ status: "open", anchor: { kind: "cells", sheetId: "other", ranges: [range(3, 0)] } });
  expect(s.getAnnotations().find(value => value.annotationId === "split")).toMatchObject({ status: "needs_reanchor", instruction: "保留跨表片段正文😀", anchor: annotations[2]!.anchor });
  const comments = s.snapshot().metadata.workbook.comments!;
  expect(comments[0]).toMatchObject({ replies: [{ nativeCell: { row: 3, column: 0 } }] }); expect(comments[1]).toMatchObject({ replies: [{ nativeCell: { row: 0, column: 0 } }] });
  s.undo(); expect(s.snapshot().annotations).toEqual(before.annotations.map(value => ({ ...value, baseRevision: 2 }))); expect(s.snapshot().metadata).toEqual(before.metadata);
});

test("cancelled cuts and stale candidates cannot publish and a zero-distance move preserves redo", async () => {
  const s = session(scalarWorkbook(Array.from({ length: 600 }, (_, row) => ({ row, column: 0, value: number(row) }))));
  apply(s, [move(range(0, 0), 0, 1)]); s.undo(); const before = s.serialize(), history = s.history();
  expect(apply(s, [move(range(0, 0), 0)]).changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  const abort = new AbortController();
  await expect(s.prepareAsync(transaction(s, [move(range(0, 0, 599, 0), 0, 1)]), undefined,
    { signal: abort.signal, yield: async () => {}, onProgress(value) { if (value.phase === "cells" && value.completed === 256) abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  const candidate = s.prepare(transaction(s, [move(range(0, 0), 0, 1)])); apply(s, [{ op: "setWorkbookSettings", changes: { calculationMode: "manual" } }]);
  expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); s.discard(candidate);
});
