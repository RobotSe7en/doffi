import { afterEach, expect, test } from "vitest";
import { createDocumentSession, parseContract } from "../src/index.js";
import type { DocumentOperation, WorkbookEditProposal, WorkbookInteractionRequest, WorkbookInteractionRequestInput, WorkbookOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
import { readFileSync, readdirSync } from "node:fs";
import { openDocument } from "../src/index.js";

const sessions: WorkbookSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const number = (value: number) => ({ type: "number" as const, value });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
function apply(session: WorkbookSession, operations: DocumentOperation[]): void { const manifest = session.snapshot().manifest; session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations })); }
function setup(cells = [{ row: 0, column: 0, value: number(4) }, { row: 0, column: 1, value: number(999), formula: "=A1+1" }], selected = range(0, 0)) {
  const session = createDocumentSession(scalarWorkbook(cells), { sessionLimits: { maxCandidates: 1 } }); sessions.push(session); const manifest = session.snapshot().manifest;
  apply(session, [{ op: "addAnnotation", annotation: { annotationId: "request", documentId: manifest.documentId, baseRevision: manifest.revision, status: "open", instruction: "更新源并复核公式😀禰󠄀", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [selected] } } }]);
  session.clearHistory(); return session;
}
function request(session: WorkbookSession, options: Partial<WorkbookInteractionRequestInput> = {}) { return session.buildInteractionRequest({ kind: "workbook", annotationIds: ["request"], ...options }); }
function proposal(request: WorkbookInteractionRequest, operations: WorkbookOperation[]): WorkbookEditProposal {
  return { protocolVersion: 1, documentId: request.documentId, baseRevision: request.baseRevision, requestId: request.requestId, proposalId: "proposal", annotationIds: ["request"], summary: "更新数值并保留区域外公式", operations };
}
const set = (row: number, column: number, value: number): WorkbookOperation => ({ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row, column, source: { kind: "value", value: number(value) } }] });

test("independently authored package, request and proposal examples agree with public source capture and atomic application", () => {
  const root = new URL("../../../examples/", import.meta.url), directory = new URL("workbook/", root);
  const files = new Map(readdirSync(directory).filter(name => /\.(?:json|html|css)$/.test(name)).map(name => [name, new Uint8Array(readFileSync(new URL(name, directory)))]));
  const model = openDocument(files, { kind: "workbook" }), session = createDocumentSession(model); sessions.push(session);
  const example = parseContract("workbook-interaction-request", JSON.parse(readFileSync(new URL("workbook.request.json", root), "utf8")));
  const actual = session.buildInteractionRequest({ kind: "workbook", annotationIds: ["review-1"], scope: example.context.scope, supportedOperations: example.supportedOperations, calculationContext: { locale: "zh-CN" } });
  expect({ ...actual, requestId: example.requestId }).toEqual(example);
  const proposal = parseContract("workbook-edit-proposal", JSON.parse(readFileSync(new URL("workbook.proposal.json", root), "utf8"))), candidate = session.prepareProposal(example, proposal);
  expect(session.getProposalReview(candidate)!.calculatedChanges.find(cell => cell.access === "read-only")!.after).toEqual({ kind: "value", value: number(11) });
  session.commit(candidate); expect(session.getCell("sheet-1", 0, 0)!.value).toEqual(number(10)); expect(session.getAnnotations()[0]!.status).toBe("resolved");
  session.undo(); expect(session.snapshot().manifest.chunks).toEqual(model.snapshot().manifest.chunks); expect(session.getAnnotations()[0]!.status).toBe("open");
});

test("public workbook proposals isolate source, distinguish consumer recalculation and atomically resolve with undo/redo", () => {
  const session = setup(), before = session.serialize(), input = request(session), candidate = session.prepareProposal(input, proposal(input, [set(0, 0, 10)]));
  expect(session.serialize()).toEqual(before); expect(candidate.model.getCell("sheet-1", 0, 0)!.value).toEqual(number(10));
  expect(candidate.model.getCell("sheet-1", 0, 1)!.value).toEqual(number(999));
  expect(session.getProposalReview(candidate)).toMatchObject({ proposalId: "proposal", calculatedChanges: [
    { row: 0, column: 0, access: "editable", before: { value: number(4) }, after: { value: number(10) } },
    { row: 0, column: 1, access: "read-only", before: { value: number(5) }, after: { value: number(11) } },
  ] });
  expect(candidate.changedRanges).toEqual([{ sheetId: "sheet-1", range: range(0, 0) }]); expect(candidate.preview.annotations[0]!.status).toBe("resolved");
  session.commit(candidate); expect(session.history().undoSteps).toBe(1); expect(session.getAnnotations()[0]!.status).toBe("resolved");
  session.undo(); expect(session.getCell("sheet-1", 0, 0)!.value).toEqual(number(4)); expect(session.getAnnotations()[0]!.status).toBe("open");
  session.redo(); expect(session.getCell("sheet-1", 0, 0)!.value).toEqual(number(10)); expect(session.getAnnotations()[0]!.status).toBe("resolved");
});

test("no-op proposals preserve the instruction and do not create a history step", () => {
  const session = setup(), input = request(session), before = session.serialize();
  const candidate = session.prepareProposal(input, proposal(input, [set(0, 0, 4)]));
  expect(candidate.preview.annotations[0]!.status).toBe("open"); expect(session.getProposalReview(candidate)!.calculatedChanges).toEqual([]);
  expect(session.commit(candidate).changed).toBe(false); expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
});

test("tampered frozen coverage and no-op writes to read-only dependencies fail before publication", () => {
  const session = setup(), input = request(session), before = session.serialize();
  for (const mutate of [
    (value: WorkbookInteractionRequest) => { value.context.cells.pop(); },
    (value: WorkbookInteractionRequest) => { value.context.cells[1]!.access = "editable"; },
    (value: WorkbookInteractionRequest) => { value.context.dependencies = []; },
  ]) {
    const changed = structuredClone(input); mutate(changed);
    expect(() => session.prepareProposal(changed, proposal(input, [set(0, 0, 10)]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  }
  expect(() => session.prepareProposal(input, proposal(input, [set(0, 1, 999)]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  expect(session.serialize()).toEqual(before); const candidate = session.prepareProposal(input, proposal(input, [set(0, 0, 10)])); session.discard(candidate);
});

test("a cut that rewrites outside formula source is rejected while row sorting preserves its address and recalculates", () => {
  const session = setup([{ row: 0, column: 0, value: number(4) }, { row: 1, column: 0, value: number(8) }, { row: 0, column: 1, value: number(999), formula: "=A1+1" }], range(0, 0, 1, 0));
  const input = request(session), before = session.serialize();
  expect(() => session.prepareProposal(input, proposal(input, [{ op: "moveWorkbookRange", sheetId: "sheet-1", range: range(0, 0), targetSheetId: "sheet-1", target: { row: 1, column: 0 }, overwrittenContent: "clear" }]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  expect(session.serialize()).toEqual(before);
  const candidate = session.prepareProposal(input, proposal(input, [{ op: "reorderWorkbookRows", sheetId: "sheet-1", range: range(0, 0, 1, 0), rows: [1, 0] }]));
  expect(candidate.model.getCell("sheet-1", 0, 1)!.formula!.expression).toBe("=A1+1");
  expect(session.getProposalReview(candidate)!.calculatedChanges.find(cell => cell.access === "read-only")!.after).toEqual({ kind: "value", value: number(9) }); session.discard(candidate);
});

test("new formula inputs require explicit read expansion and never gain write permission", () => {
  const session = setup([{ row: 0, column: 0, value: number(4) }, { row: 19, column: 0, value: number(20) }]);
  const formula: WorkbookOperation = { op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=A20+1" } }] };
  const input = request(session); expect(() => session.prepareProposal(input, proposal(input, [formula]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  const expanded = request(session, { readRanges: [{ sheetId: "sheet-1", range: range(19, 0) }] });
  expect(expanded.context.cells[1]!.access).toBe("read-only");
  const candidate = session.prepareProposal(expanded, proposal(expanded, [formula])); expect(session.getProposalReview(candidate)!.calculatedChanges[0]!.after).toEqual({ kind: "value", value: number(21) }); session.discard(candidate);
  expect(() => session.prepareProposal(expanded, proposal(expanded, [set(19, 0, 20)]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
});

test("a stale request and a forged proposal identity cannot create a candidate", () => {
  const session = setup(), input = request(session), edit = proposal(input, [set(0, 0, 10)]);
  expect(() => session.prepareProposal(input, { ...edit, requestId: "foreign" })).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  apply(session, [set(0, 0, 7)]); expect(() => session.prepareProposal(input, edit)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  expect(session.getCell("sheet-1", 0, 0)!.value).toEqual(number(7));
});

test("scoped source formatting appends local CSS without changing shared style consumers", () => {
  const session = setup(), input = request(session), candidate = session.prepareProposal(input, proposal(input, [{ op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [range(0, 0)], properties: { color: "#ff0000", "font-weight": "bold" } }]));
  expect(candidate.model.getCell("sheet-1", 0, 0)!.cssClass).toContain("workbook-local-"); expect(candidate.model.getCell("sheet-1", 0, 1)!.html).toBe(session.getCell("sheet-1", 0, 1)!.html); session.discard(candidate);
  expect(() => session.prepareProposal(input, proposal(input, [{ op: "setWorkbookStyleClass", cssClass: "cell-default", properties: { color: "red" } }]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
});

test("new drawings use explicit worksheet creation permission without granting the whole grid", () => {
  const session = setup(), op: WorkbookOperation = { op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "absolute", xPt: 20, yPt: 30, widthPt: 80, heightPt: 40 }, object: { kind: "text", html: "<p>新文本😀禰󠄀</p>" } };
  const input = request(session); expect(() => session.prepareProposal(input, proposal(input, [op]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  const expanded = request(session, { scope: { kind: "ranges", ranges: [{ sheetId: "sheet-1", range: range(0, 0) }], newObjectSheetIds: ["sheet-1"] } });
  const candidate = session.prepareProposal(expanded, proposal(expanded, [op])); expect(candidate.preview.metadata.workbook.objects).toHaveLength(1); expect(expanded.context.writableRanges).toEqual([{ sheetId: "sheet-1", range: range(0, 0) }]); session.discard(candidate);
});

test("asynchronous context capture rejects cancellation and source changes after a real checkpoint", async () => {
  const session = setup(Array.from({ length: 300 }, (_, row) => ({ row, column: 0, value: number(row) })), range(0, 0, 299, 0)), before = session.serialize(), signal = { aborted: false };
  await expect(session.buildInteractionRequestAsync({ kind: "workbook", annotationIds: ["request"] }, { signal, onProgress: progress => { if (progress.phase === "cells") signal.aborted = true; }, yield: async () => {} })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(session.serialize()).toEqual(before);
  let changed = false;
  await expect(session.buildInteractionRequestAsync({ kind: "workbook", annotationIds: ["request"] }, { yield: async () => { if (!changed) { changed = true; apply(session, [set(0, 0, 500)]); } } })).rejects.toMatchObject({ code: "RevisionConflict" });
  expect(session.getCell("sheet-1", 0, 0)!.value).toEqual(number(500));
});

test("asynchronous proposal cancellation discards an already prepared candidate and preserves the single candidate budget", async () => {
  const session = setup(Array.from({ length: 150 }, (_, row) => ({ row, column: 0, value: number(row) })), range(0, 0, 149, 0)), input = request(session), before = session.serialize(), signal = { aborted: false };
  await expect(session.prepareProposalAsync(input, proposal(input, [set(0, 0, 500)]), undefined, { signal, yield: async () => {}, onProgress: progress => { if (progress.phase === "calculation") signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  const candidate = session.prepareProposal(input, proposal(input, [set(0, 0, 600)])); session.discard(candidate);
});

test("public workbook contracts reject future protocol/context versions and foreign content operations", () => {
  const session = setup(), input = request(session);
  expect(() => parseContract("workbook-interaction-request", { ...input, protocolVersion: 2 })).toThrowError(expect.objectContaining({ code: "UnsupportedVersion" }));
  expect(() => parseContract("workbook-interaction-request", { ...input, context: { ...input.context, formatVersion: 2 } })).toThrowError(expect.objectContaining({ code: "UnsupportedVersion" }));
  expect(() => parseContract("workbook-edit-proposal", { ...proposal(input, []), operations: [{ op: "setText", nodeId: "foreign", text: "foreign" }] })).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
});
