import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, WorkbookAxisDimensions } from "../src/index.js";
import type { DocumentOperation, WorkbookDrawingAnchor, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookModel, WorkbookOperation, WorkbookSession } from "../src/index.js";

const sessions: WorkbookSession[] = [], sourceId = "sheet-1";
afterEach(() => sessions.splice(0).forEach(s => s.dispose()));
const point = (row: number, column: number) => ({ row, column });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: point(row, column), to: point(lastRow, lastColumn) });
const marker = (row: number, column: number) => ({ row, column, offsetXPt: 1.25, offsetYPt: 2.5 });
const anchors: WorkbookDrawingAnchor[] = [
  { kind: "absolute", xPt: 81.25, yPt: 32.5, widthPt: 40, heightPt: 20 },
  { kind: "one-cell", from: marker(2, 2), widthPt: 82.5, heightPt: 47.5 },
  ...[false, true].flatMap(moveWithCells => [false, true].map(sizeWithCells => ({ kind: "two-cell" as const, from: marker(2, 2), to: { ...marker(5, 4), offsetXPt: 3.75, offsetYPt: 5 }, moveWithCells, sizeWithCells }))),
];
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function metrics(request: WorkbookLayoutRequest): WorkbookLayoutMeasurement {
  return { sourceDigest: request.sourceDigest, fontEnvironment: "drawing-copy-explicit-point-metrics", sheets: request.sheetIds.map(sheetId => {
    const sheet = request.model.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!;
    return { sheetId, rows: { defaultSizePt: sheet.defaultRowHeightPt, overrides: [] }, columns: { defaultSizePt: sheet.name === "Other" ? 70 : 40, overrides: [] } };
  }) };
}
const apply = (s: WorkbookSession, operations: DocumentOperation[]) => s.commit(s.prepare(transaction(s, operations), { measureWorkbookLayout: metrics }));
function example(anchor: WorkbookDrawingAnchor) {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "object-copy" }), { kind: "workbook" })); sessions.push(s);
  apply(s, [{ op: "insertWorkbookSheet", name: "Other" }, { op: "insertWorkbookObject", sheetId: sourceId, anchor, name: "原图形", object: { kind: "geometry", preset: "rect", html: "<p>禰󠄀😀<strong>预算</strong></p>" } }]);
  const targetId = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
  apply(s, [{ op: "updateWorkbookSheet", sheetId: targetId, changes: { defaultRowHeightPt: 30 } }]); s.clearHistory(); return { s, targetId };
}
function rectangle(model: WorkbookModel, index: number) {
  const object = model.snapshot().metadata.workbook.objects![index]!, anchor = object.anchor;
  const measured = metrics({ model, sourceDigest: "assertion", sheetIds: [object.sheetId] }).sheets[0]!, rows = new WorkbookAxisDimensions(1_048_576, measured.rows.defaultSizePt, []), columns = new WorkbookAxisDimensions(16_384, measured.columns.defaultSizePt, []);
  if (anchor.kind === "absolute") return { x: anchor.xPt, y: anchor.yPt, width: anchor.widthPt, height: anchor.heightPt };
  const x = columns.position(anchor.from.column) + anchor.from.offsetXPt, y = rows.position(anchor.from.row) + anchor.from.offsetYPt;
  return { x, y, width: anchor.kind === "one-cell" ? anchor.widthPt : columns.position(anchor.to.column) + anchor.to.offsetXPt - x, height: anchor.kind === "one-cell" ? anchor.heightPt : rows.position(anchor.to.row) + anchor.to.offsetYPt - y };
}
const operation = (s: WorkbookSession, targetId: string): Extract<WorkbookOperation, { op: "copyWorkbookObjects" }> => ({ op: "copyWorkbookObjects", nodeIds: s.snapshot().metadata.workbook.objects!.filter(object => object.sheetId === sourceId).map(object => object.nodeId), sheetId: sourceId, range: range(1, 1, 6, 5), targetSheetId: targetId, target: point(259, 7) });
test.each(anchors)("copying $kind/$moveWithCells/$sizeWithCells preserves measured extents and rich source with fresh identities", anchor => {
  const { s, targetId } = example(anchor), before = s.snapshot(), bounds = rectangle(s, 0), original = before.metadata.workbook.objects![0]!, originalNode = s.getNode(original.nodeId)!;
  const candidate = s.prepare(transaction(s, [operation(s, targetId)]), { measureWorkbookLayout: metrics }); expect(s.snapshot()).toEqual(before); s.commit(candidate);
  const current = s.snapshot().metadata.workbook.objects!, copy = current[1]!; expect(current).toHaveLength(2); expect(current[0]).toEqual(original); expect(copy.nodeId).not.toBe(original.nodeId); expect(copy.sheetId).toBe(targetId);
  const after = rectangle(s, 1); expect(after.width).toBeCloseTo(bounds.width, 8); expect(after.height).toBeCloseTo(bounds.height, 8);
  expect(after.x).toBeCloseTo(bounds.x + (anchor.kind === "absolute" ? 450 : 480), 8); expect(after.y).toBeCloseTo(bounds.y + (anchor.kind === "absolute" ? 7755 : 7770), 8);
  expect(s.getNode(copy.nodeId)!.text).toBe(originalNode.text); expect(s.getNode(copy.nodeId)!.childNodeIds).not.toEqual(originalNode.childNodeIds); expect(s.getNode(original.nodeId)!.html).toBe(originalNode.html);
  expect(copy.anchor).toMatchObject({ kind: anchor.kind, ...(anchor.kind === "two-cell" ? { moveWithCells: anchor.moveWithCells, sizeWithCells: anchor.sizeWithCells } : {}) });
  expect(s.history().undoSteps).toBe(1); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
});
test("absolute boxes require whole physical coverage, including a union of adjacent copied masks", () => {
  const { s, targetId } = example({ kind: "absolute", xPt: 50, yPt: 20, widthPt: 40, heightPt: 20 });
  const op = { ...operation(s, targetId), range: range(0, 0, 3, 3), target: point(19, 7), ranges: [range(20, 8, 21, 8)] };
  expect(apply(s, [op]).changed).toBe(false); expect(s.history().undoSteps).toBe(0);
  apply(s, [{ ...op, ranges: [range(20, 8, 21, 8), range(20, 9, 21, 9)] }]); expect(s.snapshot().metadata.workbook.objects).toHaveLength(2);
});
test("same-position and empty-mask copies are no-ops, and invalid masks, missing measurements and removed targets roll back", () => {
  const { s, targetId } = example(anchors[1]!), before = s.serialize();
  for (const op of [{ ...operation(s, sourceId), target: point(1, 1) }, { ...operation(s, targetId), ranges: [] }]) expect(s.commit(s.prepare(transaction(s, [op]))).changed).toBe(false);
  expect(() => s.prepare(transaction(s, [operation(s, targetId)]))).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(() => apply(s, [{ ...operation(s, targetId), ranges: [range(0, 0)] }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(() => apply(s, [{ op: "removeWorkbookSheet", sheetId: targetId }, operation(s, targetId)])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before);
});
test("source text budgets reject hidden clone expansion before publication and do not charge excluded absolute boxes", () => {
  const { s, targetId } = example(anchors[1]!), before = s.serialize();
  expect(() => apply(s, [{ ...operation(s, targetId), maxSourceCodeUnits: 0 }])).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); expect(s.serialize()).toEqual(before);
  const { s: a, targetId: other } = example(anchors[0]!); expect(apply(a, [{ ...operation(a, other), range: range(30, 0, 31, 1), maxSourceCodeUnits: 0 }]).changed).toBe(false);
});
test("one candidate bounds cumulative clone work even when its final operation would remove every copy", () => {
  const { s } = example(anchors[1]!), id = s.snapshot().metadata.workbook.objects![0]!.nodeId;
  apply(s, [{ op: "setWorkbookObjectText", nodeId: id, html: "<p>" + "x".repeat(150) + "</p>" }]);
  const files = s.serialize(), encode = new TextEncoder(), cost = encode.encode(s.getNode(id)!.html).byteLength + encode.encode(JSON.stringify([s.snapshot().metadata.workbook.objects![0], undefined])).byteLength;
  const maxFileBytes = Math.max(cost * 10 - 1, ...[...files.values()].map(bytes => bytes.byteLength + 1));
  const limited = createDocumentSession(openDocument(files, { kind: "workbook" }), { limits: { maxFileBytes } }); sessions.push(limited); const before = limited.serialize();
  const copies: DocumentOperation[] = Array.from({ length: Math.ceil(maxFileBytes / cost) + 2 }, (_, row) => ({ op: "copyWorkbookObjects", nodeIds: [id], sheetId: sourceId, range: range(2, 2), targetSheetId: sourceId, target: point(100 + row, 5) }));
  expect(() => apply(limited, [...copies, { op: "removeWorkbookSheet", sheetId: sourceId }])).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded", message: expect.stringContaining("cumulative") }));
  expect(limited.serialize()).toEqual(before); expect(limited.history().undoSteps).toBe(0);
});
test.each(["chart-only", "partial", "full"])("chart copy %s redirects only complete direct vectors and preserves names and literals", kind => {
  const { s, targetId } = example(anchors[1]!);
  apply(s, [{ op: "setWorkbookDefinedName", name: { nameId: "data", name: "DataValues", origin: point(0, 0), expression: "=Sheet1!$B$2:$B$4" } }, { op: "insertWorkbookChart", sheetId: sourceId,
    anchor: { kind: "one-cell", from: marker(6, 0), widthPt: 300, heightPt: 190 }, kind: "line", data: { categories: { expression: "=A2:A4", valueType: "text" }, series: [
      { values: { expression: "=B2:B4", valueType: "number" } }, { values: { expression: "=DataValues", valueType: "number" } }, { values: { kind: "literal", valueType: "number", count: 3, points: [{ index: 0, value: -3 }, { index: 2, value: 9 }] } },
    ] } }]); s.clearHistory(); const before = s.snapshot().metadata.workbook.charts![0]!, first = kind === "chart-only" ? 6 : kind === "partial" ? 2 : 0;
  apply(s, [{ op: "copyWorkbookObjects", nodeIds: s.snapshot().metadata.workbook.objects!.map(object => object.nodeId), sheetId: sourceId, range: range(first, 0, 18, 4), targetSheetId: targetId, target: point(59, 7) }]);
  const [old, copy] = s.snapshot().metadata.workbook.charts!; expect(old).toEqual(before); expect(copy!.chartId).not.toBe(before.chartId); expect(copy!.series[0]!.seriesId).not.toBe(before.series[0]!.seriesId);
  expect(copy!.series[0]!.values).toEqual(kind === "full" ? { expression: "=I61:I63", valueType: "number" } : { expression: "=B2:B4", valueType: "number", originSheetId: sourceId });
  expect(copy!.series[1]!.values).toEqual({ expression: "=DataValues", valueType: "number", originSheetId: sourceId }); expect(copy!.series[2]!.values).toEqual(before.series[2]!.values);
});
test("range proposals require destination creation and actual copied source disclosure", () => {
  const { s, targetId } = example(anchors[1]!), m = s.snapshot().manifest;
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "request", documentId: m.documentId, baseRevision: m.revision, status: "open", instruction: "复制图形", anchor: { kind: "cells", sheetId: sourceId, ranges: [range(1, 1, 6, 5)] } } },
    { op: "addAnnotation", annotation: { annotationId: "target-request", documentId: m.documentId, baseRevision: m.revision, status: "open", instruction: "从只读来源复制", anchor: { kind: "cells", sheetId: targetId, ranges: [range(259, 7)] } } }]);
  const before = s.serialize(), original = s.snapshot().metadata.workbook.objects![0]!, op = operation(s, targetId);
  const request = (objectIds: string[], create: boolean) => s.buildInteractionRequest({ kind: "workbook", annotationIds: ["request"], scope: { kind: "ranges", ranges: [{ sheetId: sourceId, range: op.range }, { sheetId: targetId, range: range(259, 7, 264, 11) }], objectIds, newObjectSheetIds: create ? [targetId] : [] } });
  const proposal = (input: ReturnType<typeof request>) => ({ protocolVersion: 1 as const, documentId: input.documentId, baseRevision: input.baseRevision, requestId: input.requestId, proposalId: "copy", annotationIds: ["request"], summary: "复制图形", operations: [op] });
  for (const input of [request([], true), request([original.nodeId], false)]) expect(() => s.prepareProposal(input, proposal(input), { measureWorkbookLayout: metrics })).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  const hidden = request([], true); let measured = false;
  expect(() => s.prepareProposal(hidden, { ...proposal(hidden), operations: [{ ...op, maxSourceCodeUnits: 0 }] }, { measureWorkbookLayout: value => { measured = true; return metrics(value); } })).toThrowError(expect.objectContaining({ code: "InvalidProposal" })); expect(measured).toBe(false);
  expect(s.serialize()).toEqual(before); const input = request([original.nodeId], true), candidate = s.prepareProposal(input, proposal(input), { measureWorkbookLayout: metrics });
  expect(candidate.preview.metadata.workbook.objects).toHaveLength(2); expect(s.serialize()).toEqual(before); s.discard(candidate);
  const readOnly = s.buildInteractionRequest({ kind: "workbook", annotationIds: ["target-request"], scope: { kind: "ranges", ranges: [{ sheetId: targetId, range: range(259, 7, 264, 11) }], objectIds: [original.nodeId], newObjectSheetIds: [targetId] }, readRanges: [{ sheetId: sourceId, range: op.range }] });
  const copied = s.prepareProposal(readOnly, { ...proposal(readOnly), annotationIds: ["target-request"] }, { measureWorkbookLayout: metrics }); expect(copied.preview.metadata.workbook.objects).toHaveLength(2); s.discard(copied); expect(s.serialize()).toEqual(before);
});
test("cancelling after cloned source has been generated discards identities and history", async () => {
  const { s, targetId } = example(anchors[1]!); apply(s, Array.from({ length: 64 }, (_, row) => ({ op: "insertWorkbookObject", sheetId: sourceId, anchor: { kind: "one-cell", from: marker(row, 0), widthPt: 20, heightPt: 10 }, object: { kind: "text", html: "<p>禰󠄀😀</p>" } }))); s.clearHistory();
  const before = s.serialize(), signal = { aborted: false }; let reached = false;
  await expect(s.prepareAsync(transaction(s, [{ ...operation(s, targetId), range: range(0, 0, 64, 3), target: point(100, 5) }]), { measureWorkbookLayout: metrics }, { signal, yield: async () => {}, onProgress(progress) { if (progress.phase === "cells" && progress.completed >= 64) { reached = true; signal.aborted = true; } } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(reached).toBe(true); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
