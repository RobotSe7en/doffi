import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookSession } from "@doffi/core";
import { connectWorkbookSession, createWorkbookFillTransaction, createWorkbookRangeCopyTransaction, serveWorkbookSession } from "../src/index.js";
import { workbookFillSeedCount } from "../src/workbook/fill-selection.js";

const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const cleanups: (() => void)[] = [];
afterEach(() => { for (const cleanup of cleanups.splice(0).reverse()) cleanup(); });
function apply(session: WorkbookSession, operations: DocumentOperation[]) {
  const manifest = session.snapshot().manifest;
  return session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations }));
}
function workbook() {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: "copy-merges" }), { kind: "workbook" })); cleanups.push(() => session.dispose());
  apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 255, column: 31, source: { kind: "richText", html: '<span class="cell-default">合并&#13;😀禰󠄀</span>' } },
    { row: 257, column: 31, source: { kind: "formula", expression: "=LEN(AF256)+$A$1" } },
  ] }, { op: "mergeWorkbookCells", sheetId: "sheet-1", range: range(255, 31, 256, 32) },
  { op: "insertWorkbookSheet", name: "Other" }]); session.clearHistory(); return session;
}

test("fill seed bands close across staggered merges and ignore unrelated columns", () => {
  const merges = [range(2, 4, 3, 5), range(1, 2, 2, 3), range(0, 0, 1, 1), range(1, 6, 8, 7)];
  expect(workbookFillSeedCount(range(0, 0, 11, 5), merges, "row")).toBe(4);
  const transpose = (value: ReturnType<typeof range>) => ({ from: { row: value.from.column, column: value.from.row }, to: { row: value.to.column, column: value.to.row } });
  expect(workbookFillSeedCount(transpose(range(0, 0, 11, 5)), merges.map(transpose), "column")).toBe(4);
  expect(() => workbookFillSeedCount(range(0, 0, 2, 5), merges, "row")).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(() => workbookFillSeedCount(range(0, 0, 11, 0), merges, "row")).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
});

test.each([false, true])("complete cross-chunk merged copy keeps rich source and formula rebasing with one undo (cross-sheet %s)", async crossSheet => {
  const session = workbook(), targetSheetId = crossSheet ? session.snapshot().metadata.workbook.sheets[1]!.sheetId : "sheet-1";
  const before = session.serialize(), source = session.getCell("sheet-1", 255, 31)!;
  const transaction = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 257, 32), { row: 511, column: 63 }, { targetSheetId });
  expect(session.serialize()).toEqual(before); apply(session, transaction.operations);
  expect(session.getCell(targetSheetId, 511, 63)).toMatchObject({ value: source.value, text: source.text });
  expect(session.getCell(targetSheetId, 511, 63)!.html).toContain('<span class="cell-default">合并&#13;😀禰󠄀</span>');
  expect(session.getCell(targetSheetId, 513, 63)!.formula!.expression).toBe("=(LEN(BL512)+$A$1)");
  expect(session.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === targetSheetId)!.merges).toContainEqual(range(511, 63, 512, 64));
  expect(openDocument(session.serialize(), { kind: "workbook" }).getCell(targetSheetId, 511, 63)!.text).toBe(source.text);
  expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
  expect(session.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
  session.redo(); expect(session.getCell(targetSheetId, 511, 63)!.text).toBe(source.text);
});

test("copy uses immutable overlapping merged seeds and identical copies remain no-ops", async () => {
  const session = workbook();
  apply(session, [{ op: "mergeWorkbookCells", sheetId: "sheet-1", range: range(257, 31, 258, 32) }]); session.clearHistory();
  const before = session.serialize();
  apply(session, (await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 258, 32), { row: 255, column: 31 })).operations);
  expect(session.history().undoSteps).toBe(0); expect(session.serialize()).toEqual(before);
  const copied = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 258, 32), { row: 257, column: 31 });
  // The existing destination merge is completely covered; captured sources are
  // read before any unmerge or cell write can affect an overlapping seed.
  apply(session, copied.operations);
  expect(session.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([range(255, 31, 256, 32), range(257, 31, 258, 32), range(259, 31, 260, 32)]);
  expect(session.getCell("sheet-1", 257, 31)!.text).toBe("合并\r😀禰󠄀");
  expect(session.getCell("sheet-1", 259, 31)!.formula!.expression).toBe("=(LEN(AF258)+$A$1)");
  session.undo(); expect(session.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
});

test("partial merges, failed validation, cancellation and conflicts cannot publish topology", async () => {
  const session = workbook(), before = session.serialize();
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31), { row: 0, column: 0 })).rejects.toMatchObject({ code: "InvalidOperation" });
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 0), { row: 255, column: 31 })).rejects.toMatchObject({ code: "InvalidOperation" });
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 512, 32), { row: 1000, column: 0 }, { signal: { aborted: true } })).rejects.toMatchObject({ code: "OperationCancelled" });
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 }, { maxCodeUnits: 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  apply(session, [{ op: "mergeWorkbookCells", sheetId: "sheet-1", range: range(0, 0, 1, 0) }]);
  const incompatible = session.serialize();
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 })).rejects.toMatchObject({ code: "InvalidOperation" });
  await expect(createWorkbookFillTransaction(session, "sheet-1", range(255, 31, 256, 32), range(0, 0, 1, 1))).rejects.toMatchObject({ code: "InvalidOperation" });
  expect(session.serialize()).toEqual(incompatible);
  session.undo(); session.clearHistory();
  const generated = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 });
  // Ordinary paste replaces destination validation. This source rule is valid
  // in row 256 and invalid after its template moves to row 1, so the complete
  // cell/merge/rule candidate must still fail without publishing topology.
  apply(session, [{ op: "setWorkbookDataValidation", sheetId: "sheet-1", rule: { ruleId: "stop", kind: "custom", origin: { row: 0, column: 0 }, ranges: [range(0, 0), range(255, 31)], expression: "=ROW()>1", allowBlank: true, showInputMessage: false, showErrorMessage: true, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" } }]);
  expect(() => session.prepare(generated)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  const protectedSource = session.serialize(), blocked = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 });
  expect(() => session.prepare(blocked)).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(session.serialize()).toEqual(protectedSource);
});

test.each((["contents", "values", "formulas", "formats"] as const).flatMap(mode => [false, true].flatMap(skipBlanks => [false, true].map(matching => ({ mode, skipBlanks, matching })))))
  ("merged $mode copy, skipBlanks=$skipBlanks, matching=$matching preserves the selected content/format/topology and one undo", async ({ mode, skipBlanks, matching }) => {
    const s = workbook(), target = range(511, 63, 512, 64);
    apply(s, [
      { op: "setWorkbookNumberFormat", format: { formatId: "source-format", builtinId: 4 } },
      { op: "setWorkbookNumberFormat", format: { formatId: "target-format", builtinId: 2 } },
      { op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [255, 257].map(row => ({ row, column: 31, formatId: "source-format" })) },
      { op: "setWorkbookCells", sheetId: "sheet-1", cells: [511, 513].map(row => ({ row, column: 63, source: { kind: "value", value: { type: "number", value: row === 511 ? 88 : 99 } }, formatId: "target-format" })) },
      ...(matching ? [{ op: "mergeWorkbookCells" as const, sheetId: "sheet-1", range: target }] : []),
    ]); s.clearHistory(); const before = s.snapshot(), source = s.getCell("sheet-1", 255, 31)!, calculation = s.calculateSync();
    const transaction = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(255, 31, 257, 32), target.from, { mode, skipBlanks, calculation });
    expect(s.snapshot()).toEqual(before); apply(s, transaction.operations); s.calculateSync();
    const root = s.getCell("sheet-1", 511, 63)!, third = s.getCell("sheet-1", 513, 63)!;
    expect(root.value).toEqual(mode === "formats" ? { type: "number", value: 88 } : source.value);
    expect(root.formatId).toBe(mode === "contents" || mode === "formats" ? "source-format" : "target-format");
    expect(third.formatId).toBe(mode === "contents" || mode === "formats" ? "source-format" : "target-format");
    if (mode === "contents") expect(root.html).toContain('<span class="cell-default">合并&#13;😀禰󠄀</span>');
    expect(s.getCalculatedCell("sheet-1", 513, 63)).toMatchObject({ state: "valid", value: { type: "number", value: mode === "formats" ? 99 : 8 } });
    expect(third.formula !== undefined).toBe(mode === "contents" || mode === "formulas");
    expect(s.snapshot().metadata.workbook.sheets[0]!.merges.some(merge => JSON.stringify(merge) === JSON.stringify(target)))
      .toBe(matching || mode === "contents" || mode === "formats");
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
  });

test.each(["contents", "formats"] as const)("%s merged paste cannot discard a skipped/non-content destination value", async mode => {
  const s = workbook();
  apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 1, source: { kind: "value", value: { type: "number", value: 98 } } }] }]);
  s.clearHistory(); const before = s.serialize();
  const transaction = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 }, { mode, skipBlanks: true });
  expect(() => s.prepare(transaction)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test.each([false, true])("blank merged seeds with skipBlanks retain destination content and merge identity (matching %s)", async matching => {
  const s = workbook(), target = range(0, 0, 1, 1);
  apply(s, [{ op: "clearWorkbookRanges", sheetId: "sheet-1", ranges: [range(255, 31, 256, 32)], mode: "contents" },
    { op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 88 } } }] },
    ...(matching ? [{ op: "mergeWorkbookCells" as const, sheetId: "sheet-1", range: target }] : [])]);
  s.clearHistory(); const before = s.serialize();
  for (const mode of ["contents", "values", "formulas", "formats"] as const) {
    const transaction = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(255, 31, 256, 32), target.from, { mode, skipBlanks: true });
    expect(apply(s, transaction.operations).changed).toBe(false); expect(s.serialize()).toEqual(before);
  }
  expect(s.history().undoSteps).toBe(0);
});

test("the sole Worker owns merged copy candidates and discards their structure without touching source", async () => {
  const session = workbook(), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, session); cleanups.push(() => server.dispose());
  const client = await connectWorkbookSession(channel.port2); cleanups.push(() => client.dispose());
  const before = session.serialize(), command = { kind: "copy" as const, documentId: client.state().documentId, revision: client.state().revision, sheetId: "sheet-1", source: range(255, 31, 256, 32), target: { row: 0, column: 0 } };
  const candidate = await client.prepare(command), view = await client.view(candidate);
  expect(view.cells.snapshot.metadata.workbook.sheets[0]!.merges).toContainEqual(range(0, 0, 1, 1)); view.dispose();
  expect(session.serialize()).toEqual(before); await client.discard(candidate); expect(session.serialize()).toEqual(before);
  await client.commit(await client.prepare(command)); expect(session.history().undoSteps).toBe(1);
  await client.undo(); expect(session.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
});

test.each([253, 257])("fill repeats merged rich seeds in both directions without reading generated cells (row %s)", async row => {
  const session = workbook(), before = session.serialize(), target = range(row, 31, row + 3, 32);
  const transaction = await createWorkbookFillTransaction(session, "sheet-1", range(255, 31, 256, 32), target);
  expect(session.serialize()).toEqual(before); apply(session, transaction.operations);
  for (const offset of [0, 2]) {
    expect(session.snapshot().metadata.workbook.sheets[0]!.merges).toContainEqual(range(row + offset, 31, row + offset + 1, 32));
    expect(session.getCell("sheet-1", row + offset, 31)!.text).toBe("合并\r😀禰󠄀");
  }
  expect(session.history().undoSteps).toBe(1); session.undo();
  expect(session.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
  expect(session.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
});

test("fill rejects clipped repeated merges and cancels during topology preparation before publishing any cells", async () => {
  const session = workbook(), before = session.serialize();
  await expect(createWorkbookFillTransaction(session, "sheet-1", range(255, 31, 256, 32), range(257, 31, 259, 32))).rejects.toMatchObject({ code: "InvalidOperation" });
  const signal = { aborted: false }; let yields = 0;
  await expect(createWorkbookFillTransaction(session, "sheet-1", range(255, 31, 256, 32), range(257, 31, 1280, 32), { signal, yield: async () => { yields++; signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(yields).toBe(1); expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
});

test("fill preserving target format flattens complete rich merged seeds without manufacturing target merges", async () => {
  const s = workbook();
  apply(s, [{ op: "setWorkbookNumberFormat", format: { formatId: "target", builtinId: 2 } },
    { op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 257, column: 31, source: { kind: "richText", html: '<span class="cell-default">旧文字</span>' }, formatId: "target" }] }]);
  s.clearHistory(); const before = s.snapshot();
  // An odd target length is valid because no merge geometry is copied.
  const transaction = await createWorkbookFillTransaction(s, "sheet-1", range(255, 31, 256, 32), range(257, 31, 259, 32), { format: "preserve" });
  apply(s, transaction.operations);
  expect(s.getCell("sheet-1", 257, 31)).toMatchObject({ text: "合并\r😀禰󠄀", formatId: "target" });
  expect(s.getCell("sheet-1", 259, 31)).toMatchObject({ text: "合并\r😀禰󠄀" });
  expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual(before.metadata.workbook.sheets[0]!.merges);
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
});

test.each((["row", "column"] as const).flatMap(axis => ["source", "preserve"].map(format => ({ axis, format: format as "source" | "preserve" }))))
  ("merged numeric series uses logical $axis seeds with $format formats across both chunk boundaries", async ({ axis, format }) => {
    const s = createDocumentSession(openDocument(createWorkbook({ documentId: "merged-series" }), { kind: "workbook" })); cleanups.push(() => s.dispose());
    const at = (row: number, column: number) => axis === "row" ? { row, column } : { row: column, column: row };
    const area = (row: number, last: number) => ({ from: at(row, 31), to: at(last, 32) });
    apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [255, 257].map((row, index) => ({ ...at(row, 31), source: { kind: "value", value: { type: "number", value: (index + 1) * 10 } } })) },
      ...[255, 257].map(row => ({ op: "mergeWorkbookCells" as const, sheetId: "sheet-1", range: area(row, row + 1) }))]);
    s.clearHistory(); const before = s.snapshot();
    apply(s, (await createWorkbookFillTransaction(s, "sheet-1", area(255, 258), area(259, 264), { mode: "series", axis, format })).operations);
    for (const [index, start] of [259, 261, 263].entries()) {
      const anchor = at(start, 31), covered = at(start + 1, 32);
      expect(s.getCell("sheet-1", anchor.row, anchor.column)).toMatchObject({ value: { type: "number", value: (index + 3) * 10 } });
      expect(s.getCell("sheet-1", covered.row, covered.column)?.value ?? { type: "blank" }).toEqual({ type: "blank" });
      expect(s.snapshot().metadata.workbook.sheets[0]!.merges.some(merge => JSON.stringify(merge) === JSON.stringify(area(start, start + 1)))).toBe(format === "source");
    }
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  });

test("numeric series rejects nonuniform merged seed cells before publishing any part", async () => {
  const s = workbook(), before = s.serialize();
  await expect(createWorkbookFillTransaction(s, "sheet-1", range(255, 31, 258, 32), range(259, 31, 260, 32), { mode: "series", axis: "row" }))
    .rejects.toMatchObject({ code: "InvalidOperation" });
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("large merged-copy topology yields before scanning cells, cancels privately and rejects a changed source", async () => {
  const files = new Map(createWorkbook({ documentId: "merged-copy-topology" })), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.merges = Array.from({ length: 500 }, (_, index) => range(index * 2, 0, index * 2 + 1, 1));
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); cleanups.push(() => s.dispose());
  const before = s.serialize(), signal = { aborted: false }, progress: number[] = [];
  await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 999, 1), { row: 1000, column: 0 }, {
    skipBlanks: true, signal, onProgress(value) { progress.push(value.completed); }, yield: async () => { signal.aborted = true; },
  })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(progress).toEqual([]); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 999, 1), { row: 1000, column: 0 }, {
    skipBlanks: true, yield: async () => { apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 2000, column: 0, source: { kind: "value", value: { type: "number", value: 7 } } }] }]); },
  })).rejects.toMatchObject({ code: "RevisionConflict" });
  expect(s.getCell("sheet-1", 1000, 0)).toBeUndefined(); expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toHaveLength(500);
  await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 999, 1), { row: 1000, column: 0 }, { maxSteps: 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
});

test("a native merged region with a blank anchor and populated covered cell is not mistaken for a wholly blank seed", async () => {
  const input = createDocumentSession(openDocument(createWorkbook({ documentId: "native-covered-source" }), { kind: "workbook" })); cleanups.push(() => input.dispose());
  apply(input, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 1, source: { kind: "value", value: { type: "number", value: 98 } } }] }]);
  const files = new Map(input.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  // Native producers may persist content below a merge. Reading preserves it;
  // constructing a new merge still requires explicit handling of that data.
  metadata.workbook.sheets[0]!.merges = [range(0, 0, 1, 1)]; files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); cleanups.push(() => s.dispose()); const before = s.serialize();
  const identical = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 1, 1), { row: 0, column: 0 }, { skipBlanks: true });
  expect(apply(s, identical.operations).changed).toBe(false); expect(s.serialize()).toEqual(before);
  const sameFill = await createWorkbookFillTransaction(s, "sheet-1", range(0, 0, 1, 1), range(0, 0, 1, 1));
  expect(apply(s, sameFill.operations).changed).toBe(false); expect(s.serialize()).toEqual(before);
  const transaction = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 1, 1), { row: 3, column: 0 }, { skipBlanks: true });
  expect(transaction.operations).toContainEqual({ op: "mergeWorkbookCells", sheetId: "sheet-1", range: range(3, 0, 4, 1), coveredContent: "reject" });
  expect(() => s.prepare(transaction)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("Worker fill publishes repeated merged topology through the same candidate and history", async () => {
  const session = workbook(), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, session); cleanups.push(() => server.dispose());
  const client = await connectWorkbookSession(channel.port2); cleanups.push(() => client.dispose()); const before = session.serialize();
  const candidate = await client.prepare({ kind: "fill", documentId: client.state().documentId, revision: client.state().revision, sheetId: "sheet-1", source: range(255, 31, 256, 32), target: range(257, 31, 260, 32) });
  expect(session.serialize()).toEqual(before); await client.commit(candidate);
  expect(session.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([range(255, 31, 256, 32), range(257, 31, 258, 32), range(259, 31, 260, 32)]);
  await client.undo(); expect(session.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
});
