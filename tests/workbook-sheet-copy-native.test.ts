import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createHash } from "node:crypto";
import { createDocumentSession, openDocument, parseWorkbookFormula } from "@doffi/core";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { connectWorkbookSession, serveWorkbookSession } from "@doffi/editor";

function apply(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations })); }
async function session(path: string) { return createDocumentSession(openDocument(await importXlsx(readFileSync(path), { documentId: "native-sheet-copy" }), { kind: "workbook" })); }
const number = (value: number) => ({ type: "number", value });
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-copy-sheet.json", "utf8")) as { cases: { name: string; source: string; sourceSha256: string; producerSha256: string; observed: { sheets: { name: string }[] } }[] };
for (const record of evidence.cases) test("pinned actual " + record.name + " copy reopens independently with recorded producer differences", async () => {
  const producer = readFileSync("packages/office/tests/fixtures/workbook-copy-sheet-" + record.name + "-calc.xlsx");
  expect(createHash("sha256").update(readFileSync(record.source)).digest("hex")).toBe(record.sourceSha256); expect(createHash("sha256").update(producer).digest("hex")).toBe(record.producerSha256);
  const model = openDocument(await importXlsx(producer, { documentId: "native-sheet-copy-producer" }), { kind: "workbook" }), snapshot = model.snapshot();
  expect(snapshot.metadata.workbook.sheets.map(sheet => sheet.name)).toEqual(record.observed.sheets.map(sheet => sheet.name)); const copied = snapshot.metadata.workbook.sheets.find(sheet => sheet.name === "Copied")!;
  if (record.name === "tables") { expect(copied.tables![0]!.name).toBe("SalesTable_2"); expect(model.getCell(copied.sheetId, 1, 3)).toMatchObject({ value: number(20), formula: { expression: "=SalesTable[[#This Row],[数量]]*SalesTable[[#This Row],[单价]]" } }); }
  if (record.name === "modern") expect(snapshot.metadata.workbook.comments ?? []).toHaveLength(0);
  expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(snapshot);
});
test("independent native name source copies source locals and reachable global definitions, preserving outside sheets and one undo", async () => {
  const s = await session("packages/office/tests/fixtures/workbook-name-prefix.xlsx");
  try {
    const before = s.snapshot(), source = before.metadata.workbook.sheets[0]!, tx = { protocolVersion: 1 as const, documentId: before.manifest.documentId, baseRevision: before.manifest.revision, operations: [{ op: "copyWorkbookSheet" as const, sheetId: source.sheetId, name: "Copied", beforeSheetId: source.sheetId }] };
    const candidate = s.prepare(tx); expect(s.snapshot()).toEqual(before); expect(candidate.preview.metadata.workbook.sheets[0]!.name).toBe("Copied"); s.commit(candidate);
    const added = s.snapshot().metadata.workbook.sheets[0]!, names = s.snapshot().metadata.workbook.names!.filter(name => name.sheetId === added.sheetId);
    expect(names.map(name => name.name).sort()).toEqual(["Local", "Rate", "Tax"]); expect(s.snapshot().metadata.workbook.names!.filter(name => name.sheetId !== added.sheetId)).toEqual(before.metadata.workbook.names);
    for (const [name, address] of [["Tax", "H1"], ["Rate", "A5"], ["Local", "A6"]]) expect(parseWorkbookFormula(names.find(value => value.name === name)!.expression, { row: 0, column: 0 })).toEqual(parseWorkbookFormula("=Copied!$" + address![0] + "$" + address!.slice(1), { row: 0, column: 0 }));
    apply(s, [{ op: "setWorkbookCells", sheetId: added.sheetId, cells: [[4, 0, 99], [5, 0, 123], [0, 7, 88]].map(([row, column, value]) => ({ row: row!, column: column!, source: { kind: "value", value: { type: "number", value: value! } } })) }]); s.calculateSync();
    expect([4, 5, 6, 7].map(row => s.getCalculatedCell(added.sheetId, row, 1))).toEqual([187, 345, 88, 88].map(value => expect.objectContaining({ state: "valid", value: number(value) })));
    expect(s.getCalculatedCell(before.metadata.workbook.sheets[1]!.sheetId, 7, 1)).toMatchObject({ state: "valid", value: number(7) });
    expect(s.snapshot().manifest.chunks.filter(chunk => chunk.sheetId !== added.sheetId)).toEqual(before.manifest.chunks);
    expect(s.snapshot().metadata.workbook.sheets.filter(sheet => sheet.sheetId !== added.sheetId)).toEqual(before.metadata.workbook.sheets);
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
test("copied native tables own new identities, recipes and structured references while the original table and outside formulas stay fixed", async () => {
  const s = await session("packages/office/tests/fixtures/workbook-axis-tables.xlsx");
  try {
    const before = s.snapshot(), original = before.metadata.workbook.sheets[0]!, table = original.tables![0]!;
    apply(s, [{ op: "copyWorkbookSheet", sheetId: original.sheetId, name: "Copied" }]); const added = s.snapshot().metadata.workbook.sheets.at(-1)!, copied = added.tables![0]!;
    expect(copied.tableId).not.toBe(table.tableId); expect(copied.name).not.toBe(table.name); expect(copied.range).toEqual(table.range);
    expect(copied.columns.map(column => column.columnId).some(id => table.columns.some(old => old.columnId === id))).toBe(false);
    expect(s.getCell(added.sheetId, 1, 3)!.formula!.expression).toContain(copied.name); expect(copied.columns[3]!.calculatedFormula).toBeDefined();
    expect(s.getCell(added.sheetId, 1, 3)!.value).toEqual(s.getCell(original.sheetId, 1, 3)!.value);
    apply(s, [{ op: "setWorkbookCells", sheetId: added.sheetId, cells: [{ row: 1, column: 1, source: { kind: "value", value: { type: "number", value: 20 } } }] }]); s.calculateSync();
    expect(s.getCalculatedCell(added.sheetId, 1, 3)).toMatchObject({ state: "valid", value: number(200) }); expect(s.getCalculatedCell(original.sheetId, 1, 3)).toMatchObject({ state: "valid", value: number(20) });
    expect(s.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === original.sheetId)).toEqual(original);
    expect(s.snapshot().manifest.chunks.filter(chunk => chunk.sheetId !== added.sheetId)).toEqual(before.manifest.chunks);
    expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});

test("copy allocates an independent native table object name when it differs from the formula name", async () => {
  const files = new Map(await importXlsx(readFileSync("packages/office/tests/fixtures/workbook-axis-tables.xlsx"), { documentId: "sheet-copy-table-object-name" }));
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.tables![0]!.objectName = "SalesTable_2";
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); const s = createDocumentSession(openDocument(files, { kind: "workbook" }));
  try {
    const before = s.snapshot(); apply(s, [{ op: "copyWorkbookSheet", sheetId: before.metadata.workbook.sheets[0]!.sheetId, name: "Copied" }]);
    expect(s.snapshot().metadata.workbook.sheets.at(-1)!.tables![0]).toMatchObject({ name: "SalesTable_3", objectName: "SalesTable_3" });
    expect(s.snapshot().metadata.workbook.sheets[0]).toEqual(before.metadata.workbook.sheets[0]); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});

test("copied unqualified chart references become stale while literal vectors retain their sole source and exact values", async () => {
  const files = new Map(await importXlsx(readFileSync("examples/xlsx/07-three-charts.xlsx"), { documentId: "sheet-copy-chart-cache" }));
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata, chart = metadata.workbook.charts![0]!, values = chart.series[0]!.values;
  if (!("expression" in values) || !values.cache) throw Error("The independent chart must contain a reference and historical cache."); values.expression = "=B2:B4";
  chart.series.push({ seriesId: "copy-literal-series", values: { kind: "literal", valueType: "number", count: 3, points: [{ index: 0, value: 9 }, { index: 2, value: null }], numberFormat: "0.00" } });
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); const s = createDocumentSession(openDocument(files, { kind: "workbook" }));
  try {
    const before = s.snapshot(); apply(s, [{ op: "copyWorkbookSheet", sheetId: before.metadata.workbook.sheets[0]!.sheetId, name: "Copied" }]);
    const added = s.snapshot().metadata.workbook.sheets.at(-1)!, nodeId = s.snapshot().metadata.workbook.objects!.find(object => object.sheetId === added.sheetId && object.kind === "chart")!.nodeId;
    const copied = s.snapshot().metadata.workbook.charts!.find(chart => chart.nodeId === nodeId)!;
    expect(copied.series[0]!.values).toMatchObject({ expression: "=B2:B4", cache: { ...values.cache, state: "stale" } });
    expect(copied.series.at(-1)!.values).toEqual(chart.series.at(-1)!.values); expect(s.snapshot().metadata.workbook.charts![0]).toEqual(before.metadata.workbook.charts![0]);
    expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
for (const name of ["03-structure", "07-three-charts", "08-images-shapes", "09-classic-notes", "10-modern-threads", "11-print-layout", "12-multi-sheet-filter"]) test("copying native " + name + " allocates rich source identities, preserves associations and reopens/undoes atomically", async () => {
  const s = await session("examples/xlsx/" + name + ".xlsx");
  try {
    const before = s.snapshot(), source = before.metadata.workbook.sheets[0]!; apply(s, [{ op: "copyWorkbookSheet", sheetId: source.sheetId, name: "Copied" }]);
    const after = s.snapshot(), added = after.metadata.workbook.sheets.at(-1)!, oldObjects = before.metadata.workbook.objects ?? [], objects = after.metadata.workbook.objects?.filter(object => object.sheetId === added.sheetId) ?? [];
    expect({ rows: added.rowDimensions, columns: added.columnDimensions, merges: added.merges, frozen: added.frozen, print: added.print, filter: added.autoFilter }).toEqual({ rows: source.rowDimensions, columns: source.columnDimensions, merges: source.merges, frozen: source.frozen, print: source.print, filter: source.autoFilter });
    expect(objects).toHaveLength(oldObjects.length); expect(objects.map(object => object.anchor)).toEqual(oldObjects.map(object => object.anchor));
    expect(after.manifest.assets).toEqual(before.manifest.assets); expect(objects.every(object => !oldObjects.some(old => old.nodeId === object.nodeId))).toBe(true);
    for (const [index, object] of objects.entries()) expect(s.getNode(object.nodeId)!.text).toBe(s.getNode(oldObjects[index]!.nodeId)!.text);
    const annotations = after.annotations.filter(annotation => annotation.anchor.kind === "cells" && annotation.anchor.sheetId === added.sheetId);
    expect(annotations).toHaveLength(before.annotations.length); expect(annotations.map(annotation => annotation.status)).toEqual(before.annotations.map(annotation => annotation.status));
    for (const [index, annotation] of annotations.entries()) { const old = before.annotations[index]!; expect(annotation.annotationId).not.toBe(old.annotationId); expect(annotation.replies ?? []).toHaveLength(old.replies?.length ?? 0); expect(annotation.replies?.some(reply => old.replies?.some(previous => previous.replyId === reply.replyId)) ?? false).toBe(false); if (annotation.bodyNodeId) expect(s.getNode(annotation.bodyNodeId)!.text).toBe(s.getNode(old.bodyNodeId!)!.text); }
    for (const chart of after.metadata.workbook.charts?.filter(chart => objects.some(object => object.nodeId === chart.nodeId)) ?? []) for (const series of chart.series) if ("expression" in series.values) expect(series.values.expression).toContain("Copied");
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(after);
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().styles).toEqual(before.styles); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
test("the sole Worker holds a whole-sheet copy privately, discards it, then commits and undoes once", async () => {
  const s = await session("examples/xlsx/10-modern-threads.xlsx"), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), state = client.state(), command = { kind: "transaction" as const, transaction: { protocolVersion: 1 as const, documentId: state.documentId, baseRevision: state.revision, operations: [{ op: "copyWorkbookSheet" as const, sheetId: before.metadata.workbook.sheets[0]!.sheetId, name: "Copied" }] } };
    const candidate = await client.prepare(command), view = await client.view(candidate); expect(view.cells.snapshot.metadata.workbook.sheets).toHaveLength(2); expect(s.snapshot()).toEqual(before); view.dispose(); await client.discard(candidate); expect(s.snapshot()).toEqual(before);
    await client.commit(await client.prepare(command)); expect(s.history().undoSteps).toBe(1); await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});

test("native hyperlink copies retain literal qualified destinations, unqualified locations and authorized external URLs", async () => {
  const s = await session("packages/office/tests/fixtures/workbook-copy-sheet-hyperlinks.xlsx");
  try {
    const before = s.snapshot(), source = before.metadata.workbook.sheets[0]!; apply(s, [{ op: "copyWorkbookSheet", sheetId: source.sheetId, name: "Copied" }]);
    const copied = s.snapshot().metadata.workbook.sheets.at(-1)!; expect(copied.hyperlinks).toEqual(source.hyperlinks);
    expect(copied.hyperlinks).toEqual(expect.arrayContaining([expect.objectContaining({ location: "Data!$A$5" }), expect.objectContaining({ location: "A5" }), expect.objectContaining({ location: "Rate" }), expect.objectContaining({ location: "Other!A1" }), expect.objectContaining({ url: "https://example.com/authorized-link" })]));
    expect(s.snapshot().metadata.workbook.names!.find(name => name.sheetId === copied.sheetId && name.name === "Rate")!.expression).toBe("='Copied'!$A$5");
    expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
