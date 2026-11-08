import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { connectWorkbookSession, serveWorkbookSession } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { writeOpcZip } from "../packages/office/src/document-zip.js";

const bytes = readFileSync("packages/office/tests/fixtures/workbook-name-prefix.xlsx");
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-name-prefix.json", "utf8")) as { sourceSha256: string; sourceCurrentWorkbookNameExpected: Record<string, number> };
test("the native current-workbook name source and four producer records are independent and reproducible", () => {
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-name-prefix.py"], { encoding: "utf8" })).toContain("four independent");
});
test("independent [0] names survive native import, current editing, package reopen and undo without adopting local shadows", async () => {
  const model = openDocument(await importXlsx(bytes, { documentId: "native-book-name" }), { kind: "workbook" }), s = createDocumentSession(model);
  try {
    const before = s.snapshot(), [source, other] = before.metadata.workbook.sheets;
    s.calculateSync();
    for (const sheet of before.metadata.workbook.sheets) {
      expect(s.getCell(sheet.sheetId, 7, 1)).toMatchObject({ value: { type: "number", value: 7 }, formula: { expression: "=[0]!Tax", cacheState: "imported" } });
      expect(s.getCalculatedCell(sheet.sheetId, 7, 1)).toMatchObject({ state: "valid", value: { type: "number", value: evidence.sourceCurrentWorkbookNameExpected[sheet.name + "!B8"] } });
    }
    s.commit(s.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: source!.sheetId, cells: [{ row: 0, column: 7, source: { kind: "value", value: { type: "number", value: 9 } } }] }] })); s.calculateSync();
    expect(s.getCalculatedCell(other!.sheetId, 7, 1)).toMatchObject({ value: { type: "number", value: 9 } });
    expect(s.getCalculatedCell(other!.sheetId, 6, 1)).toMatchObject({ value: { type: "number", value: 70 } });
    expect(s.getCell(other!.sheetId, 7, 1)!.value).toEqual({ type: "number", value: 7 });
    const reopened = openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }); expect(reopened.snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
test("the recorded Calc prefix loss remains a source/cache difference and is confirmed by actual second reopen", async () => {
  const copied = readFileSync("packages/office/tests/fixtures/workbook-name-prefix-copy-other-calc.xlsx"), reopened = readFileSync("packages/office/tests/fixtures/workbook-name-prefix-reopen-copy-other-calc.xlsx");
  const s = createDocumentSession(openDocument(await importXlsx(copied, { documentId: "copied-name-producer" }), { kind: "workbook" }));
  const second = createDocumentSession(openDocument(await importXlsx(reopened, { documentId: "reopened-name-producer" }), { kind: "workbook" }));
  try {
    const other = s.snapshot().metadata.workbook.sheets[1]!.sheetId, next = second.snapshot().metadata.workbook.sheets[1]!.sheetId;
    expect(s.getCell(other, 7, 1)).toMatchObject({ value: { type: "number", value: 7 }, formula: { expression: "=Tax", cacheState: "imported" } });
    s.calculateSync(); expect(s.getCalculatedCell(other, 7, 1)).toMatchObject({ state: "valid", value: { type: "number", value: 70 } });
    expect(second.getCell(next, 7, 1)!.value).toEqual({ type: "number", value: 70 });
    expect(s.getCell(other, 4, 1)!.value).toEqual({ type: "number", value: 107 });
    expect(s.getCalculatedCell(other, 4, 1)).toMatchObject({ value: { type: "number", value: 170 } }); expect(second.getCell(next, 4, 1)!.value).toEqual({ type: "number", value: 170 });
  } finally { s.dispose(); second.dispose(); }
});
test("the sole Worker evaluates and discards native prefix input privately before one commit and undo", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "worker-book-name" }), { kind: "workbook" })), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), sheetId = before.metadata.workbook.sheets[1]!.sheetId, state = client.state(), command = { kind: "input" as const, documentId: state.documentId, revision: state.revision, sheetId, at: { row: 7, column: 1 }, text: "=[0]!Tax+1" };
    const candidate = await client.prepare(command), view = await client.view(candidate);
    expect(s.snapshot()).toEqual(before);
    const current = await view.rules!.read({ documentId: state.documentId, revision: state.revision, sheetId, coordinates: [{ row: 7, column: 1 }], includeCurrent: true, maxCells: 100, maxRules: 100, maxValues: 100, maxSteps: 10000, maxCodeUnits: 10000 }, new AbortController().signal);
    expect(current.cells[0]!.current).toEqual({ kind: "value", value: { type: "number", value: 8 } }); view.dispose(); await client.discard(candidate);
    expect(s.snapshot()).toEqual(before); await client.commit(await client.prepare(command)); s.calculateSync();
    expect(s.getCalculatedCell(sheetId, 7, 1)).toMatchObject({ value: { type: "number", value: 8 } }); expect(s.history().undoSteps).toBe(1);
    await client.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});

test("the empty producer property container does not permit unmapped populated properties or invalid roots", async () => {
  const original = readDocumentZip(readFileSync("packages/office/tests/fixtures/workbook-name-prefix-reopen-copy-other-calc.xlsx")), decoder = new TextDecoder(), encoder = new TextEncoder();
  const text = decoder.decode(original.get("docProps/custom.xml")!);
  for (const [value, code] of [[text.replace("</Properties>", '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MustPreserve"><vt:lpwstr>not silently discarded</vt:lpwstr></property></Properties>'), "UnsupportedFeature"],
    [text.replaceAll("Properties", "NotProperties"), "InvalidDocument"]] as const) {
    const files = new Map(original); files.set("docProps/custom.xml", encoder.encode(value));
    await expect(importXlsx(writeOpcZip(files), { documentId: "unsafe-custom-container" })).rejects.toMatchObject({ code });
  }
});
