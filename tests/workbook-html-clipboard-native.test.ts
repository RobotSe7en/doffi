import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { connectWorkbookSession, createWorkbookHtmlPastePlan, serveWorkbookSession } from "@doffi/editor";

const evidence = JSON.parse(readFileSync("docs/evidence/workbook-html-clipboard.json", "utf8")) as { cases: { name: string; source: string; htmlFile: string; sourceRows: number[]; sourceColumns: number[] }[] };
test("native controller transfers retain immutable independently authored sources and HTML byte hashes", () => {
  expect(execFileSync("python3", ["scripts/check-workbook-html-clipboard.py"], { encoding: "utf8" })).toContain("three independent");
});
test.each(evidence.cases)("actual $name native HTML preserves populated source values, package reopen and one-step undo", async item => {
  const original = openDocument(await importXlsx(readFileSync(item.source), { documentId: "native-html-input" }), { kind: "workbook" });
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "native-html-target" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), at = { row: 255, column: 31 }, sourceSheet = original.snapshot().metadata.workbook.activeSheetId;
    const plan = createWorkbookHtmlPastePlan(before, "sheet-1", at, readFileSync(item.htmlFile, "utf8"));
    s.commit(s.prepare(plan.transaction)); let checked = 0;
    for (let row = 0; row < item.sourceRows.length; row++) for (let column = 0; column < item.sourceColumns.length; column++) {
      const source = original.getCell(sourceSheet, item.sourceRows[row]!, item.sourceColumns[column]!);
      if (source && source.value.type !== "blank" && !source.formula) {
        expect(s.getCell("sheet-1", row + at.row, column + at.column)?.value, `${item.name}:${row}:${column}`).toEqual(source.value); checked++;
      }
    }
    expect(checked).toBeGreaterThan(0); expect(s.history().undoSteps).toBe(1);
    expect(openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().styles).toEqual(before.styles); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
test("the sole Worker privately prepares native HTML with bounded diagnostics, discards, retries and commits one history step", async () => {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "worker-native-html" }), { kind: "workbook" })), channel = new MessageChannel();
  const server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.serialize(), state = client.state(), command = { kind: "paste" as const, documentId: state.documentId, revision: state.revision,
      sheetId: "sheet-1", at: { row: 255, column: 31 }, text: "fallback must not override native HTML", html: readFileSync(evidence.cases[0]!.htmlFile, "utf8") };
    const candidate = await client.prepare(command); expect(s.serialize()).toEqual(before); expect(Object.isFrozen(candidate.clipboardDiagnostics)).toBe(true);
    const view = await client.view(candidate); expect(view.cells.snapshot.manifest.chunks.length).toBeGreaterThan(0); view.dispose(); await client.discard(candidate); expect(s.serialize()).toEqual(before);
    await client.commit(await client.prepare(command)); expect(s.getCell("sheet-1", 255, 31)?.value).toEqual({ type: "text", value: "类型与中文😀禰󠄀" }); expect(s.history().undoSteps).toBe(1);
    await client.undo(); expect(s.snapshot().manifest.chunks).toEqual([]);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
