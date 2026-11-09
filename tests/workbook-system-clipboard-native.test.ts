import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { connectWorkbookSession, createWorkbookClipboardData, createWorkbookHtmlPastePlan, serveWorkbookSession } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

test("the native clipboard consumer keeps independent source/producer byte provenance", () => {
  expect(execFileSync("python3", ["scripts/check-workbook-system-clipboard.py"], { encoding: "utf8" })).toContain("immutable independent source");
});
test("actual native paste and edit retain formulas, precise numbers, types, producer differences, Worker reads and one-step SDK history", async () => {
  const source = openDocument(await importXlsx(readFileSync("packages/office/tests/fixtures/workbook-system-clipboard-calc.xlsx"), { documentId: "system-native-reopen" }), { kind: "workbook" });
  const s = createDocumentSession(source), sheetId = source.snapshot().metadata.workbook.sheets[0]!.sheetId, channel = new MessageChannel();
  const server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    expect(s.getCell(sheetId, 11, 0)?.value).toEqual({ type: "number", value: 15 });
    expect(s.getCell(sheetId, 11, 1)?.formula?.expression).toBe("=(A12*2)");
    expect(s.getCell(sheetId, 11, 2)?.value).toEqual({ type: "text", value: "001234567890123456789" });
    expect(s.getCell(sheetId, 12, 3)?.value).toEqual({ type: "text", value: "=1+1" });
    expect(s.getCell(sheetId, 13, 0)?.value).toEqual({ type: "text", value: '中文\n\n"甲 乙"😀禰󠄀' });
    expect(s.getCell(sheetId, 13, 1)?.value).toEqual({ type: "number", value: 60 });
    expect(s.getCell(sheetId, 13, 2)?.formula?.expression).toBe("=TRUE()");
    expect(s.getCell(sheetId, 13, 3)?.value).toEqual({ type: "number", value: 12345.6789 });
    const before = s.snapshot(), at = { row: 11, column: 0 };
    const candidate = await client.prepare({ kind: "input", documentId: client.state().documentId, revision: client.state().revision, sheetId, at, text: "18" });
    expect(s.snapshot()).toEqual(before); await client.discard(candidate); expect(s.snapshot()).toEqual(before);
    await client.commit(await client.prepare({ kind: "input", documentId: client.state().documentId, revision: client.state().revision, sheetId, at, text: "18" }));
    const view = await client.view();
    try {
      const data = await client.readClipboard(sheetId, { from: at, to: { row: 13, column: 3 } }, { view });
      expect(data.text).toContain("18\t36\t001234567890123456789");
      const direct = await createWorkbookClipboardData(s, sheetId, { from: at, to: { row: 13, column: 3 } }, { calculation: await s.calculate() });
      expect(data).toEqual(direct);
      // Preparing the returned public bytes is still an ordinary isolated
      // candidate, without treating clipboard coordinates as capabilities.
      const plan = createWorkbookHtmlPastePlan(s.snapshot(), sheetId, { row: 255, column: 31 }, data.html);
      const prepared = s.prepare(plan.transaction); expect(s.history().undoSteps).toBe(1); s.discard(prepared);
    } finally { view.dispose(); }
    expect(openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    await client.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
