import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument, parseWorkbookFormula } from "@doffi/core";
import { connectWorkbookSession, serveWorkbookSession } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

const bytes = readFileSync("packages/office/tests/fixtures/workbook-axis-names.xlsx");
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-axis-names.json", "utf8")) as { sourceSha256: string; cases: {
  case: ["row" | "column", "insert" | "delete", number]; producerSha256: string; observed: {
    names: { name: string; sheet: string | null; expression: string }[];
    sheets: { address: string; expression: string; type: string; value: string }[][];
  };
}[] };
function at(address: string) {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(address)!;
  let column = 0; for (const letter of letters!) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(digits) - 1, column: column - 1 };
}
test("relative-name sources and five pinned native operations retain independent provenance", () => {
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-axis-names.py"], { encoding: "utf8" })).toContain("five independent");
});
test.each(evidence.cases)("native name-axis $case agrees on bound relative/absolute recipes, keeps source caches and reopens with history", async item => {
  const source = openDocument(await importXlsx(bytes, { documentId: "native-axis-names" }), { kind: "workbook" }), s = createDocumentSession(source);
  try {
    const before = s.snapshot(), sheetId = before.metadata.workbook.sheets[0]!.sheetId, [axis, action, index] = item.case;
    s.commit(s.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: 0,
      operations: [{ op: "editWorkbookAxis", sheetId, axis, action, index, count: 1 }] })); s.calculateSync();
    // Bang references have separately observed Excel calculation semantics.
    // Calc rewrites their global sheet context and calculates #VALUE!, so its
    // cache is recorded rather than adopted as the expected SDK result.
    for (const name of s.snapshot().metadata.workbook.names ?? []) {
      if (["Local", "Dynamic"].includes(name.name)) continue;
      const native = item.observed.names.find(value => value.name === name.name)!;
      expect(parseWorkbookFormula(name.expression, name.origin)).toEqual(parseWorkbookFormula("=" + native.expression, name.origin));
    }
    for (const [sheetIndex, formulas] of item.observed.sheets.entries()) for (const formula of formulas) {
      const id = before.metadata.workbook.sheets[sheetIndex]!.sheetId, position = at(formula.address), cell = s.getCell(id, position.row, position.column)!;
      expect(cell.formula).toBeDefined(); expect(cell.value).toEqual({ type: "number", value: 999 });
      if (["Local", "Dynamic"].includes(formula.expression)) continue;
      if (axis === "column" && formula.expression === "Alias") {
        // Calc loses the calling coordinate through this nested name after a
        // column edit (10). The name is Relative+10; Relative is 15 on either
        // sheet after insertion, or #REF! when its absolute column is deleted.
        expect(Number(formula.value)).toBe(10);
        expect(s.getCalculatedCell(id, position.row, position.column)).toMatchObject({ state: "valid", value: action === "insert"
          ? { type: "number", value: 25 } : { type: "error", value: "#REF!" } });
        continue;
      }
      const value = formula.type === "e" ? { type: "error", value: formula.value } : { type: "number", value: Number(formula.value) };
      expect(s.getCalculatedCell(id, position.row, position.column)).toMatchObject({ state: "valid", value });
    }
    const reopened = openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" });
    expect(reopened.snapshot()).toEqual(s.snapshot()); expect(s.history().undoSteps).toBe(1);
    s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
    s.redo(); expect(s.snapshot().metadata).toEqual(reopened.snapshot().metadata); expect(source.snapshot()).toEqual(before);
  } finally { s.dispose(); }
});

test.each(evidence.cases)("the actual producer $case retains native caches, supports a new edit, package reopen and exact undo", async item => {
  const path = "packages/office/tests/fixtures/workbook-axis-names-" + item.case.join("-") + "-calc.xlsx", producer = readFileSync(path);
  expect(createHash("sha256").update(producer).digest("hex")).toBe(item.producerSha256);
  const s = createDocumentSession(openDocument(await importXlsx(producer, { documentId: "producer-axis-names" }), { kind: "workbook" }));
  try {
    const before = s.snapshot();
    for (const [index, formulas] of item.observed.sheets.entries()) for (const formula of formulas) {
      const position = at(formula.address), value = formula.type === "e" ? { type: "error", value: formula.value } : { type: "number", value: Number(formula.value) };
      expect(s.getCell(before.metadata.workbook.sheets[index]!.sheetId, position.row, position.column)).toMatchObject({ value, formula: { cacheState: "imported" } });
    }
    s.commit(s.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: before.metadata.workbook.sheets[1]!.sheetId,
      cells: [{ row: 30, column: 10, source: { kind: "value", value: { type: "text", value: "原生再编辑禰󠄀😀" } } }] }] }));
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});

test("the sole Worker prepares, discards, retries and commits contextual names with current results and stale-candidate protection", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "worker-axis-names" }), { kind: "workbook" })),
    channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), [data, other] = before.metadata.workbook.sheets, state = client.state();
    const command = { kind: "transaction" as const, transaction: { protocolVersion: 1 as const, documentId: state.documentId, baseRevision: state.revision,
      operations: [{ op: "editWorkbookAxis" as const, sheetId: data!.sheetId, axis: "row" as const, action: "insert" as const, index: 0, count: 1 }] } };
    const candidate = await client.prepare(command), view = await client.view(candidate);
    expect(s.snapshot()).toEqual(before);
    const current = await view.rules!.read({ documentId: state.documentId, revision: state.revision, sheetId: other!.sheetId, coordinates: [at("C5"), at("F5")],
      includeCurrent: true, maxCells: 100, maxRules: 100, maxValues: 100, maxSteps: 10000, maxCodeUnits: 10000 }, new AbortController().signal);
    expect(current.cells.map(cell => cell.current)).toEqual([{ kind: "value", value: { type: "number", value: 14 } }, { kind: "value", value: { type: "number", value: 101 } }]);
    view.dispose(); await client.discard(candidate); expect(s.snapshot()).toEqual(before);
    await client.commit(await client.prepare(command)); expect(s.history().undoSteps).toBe(1);
    await expect(client.prepare(command)).rejects.toMatchObject({ code: "RevisionConflict" });
    await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
