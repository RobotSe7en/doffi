import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument, parseWorkbookFormula, WorkbookConditionalRuleSchema } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { connectWorkbookSession, createWorkbookRangeCopyTransaction, serveWorkbookSession } from "@doffi/editor";
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
for (const path of ["examples/xlsx/06-rules-validation.xlsx", "examples/xlsx/producer/06-rules-validation.xlsx", ...["ordinary", "partial-scale", "overlap", "validations", "other", "invalid-stop"].map(name => `packages/office/tests/fixtures/workbook-copy-rules-${name}-calc.xlsx`)]) {
  test(`independent/native rule source remains editable and reopens with exact current metadata: ${path}`, async () => {
    const s = createDocumentSession(openDocument(await importXlsx(readFileSync(path), { documentId: "native-copy-rules" }), { kind: "workbook" }));
    try {
      const before = s.snapshot(), original = before.metadata.workbook.sheets[0]!, tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(1, 0, 1, 4), { row: 269, column: 32 });
      const candidate = s.prepare(tx); expect(s.snapshot()).toEqual(before); s.discard(candidate); s.commit(s.prepare(tx));
      const current = s.snapshot(), sheet = current.metadata.workbook.sheets[0]!, copied = sheet.conditionalRules!.filter(rule => rule.ranges.some(area => area.from.row === 269));
      expect(copied.map(rule => rule.kind)).toEqual(["cellIs", "expression", "colorScale", "colorScale"]);
      const expression = copied.find(rule => rule.kind === "expression")!; if (expression.kind !== "expression") throw Error("Missing copied expression");
      expect(parseWorkbookFormula(expression.expression, expression.origin)).toEqual(parseWorkbookFormula("=AND(AH270>=50,AH270<=150)", { row: 269, column: 33 }));
      expect(sheet.dataValidations!.find(rule => rule.ranges.some(area => area.from.row === 269 && area.from.column === 36))).toMatchObject({ kind: "list", expression: original.dataValidations!.find(rule => rule.kind === "list")!.expression });
      expect(openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot().metadata).toEqual(current.metadata);
      expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); s.redo(); expect(s.snapshot().metadata).toEqual(current.metadata);
    } finally { s.dispose(); }
  });
}
test("native copy accepts an imported stop violation while the SDK transaction rejects its publication", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/06-rules-validation.xlsx"), { documentId: "native-copy-stop" }), { kind: "workbook" }));
  try {
    const before = s.serialize(), tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(2, 4), { row: 256, column: 32 });
    expect(() => s.prepare(tx)).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
    const native = openDocument(await importXlsx(readFileSync("packages/office/tests/fixtures/workbook-copy-rules-invalid-stop-calc.xlsx"), { documentId: "native-stop-observation" }), { kind: "workbook" });
    expect(native.getCell("sheet-1", 256, 32)?.value).toEqual({ type: "number", value: 99 }); expect(native.snapshot().metadata.workbook.sheets[0]!.dataValidations!.some(rule => rule.ranges.some(area => area.from.row === 256))).toBe(true);
  } finally { s.dispose(); }
});
test("Worker copy keeps cells and rules in one private candidate and one authoritative history", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/06-rules-validation.xlsx"), { documentId: "worker-copy-rules" }), { kind: "workbook" }), { sessionLimits: { maxCandidates: 1 } });
  const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), state = client.state(), candidate = await client.prepare({ kind: "copy", documentId: state.documentId, revision: state.revision, sheetId: "sheet-1", source: range(1, 0, 1, 4), target: { row: 256, column: 32 } });
    const source = await client.view(), preview = await client.view(candidate);
    try { expect(source.cells.snapshot.metadata).toEqual(before.metadata); expect(preview.cells.snapshot.metadata.workbook.sheets[0]!.conditionalRules).toHaveLength(8); expect(preview.cells.snapshot.metadata.workbook.sheets[0]!.dataValidations).toHaveLength(8); expect(s.history().undoSteps).toBe(0); } finally { source.dispose(); preview.dispose(); }
    await client.commit(candidate); expect(s.history().undoSteps).toBe(1); await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
test("formula color thresholds rebase from an origin outside the copied intersection", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(readFileSync("examples/xlsx/06-rules-validation.xlsx"), { documentId: "copy-scale-formula" }), { kind: "workbook" }));
  try {
    const original = WorkbookConditionalRuleSchema.parse(s.snapshot().metadata.workbook.sheets[0]!.conditionalRules![2]!); if (original.kind !== "colorScale") throw Error("Missing native color scale");
    original.stops[0] = { kind: "formula", expression: "=A2", cssClass: original.stops[0]!.cssClass };
    const m = s.snapshot().manifest; s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [{ op: "setWorkbookConditionalRule", sheetId: "sheet-1", rule: original }] })); s.clearHistory();
    s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(2, 2, 3, 2), { row: 256, column: 32 })));
    const copied = s.snapshot().metadata.workbook.sheets[0]!.conditionalRules!.find(rule => rule.ruleId.startsWith("copied-rule"))!;
    expect(copied).toMatchObject({ kind: "colorScale", origin: { row: 256, column: 32 }, ranges: [range(256, 32, 257, 32)] });
    if (copied.kind !== "colorScale") throw Error("Missing copied color scale"); expect(copied.stops[0]).toEqual({ kind: "formula", expression: "=AE257", cssClass: original.stops[0]!.cssClass }); expect(s.history().undoSteps).toBe(1);
  } finally { s.dispose(); }
});
