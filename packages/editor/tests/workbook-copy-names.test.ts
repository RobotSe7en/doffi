import { expect, test } from "vitest";
import { createDocumentSession, parseWorkbookFormula } from "@doffi/core";
import type { WorkbookDefinedName } from "@doffi/core";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { createWorkbookRangeCopyTransaction } from "../src/workbook/copy-range.js";

const at = { row: 1, column: 1 }, area = { from: at, to: at }, destination = { row: 4, column: 3 };
const number = (value: number) => ({ type: "number" as const, value });
function definition(name: string, expression: string, sheetId?: string): WorkbookDefinedName {
  return { nameId: (sheetId ?? "global") + "-" + name, name, origin: { row: 0, column: 0 }, expression, ...(sheetId ? { sheetId } : {}) };
}
function setup(expression: string, names: WorkbookDefinedName[]) {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(3) }, { ...at, value: number(999), formula: expression }], metadata => {
    const sheet = metadata.workbook.sheets[0]!;
    metadata.workbook.sheets.push({ ...structuredClone(sheet), sheetId: "target", name: "Target" }, { ...structuredClone(sheet), sheetId: "outside", name: "Outside" }); metadata.workbook.names = names;
  }));
  session.commit(session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: session.snapshot().manifest.revision, operations: [
    { op: "setWorkbookCells", sheetId: "target", cells: [{ row: 0, column: 0, source: { kind: "value", value: number(11) } }, { row: 3, column: 2, source: { kind: "value", value: number(2) } }] },
    { op: "setWorkbookCells", sheetId: "outside", cells: [{ row: 0, column: 0, source: { kind: "value", value: number(5) } }] },
  ] })); session.clearHistory(); return session;
}
test("copied local alias graphs rebase their own sheet, retain external and global bindings, and reuse existing target locals", async () => {
  const main = { ...definition("Main", "=Sheet1!Branch+Tax+Outside!$A$1", "sheet-1"), hidden: true, description: "中文禰󠄀" };
  const s = setup("=Sheet1!Leaf+[0]!Tax+Main+A1", [main, definition("Branch", "=Leaf+Sheet1!$A$1", "sheet-1"), definition("Leaf", "=$A$1+1", "sheet-1"), definition("Tax", "=3"), definition("Tax", "=900", "target"), definition("Leaf", "=17", "target")]);
  try {
    const before = s.snapshot(), tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", area, destination, { targetSheetId: "target" });
    const names = tx.operations.flatMap(op => op.op === "setWorkbookDefinedName" ? [op.name] : []); expect(names.map(name => name.name)).toEqual(["Main", "Branch"]);
    expect(names[0]).toMatchObject({ hidden: true, description: main.description, sheetId: "target", origin: main.origin });
    expect(parseWorkbookFormula(names[0]!.expression, main.origin)).toEqual(parseWorkbookFormula("=Target!Branch+[0]!Tax+Outside!$A$1", main.origin));
    expect(parseWorkbookFormula(names[1]!.expression, main.origin)).toEqual(parseWorkbookFormula("=Leaf+Target!$A$1", main.origin));
    s.commit(s.prepare(tx)); s.calculateSync(); expect(s.getCalculatedCell("target", destination.row, destination.column)).toMatchObject({ value: number(45) });
    expect(s.snapshot().metadata.workbook.names!.filter(name => before.metadata.workbook.names!.some(old => old.nameId === name.nameId))).toEqual(before.metadata.workbook.names);
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
test("explicit source locals and existing target aliases do not clone unused name subgraphs", async () => {
  const s = setup("=Sheet1!Leaf+Main", [definition("Leaf", "=4", "sheet-1"), definition("Main", "=Leaf", "sheet-1"), definition("Main", "=17", "target")]);
  try {
    const tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", area, destination, { targetSheetId: "target" }); expect(tx.operations.some(op => op.op === "setWorkbookDefinedName")).toBe(false);
    s.commit(s.prepare(tx)); s.calculateSync(); expect(s.getCalculatedCell("target", destination.row, destination.column)).toMatchObject({ value: number(21) });
  } finally { s.dispose(); }
});
test("cyclic aliases reserve distinct identities once and preserve the cycle in one undo", async () => {
  const s = setup("=CycleA", [definition("CycleA", "=CycleB", "sheet-1"), definition("CycleB", "=Sheet1!CycleA", "sheet-1"), { ...definition("Unused", "=1"), nameId: "copied-name-1" }]);
  try {
    const before = s.snapshot(), tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", area, destination, { targetSheetId: "target" });
    const names = tx.operations.flatMap(op => op.op === "setWorkbookDefinedName" ? [op.name] : []); expect(names.map(name => name.nameId)).toEqual(["copied-name-2", "copied-name-3"]);
    s.commit(s.prepare(tx)); s.calculateSync(); expect(s.getCalculatedCell("target", destination.row, destination.column)).toMatchObject({ state: "unavailable", issue: { reason: "circular-name" } });
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
test("large local-name graph work is cancellable, bounded and revision guarded before any edits publish", async () => {
  const names = Array.from({ length: 320 }, (_, index) => definition("Alias_" + index, index === 319 ? "=$A$1" : "=Alias_" + (index + 1), "sheet-1"));
  const s = setup("=Alias_0", names);
  try {
    const before = s.serialize(), signal = { aborted: false };
    await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", area, destination, { targetSheetId: "target", signal, yield: async () => { signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
    for (const budget of [{ maxSteps: 400 }, { maxCodeUnits: 128 }]) await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", area, destination, { targetSheetId: "target", ...budget })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
    let changed = false;
    await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", area, destination, { targetSheetId: "target", yield: async () => {
      if (changed) return; changed = true; s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations: [{ op: "setWorkbookCells", sheetId: "outside", cells: [{ row: 2, column: 2, source: { kind: "value", value: number(8) } }] }] }));
    } })).rejects.toMatchObject({ code: "RevisionConflict" });
    expect(s.getCell("target", destination.row, destination.column)).toBeUndefined(); expect(s.snapshot().metadata.workbook.names).toEqual(names); expect(s.history().undoSteps).toBe(1);
  } finally { s.dispose(); }
});
