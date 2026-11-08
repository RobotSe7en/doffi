import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument, parseWorkbookFormula } from "@doffi/core";
import type { WorkbookOperation } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

type NativeSheet = { formulas: { address: string; formula: string; value: string; type: string }[]; merges: string[];
  conditional: { ranges: string; formulas: string[] }[]; validation: { ranges: string; formula: string }[] };
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-axis.json", "utf8")) as { sourceSha256: string;
  cases: { operation: { axis: "row" | "column"; action: "insert" | "delete"; index: number; count: number };
    observed: { sheets: NativeSheet[]; names: { name: string; scope: string; expression: string }[] } }[] };
const input = readFileSync("packages/office/tests/fixtures/workbook-axis.xlsx");
function at(address: string) {
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(address)!;
  let column = 0; for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}
function area(text: string) { const [from, to = from] = text.replaceAll("$", "").split(":"); return { from: at(from!), to: at(to!) }; }

test("the independent axis source is reproducible and tied to recorded native observations", () => {
  expect(createHash("sha256").update(input).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-axis.py"], { encoding: "utf8" })).toContain("four independent");
});
test.each(evidence.cases)("native axis $operation.axis/$operation.action preserves current formulas and mapped rules/print in one history step", async item => {
  const source = openDocument(await importXlsx(input, { documentId: "axis-native" }), { kind: "workbook" }), s = createDocumentSession(source);
  try {
    const original = s.snapshot(), target = original.metadata.workbook.sheets[0]!;
    const operation: WorkbookOperation = { op: "editWorkbookAxis", sheetId: target.sheetId, ...item.operation,
      ...(item.operation.action === "insert" ? { formatSource: "before" as const } : {}) };
    s.commit(s.prepare({ protocolVersion: 1, documentId: "axis-native", baseRevision: 0, operations: [operation] }));
    s.calculateSync();
    for (const descriptor of original.manifest.chunks) for (const offset of [0, 128]) for (const cell of source.readWindow({
      sheetId: descriptor.sheetId, row: descriptor.row + offset, column: descriptor.column, rows: 128, columns: 32,
    })) {
      const position = { row: cell.row, column: cell.column }, change = item.operation, old = position[change.axis];
      if (descriptor.sheetId === target.sheetId && old >= change.index) {
        if (change.action === "delete" && old < change.index + change.count) continue;
        position[change.axis] += change.action === "insert" ? change.count : -change.count;
      }
      expect(s.getCell(descriptor.sheetId, position.row, position.column)).toMatchObject({
        value: cell.value, text: cell.text, ...(cell.cssClass ? { cssClass: cell.cssClass } : {}), ...(cell.formatId ? { formatId: cell.formatId } : {}),
      });
    }
    for (const [index, expected] of item.observed.sheets.entries()) {
      const sheet = s.snapshot().metadata.workbook.sheets[index]!;
      const cells = s.snapshot().manifest.chunks.filter(chunk => chunk.sheetId === sheet.sheetId).flatMap(chunk =>
        [0, 128].flatMap(offset => s.readWindow({ sheetId: sheet.sheetId, row: chunk.row + offset, column: chunk.column, rows: 128, columns: 32 })));
      expect(cells.filter(cell => cell.formula)).toHaveLength(expected.formulas.length);
      for (const formula of expected.formulas) {
        const location = at(formula.address), current = s.getCell(sheet.sheetId, location.row, location.column)!;
        expect(parseWorkbookFormula(current.formula!.expression, location)).toEqual(parseWorkbookFormula("=" + formula.formula, location));
        const value = formula.type === "e" ? { type: "error", value: formula.value } : { type: "number", value: Number(formula.value) };
        expect(s.getCalculatedCell(sheet.sheetId, location.row, location.column)).toMatchObject({ state: "valid", value });
      }
      for (const [ruleIndex, rule] of (sheet.conditionalRules ?? []).entries()) {
        const native = expected.conditional[ruleIndex]!;
        expect(rule.ranges).toEqual(native.ranges.split(" ").map(area));
        if ("expression" in rule) expect(parseWorkbookFormula(rule.expression, rule.origin)).toEqual(parseWorkbookFormula("=" + native.formulas[0], rule.origin));
      }
      for (const [ruleIndex, rule] of (sheet.dataValidations ?? []).entries()) {
        const native = expected.validation[ruleIndex]!;
        expect(rule.ranges).toEqual(native.ranges.split(" ").map(area));
        if ("expression" in rule) expect(parseWorkbookFormula(rule.expression, rule.origin)).toEqual(parseWorkbookFormula("=" + native.formula, rule.origin));
      }
      const print = item.observed.names.find(name => name.name === "_xlnm.Print_Area" && Number(name.scope) === index)!;
      expect(sheet.print?.areas).toEqual([area(print.expression.split("!")[1]!)]);
      // The user selected Microsoft Excel compatibility (2026-11-08).
      // This Calc difference remains open until explicit Microsoft semantics or
      // actual Excel evidence fixes the expected topology; it is not acceptance.
      if (!(item.operation.axis === "row" && item.operation.action === "delete" && index === 0)) expect(sheet.merges).toEqual(expected.merges.map(area));
    }
    const reopened = openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" });
    expect(reopened.snapshot()).toEqual(s.snapshot()); expect(s.history().undoSteps).toBe(1);
    s.undo(); expect(s.snapshot().manifest.chunks).toEqual(original.manifest.chunks); expect(s.snapshot().metadata).toEqual(original.metadata);
    s.redo(); expect(s.snapshot().metadata).toEqual(reopened.snapshot().metadata);
    expect(source.snapshot()).toEqual(original);
  } finally { s.dispose(); }
});
