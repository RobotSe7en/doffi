import { describe, expect, it } from "vitest";
import { copyFormula, rewriteFormula, type FormulaLocation } from "../src/workbook/formula-rewrite.js";
import { parseFormula } from "../src/workbook/formula.js";
import type { FormulaExpression } from "../src/workbook/formula.js";

const place = (row = 0, column = 5, sheet = "Data"): FormulaLocation => ({ row, column, sheet });
const area = (first: number, last: number, column = 0) => ({ from: { row: first, column }, to: { row: last, column } });
it("whole-sheet copies rewrite only mapped sheet, name and table identities while preserving external references and literal strings", () => {
  const from = place(4, 2), to = { ...from, sheet: "Copied" }, change = { kind: "copySheet", from: "Data", to: "Copied" } as const;
  const expression = '=Data!$A1+Other!A$1+[0]!Cloned+[0]!Fixed+SUM(Sales[Amount])+"Data!A1"';
  const bindings = { copiedNames: new Map([["clone-id", "NewLocal"]]), copiedTables: new Map([["SALES", "Sales_2"]]), resolveName: (name: string) => name === "Cloned" ? "clone-id" : "fixed-id" };
  expect(parseFormula(rewriteFormula(expression, from, to, change, bindings), to)).toEqual(parseFormula('=Copied!$A1+Other!A$1+Copied!NewLocal+[0]!Fixed+SUM(Sales_2[Amount])+"Data!A1"', to));
  expect(() => rewriteFormula(expression, from, { ...to, row: 5 }, change, bindings)).toThrow(); expect(bindings.copiedNames.get("clone-id")).toBe("NewLocal");
});
it("empty table recipes preserve relative axis offsets while absolute cells, ranges and axes still follow deleted source", () => {
  const operation = { kind: "axis", sheet: "Data", axis: "row", action: "delete", index: 1, count: 3 } as const, origin = place(1);
  const map = (source: string) => parseFormula(rewriteFormula(source, origin, origin, operation, { preserveRelativeAxis: true }), origin);
  expect(map("=A2+$B$2+A$1+Other!A2")).toEqual(parseFormula("=A2+#REF!+A$1+Other!A2", origin));
  expect(map("=SUM(1:3,$1:$3,$1:3,$2:3,A:A)")).toEqual(parseFormula("=SUM(1:3,$1:$1,$1:3,#REF!,A:A)", origin));
  expect(() => rewriteFormula("=A1", origin, origin, { kind: "renameSheet", from: "Data", to: "Other" }, { preserveRelativeAxis: true })).toThrow();
});
describe("formula reference mapping", () => {
  it("copies relative references and whole axes while preserving constants and absolute references", () => {
    expect(copyFormula('=A1+$B1+C$2+$D$4+SUM(A:C,1:$5)+"A1"', place(), place(2, 7))).toBe('=(((((C3+$B3)+E$2)+$D$4)+SUM(C:E,3:$5))+"A1")');
    expect(copyFormula("=A1:B2", place(2), place())).toBe("=#REF!:#REF!");
    expect(parseFormula(copyFormula("='Other'!A1:B2", place(1), place()), place())).toMatchObject({ operator: ":", left: { value: "#REF!" }, right: { sheet: "Other", row: 0 } });
  });
  it("inserts rows inside ranges and moves absolute coordinates without changing string literals", () => {
    const operation = { kind: "axis", sheet: "Data", axis: "row", action: "insert", index: 2, count: 2 } as const;
    expect(rewriteFormula('=SUM(A1:A5,$A$3,Other!A3)+"A3"', place(), place(), operation)).toBe('=(SUM(A1:A7,$A$5,\'Other\'!A3)+"A3")');
    expect(rewriteFormula("=A1+$B$2", place(4), place(6), operation)).toBe("=(A1+$B$2)");
    expect(rewriteFormula("=SUM(1:5,$A:$C)", place(), place(), operation)).toBe("=SUM(1:7,$A:$C)");
  });
  it("shrinks intersecting deletions and returns REF only for wholly deleted targets", () => {
    const operation = { kind: "axis", sheet: "Data", axis: "row", action: "delete", index: 1, count: 3 } as const;
    expect(rewriteFormula("=SUM(A1:A5,A2:A3,A4:A8,A8:A4,2:4)+$B$3", place(), place(), operation)).toBe("=(SUM(A1:A2,#REF!,A2:A5,A5:A2,#REF!)+#REF!)");
    expect(rewriteFormula("=SUM(A1:A2)", place(), place(), { ...operation, axis: "column", index: 1, count: 1 })).toBe("=SUM(A1:A2)");
  });
  it("maps moved formulas and referenced data across sheets without treating cut as copy", () => {
    const operation = { kind: "move", sheet: "Data", range: { from: { row: 0, column: 0 }, to: { row: 1, column: 1 } }, target: place(4, 2, "New") } as const;
    expect(rewriteFormula("=A1+$B$2+C3", place(0, 0), place(4, 2, "New"), operation)).toBe("=((C5+$D$6)+'Data'!C3)");
    expect(rewriteFormula("=SUM(A1:B2)", place(), place(), operation)).toBe("=SUM('New'!C5:D6)");
  });
  it("follows a partially cut range with exact bounded fragments, never its bounding box", () => {
    const operation = { kind: "move", sheet: "Data", range: area(2, 2), target: place(2, 2) } as const;
    expect(rewriteFormula("=SUM(A1:A5)", place(), place(), operation)).toBe("=SUM(((A1:A2,A4:A5),C3))");
    const large = rewriteFormula("=SUM(A:A)", place(), place(), operation);
    expect(large).toBe("=SUM((('Data'!A$1:A$2,'Data'!A$4:A$1048576),'Data'!C$3))");
    expect(large.length).toBeLessThan(100);
    const overlapping = { ...operation, range: area(0, 2), target: place(1, 0) };
    expect(rewriteFormula("=SUM(A1:A5)", place(), place(), overlapping)).toBe("=SUM((A4:A5,A2:A3))");
  });
  it("keeps outside sort formulas at their addresses and copies inside relative references with their records", () => {
    const operation = { kind: "sort", sheet: "Data", range: { from: { row: 1, column: 0 }, to: { row: 7, column: 2 } } } as const;
    expect(rewriteFormula("=B2", place(0, 5), place(0, 5), operation)).toBe("=B2");
    expect(rewriteFormula("=B2+$D$1", place(1, 2), place(7, 2), operation)).toBe("=(B8+$D$1)");
    expect(() => rewriteFormula("=B2", place(0, 5), place(7, 5), operation)).toThrow();
  });
  it("renames sheets and exact table bindings without changing text or similarly named identifiers", () => {
    expect(rewriteFormula('=Data!A1+Data!Rate+"Data!A1"', place(), place(0, 5, "Q1'预算"), { kind: "renameSheet", from: "Data", to: "Q1'预算" })).toBe('=((\'Q1\'\'预算\'!A1+\'Q1\'\'预算\'!Rate)+"Data!A1")');
    expect(rewriteFormula("=SUM(Data!A1:B3,Data!Rate,Other!B2)", place(0, 5, "Other"), place(0, 5, "Other"), { kind: "deleteSheet", sheet: "Data" })).toBe("=SUM(#REF!,#REF!,'Other'!B2)");
    expect(rewriteFormula('=SUM(T[Cost],T2[Cost])+"T[Cost]"', place(), place(), { kind: "renameTable", from: "T", to: "Sales" })).toBe('=(SUM(Sales[Cost],T2[Cost])+"T[Cost]")');
    expect(rewriteFormula("=SUM(T[Cost],[@Cost],T2[Cost])", place(), place(), { kind: "renameColumn", table: "T", from: "Cost", to: "Cost]#" }, { table: "T" })).toBe("=SUM(T[Cost']'#],[[#This Row],[Cost']'#]],T2[Cost])");
    expect(() => rewriteFormula("=SUM(T)", place(), place(), { kind: "renameTable", from: "T", to: "Sales" })).toThrow();
    expect(rewriteFormula("=SUM(T)", place(), place(), { kind: "renameTable", from: "T", to: "Sales" }, { isTableName: () => true })).toBe("=SUM(Sales)");
    expect(rewriteFormula("=SUM(T)", place(), place(), { kind: "renameTable", from: "T", to: "Sales" }, { isTableName: () => false })).toBe("=SUM(T)");
  });
  it("rejects invalid mapping coordinates and injected identifiers before returning source", () => {
    expect(() => rewriteFormula("=A1", place(), place(), { kind: "axis", axis: "row", sheet: "Data", action: "insert", index: 0, count: 0 })).toThrow();
    expect(() => rewriteFormula("=T[Cost]", place(), place(), { kind: "renameTable", from: "T", to: "A1+Other" })).toThrow();
    expect(() => rewriteFormula("=T[Cost]", place(), place(), { kind: "renameColumn", table: "T", from: "Cost", to: "\0" })).toThrow();
  });
  it("preserves the exact source of unaffected formulas for shared immutable chunks", () => {
    expect(rewriteFormula('=sum( Other!A1 , 2 )', place(), place(), { kind: "axis", sheet: "Data", axis: "row", action: "insert", index: 3, count: 1 })).toBe('=sum( Other!A1 , 2 )');
  });
  it("matches independent point mapping for overlapping cut rectangles and deletion boundaries", () => {
    const origin = place(), source = "=SUM(A1:D4)", original = Array.from({ length: 16 }, (_, index) => ({ row: Math.floor(index / 4), column: index % 4 }));
    const refs = (source: string): string[] => {
      const result = new Set<string>();
      const visit = (node: FormulaExpression): void => {
        const point = (cell: Extract<FormulaExpression, { kind: "cell" }>) => ({ row: cell.row + (cell.absoluteRow ? 0 : origin.row), column: cell.column + (cell.absoluteColumn ? 0 : origin.column) });
        if (node.kind === "cell") { const p = point(node); result.add(p.row + ":" + p.column); }
        else if (node.kind === "range") {
          const a = point(node.from), b = point(node.to);
          for (let row = Math.min(a.row, b.row); row <= Math.max(a.row, b.row); row++) for (let col = Math.min(a.column, b.column); col <= Math.max(a.column, b.column); col++) result.add(row + ":" + col);
        } else if (node.kind === "binary" && node.operator === ",") { visit(node.left); visit(node.right); }
        else if (node.kind === "call") node.arguments.forEach(argument => { if (argument) visit(argument); });
        else if (node.kind !== "error") throw new Error("Unexpected reference expression");
      };
      visit(parseFormula(source, origin)); return [...result].sort();
    };
    for (let row = 0; row < 5; row++) for (let column = 0; column < 5; column++) {
      const operation = { kind: "move", sheet: "Data", range: { from: { row: 1, column: 1 }, to: { row: 2, column: 2 } }, target: place(row, column) } as const;
      const expected = new Set(original.map(point => point.row >= 1 && point.row <= 2 && point.column >= 1 && point.column <= 2 ? (point.row + row - 1) + ":" + (point.column + column - 1) : point.row + ":" + point.column));
      expect(refs(rewriteFormula(source, origin, origin, operation))).toEqual([...expected].sort());
    }
    for (let index = 0; index < 5; index++) for (let count = 1; count < 5; count++) {
      const operation = { kind: "axis", action: "delete", axis: "row", sheet: "Data", index, count } as const;
      const expected = original.filter(point => point.row < index || point.row >= index + count).map(point => (point.row >= index + count ? point.row - count : point.row) + ":" + point.column).sort();
      expect(refs(rewriteFormula(source, origin, origin, operation))).toEqual(expected);
    }
  });
});
