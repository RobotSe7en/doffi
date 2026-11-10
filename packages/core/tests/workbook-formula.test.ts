import { describe, expect, it } from "vitest";
import { createFormulaTemplatePool, parseFormula } from "../src/workbook/formula.js";
import { printFormula } from "../src/workbook/formula-print.js";
const origin = { row: 4, column: 2 };
describe("first-party formula parser foundation", () => {
  it("preserves constants, missing arguments, escaping and sheet-relative references", () => {
    const ast = parseFormula('=IF(TRUE,"a""b",\'Q1\'\'收入\'!$B5)+SUM(A1:$B$2,,#N/A)', origin);
    expect(ast).toMatchObject({ kind: "binary", operator: "+", left: { kind: "call", name: "IF", arguments: [
      { kind: "boolean", value: true }, { kind: "string", value: 'a"b' },
      { kind: "cell", sheet: "Q1'收入", row: 0, column: 1, absoluteRow: false, absoluteColumn: true },
    ] }, right: { kind: "call", name: "SUM", arguments: [{ kind: "range", from: { row: -4, column: -2 }, to: { row: 1, column: 1 } }, null, { kind: "error", value: "#N/A" }] } });
    expect(Object.isFrozen(ast)).toBe(true);
  });
  it("uses Excel precedence, left associativity and explicit range qualification", () => {
    expect(parseFormula('=-2^2^3', origin)).toMatchObject({ kind: "binary", operator: "^", left: { kind: "binary", operator: "^", left: { kind: "unary", operator: "-" }, right: { value: 2 } }, right: { value: 3 } });
    expect(parseFormula('=1+2*3&"x"=TRUE', origin)).toMatchObject({ operator: "=", left: { operator: "&", left: { operator: "+", right: { operator: "*" } } } });
    expect(parseFormula("=SUM('Budget 2026'!A1:B3)", origin)).toMatchObject({ arguments: [{ from: { sheet: "Budget 2026" }, to: { sheet: "Budget 2026" } }] });
    expect(parseFormula("=LOG10(100)", origin)).toMatchObject({ kind: "call", name: "LOG10" });
  });
  it.each([",", ")", "(", "!", "%", ":", "+", " "])("preserves punctuation strings as arguments: %s", value => {
    const source = `=CONCAT("${value}","tail")`, ast = parseFormula(source, origin);
    expect(ast).toEqual({ kind: "call", name: "CONCAT", arguments: [{ kind: "string", value }, { kind: "string", value: "tail" }] });
    expect(parseFormula(printFormula(ast, origin), origin)).toEqual(ast);
    expect(() => parseFormula(`=SUM(1"${value}"2)`, origin)).toThrow();
  });
  it("preserves bang references and names separately from ordinary unqualified tokens", () => {
    for (const source of ["=!A1", "=!$A$1", "=!Scoped", "=SUM(!A1:B2)", "=SUM(!1:2)", "=SUM(!A:B)", "=!Scoped !$A$1"]) {
      const ast = parseFormula(source, origin);
      expect(parseFormula(printFormula(ast, origin), origin)).toEqual(ast);
    }
    expect(parseFormula("=!Scoped", origin)).toEqual({ kind: "name", name: "Scoped", sheet: "" });
    for (const source of ["=''!A1", "=[0]''!A1", "=!", "=!!Scoped"]) expect(() => parseFormula(source, origin)).toThrow();
  });
  it("interns copied formulas without retaining one AST per cell", () => {
    const pool = createFormulaTemplatePool();
    for (let row = 0; row < 10_000; row++) {
      expect(pool.intern(`=SUM(C${row + 1}:R${row + 1})`, { row, column: 18 })).toBe(0);
      expect(pool.intern(`=A${row + 1}+S${row + 1}`, { row, column: 19 })).toBe(1);
    }
    expect(pool.size).toBe(2); expect(pool.keyCodeUnits).toBeLessThan(1000);
    expect(pool.get(0)).toMatchObject({ arguments: [{ from: { row: 0, column: -16 }, to: { row: 0, column: -1 } }] });
    expect(() => createFormulaTemplatePool(1).intern('=1+2', origin)).not.toThrow();
    const small = createFormulaTemplatePool(1); small.intern('=1', origin);
    expect(() => small.intern('=2', origin)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(small.size).toBe(1);
  });
  it.each(['=', '=1 2', '=[book]Sheet!A1', '=Sheet1:Sheet3!A1', '=A1#', '=@A1', '={1,2}', '=SUM(', '=SUM(1;2)', '=\'no end', '="no end', '=A1:Other!B2', '=$XFE$1', '="\0"'])('rejects unimplemented or malformed syntax without changing its meaning: %s', source => {
    expect(() => parseFormula(source, origin)).toThrow();
  });
  it("rejects both nested and left-associative oversized trees with bounded errors", () => {
    for (const source of ['=' + '('.repeat(200) + '1' + ')'.repeat(200), '=' + Array(1000).fill('1').join('+'), '=1' + '%'.repeat(200)]) {
      expect(() => parseFormula(source, origin)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    }
  });
});
