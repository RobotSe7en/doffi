import { expect, test } from "vitest";
import { parseWorkbookFormula, parseWorkbookR1C1Formula, printWorkbookFormula, printWorkbookR1C1Formula } from "../src/index.js";

const at = { row: 255, column: 31 };
test.each([
  ["=RC", "=AF256"], ["=R1C1", "=$A$1"], ["=RC1", "=$A256"], ["=R1C", "=AF$1"],
  ["=R[-1]C[1]", "=AG255"], ["=SUM(R[-2]C:R[-1]C)", "=SUM(AF254:AF255)"],
  ["=SUM(R1:R3)", "=SUM($1:$3)"], ["=SUM(C[-1]:C[1])", "=SUM(AE:AG)"],
  ["='中文 R1C1'!RC[-1]", "='中文 R1C1'!AE256"], ["=!R1C1", "=!$A$1"],
  ['=IF(TRUE,"R[-1]C[2]",RC)', '=IF(TRUE,"R[-1]C[2]",AF256)'],
  ['=SUM(Table1[RC],RC)', '=SUM(Table1[RC],AF256)'],
  ['=R2D2+C3PO', '=(R2D2+C3PO)'],
  ["=SUM((RC,R1C1) R:R)", "=SUM(((AF256,$A$1) 256:256))"],
])("R1C1 %s produces independently specified A1 %s through the shared AST", (source, expected) => {
  const ast = parseWorkbookR1C1Formula(source, at); expect(printWorkbookFormula(ast, at)).toBe(expected);
  expect(ast).toEqual(parseWorkbookFormula(expected, at));
  expect(parseWorkbookR1C1Formula(printWorkbookR1C1Formula(ast, at), at)).toEqual(ast);
});
test("R1C1 preserves relative offsets at boundaries until a destination is known", () => {
  const ast = parseWorkbookR1C1Formula("=RC[-1]", { row: 0, column: 0 });
  expect(printWorkbookFormula(ast, { row: 0, column: 0 })).toBe("=#REF!");
  expect(printWorkbookFormula(ast, { row: 0, column: 1 })).toBe("=A1");
  expect(printWorkbookR1C1Formula(ast, { row: 0, column: 0 })).toBe("=RC[-1]");
});
test.each(["=R0C1", "=R1048577C1", "=R1C16385", "=R[-x]C", "=RC[", "=R[1.5]C", "=R999999999999999999999C", "=SUM(R1C1;R2C2)", "=[1]Sheet!RC"])("invalid or unsupported R1C1 fails explicitly: %s", source => {
  expect(() => parseWorkbookR1C1Formula(source, at)).toThrow();
});
