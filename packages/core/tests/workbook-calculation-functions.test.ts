import { describe, expect, test } from "vitest";
import { createWorkbookScalarEvaluator } from "../src/index.js";
import type { WorkbookValue } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const origin = { sheetId: "sheet-1", row: 10, column: 10 };
const number = (value: number): WorkbookValue => ({ type: "number", value });
const text = (value: string): WorkbookValue => ({ type: "text", value });
const error = (value: Extract<WorkbookValue, { type: "error" }>["value"]): WorkbookValue => ({ type: "error", value });

describe("selected first-party workbook calculation functions", () => {
  test("evaluates independent text, conditional and historical calendar examples", () => {
    const evaluator = createWorkbookScalarEvaluator(scalarWorkbook([]), {}, { locale: "en-US" });
    const cases: [string, WorkbookValue][] = [
      ['=IFS(FALSE,1,TRUE,2,FUTUREFUNC(),3)', number(2)], ['=IFS(FALSE,1)', error("#N/A")],
      ['=LEFT("中文ABC",2)', text("中文")], ['=RIGHT("中文ABC",2)', text("BC")], ['=MID("中文ABC",2,3)', text("文AB")],
      ['=RIGHT("abc",0)', text("")], ['=LEFT("abc",-1)', error("#VALUE!")], ['=TRIM("  a   b  ")', text("a b")],
      ['=UPPER("abc中国")', text("ABC中国")], ['=LOWER("AbC中国")', text("abc中国")], ['=EXACT("ABC","abc")', { type: "boolean", value: false }],
      ['=CONCAT("a",1,TRUE)', text("a1TRUE")], ['=CONCATENATE("a","b",3)', text("ab3")], ['=TEXTJOIN("-",TRUE,"a","","b")', text("a-b")],
      ['=TEXTJOIN("-",FALSE,"a","","b")', text("a--b")], ['=SUBSTITUTE("a-a-a","a","x",2)', text("a-x-a")],
      ['=SUBSTITUTE("a-a-a","a","x")', text("x-x-x")], ['=REPLACE("abcdef",2,3,"X")', text("aXef")],
      ['=FIND("B","aBcB",3)', number(4)], ['=SEARCH("b?","aBcB")', number(2)], ['=SEARCH("~*","a*b")', number(2)],
      ['=TEXT(1234.5,"#,##0.00")', text("1,234.50")], ['=VALUE("$1,234.50")', number(1234.5)], ['=VALUE("15%")', number(0.15)],
      ['=DATE(2024,2,29)', number(45351)], ['=DATE(1900,2,29)', number(60)], ['=DATE(1900,3,0)', number(60)],
      ['=YEAR(60)', number(1900)], ['=MONTH(60)', number(2)], ['=DAY(60)', number(29)], ['=DAY(0)', number(0)],
      ['=TIME(12,30,0)', number(12.5 / 24)], ['=HOUR(0.5)', number(12)], ['=MINUTE(0.5)', number(0)], ['=SECOND(0.5)', number(0)],
      ['=HOUR(0.9999999)', number(23)], ['=MINUTE(0.9999999)', number(59)], ['=SECOND(0.9999999)', number(59)],
      ['=VALUE("2024-02-29")', number(45351)], ['=DAYS("2024-03-01","2024-02-28")', number(2)],
      ['=HOUR("12:30:05")', number(12)], ['=MINUTE("12:30:05")', number(30)], ['=SECOND("12:30:05")', number(5)],
      ['=DAYS(DATE(2024,3,1),DATE(2024,2,28))', number(2)], ['=EDATE(DATE(2024,1,31),1)', number(45351)],
      ['=EOMONTH(DATE(2024,1,1),1)', number(45351)], ['=WEEKDAY(DATE(2024,1,1),2)', number(1)],
      ['=WEEKDAY(1)', number(1)], ['=DATE(-1,1,1)', error("#NUM!")],
    ];
    for (const [expression, value] of cases) expect(evaluator.evaluate(expression, origin), expression).toMatchObject({ kind: "value", value });
    evaluator.dispose();
  });
  test("captures explicit host time, keeps a consistent batch and never invents current time", () => {
    const model = scalarWorkbook([]), evaluator = createWorkbookScalarEvaluator(model), now = { unixMilliseconds: Date.UTC(2024, 1, 29, 4, 30), utcOffsetMinutes: 480 };
    expect(evaluator.evaluate("=TODAY()", origin)).toMatchObject({ kind: "unavailable", reason: "calculation-required" }); evaluator.dispose();
    const timed = createWorkbookScalarEvaluator(model, {}, { now }); now.unixMilliseconds = 0;
    expect(timed.evaluate("=TODAY()", origin)).toMatchObject({ kind: "value", value: number(45351) });
    const result = timed.evaluate("=NOW()", origin); expect(result.kind).toBe("value"); if (result.kind === "value" && result.value.type === "number") expect(result.value.value).toBeCloseTo(45351 + 12.5 / 24, 10);
    timed.dispose();
  });
  test("binds criteria to aligned sparse references, including implicit blank counts", () => {
    const model = scalarWorkbook([
      { row: 0, column: 0, value: text("apples") }, { row: 1, column: 0, value: text("oranges") }, { row: 2, column: 0, value: text("peaches") }, { row: 3, column: 0, value: text("APPLES") },
      { row: 0, column: 1, value: number(32) }, { row: 1, column: 1, value: number(54) }, { row: 2, column: 1, value: number(75) }, { row: 3, column: 1, value: number(86) },
    ]), evaluator = createWorkbookScalarEvaluator(model);
    const cases: [string, WorkbookValue][] = [
      ['=COUNTIF(A1:A4,"apples")', number(2)], ['=COUNTIF(A1:A4,"*es")', number(4)], ['=COUNTIF(B1:B4,">55")', number(2)],
      ['=SUMIF(A1:A4,"apples",B1:B4)', number(118)], ['=SUMIFS(B1:B4,A1:A4,"*es",B1:B4,">55")', number(161)],
      ['=COUNTIFS(A1:A4,"apples",B1:B4,">50")', number(1)], ['=AVERAGEIF(A1:A4,"apples",B1:B4)', number(59)],
      ['=AVERAGEIFS(B1:B4,A1:A4,"apples",B1:B4,">50")', number(86)], ['=SUMIF(A1:A4,"apples",B1)', number(118)],
      ['=COUNTIF(C:C,"")', number(1_048_576)], ['=COUNTIFS(A:A,"",B:B,"")', number(1_048_572)],
      ['=SUMIFS(B1:B4,A1:A3,"apples")', error("#VALUE!")], ['=AVERAGEIF(A1:A4,"none",B1:B4)', error("#DIV/0!")],
    ];
    for (const [expression, value] of cases) expect(evaluator.evaluate(expression, origin), expression).toMatchObject({ kind: "value", value });
    evaluator.dispose();
  });
  test("performs exact, wildcard and approximate lookup and reference-returning INDEX", () => {
    const evaluator = createWorkbookScalarEvaluator(scalarWorkbook([
      ...[10, 20, 30].map((value, row) => ({ row, column: 0, value: number(value) })),
      ...["a", "b", "c"].map((value, row) => ({ row, column: 1, value: text(value) })),
      ...[90, 70, 50].map((value, row) => ({ row, column: 2, value: number(value) })),
    ]));
    const cases: [string, WorkbookValue][] = [
      ['=MATCH(20,A1:A3,0)', number(2)], ['=MATCH(25,A1:A3)', number(2)], ['=MATCH(60,C1:C3,-1)', number(2)], ['=MATCH("B*",B1:B3,0)', number(2)],
      ['=VLOOKUP(20,A1:B3,2,FALSE)', text("b")], ['=VLOOKUP(25,A1:B3,2)', text("b")], ['=VLOOKUP(5,A1:B3,2)', error("#N/A")],
      ['=HLOOKUP(90,A1:C3,3,FALSE)', number(50)], ['=INDEX(A1:C3,2,3)', number(70)], ['=SUM(INDEX(A1:C3,0,1))', number(60)],
      ['=ROW(INDEX(A1:C3,2,1))', number(2)], ['=INDEX(A1:C3,4,1)', error("#REF!")],
    ];
    for (const [expression, value] of cases) expect(evaluator.evaluate(expression, origin), expression).toMatchObject({ kind: "value", value }); evaluator.dispose();
  });
  test("ignores nested subtotal sources, distinguishes hidden rows, and bounds wildcard work", () => {
    const evaluator = createWorkbookScalarEvaluator(scalarWorkbook([
      { row: 0, column: 0, value: number(10) }, { row: 1, column: 0, value: number(20) }, { row: 2, column: 0, value: number(999), formula: "=SUBTOTAL(9,A1:A2)" },
    ], metadata => { metadata.workbook.sheets[0]!.rowDimensions = [{ from: 1, to: 1, hidden: true }]; }));
    expect(evaluator.evaluate("=SUBTOTAL(9,A1:A3)", origin)).toMatchObject({ kind: "value", value: number(30) });
    expect(evaluator.evaluate("=SUBTOTAL(109,A1:A3)", origin)).toMatchObject({ kind: "value", value: number(10) }); evaluator.dispose();
    const bounded = createWorkbookScalarEvaluator(scalarWorkbook([]), { maxSteps: 100 });
    expect(() => bounded.evaluate('=SEARCH("' + "*a".repeat(60) + '","' + "a".repeat(100) + '")', origin)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); bounded.dispose();
  });
});
