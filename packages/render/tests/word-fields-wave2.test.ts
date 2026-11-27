import { expect, test } from "vitest";
import {
  applyWordFieldSwitches,
  evaluateIfOrCompare,
  evaluateSymbolField,
  formatIndexField,
  formatNumericPicture,
  formatToaField,
  parseTaInstruction,
  parseXeInstruction,
  tokenizeWordFieldArgs,
} from "../src/word-fields-eval.js";

test("Word field switches format numbers, Chinese numerals, Roman numerals, case and ordinal text", () => {
  expect(formatNumericPicture(1234.5, "#,##0.00")).toBe("1,234.50");
  expect(formatNumericPicture(-42, "#,##0.00;(#,##0.00);0")).toBe("(42.00)");
  expect(formatNumericPicture(0, "#,##0;-#,##0;零")).toBe("零");

  expect(applyWordFieldSwitches("hello world", ' QUOTE "hello world" \\* Upper ')).toBe("HELLO WORLD");
  expect(applyWordFieldSwitches("hello world", ' QUOTE "hello world" \\* Caps ')).toBe("Hello World");
  expect(applyWordFieldSwitches("hello world", ' QUOTE "hello world" \\* FirstCap ')).toBe("Hello world");
  expect(applyWordFieldSwitches("42", " = 42 \\* ROMAN ")).toBe("XLII");
  expect(applyWordFieldSwitches("42", " = 42 \\* roman ")).toBe("xlii");
  expect(applyWordFieldSwitches("2026", " = 2026 \\* CHINESENUM3 ")).toBe("二千〇二十六");
  expect(applyWordFieldSwitches("12", " = 12 \\* CHINESENUM2 ")).toBe("壹拾贰");
  expect(applyWordFieldSwitches("23", " = 23 \\* CardText ")).toBe("twenty-three");
  expect(applyWordFieldSwitches("3", " = 3 \\* OrdText ")).toBe("third");
  expect(applyWordFieldSwitches("21", " = 21 \\* Ordinal ")).toBe("21st");
});

test("IF, COMPARE and SYMBOL fields evaluate Word conditions, wildcards and codepoints", () => {
  expect(tokenizeWordFieldArgs('IF "Alpha" = "Al*" "匹配" "否"')).toEqual([
    "IF",
    "Alpha",
    "=",
    "Al*",
    "匹配",
    "否",
  ]);
  expect(evaluateIfOrCompare(' IF 10 >= 5 "通过" "失败" ', false)).toBe("通过");
  expect(evaluateIfOrCompare(' IF 3 >= 5 "通过" "失败" ', false)).toBe("失败");
  expect(evaluateIfOrCompare(' IF "Report2026" = "Report*" "是" "否" ', false)).toBe("是");
  expect(evaluateIfOrCompare(" COMPARE 100 > 50 ", true)).toBe("1");
  expect(evaluateIfOrCompare(" COMPARE 10 > 50 ", true)).toBe("0");
  expect(evaluateSymbolField(" SYMBOL 945 \\u ")).toBe("α");
  expect(evaluateSymbolField(" SYMBOL 0x2605 \\u ")).toBe("★");
});

test("XE/INDEX parses multi-level entries, compresses page ranges and emits alpha headings", () => {
  const e1 = parseXeInstruction(' XE "Algorithm:Sorting" ');
  const e2 = parseXeInstruction(' XE "Algorithm:Sorting" ');
  const e3 = parseXeInstruction(' XE "Algorithm:Sorting" ');
  const e4 = parseXeInstruction(' XE "Beta" \\t "参见 Alpha" ');
  expect(e1).toMatchObject({ main: "Algorithm", sub: "Sorting" });
  expect(e4).toMatchObject({ main: "Beta", seeAlso: "参见 Alpha" });

  const rendered = formatIndexField(
    [
      { ...e1!, pageNumber: 1, pageLabel: "1" },
      { ...e2!, pageNumber: 2, pageLabel: "2" },
      { ...e3!, pageNumber: 3, pageLabel: "3" },
      { ...e4!, pageNumber: 5, pageLabel: "5" },
    ],
    ' INDEX \\h "A" \\c "2" ',
  );
  expect(rendered).toContain("A");
  expect(rendered).toContain("Algorithm");
  expect(rendered).toContain("Sorting, 1-3");
  expect(rendered).toContain("B");
  expect(rendered).toContain("Beta, 参见 Alpha");
});

test("TA/TOA groups legal authorities by category and collapses 5+ pages into passim", () => {
  const parsed = parseTaInstruction(' TA \\l "中华人民共和国民法典 第577条" \\s "民法典577" \\c 2 ');
  expect(parsed).toMatchObject({
    longCitation: "中华人民共和国民法典 第577条",
    shortCitation: "民法典577",
    category: 2,
  });

  const entries = [1, 2, 3, 4, 5].map((p) => ({
    ...parsed!,
    pageNumber: p,
    pageLabel: String(p),
  }));
  const toa = formatToaField(entries, ' TOA \\h \\c "2" \\p ');
  expect(toa).toContain("法规");
  expect(toa).toContain("中华人民共和国民法典 第577条");
  expect(toa).toContain("passim");
});
