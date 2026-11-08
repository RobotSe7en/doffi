import { expect, test } from "vitest";
import { readWorkbookChartVectorDraft, workbookChartVectorDraft } from "../src/workbook/chart-vector-controls.js";

test("chart vector drafts retain sparse native literals, errors and individual number formats", () => {
  const vector = { kind: "literal" as const, valueType: "number" as const, count: 1002, numberFormat: "0.00", points: [
    { index: 1, value: 7, numberFormat: "" }, { index: 102, value: "#GETTING_DATA" }, { index: 1001, value: null, numberFormat: "0%" },
  ] };
  const draft = workbookChartVectorDraft(vector);
  expect(readWorkbookChartVectorDraft(draft)).toEqual(vector);
  draft.points[1]!.value = "-4.5";
  expect(readWorkbookChartVectorDraft(draft)).toEqual({ ...vector, points: [vector.points[0], { index: 102, value: -4.5 }, vector.points[2]] });
  expect(vector.points[1]!.value).toBe("#GETTING_DATA");
});

test("reference inputs never copy imported cache points into literal data", () => {
  const draft = workbookChartVectorDraft({ expression: "='表 1'!$A$1:$A$3", valueType: "number", cache: { state: "stale", count: 3, points: [{ index: 0, value: 99 }] } });
  expect(readWorkbookChartVectorDraft(draft)).toEqual({ expression: "='表 1'!$A$1:$A$3", valueType: "number" });
  draft.mode = "literal";
  expect(readWorkbookChartVectorDraft(draft)).toEqual({ kind: "literal", valueType: "number", count: 0, points: [] });
  draft.expression = "=AnotherName"; draft.mode = "reference";
  expect(readWorkbookChartVectorDraft(draft)).toEqual({ expression: "=AnotherName", valueType: "number" });
});

test("literal edits validate exact positions, length and finite values before submission", () => {
  const draft = workbookChartVectorDraft({ kind: "literal", valueType: "number", count: 2, points: [{ index: 1, value: 3 }] });
  draft.points.push({ position: "2", value: "4", numberFormat: "" }); expect(() => readWorkbookChartVectorDraft(draft)).toThrow(/唯一/);
  draft.points[1]!.position = "3"; expect(() => readWorkbookChartVectorDraft(draft)).toThrow(/总长度/);
  draft.count = "3"; draft.points[1]!.value = "Infinity"; expect(() => readWorkbookChartVectorDraft(draft)).toThrow(/有限/);
  draft.points[1]!.value = "1.5"; draft.points[1]!.position = "0"; expect(() => readWorkbookChartVectorDraft(draft)).toThrow(/从 1/);
  draft.points[1]!.position = "1"; expect(readWorkbookChartVectorDraft(draft)).toMatchObject({ count: 3, points: [{ index: 0, value: 1.5 }, { index: 1, value: 3 }] });
});

test("text literal drafts preserve CRLF, variation selectors, empty and omitted native points", () => {
  const vector = { kind: "literal" as const, valueType: "text" as const, count: 4, points: [{ index: 0, value: "中文禰󠄀😀\r\n" }, { index: 2, value: "" }] };
  expect(readWorkbookChartVectorDraft(workbookChartVectorDraft(vector))).toEqual(vector);
});
