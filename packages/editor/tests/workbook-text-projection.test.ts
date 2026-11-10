import { expect, test } from "vitest";
import { reconcileWorkbookInputText, workbookInputTextView } from "../src/workbook/text-projection.js";
test("native LF projection retains untouched CR and CRLF while mapping real edits", () => {
  const raw = "甲\r\n乙\r丙\n丁😀禰󠄀", visible = "甲\n乙\n丙\n丁😀禰󠄀";
  expect(workbookInputTextView(raw)).toBe(visible); expect(reconcileWorkbookInputText(raw, visible)).toBe(raw);
  expect(reconcileWorkbookInputText(raw, visible + "尾")).toBe(raw + "尾");
  expect(reconcileWorkbookInputText(raw, visible.replace("乙", "混排"))).toBe(raw.replace("乙", "混排"));
  expect(reconcileWorkbookInputText(raw, visible.replace("乙\n丙", "替换"))).toBe(raw.replace("乙\r丙", "替换"));
});
test("projection edits retain complete UTF-16 and variation sequences around changed characters", () => {
  expect(reconcileWorkbookInputText("A\r😀禰󠄀Z", "A\n😀X禰󠄀Z")).toBe("A\r😀X禰󠄀Z");
  expect(reconcileWorkbookInputText("A\r禰󠄀Z", "A\n祢󠄀Z")).toBe("A\r祢󠄀Z");
  expect(reconcileWorkbookInputText("A\r禰󠄀Z", "A\n禰Z")).toBe("A\r禰Z");
  expect(reconcileWorkbookInputText("A\r\nB", "AB")).toBe("AB");
});
