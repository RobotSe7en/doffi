import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, SourceStyleResolver, DEFAULT_DOCUMENT_READ_LIMITS } from "@doffi/core";
import type { WorkbookSession } from "@doffi/core";
import { createWorkbookHtmlPastePlan } from "../src/workbook/html-paste.js";

const sessions: WorkbookSession[] = [];
function blank() { const s = createDocumentSession(openDocument(createWorkbook({ documentId: "html-clipboard" }), { kind: "workbook" })); sessions.push(s); return s; }
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const origin = { row: 255, column: 31 }, sheetId = "sheet-1";
test("HTML tables paste rich Unicode, cell appearance, number formats and merges through one private candidate across both chunk boundaries", () => {
  const s = blank(), before = s.snapshot();
  const plan = createWorkbookHtmlPastePlan(before, sheetId, origin, '<style>.amount{background-color:#ffeeaa;mso-number-format:"0.00"}</style><table><tr><td><b>中文</b>😀禰󠄀</td><td class="amount" x:num="12.5">12.50</td></tr><tr><td colspan="2">合并</td></tr></table>');
  expect(plan.range).toEqual({ from: origin, to: { row: 256, column: 32 } }); expect(plan.diagnostics).toEqual([]);
  const candidate = s.prepare(plan.transaction); expect(s.snapshot()).toEqual(before); s.commit(candidate);
  const first = s.getCell(sheetId, 255, 31)!; expect(first.value).toEqual({ type: "text", value: "中文😀禰󠄀" }); expect(first.html).toContain("<span");
  expect(s.getCell(sheetId, 255, 32)?.value).toEqual({ type: "number", value: 12.5 });
  expect(s.snapshot().metadata.workbook.numberFormats).toContainEqual(expect.objectContaining({ code: "0.00" }));
  expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([{ from: { row: 256, column: 31 }, to: { row: 256, column: 32 } }]);
  expect(s.history().undoSteps).toBe(1); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  const resolver = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS);
  expect(resolver.get({ tag: "td", attrs: { class: s.getCell(sheetId, 255, 32)!.cssClass! }, parent: undefined })["background-color"]).toBe("#ffeeaa");
  s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().styles).toEqual(before.styles); expect(s.snapshot().metadata).toEqual(before.metadata);
});
test("copied producer value metadata preserves empty text, long IDs, booleans and errors independently of displayed number strings", () => {
  const s = blank();
  const plan = createWorkbookHtmlPastePlan(s.snapshot(), sheetId, { row: 0, column: 0 }, '<table><tr><td data-sheets-value=\'{"1":2,"2":""}\'><br></td><td data-sheets-value=\'{"1":2,"2":"001234567890123456789"}\'>001234567890123456789</td><td data-sheets-value=\'{"1":4,"4":1}\' sdval="1" data-sheets-formula="=TRUE()">TRUE</td><td data-sheets-formula="=#DIV/0!" sdval="0">#DIV/0!</td></tr></table>');
  s.commit(s.prepare(plan.transaction));
  expect([0, 1, 2, 3].map(column => s.getCell(sheetId, 0, column)?.value)).toEqual([{ type: "text", value: "" }, { type: "text", value: "001234567890123456789" }, { type: "boolean", value: true }, { type: "error", value: "#DIV/0!" }]);
});
test("external paragraphs and literal CR entities preserve exact text without appending a synthetic last line", () => {
  const s = blank(), plan = createWorkbookHtmlPastePlan(s.snapshot(), sheetId, { row: 0, column: 0 }, '<p>甲&#13;乙</p><p><i>😀禰󠄀</i></p>');
  s.commit(s.prepare(plan.transaction)); expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "text", value: "甲\r乙\n😀禰󠄀" });
});
test.each(['<script>alert(1)</script><table><tr><td>x</td></tr></table>', '<table><tr><td onclick="x()">x</td></tr></table>', '<img src="https://example.com/private">', '<span style="color:var(--private)">x</span>', '<span>&#xD800;</span>', '<table><tr><td rowspan="2">x</td></tr></table>', '<table><tr><td><table><tr><td>x</td></tr></table></td></tr></table>'])("unsafe or invalid clipboard HTML fails before publication: %s", html => {
  const s = blank(), before = s.serialize(); expect(() => createWorkbookHtmlPastePlan(s.snapshot(), sheetId, origin, html)).toThrow(); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
test("HTML input and span budgets fail before expanding sparse rectangles or changing destination merges", () => {
  const s = blank();
  expect(() => createWorkbookHtmlPastePlan(s.snapshot(), sheetId, origin, '<table><tr><td colspan="16384">x</td></tr></table>', { maxCells: 10 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => createWorkbookHtmlPastePlan(s.snapshot(), sheetId, origin, '<p>large</p>', { maxCodeUnits: 5 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => createWorkbookHtmlPastePlan(s.snapshot(), sheetId, { row: 1_048_575, column: 16_383 }, '<table><tr><td>x</td><td>y</td></tr></table>')).toThrow();
});
test("full numeric font wrappers preserve font lists and precision instead of turning numbers into rich text", () => {
  const s = blank(), plan = createWorkbookHtmlPastePlan(s.snapshot(), sheetId, { row: 0, column: 0 }, '<table><tr><td sdval="12345.678" data-sheets-value=\'{"1":3,"3":12345}\'><b><font face="Arial,宋体" color="#cc2200">12,345.68</font></b></td></tr></table>');
  s.commit(s.prepare(plan.transaction)); const cell = s.getCell(sheetId, 0, 0)!;
  expect(cell.value).toEqual({ type: "number", value: 12345.678 }); expect(cell.html).not.toContain("<span");
  const resolver = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS), style = resolver.get({ tag: "td", attrs: { class: cell.cssClass! }, parent: undefined });
  expect(style).toMatchObject({ "font-weight": "bold", color: "#cc2200", "font-family": "Arial,宋体" });
});
test("shared whole-cell font wrappers across many rows remain grouped in one bounded transaction", () => {
  const s = blank(), plan = createWorkbookHtmlPastePlan(s.snapshot(), sheetId, { row: 0, column: 0 }, '<table>' + '<tr><td><b>共享字体</b></td></tr>'.repeat(1100) + '</table>');
  expect(plan.transaction.operations).toHaveLength(3); expect(plan.transaction.operations.some(operation => operation.op === "formatWorkbookText")).toBe(false);
  s.commit(s.prepare(plan.transaction)); expect(s.getCell(sheetId, 1099, 0)?.value).toEqual({ type: "text", value: "共享字体" }); expect(s.history().undoSteps).toBe(1);
});
test("complete clipboard HTML retains body fonts, direction and native vertical alignment, and never drops outside prose", () => {
  const s = blank(), plan = createWorkbookHtmlPastePlan(s.snapshot(), sheetId, { row: 0, column: 0 }, '<html><head><title>clipboard</title></head><body dir="rtl" style="font-family:Arial;color:#112233"><table><tr><td valign="middle">完整页面</td></tr></table></body></html>');
  s.commit(s.prepare(plan.transaction)); const cell = s.getCell(sheetId, 0, 0)!;
  const resolver = new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS);
  expect(resolver.get({ tag: "td", attrs: { class: cell.cssClass! }, parent: undefined })).toMatchObject({ "font-family": "Arial", color: "#112233", "--office-cell-reading-order": "rtl", "--office-cell-vertical": "center" });
  const before = s.serialize(); expect(() => createWorkbookHtmlPastePlan(s.snapshot(), sheetId, origin, '<p>must not disappear</p><table><tr><td>cell</td></tr></table>')).toThrow(); expect(s.serialize()).toEqual(before);
});
test("explicit native A1 and R1C1 dialects share destination-relative formula ASTs without trusting the display cache", async () => {
  const s = blank(), plan = createWorkbookHtmlPastePlan(s.snapshot(), sheetId, origin,
    '<table><tr><td x:num="3">3</td><td x:fmla="=A1+$B$2">9999</td><td data-sheets-value=\'{"1":3,"3":9999}\' data-sheets-formula="=RC[-2]*2">9999</td></tr></table>');
  s.commit(s.prepare(plan.transaction));
  expect(s.getCell(sheetId, 255, 32)?.formula?.expression).toBe("=(AF256+$B$2)");
  expect(s.getCell(sheetId, 255, 33)?.formula?.expression).toBe("=(AF256*2)");
  await s.calculate(); expect(s.getCalculatedCell(sheetId, 255, 32)).toMatchObject({ state: "valid", value: { type: "number", value: 3 } });
  expect(s.getCalculatedCell(sheetId, 255, 33)).toMatchObject({ state: "valid", value: { type: "number", value: 6 } });
});
test("literal formula-like text is typed before parsing its display while conflicting native dialects fail atomically", () => {
  const s = blank(), plan = createWorkbookHtmlPastePlan(s.snapshot(), sheetId, origin,
    '<table><tr><td x:str="">=unclosed(</td><td data-sheets-value=\'{"1":2,"2":"=alsoUnclosed("}\'>=alsoUnclosed(</td></tr></table>');
  s.commit(s.prepare(plan.transaction)); expect(s.getCell(sheetId, 255, 31)?.value).toEqual({ type: "text", value: "=unclosed(" });
  expect(s.getCell(sheetId, 255, 32)?.value).toEqual({ type: "text", value: "=alsoUnclosed(" });
  const before = s.serialize();
  expect(() => createWorkbookHtmlPastePlan(s.snapshot(), sheetId, origin, '<table><tr><td x:fmla="=A1+1" data-sheets-formula="=RC+2">0</td></tr></table>')).toThrow();
  expect(s.serialize()).toEqual(before);
});
test("CSS number-format strings decode once and retain quoted native positive, negative and zero literals", () => {
  const s = blank(), plan = createWorkbookHtmlPastePlan(s.snapshot(), sheetId, origin,
    '<table><tr><td x:num="12.5" style="mso-number-format:\'&quot;positive&quot;;&quot;negative&quot;;&quot;zero&quot;\'">positive</td></tr></table>');
  expect(plan.transaction.operations).toContainEqual(expect.objectContaining({ op: "setWorkbookNumberFormat", format: expect.objectContaining({ code: '"positive";"negative";"zero"' }) }));
  s.commit(s.prepare(plan.transaction)); expect(s.getCell(sheetId, 255, 31)?.value).toEqual({ type: "number", value: 12.5 });
});
