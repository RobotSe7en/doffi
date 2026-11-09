import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, createWorkbookNumberFormatter, DEFAULT_DOCUMENT_READ_LIMITS, openDocument, SourceStyleResolver } from "@doffi/core";
import type { SourceStyleNode, WorkbookSession } from "@doffi/core";
import { parseFragment } from "parse5";
import type { DefaultTreeAdapterTypes } from "parse5";
import { connectWorkbookSession, createWorkbookClipboardData, createWorkbookHtmlPastePlan, serveWorkbookSession } from "../src/index.js";
import { WorkbookSystemClipboard } from "../src/workbook/system-clipboard.js";

const sessions: WorkbookSession[] = [], sheetId = "sheet-1";
function workbook() { const s = createDocumentSession(openDocument(createWorkbook({ documentId: "portable-clipboard" }), { kind: "workbook" })); sessions.push(s); return s; }
afterEach(() => { for (const s of sessions.splice(0)) s.dispose(); });
const range = { from: { row: 255, column: 31 }, to: { row: 257, column: 32 } };
function source() {
  const s = workbook();
  s.commit(s.prepare({ protocolVersion: 1, documentId: "portable-clipboard", baseRevision: 0, operations: [
    { op: "setWorkbookCells", sheetId, cells: [
      { row: 255, column: 31, source: { kind: "value", value: { type: "number", value: 12.5 } } },
      { row: 255, column: 32, source: { kind: "formula", expression: "=AF256*2" } },
      { row: 256, column: 31, source: { kind: "value", value: { type: "text", value: '中文😀禰󠄀\r\n"甲\t乙"' } } },
      { row: 256, column: 32, source: { kind: "value", value: { type: "text", value: "001234567890123456789" } } },
      { row: 257, column: 31, source: { kind: "value", value: { type: "text", value: "合并" } } },
    ] }, { op: "formatWorkbookText", sheetId, at: { row: 256, column: 31 }, from: 0, to: "中文😀禰󠄀".length, expectedText: "中文😀禰󠄀", properties: { "font-weight": "bold" } },
    { op: "mergeWorkbookCells", sheetId, range: { from: { row: 257, column: 31 }, to: { row: 257, column: 32 } }, coveredContent: "reject" },
  ] })); return s;
}
test("portable clipboard reads current formula values, exact rich text and complete merges across both chunk boundaries without editing source", async () => {
  const s = source(), before = s.serialize(), history = s.history(), calculation = await s.calculate();
  const data = await createWorkbookClipboardData(s, sheetId, range, { calculation });
  expect(data.text).toBe('12.5\t25\r\n"中文😀禰󠄀\r\n""甲\t乙"""\t001234567890123456789\r\n合并\t');
  expect(data.html).toContain('data-sheets-formula="=(RC[-1]*2)"'); expect(data.html).toContain('colspan="2"');
  expect(data.html).not.toMatch(/data-node-id|portable-clipboard|data-sheet-id|data-cache-state/);
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  const target = workbook(), plan = createWorkbookHtmlPastePlan(target.snapshot(), sheetId, { row: 0, column: 0 }, data.html);
  const candidate = target.prepare(plan.transaction); expect(target.history().undoSteps).toBe(0); target.commit(candidate);
  expect(target.getCell(sheetId, 0, 0)?.value).toEqual({ type: "number", value: 12.5 });
  expect(target.getCell(sheetId, 0, 1)?.formula?.expression).toBe("=(A1*2)"); await target.calculate();
  expect(target.getCalculatedCell(sheetId, 0, 1)).toMatchObject({ state: "valid", value: { type: "number", value: 25 } });
  expect(target.getCell(sheetId, 1, 0)?.value).toEqual({ type: "text", value: '中文😀禰󠄀\r\n"甲\t乙"' });
  expect(target.getCell(sheetId, 1, 1)?.value).toEqual({ type: "text", value: "001234567890123456789" });
  expect(target.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([{ from: { row: 2, column: 0 }, to: { row: 2, column: 1 } }]);
  expect(openDocument(target.serialize(), { kind: "workbook" }).snapshot()).toEqual(target.snapshot());
  target.undo(); expect(target.snapshot().manifest.chunks).toEqual([]);
});
test("clipboard output fails atomically on unavailable calculations, partial merges, budgets, cancellation and concurrent source changes", async () => {
  const s = source(), before = s.serialize();
  await expect(createWorkbookClipboardData(s, sheetId, range)).rejects.toMatchObject({ code: "UnsupportedFeature" });
  await expect(createWorkbookClipboardData(s, sheetId, { from: range.to, to: range.to })).rejects.toMatchObject({ code: "InvalidOperation" });
  await expect(createWorkbookClipboardData(s, sheetId, range, { maxCells: 2 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  await expect(createWorkbookClipboardData(s, sheetId, range, { maxCodeUnits: 10 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  await expect(createWorkbookClipboardData(s, sheetId, range, { signal: { aborted: true } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(s.serialize()).toEqual(before);
  const large = { from: { row: 10, column: 0 }, to: { row: 300, column: 1 } }, signal = { aborted: false };
  await expect(createWorkbookClipboardData(s, sheetId, large, { signal, yield: async () => { signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  await expect(createWorkbookClipboardData(s, sheetId, large, { yield: async () => { s.undo(); } })).rejects.toMatchObject({ code: "RevisionConflict" });
});
test("Worker clipboard reuses its retained current calculation and rejects foreign, released and stale views without transferring its model", async () => {
  const s = source(), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.serialize(), view = await client.view(), data = await client.readClipboard(sheetId, range, { view });
    expect(data.text).toContain("12.5\t25"); expect(Object.keys(data).sort()).toEqual(["html", "text"]); expect(s.serialize()).toEqual(before);
    await expect(client.readClipboard(sheetId, range, { view: { ...view } })).rejects.toMatchObject({ code: "InvalidOperation" });
    view.dispose(); await expect(client.readClipboard(sheetId, range, { view })).rejects.toMatchObject({ code: "InvalidOperation" });
    const old = await client.view(); await client.undo(); await expect(client.readClipboard(sheetId, range, { view: old })).rejects.toMatchObject({ code: "InvalidOperation" }); old.dispose();
    await expect(client.readClipboard(sheetId, range, { signal: AbortSignal.abort() })).rejects.toMatchObject({ code: "OperationCancelled" });
  } finally { client.dispose(); server.dispose(); }
});
test("native clipboard prefetch retains only the latest selection and never publishes a cancelled old read", async () => {
  const first = { documentId: "d", baseRevision: 1, anchor: { kind: "cells" as const, sheetId, ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] } };
  const second = { ...first, baseRevision: 2 }, jobs: { resolve(data: { text: string; html: string }): void; signal: AbortSignal }[] = [];
  const cache = new WorkbookSystemClipboard((_selection, signal) => new Promise(resolve => { jobs.push({ resolve, signal }); }));
  cache.prepare(first); await Promise.resolve(); cache.prepare(second); await Promise.resolve();
  expect(jobs[0]!.signal.aborted).toBe(true); jobs[0]!.resolve({ text: "old", html: "old" }); jobs[1]!.resolve({ text: "new", html: "new" }); await cache.settled();
  expect(cache.data(second)).toEqual({ text: "new", html: "new" }); expect(() => cache.data(first)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  cache.dispose(); expect(jobs[1]!.signal.aborted).toBe(true);
});
test("portable values retain blank versus empty text, booleans, errors, literal formulas, native date identities and formula types", async () => {
  const s = workbook(), values = [{ type: "blank" as const }, { type: "text" as const, value: "" }, { type: "boolean" as const, value: false }, { type: "error" as const, value: "#N/A" as const }, { type: "text" as const, value: "=1+1" }];
  s.commit(s.prepare({ protocolVersion: 1, documentId: "portable-clipboard", baseRevision: 0, operations: [
    { op: "setWorkbookNumberFormat", format: { formatId: "native-date", builtinId: 14 } },
    { op: "setWorkbookCells", sheetId, cells: [...values.map((value, column) => ({ row: 0, column, source: { kind: "value" as const, value } })),
      { row: 0, column: 5, source: { kind: "value", value: { type: "number", value: 60 } }, formatId: "native-date" },
      { row: 0, column: 6, source: { kind: "formula", expression: "=TRUE()" } }, { row: 0, column: 7, source: { kind: "formula", expression: "=#N/A" } },
    ] },
  ] }));
  const data = await createWorkbookClipboardData(s, sheetId, { from: { row: 0, column: 0 }, to: { row: 0, column: 7 } }, { calculation: await s.calculate() });
  const target = workbook(), pasted = createWorkbookHtmlPastePlan(target.snapshot(), sheetId, { row: 0, column: 0 }, data.html);
  target.commit(target.prepare(pasted.transaction));
  for (const [column, value] of values.entries()) expect(target.getCell(sheetId, 0, column)?.value ?? { type: "blank" }).toEqual(value);
  expect(target.getCell(sheetId, 0, 5)?.value).toEqual({ type: "number", value: 60 });
  const format = target.snapshot().metadata.workbook.numberFormats.find(format => format.formatId === target.getCell(sheetId, 0, 5)!.formatId); expect(format).toMatchObject({ builtinId: 14 });
  expect(target.getCell(sheetId, 0, 6)?.formula?.expression).toBe("=TRUE()"); expect(target.getCell(sheetId, 0, 7)?.formula?.expression).toBe("=#N/A");
  await target.calculate(); expect(target.getCalculatedCell(sheetId, 0, 6)).toMatchObject({ state: "valid", value: { type: "boolean", value: true } });
  expect(target.getCalculatedCell(sheetId, 0, 7)).toMatchObject({ state: "valid", value: { type: "error", value: "#N/A" } });
});
test("clipboard appearance resolves the real chunk ancestry without relying on story node IDs", async () => {
  const original = source(), files = new Map(original.serialize()), encoder = new TextEncoder(), decoder = new TextDecoder();
  const path = original.snapshot().styles[0]!.path;
  files.set(path, encoder.encode(decoder.decode(files.get(path)!) + '\ntable[data-sheet-id="sheet-1"] tr[data-row="256"] td{color:#123456}'));
  const model = openDocument(files, { kind: "workbook" }), data = await createWorkbookClipboardData(model, sheetId, { from: { row: 256, column: 31 }, to: { row: 256, column: 32 } });
  expect(data.html).toContain("color:#123456"); expect(data.html).not.toContain('data-row="256"');
});
test("portable format codes use the same explicit locale and lifecycle as visible number formatting", () => {
  const formatter = createWorkbookNumberFormatter({ dateSystem: "1900", hostLocale: "en-US" });
  expect(formatter.code({ builtinId: 2 })).toBe("0.00"); expect(formatter.code({ code: "0.0000" })).toBe("0.0000");
  expect(formatter.format(12.5, { code: formatter.code({ builtinId: 2 }) })).toMatchObject({ kind: "parts" });
  expect(() => formatter.code({ builtinId: 65536 })).toThrow(); formatter.dispose(); expect(() => formatter.code()).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
});
test("portable rich text keeps partial ancestor-decoration clearing after an independent paste", async () => {
  const s = workbook();
  s.commit(s.prepare({ protocolVersion: 1, documentId: "portable-clipboard", baseRevision: 0, operations: [
    { op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "ABCD" } } }] },
    { op: "formatWorkbookRanges", sheetId, ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], properties: { "text-decoration-line": "underline" } },
    { op: "formatWorkbookText", sheetId, at: { row: 0, column: 0 }, from: 1, to: 3, expectedText: "BC", properties: { "text-decoration-line": "none" } },
  ] }));
  const paints = (session: WorkbookSession): string[] => {
    const resolver = new SourceStyleResolver(session.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS), result: string[] = [];
    const visit = (node: DefaultTreeAdapterTypes.Node, parent?: SourceStyleNode): void => {
      if (node.nodeName === "#text" && "value" in node && parent) for (const character of node.value) result.push(character + ":" + resolver.get(parent)["text-decoration-line"]);
      const current = "tagName" in node ? { tag: node.tagName, attrs: Object.fromEntries(node.attrs.map(attr => [attr.name, attr.value])), parent } : parent;
      if ("childNodes" in node) for (const child of node.childNodes) visit(child, current);
    };
    visit(parseFragment("<table><tr>" + session.getCell(sheetId, 0, 0)!.html + "</tr></table>")); return result;
  };
  expect(paints(s)).toEqual(["A:underline", "B:none", "C:none", "D:underline"]);
  const data = await createWorkbookClipboardData(s, sheetId, { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }), target = workbook();
  target.commit(target.prepare(createWorkbookHtmlPastePlan(target.snapshot(), sheetId, { row: 0, column: 0 }, data.html).transaction));
  expect(paints(target)).toEqual(["A:underline", "B:none", "C:none", "D:underline"]);
});
test("portable font roles keep independent Latin and east-Asian sources and native decoration defaults", async () => {
  const s = workbook();
  s.commit(s.prepare({ protocolVersion: 1, documentId: "portable-clipboard", baseRevision: 0, operations: [
    { op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "A中禰󠄀" } } }] },
    { op: "formatWorkbookRanges", sheetId, ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], properties: {
      "--office-font-ascii": '"Liberation Sans"', "--office-font-h-ansi": '"DejaVu Sans"', "--office-font-east-asia": '"Noto Serif CJK SC"',
      "--office-font-complex-script": '"DejaVu Serif"', "--office-font-underline": "double", "--office-font-strike": "true",
    } },
  ] }));
  const data = await createWorkbookClipboardData(s, sheetId, { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }), target = workbook();
  target.commit(target.prepare(createWorkbookHtmlPastePlan(target.snapshot(), sheetId, { row: 0, column: 0 }, data.html).transaction));
  expect(target.getCell(sheetId, 0, 0)?.value).toEqual({ type: "text", value: "A中禰󠄀" });
  const cell = target.getCell(sheetId, 0, 0)!, resolver = new SourceStyleResolver(target.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS);
  const style = resolver.get({ tag: "td", attrs: { class: cell.cssClass! }, parent: undefined });
  expect(style).toMatchObject({ "--office-font-ascii": '"Liberation Sans"', "--office-font-h-ansi": '"DejaVu Sans"', "--office-font-east-asia": '"Noto Serif CJK SC"', "--office-font-complex-script": '"DejaVu Serif"', "--office-font-underline": "double", "--office-font-strike": "true" });
});
