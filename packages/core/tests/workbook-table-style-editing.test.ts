import { afterEach, expect, test } from "vitest";
import { createDocumentSession, DEFAULT_DOCUMENT_READ_LIMITS, openDocument, SourceStyleResolver } from "../src/index.js";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookEditProposal, WorkbookInteractionRequest, WorkbookOperation, WorkbookSession, WorkbookTableStyle } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sessions: WorkbookSession[] = [], encoder = new TextEncoder(), decoder = new TextDecoder();
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const operation = { op: "formatWorkbookTableStyleElement" as const, styleId: "chosen", region: "wholeTable" as const, properties: { "background-color": "#00aa77" } };
const transaction = (s: WorkbookSession, operations: readonly DocumentOperation[]) => ({ protocolVersion: 1 as const, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations: [...operations] });
const apply = (s: WorkbookSession, operations: readonly DocumentOperation[]) => s.commit(s.prepare(transaction(s, operations)));
const element = (s: WorkbookSession, styleId = "chosen", region = "wholeTable") => s.snapshot().metadata.workbook.tableStyles!.find(style => style.styleId === styleId)!.elements.find(element => element.region === region)!;
const resolver = (s: WorkbookSession) => new SourceStyleResolver(s.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set());
function fixture(): WorkbookSession {
  const initial = scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "源文本😀\r\n尾" } }]), files = new Map(initial.serialize()), snapshot = initial.snapshot();
  const metadata = JSON.parse(decoder.decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.tableStyles = [{ styleId: "chosen", name: "自定义中文", pivotCompatible: true, elements: [
    { region: "wholeTable", size: 1, style: { cssClass: "shared-table", numberFormatId: snapshot.metadata.workbook.numberFormats[0]!.formatId } },
    { region: "headerRow", size: 1, style: { cssClass: "shared-table" } }, { region: "totalRow", size: 1 },
    { region: "firstRowStripe", size: 2, style: { cssClass: "shared-table" } },
  ] }, { styleId: "other", name: "Other", pivotCompatible: false, elements: [{ region: "wholeTable", size: 1, style: { cssClass: "shared-table" } }] }];
  metadata.workbook.defaultTableStyle = { kind: "custom", styleId: "chosen" }; metadata.workbook.sheets[0]!.defaultCssClass = "shared-table";
  files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  const style = snapshot.styles[0]!; files.set(style.path, encoder.encode(style.css + '.shared-table{background-color:#cc0000;color:#112233;font-weight:bold;--office-font-east-asia:"Noto Serif CJK SC"}td.shared-table{color:#ffffff!important}'));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" }), { sessionLimits: { maxCandidates: 1 } }); sessions.push(session); return session;
}
test("a table region isolates shared contextual appearance and retains other regions, defaults and cell bytes", () => {
  const s = fixture(), before = s.snapshot(), original = resolver(s).classStyle("shared-table");
  apply(s, [{ ...operation, properties: { "background-color": "#00aa77", color: "#112233", "font-weight": "normal" } }]);
  const current = element(s).style!, styles = resolver(s); expect(current.cssClass).not.toBe("shared-table");
  expect(styles.classStyle(current.cssClass)).toMatchObject({ "background-color": "#00aa77", "font-weight": "normal", "--office-font-east-asia": '"Noto Serif CJK SC"' });
  expect(styles.specified({ tag: "td", attrs: { class: current.cssClass }, parent: undefined }).color).toBe("#112233");
  expect(styles.classStyle("shared-table")).toEqual(original); expect(element(s, "other").style!.cssClass).toBe("shared-table"); expect(element(s, "chosen", "headerRow").style!.cssClass).toBe("shared-table");
  expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata.workbook.defaultTableStyle).toEqual(before.metadata.workbook.defaultTableStyle);
  const saved = s.serialize(); expect(apply(s, [{ ...operation, properties: { "background-color": "#00aa77", color: "#112233", "font-weight": "normal" } }]).changed).toBe(false); expect(s.serialize()).toEqual(saved);
  expect(openDocument(saved, { kind: "workbook" }).snapshot().metadata).toEqual(s.snapshot().metadata); expect(s.history().undoSteps).toBe(1);
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); s.redo(); expect(element(s).style!.cssClass).toBe(current.cssClass);
});
test("an unpainted region gains a number format without inheriting cell paint and supports explicit clearing", () => {
  const s = fixture(), formatId = s.snapshot().metadata.workbook.numberFormats[0]!.formatId;
  expect(apply(s, [{ ...operation, region: "totalRow", properties: {}, numberFormatId: null }]).changed).toBe(false);
  apply(s, [{ ...operation, region: "totalRow", properties: {}, numberFormatId: formatId }]);
  const style = element(s, "chosen", "totalRow").style!; expect(style.numberFormatId).toBe(formatId); expect(resolver(s).specifiedClass(style.cssClass)).toEqual({});
  apply(s, [{ ...operation, region: "totalRow", properties: {}, numberFormatId: null }]); expect(element(s, "chosen", "totalRow").style).toEqual({ cssClass: style.cssClass });
  expect(apply(s, [{ ...operation, region: "totalRow", properties: {}, numberFormatId: null }]).changed).toBe(false); s.undo(); expect(element(s, "chosen", "totalRow").style!.numberFormatId).toBe(formatId);
});
test("unknown regions, formats and illegal native paint atomically reject preceding source edits", () => {
  const s = fixture(), original = s.serialize();
  for (const edit of [{ ...operation, styleId: "missing" }, { ...operation, region: "firstColumn" as const }, { ...operation, numberFormatId: "missing" }, { ...operation, properties: { "--office-font-strike": "invalid" } }]) {
    expect(() => apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 1, column: 0, source: { kind: "value", value: { type: "number", value: 9 } } }] }, edit])).toThrow();
    expect(s.serialize()).toEqual(original); expect(s.history().undoSteps).toBe(0);
  }
});
test("style definition, regional appearance and default associations share one candidate and undo", () => {
  const s = fixture(), before = s.snapshot().metadata, style: WorkbookTableStyle = { styleId: "new-style", name: "新样式", pivotCompatible: false, elements: [{ region: "headerRow", size: 1 }] };
  apply(s, [{ op: "setWorkbookTableStyle", style }, { ...operation, styleId: style.styleId, region: "headerRow" }, { op: "setWorkbookSettings", changes: { defaultTableStyle: { kind: "custom", styleId: style.styleId } } }]);
  expect(element(s, style.styleId, "headerRow").style).toBeDefined(); expect(s.snapshot().metadata.workbook.defaultTableStyle).toEqual({ kind: "custom", styleId: style.styleId }); expect(s.history().undoSteps).toBe(1);
  s.undo(); expect(s.snapshot().metadata).toEqual(before); s.redo();
  const current = s.serialize(); expect(() => apply(s, [{ op: "removeWorkbookTableStyle", styleId: style.styleId }])).toThrow(); expect(s.serialize()).toEqual(current);
  apply(s, [{ op: "setWorkbookSettings", changes: { defaultTableStyle: null } }, { op: "removeWorkbookTableStyle", styleId: style.styleId }]);
  expect(s.snapshot().metadata.workbook.defaultTableStyle).toBeUndefined(); expect(s.snapshot().metadata.workbook.tableStyles!.some(value => value.styleId === style.styleId)).toBe(false); s.undo(); expect(element(s, style.styleId, "headerRow").style).toBeDefined();
});
test("cancelled regional preparation releases its sole lease and stale candidates cannot publish", async () => {
  const s = fixture(), original = s.serialize(), abort = new AbortController();
  await expect(s.prepareAsync(transaction(s, [operation]), {}, { signal: abort.signal, yield: async () => {}, onProgress(progress) { if (progress.phase === "operations" && progress.completed === 1) abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  expect(s.serialize()).toEqual(original); apply(s, [operation]);
  const candidate = s.prepare(transaction(s, [{ ...operation, properties: { color: "#aa5500" } }])); s.undo();
  expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); expect(s.snapshot().metadata.workbook.tableStyles).toEqual(fixture().snapshot().metadata.workbook.tableStyles);
});

test("shared table definition appearance requires global proposal ownership", () => {
  const s = fixture(), manifest = s.snapshot().manifest;
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "review", documentId: manifest.documentId, baseRevision: manifest.revision, status: "open", instruction: "复核智能表样式", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] } } }]); s.clearHistory();
  const scoped = s.buildInteractionRequest({ kind: "workbook", annotationIds: ["review"], scope: { kind: "ranges", ranges: [{ sheetId: "sheet-1", range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } }] } });
  const proposal = (request: WorkbookInteractionRequest, operations: WorkbookOperation[]): WorkbookEditProposal => ({ protocolVersion: 1, documentId: request.documentId, baseRevision: request.baseRevision, requestId: request.requestId, proposalId: "style-proposal", annotationIds: ["review"], summary: "更新智能表样式", operations });
  expect(scoped.supportedOperations).not.toContain(operation.op); const original = s.serialize();
  expect(() => s.prepareProposal(scoped, proposal(scoped, [operation]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" })); expect(s.serialize()).toEqual(original);
  const global = s.buildInteractionRequest({ kind: "workbook", annotationIds: ["review"], scope: { kind: "global" } }); expect(global.supportedOperations).toContain(operation.op);
  s.commit(s.prepareProposal(global, proposal(global, [operation]))); expect(element(s).style!.cssClass).not.toBe("shared-table"); expect(s.getAnnotations()[0]!.status).toBe("resolved"); expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.getAnnotations()[0]!.status).toBe("open");
});
