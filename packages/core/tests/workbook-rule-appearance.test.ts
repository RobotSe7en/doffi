import { createHash } from "node:crypto";
import { afterEach, expect, test } from "vitest";
import { createDocumentSession, DEFAULT_DOCUMENT_READ_LIMITS, openDocument, parseContract, SourceStyleResolver } from "../src/index.js";
import type { DocumentOperation, WorkbookConditionalRule, WorkbookDocumentMetadata, WorkbookEditProposal, WorkbookManifest, WorkbookOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sessions: WorkbookSession[] = [], encoder = new TextEncoder(), decoder = new TextDecoder();
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const transaction = (session: WorkbookSession, operations: readonly DocumentOperation[]) => ({ protocolVersion: 1 as const, documentId: session.snapshot().manifest.documentId, baseRevision: session.snapshot().manifest.revision, operations: [...operations] });
const apply = (session: WorkbookSession, operations: readonly DocumentOperation[]) => session.commit(session.prepare(transaction(session, operations)));
const rule = (session: WorkbookSession, id = "selected") => session.snapshot().metadata.workbook.sheets[0]!.conditionalRules!.find(rule => rule.ruleId === id)!;
const appearance = (session: WorkbookSession, cssClass: string) => new SourceStyleResolver(session.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set()).classStyle(cssClass);
const operation: Extract<WorkbookOperation, { op: "formatWorkbookConditionalRule" }> = { op: "formatWorkbookConditionalRule", sheetId: "sheet-1", ruleId: "selected", properties: { "background-color": "#00aa77" } };
function fixture(): WorkbookSession {
  const files = new Map(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 5 } }, { row: 1, column: 0, value: { type: "number", value: 10 } }, { row: 0, column: 3, value: { type: "number", value: 999 }, formula: "=A1*2" }]).serialize());
  const manifest = JSON.parse(decoder.decode(files.get("manifest.json")!)) as WorkbookManifest, metadata = JSON.parse(decoder.decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  files.set(manifest.styles[0]!, encoder.encode(decoder.decode(files.get(manifest.styles[0]!)!) + '.shared-rule{background-color:#cc0000;color:#112233;font-weight:bold;--office-font-east-asia:"Noto Serif CJK SC"}.blue-rule{background-color:#0000ee}td.shared-rule{color:#ffffff!important}'));
  metadata.workbook.sheets[0]!.conditionalRules = [
    { ruleId: "selected", kind: "cellIs", ranges: [range(0, 0, 1, 0)], origin: { row: 0, column: 0 }, priority: 1, stopIfTrue: false, operator: "greaterThan", formulas: ["=0"], cssClass: "shared-rule" },
    { ruleId: "foreign", kind: "expression", ranges: [range(0, 2)], origin: { row: 0, column: 2 }, priority: 2, stopIfTrue: false, expression: "=TRUE", cssClass: "shared-rule" },
    { ruleId: "scale", kind: "colorScale", ranges: [range(0, 1, 1, 1)], origin: { row: 0, column: 1 }, priority: 3, stopIfTrue: false, stops: [{ kind: "min", cssClass: "shared-rule" }, { kind: "max", cssClass: "blue-rule" }] },
  ]; files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  const chunk = manifest.chunks[0]!, bytes = encoder.encode(decoder.decode(files.get(chunk.path)!).replace('data-column="3"', 'data-column="3" class="shared-rule"')); chunk.sha256 = createHash("sha256").update(bytes).digest("hex"); files.set(chunk.path, bytes); files.set("manifest.json", encoder.encode(JSON.stringify(manifest)));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" }), { sessionLimits: { maxCandidates: 1 } }); sessions.push(session);
  apply(session, [{ op: "addAnnotation", annotation: { annotationId: "review", documentId: manifest.documentId, baseRevision: manifest.revision, status: "open", instruction: "复核规则外观禰󠄀😀", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [range(0, 0, 1, 0)] } } }]); session.clearHistory(); return session;
}
test("conditional appearance isolates shared classes, follows the class-only cascade, preserves source and owns one undo", () => {
  const session = fixture(), before = session.serialize(), snapshot = session.snapshot(), foreign = appearance(session, "shared-rule"), candidate = session.prepare(transaction(session, [operation]));
  expect(session.serialize()).toEqual(before); expect(candidate.preview.styles).not.toEqual(snapshot.styles); session.commit(candidate);
  const selected = rule(session); expect(selected.kind === "colorScale" ? undefined : selected.cssClass).not.toBe("shared-rule");
  expect(appearance(session, "cssClass" in selected ? selected.cssClass! : "")).toMatchObject({ "background-color": "#00aa77", color: "#112233", "font-weight": "bold", "--office-font-east-asia": '"Noto Serif CJK SC"' });
  expect(appearance(session, "shared-rule")).toEqual(foreign); expect(rule(session, "foreign")).toEqual(snapshot.metadata.workbook.sheets[0]!.conditionalRules![1]); expect(session.snapshot().manifest.chunks).toEqual(snapshot.manifest.chunks);
  expect(session.getCell("sheet-1", 0, 3)).toMatchObject({ cssClass: "shared-rule", value: { value: 999 }, formula: { cacheState: "imported", expression: "=A1*2" } });
  expect(openDocument(session.serialize(), { kind: "workbook" }).snapshot()).toEqual(session.snapshot());
  const saved = session.serialize(), history = session.history(); expect(apply(session, [operation]).changed).toBe(false); expect(session.serialize()).toEqual(saved); expect(session.history()).toEqual(history);
  session.undo(); expect(session.snapshot().metadata).toEqual(snapshot.metadata); expect(session.snapshot().styles).toEqual(snapshot.styles); session.redo(); expect(session.snapshot().styles).toEqual(candidate.preview.styles);
});
test("indexed color stops isolate their shared paint and number formats can be set or cleared without restyling", () => {
  const session = fixture(), snapshot = session.snapshot();
  apply(session, [{ ...operation, ruleId: "scale", stopIndex: 1, properties: { "background-color": "#ff8800" } }]);
  const scale = rule(session, "scale"); if (scale.kind !== "colorScale") throw Error("Missing scale");
  expect(scale.stops[0]).toEqual((snapshot.metadata.workbook.sheets[0]!.conditionalRules![2] as Extract<WorkbookConditionalRule, { kind: "colorScale" }>).stops[0]); expect(appearance(session, scale.stops[1]!.cssClass)).toMatchObject({ "--office-scale-color": "#ff8800", "--office-scale-tint": "0" });
  expect(apply(session, [{ ...operation, ruleId: "scale", stopIndex: 1, properties: { "background-color": "#ff8800" } }]).changed).toBe(false);
  const format = snapshot.metadata.workbook.numberFormats[0]!.formatId, styles = session.snapshot().styles;
  apply(session, [{ ...operation, properties: {}, numberFormatId: format }]); expect(rule(session)).toMatchObject({ numberFormatId: format, cssClass: "shared-rule" }); expect(session.snapshot().styles).toEqual(styles);
  expect(apply(session, [{ ...operation, properties: {}, numberFormatId: format }]).changed).toBe(false); apply(session, [{ ...operation, properties: {}, numberFormatId: null }]); expect(rule(session)).not.toHaveProperty("numberFormatId"); expect(session.snapshot().styles).toEqual(styles);
});
test("invalid targets, stop slots, native paints and styles roll back preceding source writes", () => {
  const session = fixture(), before = session.serialize();
  for (const patch of [{ sheetId: "foreign" }, { ruleId: "foreign-id" }, { stopIndex: 0 }, { numberFormatId: "missing" }, { ruleId: "scale" }, { ruleId: "scale", stopIndex: 2 }, { ruleId: "scale", stopIndex: 0, numberFormatId: null }, { ruleId: "scale", stopIndex: 0, properties: { color: "red" } }, { ruleId: "scale", stopIndex: 0, properties: { "--office-cell-fill-kind": "gradient" } }, { properties: { width: "50pt" } }, { properties: { "background-color": "url(https://invalid.example/)" } }]) {
    expect(() => apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 100 } } }] }, { ...operation, ...patch }])).toThrow(); expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  }
  expect(() => parseContract("document-operation", { ...operation, stopIndex: -1 })).toThrow(); expect(() => parseContract("document-operation", { ...operation, extra: true })).toThrow();
});
test("private painted source can be cancelled or discarded, releases the sole slot and rejects stale commit", async () => {
  const session = fixture(), before = session.serialize(), signal = { aborted: false }; let painted = false;
  await expect(session.prepareAsync(transaction(session, [operation]), undefined, { signal, yield: async () => {}, onProgress(progress) { if (progress.phase === "operations" && progress.completed === 1) { painted = true; signal.aborted = true; } } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(painted).toBe(true); expect(session.serialize()).toEqual(before); const discarded = session.prepare(transaction(session, [operation])); session.discard(discarded); expect(session.serialize()).toEqual(before);
  apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 8 } } }] }]); const stale = session.prepare(transaction(session, [operation])); session.undo(); const current = session.serialize(); expect(() => session.commit(stale)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); expect(session.serialize()).toEqual(current);
});
test("range proposals format only fully owned rules and can create then format one new rule atomically", () => {
  const session = fixture(), request = session.buildInteractionRequest({ kind: "workbook", annotationIds: ["review"], scope: { kind: "ranges", ranges: [{ sheetId: "sheet-1", range: range(0, 0, 1, 0) }] } });
  const proposal = (operations: WorkbookOperation[]): WorkbookEditProposal => ({ protocolVersion: 1, documentId: request.documentId, baseRevision: request.baseRevision, requestId: request.requestId, proposalId: "rule-proposal", annotationIds: ["review"], summary: "更新规则外观", operations });
  const before = session.serialize(); expect(() => session.prepareProposal(request, proposal([{ ...operation, ruleId: "foreign", properties: {} }]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" })); expect(session.serialize()).toEqual(before);
  const candidate = session.prepareProposal(request, proposal([operation])); session.commit(candidate); expect(session.getAnnotations()[0]!.status).toBe("resolved"); session.undo();
  const nextRequest = session.buildInteractionRequest({ kind: "workbook", annotationIds: ["review"] });
  const next = { ...proposal([{ op: "setWorkbookConditionalRule", sheetId: "sheet-1", rule: { ruleId: "new-rule", kind: "expression", expression: "=TRUE", origin: { row: 0, column: 0 }, ranges: [range(0, 0)], priority: 4, stopIfTrue: false } }, { ...operation, ruleId: "new-rule" }]), requestId: nextRequest.requestId, baseRevision: nextRequest.baseRevision };
  session.commit(session.prepareProposal(nextRequest, next)); expect(rule(session, "new-rule")).toHaveProperty("cssClass"); expect(session.history().undoSteps).toBe(1);
});
test("continued conditional painting retains bounded CSS growth and no-op history", () => {
  const session = fixture(); let first = 0;
  for (let index = 0; index < 100; index++) { apply(session, [{ ...operation, properties: { "background-color": "#" + (0x100000 + index).toString(16) } }]); if (index === 49) first = session.snapshot().styles.reduce((size, style) => size + style.css.length, 0); }
  expect(session.snapshot().styles.reduce((size, style) => size + style.css.length, 0)).toBeLessThan(first * 2.5); expect(session.history().undoSteps).toBe(100);
});
test("contextual native paint cannot make a differing visible rule look like a no-op", () => {
  const initial = fixture(), files = new Map(initial.serialize()), snapshot = initial.snapshot(), style = snapshot.styles[0]!;
  files.set(style.path, encoder.encode(style.css + ".shared-rule{--office-cell-foreground-color:#112233;--office-cell-foreground-tint:0;--office-cell-foreground-native-alpha:initial}td.shared-rule{--office-cell-foreground-color:#ffffff!important}.shared-rule[data-column=\"0\"]{--office-cell-foreground-color:#aabbcc!important}"));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" })); sessions.push(session);
  const edit = { ...operation, properties: { color: "#112233" } }, resolver = new SourceStyleResolver(session.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set());
  expect(resolver.hasSpecifiedClass("shared-rule", { "--office-cell-foreground-color": "#112233" })).toBe(false);
  expect(apply(session, [edit]).changed).toBe(true); const selected = rule(session); if (!("cssClass" in selected)) throw Error("Missing rule class");
  const after = new SourceStyleResolver(session.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set());
  expect(after.specified({ tag: "td", attrs: { class: selected.cssClass!, "data-column": "0" }, parent: undefined })["--office-cell-foreground-color"]).toBe("#112233");
  const saved = session.serialize(); expect(apply(session, [edit]).changed).toBe(false); expect(session.serialize()).toEqual(saved);
  const global = new SourceStyleResolver([".paint{color:#112233}td{color:#ffffff!important}"], DEFAULT_DOCUMENT_READ_LIMITS, new Set());
  expect(global.hasSpecifiedClass("paint", { color: "#112233" }, "td")).toBe(false); expect(global.hasSpecifiedClass("paint", { color: "#112233" }, "span")).toBe(true);
});
test("a local decoration longhand wins a contextual important shorthand and keeps repeated history unchanged", () => {
  const initial = fixture(), files = new Map(initial.serialize()), style = initial.snapshot().styles[0]!;
  files.set(style.path, encoder.encode(style.css + ".shared-rule{text-decoration-line:none}td.shared-rule{text-decoration:underline!important}"));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" })); sessions.push(session);
  const edit = { ...operation, properties: { "text-decoration-line": "none" } }; expect(apply(session, [edit]).changed).toBe(true);
  const selected = rule(session); if (!("cssClass" in selected)) throw Error("Missing rule class");
  const resolver = new SourceStyleResolver(session.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set());
  expect(resolver.specified({ tag: "td", attrs: { class: selected.cssClass! }, parent: undefined })["text-decoration-line"]).toBe("none");
  const saved = session.serialize(); expect(apply(session, [edit]).changed).toBe(false); expect(session.serialize()).toEqual(saved);
});
