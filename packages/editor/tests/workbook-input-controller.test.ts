import { expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentSelection, WorkbookSession } from "@doffi/core";
import { createWorkbookInputController } from "../src/workbook/input-controller.js";
function session(): WorkbookSession { return createDocumentSession(openDocument(createWorkbook({ documentId: "input-controller" }), { kind: "workbook" })); }
function selection(session: WorkbookSession, row = 0, column = 0): DocumentSelection { return { documentId: "input-controller", baseRevision: session.snapshot().manifest.revision,
  anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row, column }, to: { row, column } }] } }; }
test("async input preserves rich UTF-16 text and returns one committed history step", async () => {
  const source = session(), controller = createWorkbookInputController(source);
  try {
    controller.select(selection(source)); controller.setText("中文禰󠄀😀\nnext"); const receipt = await controller.applyAsync();
    expect(receipt.changed).toBe(true); expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "中文禰󠄀😀\nnext" });
    expect(controller.state()).toMatchObject({ dirty: false, selection: { baseRevision: receipt.revision } }); expect(source.history().undoSteps).toBe(1);
    source.undo(); expect(source.getCell("sheet-1", 0, 0)).toBeUndefined();
  } finally { controller.dispose(); source.dispose(); }
});
test.each(["text", "composition", "cancel", "dispose"] as const)("%s during async input cancels the old candidate and preserves the new transient state", async action => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.setText("old draft");
  let started!: () => void, release!: () => void;
  const begin = new Promise<void>(resolve => { started = resolve; }), wait = new Promise<void>(resolve => { release = resolve; });
  try {
    const job = controller.applyAsync(undefined, { yield: async () => { started(); await wait; } }); await begin;
    if (action === "text") controller.setText("new 禰󠄀😀"); else if (action === "composition") controller.compositionStart(); else controller[action]();
    release(); await expect(job).rejects.toMatchObject({ code: "OperationCancelled" }); expect(source.history().undoSteps).toBe(0); expect(source.getCell("sheet-1", 0, 0)).toBeUndefined();
    if (action === "text") expect(controller.state()).toMatchObject({ text: "new 禰󠄀😀", dirty: true });
    if (action === "composition") expect(controller.state()).toMatchObject({ text: "old draft", composing: true });
    if (action === "cancel") expect(controller.state()).toMatchObject({ text: "", dirty: false });
    const next = createWorkbookInputController(source); next.select(selection(source)); next.setText("usable"); await next.applyAsync(); next.dispose();
    expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "usable" });
  } finally { controller.dispose(); source.dispose(); }
});
test("external source commits reject late async input and retain the recoverable draft", async () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.setText("local draft");
  let started!: () => void, release!: () => void;
  const begin = new Promise<void>(resolve => { started = resolve; }), wait = new Promise<void>(resolve => { release = resolve; });
  try {
    const job = controller.applyAsync(undefined, { yield: async () => { started(); await wait; } }); await begin;
    source.commit(source.prepare({ protocolVersion: 1, documentId: "input-controller", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "external" } } }] }] }));
    release(); await expect(job).rejects.toMatchObject({ code: "RevisionConflict" });
    expect(controller.state()).toMatchObject({ text: "local draft", dirty: true, selection: { baseRevision: 0 } }); expect(source.history().undoSteps).toBe(1);
    expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "external" });
  } finally { controller.dispose(); source.dispose(); }
});
test("continued typing from the known commit callback becomes a new draft bound to that receipt", async () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.setText("submitted");
  const stop = source.subscribe(event => { if (event.type === "documentChanged") controller.setText("next draft"); });
  try {
    const receipt = await controller.applyAsync(); expect(controller.state()).toMatchObject({ text: "next draft", dirty: true, selection: { baseRevision: receipt.revision } });
    expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "submitted" }); controller.cancel(); expect(controller.state()).toMatchObject({ text: "submitted", dirty: false });
  } finally { stop(); controller.dispose(); source.dispose(); }
});
test("composition is transient and publishes exactly one atomic undo step", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.compositionStart();
  controller.setText("中"); expect(() => controller.apply()).toThrow(); expect(source.history().undoSteps).toBe(0);
  controller.compositionEnd("中文😀禰󠄀\nabc"); expect(source.getCell("sheet-1", 0, 0)).toBeUndefined();
  controller.apply(); expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "中文😀禰󠄀\nabc" });
  expect(controller.state()).toMatchObject({ dirty: false, composing: false, selection: { baseRevision: 1 } }); expect(source.history().undoSteps).toBe(1);
  source.undo(); expect(source.getCell("sheet-1", 0, 0)).toBeUndefined(); controller.dispose(); expect(() => controller.state()).toThrow(); expect(source.snapshot().manifest.revision).toBe(2);
});
test("a changed source rejects a late composition while retaining recoverable input", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.compositionStart(); controller.setText("暂存");
  source.commit(source.prepare({ protocolVersion: 1, documentId: "input-controller", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "external" } } }] }] }));
  controller.compositionEnd("本地输入😀"); expect(() => controller.apply()).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  expect(controller.state()).toMatchObject({ dirty: true, text: "本地输入😀", selection: { baseRevision: 0 } }); expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "external" });
});
test("failed formulas, selection replacement and explicit conversion preserve the draft boundary", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.setText("=");
  expect(() => controller.apply()).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(controller.state().text).toBe("=");
  expect(() => controller.select(selection(source, 1))).toThrow(); controller.cancel(); controller.select(selection(source, 1));
  controller.setText("00123"); controller.apply(); expect(source.getCell("sheet-1", 1, 0)?.value).toEqual({ type: "text", value: "00123" });
  controller.setMode("number"); controller.apply(); expect(source.getCell("sheet-1", 1, 0)?.value).toEqual({ type: "number", value: 123 });
  expect(controller.state()).toMatchObject({ mode: "auto", dirty: false });
});
test("idle apply preserves explicitly typed text and empty strings without a new revision", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.setMode("text"); controller.setText("12"); controller.apply();
  const revision = source.snapshot().manifest.revision; expect(controller.apply()).toMatchObject({ changed: false, revision });
  expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "12" });
  controller.select(selection(source, 1)); controller.setMode("text"); controller.apply(); const emptyRevision = source.snapshot().manifest.revision;
  expect(controller.apply()).toMatchObject({ changed: false, revision: emptyRevision }); expect(source.getCell("sheet-1", 1, 0)?.value).toEqual({ type: "text", value: "" });
});
test("editing literal formula-like text keeps its type while whole replacement can enter a formula", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.setMode("text"); controller.setText("=1"); controller.apply();
  expect(controller.state()).toMatchObject({ text: "'=1", dirty: false }); controller.setText(controller.state().text + "X"); controller.apply();
  expect(source.getCell("sheet-1", 0, 0)).toMatchObject({ value: { type: "text", value: "=1X" } });
  controller.setText("=1+1"); controller.apply(); expect(source.getCell("sheet-1", 0, 0)?.formula?.expression).toBe("=1+1");
  controller.setMode("text"); controller.setText("12"); controller.apply(); expect(controller.state().text).toBe("'12");
  controller.setMode("number"); expect(controller.state().text).toBe("12"); controller.apply(); expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 12 });
});
test("rich input preserves unaffected font runs, CR and variation sequences through one IME commit", () => {
  const files = new Map(createWorkbook({ documentId: "input-controller" })); files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + '.emphasis{font-style:italic;--office-font-east-asia:"Noto Serif CJK SC"}'));
  const source = createDocumentSession(openDocument(files, { kind: "workbook" }));
  source.commit(source.prepare({ protocolVersion: 1, documentId: "input-controller", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0,
    source: { kind: "richText", html: '<span class="cell-default">ABC&#13;</span><span class="emphasis">中文禰󠄀</span>' } }] }] })); source.clearHistory();
  const original = source.getCell("sheet-1", 0, 0)!.html, controller = createWorkbookInputController(source); controller.select(selection(source)); controller.compositionStart(); controller.compositionEnd(controller.state().text + "尾"); controller.apply();
  expect(source.getCell("sheet-1", 0, 0)?.html).toContain('<span class="cell-default">ABC&#13;</span><span class="emphasis">中文禰󠄀尾</span>');
  expect(source.history().undoSteps).toBe(1); source.undo(); expect(source.getCell("sheet-1", 0, 0)?.html).toBe(original);
});
test("literal input escaping supports the full source text budget without an idle edit", () => {
  const source = session(), controller = createWorkbookInputController(source); controller.select(selection(source)); controller.setMode("text"); controller.setText("=" + "A".repeat(32766)); controller.apply();
  expect(controller.state().text).toHaveLength(32768); const revision = source.snapshot().manifest.revision;
  expect(controller.apply()).toMatchObject({ changed: false, revision }); expect(source.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "=" + "A".repeat(32766) });
});
