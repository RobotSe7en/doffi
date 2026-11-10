import { afterEach, expect, test, vi } from "vitest";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { createDocumentSession, DocumentError } from "@doffi/core";
import type { DocumentSelection, DocumentTransaction, TransactionPrepareOptions, WorkbookDataValidation, WorkbookSession } from "@doffi/core";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { createWorkbookInputController } from "../src/workbook/input-controller.js";
import { createWorkbookRemoteInputController } from "../src/workbook/remote-input.js";
import { connectWorkbookSession } from "../src/workbook/session-client.js";
import { serveWorkbookSession } from "../src/workbook/session-server.js";
import { withWorkbookValidationConfirmation as confirm, withWorkbookAsyncValidationConfirmation as confirmAsync, workbookValidationInputKey } from "../src/workbook/validation-confirmation.js";

const cleanups: (() => void)[] = [];
afterEach(() => { for (const cleanup of cleanups.splice(0).reverse()) cleanup(); });
const range = { from: { row: 0, column: 0 }, to: { row: 9, column: 0 } };
const base = { ruleId: "rule", ranges: [range], origin: range.from, allowBlank: true, showErrorMessage: true, showInputMessage: true,
  errorStyle: "warning" as const, errorTitle: "范围", error: "必须小于五", promptTitle: "", prompt: "" };
function session(rules: WorkbookDataValidation[] = [{ ...base, kind: "decimal", operator: "lessThan", formulas: ["=5"] }]) {
  const s = createDocumentSession(scalarWorkbook([], metadata => { metadata.workbook.sheets[0]!.dataValidations = rules; }), { sessionLimits: { maxCandidates: 1 } });
  cleanups.push(() => s.dispose()); return s;
}
function transaction(s: WorkbookSession, rows = [0]): DocumentTransaction {
  return { protocolVersion: 1, documentId: "scalar-kernel", baseRevision: s.snapshot().manifest.revision,
    operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: rows.map(row => ({ row, column: 0, source: { kind: "value", value: { type: "number", value: 9 } } })) }] };
}
function selection(s: WorkbookSession): DocumentSelection { return { documentId: "scalar-kernel", baseRevision: s.snapshot().manifest.revision,
  anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: range.from, to: range.from }] } }; }
function context(s: WorkbookSession, answer: boolean, draft?: () => string) {
  const revision = s.snapshot().manifest.revision, original = draft?.(), confirm = vi.fn(() => answer);
  return { documentId: "scalar-kernel", baseRevision: revision, confirm, assertCurrent() {
    if (s.snapshot().manifest.revision !== revision) throw new DocumentError("RevisionConflict", "Source changed.");
    if (draft && draft() !== original) throw new DocumentError("OperationCancelled", "Draft changed.");
  } };
}
test("declining or omitting the host decision retains the exact draft and source; acceptance creates one undo", () => {
  const s = session(), input = createWorkbookInputController(s); cleanups.push(() => input.dispose()); input.select(selection(s)); input.setText("9");
  const original = s.serialize(), declined = context(s, false);
  expect(() => confirm(options => input.apply(options), {}, { ...declined, confirm: undefined })).toThrow(); expect(declined.confirm).not.toHaveBeenCalled();
  expect(() => confirm(options => input.apply(options), {}, declined)).toThrow(); expect(s.serialize()).toEqual(original);
  expect(input.state()).toMatchObject({ text: "9", dirty: true }); expect(s.history().undoSteps).toBe(0);
  const accepted = context(s, true); const receipt = confirm(options => input.apply(options), {}, accepted);
  expect(receipt.changed).toBe(true); expect(accepted.confirm).toHaveBeenCalledTimes(1);
  const event = accepted.confirm.mock.calls[0] as unknown as [unknown];
  expect(event[0]).toMatchObject({ documentId: "scalar-kernel", baseRevision: 0, warnings: [{ ruleId: "rule", sheetId: "sheet-1", row: 0, column: 0, message: "必须小于五" }] });
  expect(Object.isFrozen(event[0])).toBe(true); expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.getCell("sheet-1", 0, 0)).toBeUndefined();
});
test.each(["stop", "unavailable"])("%s validation cannot be acknowledged", mode => {
  const rule: WorkbookDataValidation = mode === "stop" ? { ...base, kind: "decimal", operator: "lessThan", formulas: ["=5"], errorStyle: "stop" }
    : { ...base, kind: "custom", expression: "=FUTUREFUNC(A1)" };
  const s = session([rule]), decision = context(s, true);
  expect(() => confirm(options => s.prepare(transaction(s), options), {}, decision)).toThrow(); expect(decision.confirm).not.toHaveBeenCalled(); expect(s.history().undoSteps).toBe(0);
});
test("all warning rules must pass private retries before one bulk publication", async () => {
  const s = session([ { ...base, kind: "decimal", operator: "lessThan", formulas: ["=5"] },
    { ...base, ruleId: "second", ranges: [{ from: { row: 10, column: 0 }, to: { row: 10, column: 0 } }], kind: "whole", operator: "equal", formulas: ["=2"], errorStyle: "information" } ]);
  const tx = transaction(s, [0, 10]), decision = context(s, true);
  decision.confirm.mockImplementation(() => { expect(s.history().undoSteps).toBe(0); expect(s.getCell("sheet-1", 0, 0)).toBeUndefined(); return true; });
  const candidate = await confirmAsync(options => s.prepareAsync(tx, options), {}, decision); s.commit(candidate);
  expect(decision.confirm).toHaveBeenCalledTimes(2); expect(s.history().undoSteps).toBe(1); expect(s.getCell("sheet-1", 10, 0)?.value).toEqual({ type: "number", value: 9 });
  s.undo(); expect(s.getCell("sheet-1", 0, 0)).toBeUndefined(); expect(s.getCell("sheet-1", 10, 0)).toBeUndefined();
});
test("a decision cannot acknowledge a changed input or revision", () => {
  for (const changeSource of [false, true]) {
    const s = session(), input = createWorkbookInputController(s); cleanups.push(() => input.dispose()); input.select(selection(s)); input.setText("9");
    const decision = context(s, true, () => workbookValidationInputKey(input.state()));
    decision.confirm.mockImplementation(() => { if (!changeSource) input.setText("latest");
      if (changeSource) s.commit(s.prepare({ ...transaction(s), operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 1, source: { kind: "value", value: { type: "text", value: "external" } } }] }] })); return true; });
    expect(() => confirm(options => input.apply(options), {}, decision)).toThrowError(expect.objectContaining({ code: changeSource ? "RevisionConflict" : "OperationCancelled" }));
    expect(s.getCell("sheet-1", 0, 0)).toBeUndefined(); expect(input.state()).toMatchObject({ text: changeSource ? "9" : "latest", dirty: true });
  }
});
test("frozen host time, authorized bytes and prior acknowledgments survive confirmation", () => {
  const options: TransactionPrepareOptions = { calculationContext: { now: { unixMilliseconds: 1, utcOffsetMinutes: 480 }, locale: "zh-CN" },
    validation: { acceptedRuleIds: ["existing"] }, assets: [{ descriptor: { path: "assets/authorized.png", mimeType: "image/png", sha256: "0".repeat(64) }, bytes: new Uint8Array([1, 2, 3]) }] };
  const error = new DocumentError("InvalidOperation", "warning", [{ code: "validation-confirmation-required", path: ["sheet-1", 0, 0, "rule"], message: "warning" }]);
  const apply = vi.fn((value: TransactionPrepareOptions) => { if (apply.mock.calls.length === 1) throw error; return value; });
  const s = session(), decision = context(s, true);
  decision.confirm.mockImplementation(() => { Object.assign(options.calculationContext!.now!, { unixMilliseconds: 99 }); (options.validation!.acceptedRuleIds as string[]).push("unrelated"); options.assets![0]!.bytes[0] = 99; Object.assign(options.assets![0]!.descriptor, { path: "assets/changed.png" }); return true; });
  expect(confirm(apply, options, decision)).toMatchObject({ calculationContext: { now: { unixMilliseconds: 1, utcOffsetMinutes: 480 }, locale: "zh-CN" }, validation: { acceptedRuleIds: ["existing", "rule"] }, assets: [{ descriptor: { path: "assets/authorized.png" }, bytes: new Uint8Array([1, 2, 3]) }] });
});
test("repeated, mixed or malformed diagnostics and a truthy async decision never auto-accept", async () => {
  const s = session(), decision = context(s, true), detail = { code: "validation-confirmation-required", path: ["sheet-1", 0, 0, "rule"], message: "warning" };
  for (const details of [[detail, { ...detail, code: "validation-blocked" }], [{ ...detail, path: ["rule"] }]]) {
    const error = new DocumentError("InvalidOperation", "bad", details); expect(() => confirm(() => { throw error; }, {}, decision)).toThrow(error);
  }
  expect(decision.confirm).not.toHaveBeenCalled();
  const repeated = new DocumentError("InvalidOperation", "warning", [detail]);
  expect(() => confirm(() => { throw repeated; }, {}, decision)).toThrow(repeated); expect(decision.confirm).toHaveBeenCalledTimes(1);
  const asyncDecision = { ...decision, confirm: (() => Promise.resolve(true)) as unknown as typeof decision.confirm };
  await expect(confirmAsync(async () => { throw repeated; }, {}, asyncDecision)).rejects.toThrow(repeated);
});
test("Worker warning rejection and acceptance retain the sole session and one history step", async () => {
  const s = session(), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s); cleanups.push(() => server.dispose());
  const client = await connectWorkbookSession(channel.port2); cleanups.push(() => client.dispose());
  const input = createWorkbookRemoteInputController(client); cleanups.push(() => input.dispose()); await input.select(selection(s)); input.setText("9");
  const declined = context(s, false); await expect(confirmAsync(options => input.apply(options), {}, declined)).rejects.toThrow();
  expect(input.state()).toMatchObject({ text: "9", dirty: true, applying: false }); expect(s.history().undoSteps).toBe(0);
  const accepted = context(s, true, () => workbookValidationInputKey(input.state())); await confirmAsync(options => input.apply(options), {}, accepted);
  expect(s.history().undoSteps).toBe(1); expect(input.state()).toMatchObject({ text: "9", dirty: false, applying: false }); await client.undo(); expect(s.getCell("sheet-1", 0, 0)).toBeUndefined();
});

test("a delayed Worker warning cannot prompt or retry a superseding draft", async () => {
  const s = session(), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s); cleanups.push(() => server.dispose());
  const client = await connectWorkbookSession(channel.port2); cleanups.push(() => client.dispose());
  const input = createWorkbookRemoteInputController(client); cleanups.push(() => input.dispose()); await input.select(selection(s)); input.setText("9");
  const send = channel.port1.postMessage.bind(channel.port1); let held: unknown;
  Object.defineProperty(channel.port1, "postMessage", { configurable: true, value(message: { error?: unknown }, transfer: Transferable[] = []) {
    if (!held && message.error) { held = message; return; } send(message, transfer);
  } });
  const decision = context(s, true, () => workbookValidationInputKey(input.state()));
  const applying = confirmAsync(options => input.apply(options), {}, decision);
  await vi.waitFor(() => expect(held).toBeDefined()); input.setText("latest"); send(held);
  await expect(applying).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" })); expect(decision.confirm).not.toHaveBeenCalled();
  expect(input.state()).toMatchObject({ text: "latest", dirty: true, applying: false }); expect(s.history().undoSteps).toBe(0);
});

test("an actual TODAY warning uses the initial host clock on every retry", () => {
  const s = session([{ ...base, kind: "custom", expression: "=TODAY()>A1" }]), tx = transaction(s);
  const cells = tx.operations[0]; if (cells?.op !== "setWorkbookCells") throw Error("Missing input"); cells.cells[0]!.source = { kind: "value", value: { type: "number", value: 1000000000 } };
  const bytes = readFileSync(new URL("../../../examples/presentation/assets/sample.png", import.meta.url)), original = new Uint8Array(bytes);
  const descriptor = { path: "assets/authorized.png", mimeType: "image/png", sha256: createHash("sha256").update(bytes).digest("hex") };
  tx.operations.push({ op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "one-cell", from: { row: 2, column: 2, offsetXPt: 0, offsetYPt: 0 }, widthPt: 120, heightPt: 80 }, object: { kind: "image", assetPath: descriptor.path, alt: "授权图片" } });
  const options: TransactionPrepareOptions = { calculationContext: { now: { unixMilliseconds: Date.UTC(2026, 10, 9), utcOffsetMinutes: 480 } }, assets: [{ descriptor, bytes }] }, decision = context(s, true);
  decision.confirm.mockImplementation(() => { Object.assign(options.calculationContext!.now!, { unixMilliseconds: 0 }); bytes.fill(0); descriptor.path = "assets/changed.png"; return true; });
  s.commit(confirm(value => s.prepare(tx, value), options, decision)); expect(decision.confirm).toHaveBeenCalledTimes(1); expect(s.history().undoSteps).toBe(1);
  expect(s.serialize().get("assets/authorized.png")).toEqual(original); s.undo(); expect(s.serialize().has("assets/authorized.png")).toBe(false);
});
