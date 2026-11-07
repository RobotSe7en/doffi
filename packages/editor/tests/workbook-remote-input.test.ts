import { afterEach, expect, test, vi } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentSelection } from "@doffi/core";
import { connectWorkbookSession } from "../src/workbook/session-client.js";
import { serveWorkbookSession } from "../src/workbook/session-server.js";
import { createWorkbookRemoteInputController } from "../src/workbook/remote-input.js";
const cleanups: (() => void)[] = [];
afterEach(() => { for (const cleanup of cleanups.splice(0).reverse()) cleanup(); });
async function setup() {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "remote-input" }), { kind: "workbook" }), { sessionLimits: { maxCandidates: 1 } });
  cleanups.push(() => s.dispose()); const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s); cleanups.push(() => server.dispose());
  const client = await connectWorkbookSession(channel.port2); cleanups.push(() => client.dispose());
  const input = createWorkbookRemoteInputController(client); cleanups.push(() => input.dispose());
  const selection = (row = 0, column = 0): DocumentSelection => ({ documentId: "remote-input", baseRevision: client.state().revision,
    anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row, column }, to: { row, column } }] } });
  return { s, channel, client, input, selection };
}
test("remote composition, canonical text and numeric conversion share the host history", async () => {
  const { s, client, input, selection } = await setup(); await input.select(selection());
  input.compositionStart(); input.setText("中文😀禰󠄀\r\nabc"); await expect(input.apply()).rejects.toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.history().undoSteps).toBe(0); input.compositionEnd(input.state().text); await input.apply();
  expect(s.getCell("sheet-1", 0, 0)?.text).toBe("中文😀禰󠄀\r\nabc"); expect(s.history().undoSteps).toBe(1);
  input.setMode("text"); input.setText("12"); await input.apply(); expect(input.state()).toMatchObject({ text: "'12", dirty: false, applying: false, selection: { baseRevision: 2 } });
  expect(await input.apply()).toMatchObject({ revision: 2, changed: false });
  input.setMode("number"); expect(input.state().text).toBe("12"); await input.apply();
  expect(input.state()).toMatchObject({ text: "12", dirty: false, selection: { baseRevision: 3 } }); expect(s.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 12 });
  expect((await client.readInput("sheet-1", { row: 0, column: 0 })).text).toBe("12");
});
test("an external source revision leaves a recoverable composition tied to its original revision", async () => {
  const { s, client, input, selection } = await setup(); await input.select(selection()); input.compositionStart(); input.setText("本地草稿");
  s.commit(s.prepare({ protocolVersion: 1, documentId: "remote-input", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "external" } } }] }] }));
  await vi.waitFor(() => expect(client.state().revision).toBe(1)); input.compositionEnd("本地草稿😀");
  await expect(input.apply()).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  expect(input.state()).toMatchObject({ text: "本地草稿😀", dirty: true, applying: false, selection: { baseRevision: 0 } });
  expect(s.getCell("sheet-1", 0, 0)?.text).toBe("external"); input.cancel(); await input.select(selection()); expect(input.state().text).toBe("external");
});
test.each([false, true])("changing or disposing a pending draft reclaims its prepared lease (dispose %s)", async dispose => {
  const { s, channel, client, input, selection } = await setup(); await input.select(selection()); input.setText("first");
  const send = channel.port1.postMessage.bind(channel.port1); let held: unknown;
  Object.defineProperty(channel.port1, "postMessage", { configurable: true, value(message: { value?: { handle?: number } }, transfer: Transferable[] = []) {
    if (!held && message.value?.handle) { held = message; return; } send(message, transfer);
  } });
  const applying = input.apply(); await vi.waitFor(() => expect(held).toBeDefined());
  if (dispose) input.dispose(); else input.setText("latest");
  await expect(applying).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  send(held); expect(s.history().undoSteps).toBe(0); expect(s.getCell("sheet-1", 0, 0)).toBeUndefined();
  if (dispose) { const next = createWorkbookRemoteInputController(client); cleanups.push(() => next.dispose()); await next.select(selection()); next.setText("after disposal"); await next.apply(); }
  else { expect(input.state()).toMatchObject({ text: "latest", dirty: true, applying: false }); await input.apply(); }
  expect(s.getCell("sheet-1", 0, 0)?.text).toBe(dispose ? "after disposal" : "latest"); expect(s.history().undoSteps).toBe(1);
});
test("a late selection cannot overwrite a draft begun at the previous cell", async () => {
  const { s, channel, input, selection } = await setup(); await input.select(selection());
  const send = channel.port1.postMessage.bind(channel.port1); let held: unknown;
  Object.defineProperty(channel.port1, "postMessage", { configurable: true, value(message: { value?: { text?: unknown } }, transfer: Transferable[] = []) {
    if (message.value && "text" in message.value) { held = message; return; } send(message, transfer);
  } });
  const choosing = input.select(selection(1)); await vi.waitFor(() => expect(held).toBeDefined()); input.setText("stay here");
  await expect(choosing).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" })); send(held);
  expect(input.state()).toMatchObject({ text: "stay here", dirty: true, selection: { anchor: { ranges: [{ from: { row: 0, column: 0 } }] } } }); expect(s.history().undoSteps).toBe(0);
});
test("a sent commit retains its real outcome while a receipt is delayed", async () => {
  const { s, channel, input, selection } = await setup(); await input.select(selection()); input.setMode("text"); input.setText("12");
  const send = channel.port1.postMessage.bind(channel.port1); let held: unknown;
  Object.defineProperty(channel.port1, "postMessage", { configurable: true, value(message: { value?: { receipt?: unknown } }, transfer: Transferable[] = []) {
    if (message.value?.receipt) { held = message; return; } send(message, transfer);
  } });
  const applying = input.apply(); await vi.waitFor(() => expect(held).toBeDefined());
  expect(s.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "12" });
  expect(() => input.cancel()).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(() => input.setText("new draft")).toThrow();
  send(held); expect(await applying).toMatchObject({ revision: 1 }); expect(input.state()).toMatchObject({ applying: false, dirty: false, text: "'12", selection: { baseRevision: 1 } });
});
