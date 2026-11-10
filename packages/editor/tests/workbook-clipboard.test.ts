import { expect, test } from "vitest";
import { createWorkbook, openDocument } from "@doffi/core";
import type { DocumentSelection } from "@doffi/core";
import { WorkbookClipboard } from "../src/workbook/clipboard.js";

const snapshot = openDocument(createWorkbook({ documentId: "clipboard" }), { kind: "workbook" }).snapshot();
const selection = (row: number, column = 0): DocumentSelection => ({ documentId: "clipboard", baseRevision: 0, anchor: {
  kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row, column }, to: { row, column: column + 1 } }],
} });
test("clipboard ownership is local, captures immutable coordinates and rejects stale/foreign source selections", () => {
  const clipboard = new WorkbookClipboard(), second = new WorkbookClipboard(), source = selection(0);
  clipboard.capture(snapshot, source, "cut"); expect(clipboard.available).toBe(true); expect(clipboard.kind).toBe("cut"); expect(second.available).toBe(false);
  const target = clipboard.plan(snapshot, selection(4)); expect(Object.isFrozen(target.source.range.from)).toBe(true);
  expect(() => second.plan(snapshot, selection(4))).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(() => clipboard.capture(snapshot, { ...source, documentId: "foreign" }, "copy")).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  const next = { ...snapshot, manifest: { ...snapshot.manifest, revision: 1 } };
  expect(() => clipboard.plan(next, { ...selection(4), baseRevision: 1 })).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  expect(clipboard.plan(snapshot, selection(4)).source.range.from).toEqual({ row: 0, column: 0 });
});
test("only a proven nonoverlapping source-preserving copy can advance its reference after its own paste", () => {
  const clipboard = new WorkbookClipboard(); clipboard.capture(snapshot, selection(0), "copy");
  clipboard.applied(1, "sheet-1", { row: 4, column: 0 }, "contents");
  const next = { ...snapshot, manifest: { ...snapshot.manifest, revision: 1 } }, target = { ...selection(5), baseRevision: 1 };
  expect(clipboard.plan(next, target).source.baseRevision).toBe(1);
  clipboard.applied(2, "sheet-1", { row: 0, column: 1 }, "contents"); expect(clipboard.available).toBe(false);
  clipboard.capture(snapshot, selection(0), "copy"); clipboard.applied(1, "sheet-1", { row: 4, column: 0 }, "values"); expect(clipboard.available).toBe(false);
  clipboard.capture(snapshot, selection(0), "cut"); clipboard.applied(1, "sheet-1", { row: 4, column: 0 }, "contents"); expect(clipboard.available).toBe(false);
});
