import { createHash } from "node:crypto";
import { expect, test } from "vitest";
import { createWorkbook, openDocument, parseContract } from "../src/index.js";
import type { WorkbookManifest } from "../src/index.js";
import { openWorkbookModel } from "../src/workbook/model.js";
import { workbookSource } from "../src/workbook/source.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
function original() {
  const files = new Map(createWorkbook({ documentId: "validated-chunks" })), manifest = JSON.parse(decoder.decode(files.get("manifest.json")!)) as WorkbookManifest;
  const html = '<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="0" data-column="0"><tbody><tr data-row="0"><td data-column="0" class="linked" data-value-type="number" data-value="7">7</td><td data-column="1" data-value-type="number" data-value="9">9</td></tr></tbody></table></body></html>';
  const path = "sheets/sheet-1/0-0.html", bytes = encoder.encode(html);
  manifest.chunks.push({ path, sheetId: "sheet-1", row: 0, column: 0, sha256: createHash("sha256").update(bytes).digest("hex") });
  files.set(path, bytes); files.set("manifest.json", encoder.encode(JSON.stringify(manifest)));
  files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")!) + ".linked{color:red}"));
  return openDocument(files, { kind: "workbook" });
}
test("private source reuse keeps verified indexes and still enforces the current byte/cell limits", () => {
  const previous = original(), source = workbookSource(previous)!, files = new Map(source.files), manifest = parseContract("document-manifest", previous.snapshot().manifest) as WorkbookManifest;
  const next = openWorkbookModel(files, manifest, source.limits, undefined, previous), path = manifest.chunks[0]!.path;
  expect(workbookSource(next)!.chunks.get(path)).toBe(source.chunks.get(path)); expect(next.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 7 });
  expect(() => openWorkbookModel(new Map(source.files), parseContract("document-manifest", manifest) as WorkbookManifest, { ...source.limits, maxFileBytes: 100 }, undefined, previous))
    .toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => openWorkbookModel(new Map(source.files), parseContract("document-manifest", manifest) as WorkbookManifest, source.limits, { maxCells: 1 }, previous))
    .toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
});
test("a changed digest declaration cannot borrow an unchanged buffer's prior proof", () => {
  const previous = original(), source = workbookSource(previous)!, manifest = parseContract("document-manifest", previous.snapshot().manifest) as WorkbookManifest;
  manifest.chunks[0]!.sha256 = "0".repeat(64);
  expect(() => openWorkbookModel(new Map(source.files), manifest, source.limits, undefined, previous)).toThrowError(/digest/);
  expect(previous.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 7 });
});
test("new mutable buffers are revalidated even at an existing path with the old digest", () => {
  const previous = original(), source = workbookSource(previous)!, files = new Map(source.files), manifest = parseContract("document-manifest", previous.snapshot().manifest) as WorkbookManifest;
  const path = manifest.chunks[0]!.path;
  files.set(path, encoder.encode(decoder.decode(source.files.get(path)!).replace('data-value="7">7', 'data-value="8">8')));
  expect(() => openWorkbookModel(files, manifest, source.limits, undefined, previous)).toThrowError(/digest/);
  expect(previous.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 7 });
});
test("verified bytes are parsed again when a declaration moves their source position", () => {
  const previous = original(), source = workbookSource(previous)!, manifest = parseContract("document-manifest", previous.snapshot().manifest) as WorkbookManifest;
  manifest.chunks[0]!.row = 256;
  expect(() => openWorkbookModel(new Map(source.files), manifest, source.limits, undefined, previous)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});
test("verified bytes are parsed again when their formerly declared style disappears", () => {
  const previous = original(), source = workbookSource(previous)!, files = new Map(source.files), manifest = parseContract("document-manifest", previous.snapshot().manifest) as WorkbookManifest;
  files.set("styles.css", encoder.encode(decoder.decode(source.files.get("styles.css")!).replace(".linked{color:red}", "")));
  expect(() => openWorkbookModel(files, manifest, source.limits, undefined, previous)).toThrowError(/style|class/i);
});
