import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import { createWorkbook, createDocumentSession, DocumentManifestSchema, getContractJsonSchema, openDocument, parseContract } from "../src/index.js";
import type { DocumentPackage, WorkbookManifest } from "../src/index.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
function json(files: DocumentPackage, path: string): any { return JSON.parse(decoder.decode(files.get(path)!)); }
function seeded(): Map<string, Uint8Array> {
  const files = new Map(createWorkbook({ documentId: "book", sheetName: "预算" })), manifest: WorkbookManifest = json(files, "manifest.json");
  for (const [row, column, cells] of [
    [0, 0, '<tr data-row="0"><td data-column="0" data-value-type="text" class="cell-default">预算😀禰󠄀<br>第二行</td><td data-column="1" data-value-type="number" data-value="12.5" data-format-id="general">12.50</td><td data-column="2" data-value-type="number" data-value="13.5" data-formula="=B1+1" data-cache-state="imported">13.5</td></tr>'],
    [256, 32, '<tr data-row="256"><td data-column="32" data-value-type="boolean" data-value="true">TRUE</td></tr>'],
    [1_048_320, 16_352, '<tr data-row="1048575"><td data-column="16383" data-value-type="error" data-value="#N/A">#N/A</td></tr>'],
  ] as const) {
    const path = `sheets/sheet-1/${row}-${column}.html`, bytes = encoder.encode(`<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="${row}" data-column="${column}"><tbody>${cells}</tbody></table></body></html>`);
    files.set(path, bytes); manifest.chunks.push({ sheetId: "sheet-1", row, column, path, sha256: createHash("sha256").update(bytes).digest("hex") });
  }
  files.set("manifest.json", encoder.encode(JSON.stringify(manifest))); return files;
}
function mutateJson(files: Map<string, Uint8Array>, path: string, edit: (value: any) => void): void {
  const value = json(files, path); edit(value); files.set(path, encoder.encode(JSON.stringify(value)));
}

describe("unified workbook package reader", () => {
  it("creates an isolated empty package and branches its snapshot without full HTML", () => {
    const files = createWorkbook({ documentId: "new", dateSystem: "1904", sheetName: "账目" }), model = openDocument(files);
    expect(model.kind).toBe("workbook"); if (model.kind !== "workbook") throw new Error("Wrong model branch");
    expect(model.snapshot().manifest.chunks).toEqual([]); expect(model.snapshot().metadata.workbook.dateSystem).toBe("1904");
    expect("html" in model.snapshot()).toBe(false); expect("entrypoint" in model.snapshot().manifest).toBe(false);
    expect(model.readWindow({ sheetId: "sheet-1", row: 0, column: 0, rows: 10, columns: 10 })).toEqual([]);
    expect(createDocumentSession(model).snapshot()).toEqual(model.snapshot());
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(model.snapshot());
    expect(() => openDocument(files, { kind: "document" })).toThrowError(expect.objectContaining({ code: "UnsupportedFormat" }));
  });
  it("reads typed cells, exact Unicode, source formulas and sparse distant windows", () => {
    const files = seeded(), model = openDocument(files, { kind: "workbook" });
    const cells = model.readWindow({ sheetId: "sheet-1", row: 0, column: 0, rows: 1, columns: 3 });
    expect(cells[0]?.value).toEqual({ type: "text", value: "预算😀禰󠄀\n第二行" });
    expect(cells[1]?.value).toEqual({ type: "number", value: 12.5 }); expect(cells[1]?.text).toBe("12.50");
    expect(cells[2]?.formula).toEqual({ expression: "=B1+1", cacheState: "imported" });
    expect(model.getCell("sheet-1", 256, 32)?.value).toEqual({ type: "boolean", value: true });
    expect(model.getCell("sheet-1", 1_048_575, 16_383)?.value).toEqual({ type: "error", value: "#N/A" });
    expect(model.getCell("sheet-1", 100, 100)).toBeUndefined(); expect(Object.isFrozen(cells[0]?.value)).toBe(true);
    files.get("sheets/sheet-1/0-0.html")!.fill(0); model.serialize().get("sheets/sheet-1/0-0.html")!.fill(0);
    expect(model.getCell("sheet-1", 0, 1)?.value).toEqual({ type: "number", value: 12.5 });
    const reopened = openDocument(model.serialize(), { kind: "workbook" });
    expect(reopened.snapshot()).toEqual(model.snapshot()); expect(reopened.readWindow({ sheetId: "sheet-1", row: 0, column: 0, rows: 1, columns: 3 })).toEqual(cells);
  });
  it("checks digests, duplicate positions and missing sheets or referenced styles", () => {
    const variants: ((files: Map<string, Uint8Array>) => void)[] = [
      files => { const bytes = files.get("sheets/sheet-1/0-0.html")!; bytes[100] = bytes[100]! ^ 1; },
      files => mutateJson(files, "manifest.json", value => { value.chunks[1].row = 0; value.chunks[1].column = 0; }),
      files => mutateJson(files, "manifest.json", value => { value.chunks[0].sheetId = "missing"; }),
      files => mutateJson(files, "metadata.json", value => { value.workbook.numberFormats = []; }),
      files => { files.set("styles.css", encoder.encode('.other{color:red}')); },
    ];
    for (const edit of variants) { const files = seeded(); edit(files); expect(() => openDocument(files)).toThrow(); }
  });
  it.each([
    (w: any) => { w.activeSheetId = "missing"; },
    (w: any) => { w.sheets[0].visibility = "hidden"; },
    (w: any) => { w.sheets[0].name = "Invalid\n"; },
    (w: any) => { w.sheets.push({ ...w.sheets[0], sheetId: "second" }); },
    (w: any) => { w.sheets[0].rowDimensions = [{ from: 2, to: 1 }]; },
    (w: any) => { w.sheets[0].merges = [{ from: { row: 2, column: 0 }, to: { row: 1, column: 0 } }]; },
    (w: any) => { w.sheets[0].merges = [{ from: { row: 0, column: 0 }, to: { row: 3, column: 4 } }, { from: { row: 3, column: 4 }, to: { row: 4, column: 8 } }]; },
  ])("rejects inconsistent worksheet metadata", edit => {
    const files = seeded(); mutateJson(files, "metadata.json", value => edit(value.workbook)); expect(() => openDocument(files)).toThrow();
  });
  it("bounds the complete workbook and requested window independently", () => {
    expect(() => openDocument(seeded(), { workbookLimits: { maxCells: 4 } })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    const model = openDocument(seeded(), { kind: "workbook" });
    expect(() => model.readWindow({ sheetId: "sheet-1", row: 0, column: 0, rows: 10_000, columns: 20 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(() => model.getCell("missing", 0, 0)).toThrow(); expect(() => model.getCell("sheet-1", -1, 0)).toThrow();
  });
  it("uses strict contracts and a real manifest schema branch", () => {
    const manifest = json(createWorkbook({ documentId: "schema" }), "manifest.json");
    expect(parseContract("document-manifest", manifest)).toEqual(manifest);
    expect(DocumentManifestSchema.safeParse({ ...manifest, entrypoint: "content.html" }).success).toBe(false);
    expect(DocumentManifestSchema.safeParse({ ...manifest, unknown: true }).success).toBe(false);
    expect(JSON.stringify(getContractJsonSchema("document-manifest"))).toContain('"chunks"');
  });
});
