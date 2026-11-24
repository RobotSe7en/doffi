import { dailyCi } from '../../../tests/ci-mode.js';
import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createDocumentSession, DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, iterateWorkbookRows, openDocument, SourceStyleResolver } from "@doffi/core";
import { exportXlsx, importXlsx } from "@doffi/office";
import type { OfficeImportProgress, XlsxImportOptions } from "@doffi/office";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";

const source = (name: string) => readFileSync(`examples/xlsx/${name}.xlsx`);
const decoder = new TextDecoder(), encoder = new TextEncoder();
const filterName = (reference: string, localSheetId?: number) =>
  `<definedName name="_xlnm._FilterDatabase"${localSheetId === undefined ? "" : ` localSheetId="${localSheetId}"`}>${reference}</definedName>`;
function workbookWithNames(...names: string[]): Uint8Array {
  const files = new Map(readDocumentZip(source("12-multi-sheet-filter")));
  const part = "xl/workbook.xml";
  files.set(part, encoder.encode(decoder.decode(files.get(part)!).replace("</definedNames>", names.join("") + "</definedNames>")));
  return writeOpcZip(files);
}
interface Expected { file: string; sheets: string[]; cellCount: number; formulaCount: number; values?: Record<string, number | string | boolean> }
const expected = JSON.parse(readFileSync("examples/xlsx/expected.json", "utf8")).cases as Expected[];

describe("public native workbook import", () => {
  for (const scale of [false, true]) it.skipIf(dailyCi && scale)("opens all 32 independent and producer sources through unified openDocument with their complete feature package" + (scale ? " @scale" : " (small corpus)"), async () => {
    let count = 0;
    for (const directory of ["examples/xlsx", "examples/xlsx/producer"]) for (const name of readdirSync(directory).filter(name => name.endsWith(".xlsx") && name.startsWith("15-") === scale)) {
      const documentId = "public-native-" + (scale ? "scale-" : "small-") + ++count;
      const files = await importXlsx(readFileSync(`${directory}/${name}`), { documentId });
      const model = openDocument(files, { kind: "workbook" }), snapshot = model.snapshot(), book = snapshot.metadata.workbook;
      expect(snapshot.manifest).toMatchObject({ documentId, kind: "workbook", revision: 0 });
      for (const annotation of snapshot.annotations) expect(annotation.documentId).toBe(documentId);
      if (directory === "examples/xlsx") {
        const fixture = expected.find(item => item.file === name)!;
        expect(book.sheets.map(sheet => sheet.name)).toEqual(fixture.sheets);
        let cells = 0, formulas = 0;
        for (const chunk of snapshot.manifest.chunks) {
          const html = decoder.decode(files.get(chunk.path)!); cells += [...html.matchAll(/<td\b/g)].length; formulas += [...html.matchAll(/ data-formula=/g)].length;
        }
        expect({ cells, formulas }, name).toEqual({ cells: fixture.cellCount, formulas: fixture.formulaCount });
        for (const [address, value] of Object.entries(fixture.values ?? {})) {
          const [sheetName, coordinate] = address.split("!"), sheet = book.sheets.find(sheet => sheet.name === sheetName)!;
          const match = /^([A-Z]+)(\d+)$/.exec(coordinate!)!; let column = 0;
          for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64;
          expect(model.getCell(sheet.sheetId, Number(match[2]) - 1, column - 1)?.value, address).toMatchObject({ value });
        }
        if (name.startsWith("05-")) expect(book.sheets[0]!.tables?.map(table => table.name)).toEqual(["SalesTable"]);
        if (name.startsWith("06-")) { expect(book.sheets[0]!.conditionalRules).toHaveLength(4); expect(book.sheets[0]!.dataValidations).toHaveLength(7); }
        if (name.startsWith("07-")) { expect(book.charts?.map(chart => chart.kind)).toEqual(["column", "line", "pie"]); expect(book.objects).toHaveLength(3); }
        if (name.startsWith("08-")) { expect(book.objects).toHaveLength(3); expect(snapshot.manifest.assets).toHaveLength(1); }
        if (name.startsWith("09-")) expect(snapshot.annotations).toHaveLength(1);
        if (name.startsWith("10-")) { expect(snapshot.annotations).toHaveLength(2); expect(snapshot.annotations.flatMap(note => note.replies ?? [])).toHaveLength(1); }
        if (name.startsWith("14-")) { expect(book.calculationMode).toBe("manual"); expect(model.getCell("sheet-1", 1, 1)?.formula?.cacheState).toBe("imported"); expect(model.getCell("sheet-1", 2, 1)?.formula?.cacheState).toBe("unavailable"); }
      }
      if (name.startsWith("15-")) { expect(snapshot.manifest.chunks).toHaveLength(40); expect(model.getCell("sheet-1", 9999, 19)?.value).toEqual({ type: "number", value: 3_400_152 }); }
      expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(snapshot);
    }
    expect(count).toBe(scale ? 2 : 30);
  });
  it("resolves an unscoped filter name by its qualified range without activating a filter", async () => {
    const options = { documentId: "unscoped-filter" };
    const original = openDocument(await importXlsx(source("12-multi-sheet-filter"), options), { kind: "workbook" });
    // Overview has no auto-filter. Input has an active filter and is not the active tab.
    for (const reference of ["'Overview'!$A$1:$B$3", "'Input'!$A$1:$B$3"]) {
      const files = await importXlsx(workbookWithNames(filterName(reference)), options);
      expect(openDocument(files, { kind: "workbook" }).snapshot()).toEqual(original.snapshot());
    }
  });
  it("rejects conflicting or invalid filter names and keeps print names worksheet-scoped", async () => {
    const options = { documentId: "invalid-filter" };
    for (const names of [
      [filterName("$A$1:$B$3")],
      [filterName("Missing!$A$1:$B$3")],
      [filterName("Overview!$A$1:B3")],
      [filterName("Input!$A$1:$B$4")],
      [filterName("Overview!$A$1:$B$3", 0)],
      [filterName("Input!$A$1:$B$3", 1), filterName("Input!$A$1:$B$3")],
      [filterName("Overview!$A$1:$B$3"), filterName("Input!$A$1:$B$3")],
      ['<definedName name="_xlnm.Print_Area">Overview!$A$1:$B$3</definedName>'],
    ]) {
      await expect(importXlsx(workbookWithNames(...names), options)).rejects.toMatchObject({ code: "InvalidDocument" });
    }
    await expect(importXlsx(workbookWithNames(filterName("Overview!$A$1:Input!$B$3")), options))
      .rejects.toMatchObject({ code: "UnsupportedFeature" });
  });
  it("preserves ordered Excel theme saturation and luminance transforms", async () => {
    const files = new Map(readDocumentZip(source("12-multi-sheet-filter"))), part = "xl/theme/theme1.xml";
    const xml = decoder.decode(files.get(part)!);
    const transforms = '<a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/>';
    const background = '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr">';
    const edited = xml.replace(background, background + transforms);
    expect(edited).not.toBe(xml);
    files.set(part, encoder.encode(edited));
    const model = openDocument(await importXlsx(writeOpcZip(files), { documentId: "ordered-theme-color" }), { kind: "workbook" });
    const css = new SourceStyleResolver(model.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS);
    expect(css.specifiedClass("xlsx-theme-background-1")).toMatchObject({
      "--office-fill-color-transforms": "tint 93000 satMod 150000 shade 98000 lumMod 102000",
    });
  });
  it("retains row and column outlines alongside text with inactive phonetic settings through XLSX reopen", async () => {
    const files = new Map(readDocumentZip(source("01-values-text")));
    const sheetPart = "xl/worksheets/sheet1.xml", stringsPart = "xl/sharedStrings.xml";
    const phonetic = '<phoneticPr fontId="1" type="noConversion"/>';
    const sheet = decoder.decode(files.get(sheetPart)!).replace("<sheetPr>", '<sheetPr codeName="功能表">')
      .replace("<sheetFormatPr ", '<sheetFormatPr outlineLevelRow="1" outlineLevelCol="2" ')
      .replace("<sheetData>", '<cols><col min="2" max="2" outlineLevel="2" hidden="1"/><col min="3" max="3" collapsed="1"/></cols><sheetData>')
      .replace('<row r="1">', '<row r="1" outlineLevel="1" hidden="1">')
      .replace('<row r="2">', '<row r="2" outlineLevel="1" hidden="1">')
      .replace('<row r="3">', '<row r="3" outlineLevel="1" collapsed="1">')
      .replace("</sheetData>", "</sheetData>" + phonetic);
    files.set(sheetPart, encoder.encode(sheet));
    files.set(stringsPart, encoder.encode(decoder.decode(files.get(stringsPart)!).replaceAll("</si>", phonetic + "</si>")));
    const options = { documentId: "imported-outlines" };
    const model = openDocument(await importXlsx(writeOpcZip(files), options), { kind: "workbook" });
    const metadata = model.snapshot().metadata.workbook.sheets[0]!;
    expect(metadata.rowDimensions).toEqual([
      { from: 0, to: 1, outlineLevel: 1, hidden: true },
      { from: 2, to: 2, outlineLevel: 1, collapsed: true },
    ]);
    expect(metadata.columnDimensions).toEqual([
      { from: 1, to: 1, unit: "characters", outlineLevel: 2, hidden: true },
      { from: 2, to: 2, unit: "characters", collapsed: true },
    ]);
    const plain = openDocument(await importXlsx(source("01-values-text"), options), { kind: "workbook" });
    expect([...iterateWorkbookRows(model, metadata.sheetId)]).toEqual([...iterateWorkbookRows(plain, metadata.sheetId)]);
    const session = createDocumentSession(model);
    try {
      session.commit(session.prepare({ protocolVersion: 1, documentId: options.documentId, baseRevision: 0,
        operations: [{ op: "editWorkbookAxis", sheetId: metadata.sheetId, axis: "row", action: "insert", index: 2, count: 1, formatSource: "after" }] }));
      const { bytes } = await exportXlsx(session);
      const reopened = openDocument(await importXlsx(bytes, options), { kind: "workbook" }).snapshot().metadata.workbook.sheets[0]!;
      expect(reopened.rowDimensions).toEqual([metadata.rowDimensions[0], { from: 3, to: 3, outlineLevel: 1, collapsed: true }]);
      expect(reopened.columnDimensions).toEqual(metadata.columnDimensions);
    } finally { session.dispose(); }
  });
  it("rejects all 17 negative sources and late unmapped parts without exposing a partial package", async () => {
    const names = readdirSync("examples/xlsx/invalid").filter(name => name.endsWith(".xlsx")); expect(names).toHaveLength(17);
    for (const name of names) await expect(importXlsx(readFileSync(`examples/xlsx/invalid/${name}`), { documentId: "invalid" }), name).rejects.toBeInstanceOf(DocumentError);
    await expect(importXlsx(readFileSync("examples/docx/06-links-tabs.docx"), { documentId: "wrong-kind" })).rejects.toMatchObject({ code: "UnsupportedFormat" });
    const files = new Map(readDocumentZip(source("08-images-shapes"))), part = "xl/drawings/drawing1.xml";
    files.set(part, encoder.encode(decoder.decode(files.get(part)!).replace('prst="rect"', 'prst="unsupported-shape"')));
    await expect(importXlsx(writeOpcZip(files), { documentId: "late-invalid" })).rejects.toMatchObject({ code: "UnsupportedFeature" });
  });
  it("cancels at every phase including after validation, and does not swallow host callback errors", async () => {
    const bytes = source("09-classic-notes");
    for (const phase of ["zip", "metadata", "worksheet", "content", "validate"] as const) {
      const signal = { aborted: false }; let delivered = 0;
      await expect(importXlsx(bytes, { documentId: "cancelled", signal, onProgress(progress) {
        delivered++; expect(Object.isFrozen(progress)).toBe(true);
        if (progress.phase === phase && (phase !== "validate" || progress.completed === 1)) signal.aborted = true;
      } })).rejects.toMatchObject({ code: "OperationCancelled" });
      expect(delivered).toBeGreaterThan(0);
    }
    await expect(importXlsx(bytes, { documentId: "pre-cancel", signal: { aborted: true }, onProgress() { throw new Error("must not run"); } })).rejects.toMatchObject({ code: "OperationCancelled" });
    await expect(importXlsx(bytes, { documentId: "bad-host", onProgress(progress) { if (progress.phase === "validate") throw new Error("private host failure"); } })).rejects.toMatchObject({ code: "HostSubmissionFailed", message: "The host import progress callback failed." });
  });
  it("captures host identity, hooks, limits and input bytes before yielding, isolating concurrent imports", async () => {
    const bytes = source("09-classic-notes"), progress: OfficeImportProgress[] = [];
    const options = { documentId: "first", signal: { aborted: false }, limits: { maxFiles: 2048 }, onProgress(event: OfficeImportProgress) { progress.push(event); } };
    const first = importXlsx(bytes, options);
    bytes.fill(0); options.documentId = "changed"; options.signal = { aborted: true }; options.limits.maxFiles = 1;
    options.onProgress = () => { throw new Error("replaced callback must not run"); };
    const second = importXlsx(source("09-classic-notes"), { documentId: "second" });
    const [a, b] = await Promise.all([first, second]), left = openDocument(a, { kind: "workbook" }), right = openDocument(b, { kind: "workbook" });
    expect(left.snapshot().annotations[0]!.documentId).toBe("first"); expect(right.snapshot().annotations[0]!.documentId).toBe("second");
    expect(progress.at(-1)).toEqual({ phase: "validate", completed: 1, total: 1 });
    const serialized = left.serialize(), chunk = left.snapshot().manifest.chunks[0]!.path; a.get(chunk)!.fill(0); serialized.get(chunk)!.fill(0);
    expect(left.getCell("sheet-1", 1, 1)).toEqual(right.getCell("sheet-1", 1, 1));
  });
  it("checks combined generated source budgets as well as native bytes and public options", async () => {
    const bytes = source("09-classic-notes"), files = await importXlsx(bytes, { documentId: "budget" });
    const total = [...files.values()].reduce((sum, bytes) => sum + bytes.byteLength, 0);
    expect([...files.values()].every(bytes => bytes.byteLength < total - 1)).toBe(true);
    for (const options of [{ limits: { maxPackageBytes: total - 1 } }, { limits: { maxFiles: files.size - 1 } }, { nativeLimits: { maxZipBytes: bytes.length - 1 } }])
      await expect(importXlsx(bytes, { documentId: "budget", ...options })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    for (const documentId of ["", "bad\u0000identity", "bad\ud800identity"]) await expect(importXlsx(bytes, { documentId })).rejects.toMatchObject({ code: "InvalidDocument" });
    await expect(importXlsx(bytes, { documentId: "bad-option", partial: true } as unknown as XlsxImportOptions)).rejects.toMatchObject({ code: "InvalidOperation" });
  });
});
