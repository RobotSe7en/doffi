import { execFile } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import { expect, test } from "@playwright/test";
import { createDocumentSession, openDocument } from "@doffi/core";
import { createDocumentConverter } from "@doffi/node";
import { importXlsx } from "@doffi/office";

const exec = promisify(execFile);
test.describe.configure({ timeout: 60_000 });

test("workbook PDF and PNG share the fixed-font landscape pages and repeat native titles", async () => {
  const model = await openXlsx("11-print-layout");
  const directory = await mkdtemp(join(tmpdir(), "doffi-xlsx-print-")), converter = createDocumentConverter();
  try {
    const before = model.serialize();
    const pdf = await converter.renderWorkbookPdf(model);
    const png = await converter.renderWorkbookPage(model, { pageIndex: 1 });
    expect(png.report).toEqual(pdf.report);
    expect(pdf.report).toMatchObject({ kind: "workbook", revision: 0, formulaValues: "current", formulaCells: { valid: 0 } });
    const path = join(directory, "workbook.pdf");
    await writeFile(path, pdf.bytes);
    const reference = join(process.cwd(), "examples/xlsx/reference/11-print-layout.pdf");
    const [{ stdout: actualInfo }, { stdout: referenceInfo }] = await Promise.all([exec("pdfinfo", [path]), exec("pdfinfo", [reference])]);
    expect(actualInfo).toMatch(/Pages:\s+2\b/);
    expect(referenceInfo).toMatch(/Pages:\s+2\b/);
    for (const info of [actualInfo, referenceInfo]) {
      const size = /Page size:\s+([\d.]+) x ([\d.]+) pts/.exec(info)!;
      expect(Math.abs(Number(size[1]) - 841.89)).toBeLessThan(1);
      expect(Math.abs(Number(size[2]) - 595.28)).toBeLessThan(1);
    }
    const [actual, native] = await Promise.all([extractPdfPages(path, false), extractPdfPages(reference, false)]);
    for (const pages of [actual, native]) {
      expect(pages).toHaveLength(2);
      for (const [index, page] of pages.entries()) {
        const compact = page.replace(/\s/g, "");
        expect(compact).toContain("独立打印");
        expect(compact).toContain("标题");
        expect(compact).toContain(`第${index + 1}页/2页`);
      }
      expect([...pages.join("\n").matchAll(/记录[ \t]*(\d+)(?=\s|$)/g)].map(match => "记录" + match[1])).toEqual(Array.from({ length: 44 }, (_, index) => "记录" + (index + 2)));
    }
    // Calc honors the manual row break in this source. Excel's Fit To ignores it.
    expect(native[1]).toMatch(/记录\s*24\b/);
    expect(actual[0]).toMatch(/记录\s*24\b/);
    const header = Buffer.from(png.bytes);
    expect(header.subarray(1, 4).toString()).toBe("PNG");
    expect(Math.abs(header.readUInt32BE(16) - 1123)).toBeLessThanOrEqual(1);
    expect(Math.abs(header.readUInt32BE(20) - 794)).toBeLessThanOrEqual(1);
    const imagePath = join(directory, "workbook-page-2.png");
    await writeFile(imagePath, png.bytes);
    const imageReference = join(process.cwd(), "examples/xlsx/reference/11-print-layout-2.png");
    const ink = await exec("/usr/bin/python3", ["-c", `
from PIL import Image
import sys
def bounds(path):
    with Image.open(path) as source:
        image = source.convert("L")
    return [image.crop(region).point(lambda level: 255 if level < 160 else 0).getbbox()
            for region in ((0, 0, 200, 70), (200, 725, 900, 790))]
for actual, reference in zip(bounds(sys.argv[1]), bounds(sys.argv[2])):
    assert actual and reference and max(abs(a - b) for a, b in zip(actual, reference)) <= 4, (actual, reference)
`, imagePath, imageReference]);
    expect(ink.stderr).toBe("");
    expect(model.serialize()).toEqual(before);
  } finally { await converter.close(); await rm(directory, { recursive: true, force: true }); }
});

test("workbook output renders current formula results and explicitly marks saved caches", async () => {
  const model = await openXlsx("14-preserved-formulas");
  const directory = await mkdtemp(join(tmpdir(), "doffi-xlsx-values-")), converter = createDocumentConverter();
  const unsupported = { row: 1, column: 1 }, newFormula = { row: 0, column: 2 };
  const choose = (point: typeof unsupported) => [{ sheetId: "sheet-1", areas: [{ from: point, to: point }] }];
  const session = createDocumentSession(model);
  try {
    await expect(converter.renderWorkbookPdf(model, { sheets: choose(unsupported) })).rejects.toMatchObject({ code: "UnsupportedFeature" });
    const cached = await converter.renderWorkbookPdf(model, { sheets: choose(unsupported), formulaValues: "cached", context: { locale: "en-US" } });
    expect(cached.report).toMatchObject({ formulaValues: "cached", formulaCells: { imported: 1, valid: 0 } });
    const saved = join(directory, "cached.pdf");
    await writeFile(saved, cached.bytes);
    expect((await extractPdfPages(saved))[0]).toMatch(/24.*Notrecalculated.*cachedformulavalues/);
    const { documentId, revision } = session.snapshot().manifest;
    session.commit(session.prepare({ protocolVersion: 1, documentId, baseRevision: revision, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
      { row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 7 } } },
      { ...newFormula, source: { kind: "formula", expression: "=A1*3" } },
    ] }] }));
    const current = await converter.renderWorkbookPdf(session, { sheets: choose(newFormula) });
    expect(current.report).toMatchObject({ revision: 1, formulaValues: "current", formulaCells: { valid: 1, imported: 0 } });
    const updated = join(directory, "current.pdf");
    await writeFile(updated, current.bytes);
    expect(await extractPdfPages(updated)).toEqual(["21"]);
    expect(session.history().undoSteps).toBe(1);
    await expect(converter.renderWorkbookPdf(session, { sheets: choose(newFormula), maxPaperPixels: 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    await expect(converter.renderWorkbookPdf(session, { sheets: choose(newFormula), limits: { maxPrintNodes: 6 } })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    await expect(converter.renderWorkbookPdf(session, { signal: AbortSignal.abort() })).rejects.toMatchObject({ code: "Cancelled" });
  } finally { session.dispose(); await converter.close(); await rm(directory, { recursive: true, force: true }); }
});

async function openXlsx(name: string) {
  const bytes = readFileSync(new URL(`../../examples/xlsx/${name}.xlsx`, import.meta.url));
  return openDocument(await importXlsx(bytes, { documentId: name }), { kind: "workbook" });
}
async function extractPdfPages(path: string, compact = true): Promise<string[]> {
  return (await exec("pdftotext", ["-layout", path, "-"])).stdout.split("\f")
    .map(text => compact ? text.replace(/\s/g, "") : text.trim()).filter(Boolean);
}

test("workbook PDF retains drawings, and selected sheets reflect current filters", async () => {
  const drawings = await openXlsx("08-images-shapes");
  const model = await openXlsx("12-multi-sheet-filter");
  const directory = await mkdtemp(join(tmpdir(), "doffi-xlsx-drawings-")), converter = createDocumentConverter();
  const session = createDocumentSession(model);
  try {
    const drawingPath = join(directory, "drawings.pdf");
    await writeFile(drawingPath, (await converter.renderWorkbookPdf(drawings)).bytes);
    expect((await extractPdfPages(drawingPath)).join("")).toContain("绝对位置文本框ABC中文");
    expect((await exec("pdfimages", ["-list", drawingPath])).stdout).toMatch(/\bimage\s+32\s+24\b/);

    const { documentId, revision } = session.snapshot().manifest;
    const filter = model.snapshot().metadata.workbook.sheets[1]!.autoFilter!;
    session.commit(session.prepare({ protocolVersion: 1, documentId, baseRevision: revision, operations: [
      { op: "setWorkbookCells", sheetId: "sheet-2", cells: [
        { row: 1, column: 0, source: { kind: "value", value: { type: "text", value: "待排除" } } },
        { row: 1, column: 1, source: { kind: "value", value: { type: "number", value: 17 } } },
      ] },
      { op: "setWorkbookAutoFilter", sheetId: "sheet-2", filter: { range: filter.range, columns: [{
        column: 0, hiddenButton: false, showButton: true, criterion: { kind: "values", values: ["甲"], dates: [], blank: false },
      }] } },
    ] }));
    const ordered = await converter.renderWorkbookPdf(session, { sheets: [{ sheetId: "sheet-2" }, { sheetId: "sheet-1" }] });
    expect(ordered.report.pages.map(page => page.sheetId)).toEqual(["sheet-2", "sheet-1"]);
    const path = join(directory, "filtered.pdf");
    await writeFile(path, ordered.bytes);
    const [input, overview] = await extractPdfPages(path);
    expect(input).toContain("来源");
    expect(input).not.toMatch(/待排除|17|乙/);
    expect(overview).toContain("37");
    await expect(converter.renderWorkbookPdf(session, { sheets: [{ sheetId: "sheet-3" }] })).rejects.toMatchObject({ code: "InvalidOperation" });
    await expect(converter.renderWorkbookPdf(session, { limits: { maxPages: 1 } })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  } finally { session.dispose(); await converter.close(); await rm(directory, { recursive: true, force: true }); }
});
