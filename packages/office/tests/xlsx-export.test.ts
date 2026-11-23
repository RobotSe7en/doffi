import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createDocumentSession, createWorkbook, DEFAULT_DOCUMENT_READ_LIMITS, iterateWorkbookRows, openDocument, SourceStyleResolver } from "@doffi/core";
import type { DocumentOperation, WorkbookCell, WorkbookSession } from "@doffi/core";
import { exportXlsx, importXlsx } from "@doffi/office";
import { dailyCi } from "../../../tests/ci-mode.js";
import { readDocumentZip } from "../src/document-zip.js";

const open = async (bytes: Uint8Array, documentId = "export-reopen") =>
  openDocument(await importXlsx(bytes, { documentId }), { kind: "workbook" });
const source = (name: string) => readFileSync("examples/xlsx/" + name + ".xlsx");
const value = (row: number, column: number, number: number) => ({
  row, column, source: { kind: "value", value: { type: "number", value: number } } as const,
});
function commit(session: WorkbookSession, operations: readonly DocumentOperation[]): void {
  const { documentId, revision } = session.snapshot().manifest;
  session.commit(session.prepare({ protocolVersion: 1, documentId, baseRevision: revision, operations: [...operations] }));
}
function part(bytes: Uint8Array, path: string): string {
  return new TextDecoder().decode(readDocumentZip(bytes).get(path)!);
}

describe("editable native XLSX export", () => {
  it("uses the current revision and reports omitted caches without replaying imported values", async () => {
    const session = createDocumentSession(await open(source("14-preserved-formulas"), "current-revision"));
    try {
      const old = session.calculateSync();
      commit(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [value(0, 0, 7), {
        row: 0, column: 2, source: { kind: "formula", expression: '=A1*3' },
      }] }]);
      await expect(exportXlsx(session, { calculation: old })).rejects.toMatchObject({ code: "RevisionConflict" });
      const before = session.serialize();
      const change = session.calculateSync();
      const result = await exportXlsx(session, { calculation: change });
      expect(result.report).toMatchObject({ documentId: "current-revision", revision: 1, formulaCaches: { written: 1, omitted: 3 } });
      const reopened = await open(result.bytes);
      expect(reopened.getCell("sheet-1", 0, 2)).toMatchObject({ value: { type: "number", value: 21 }, formula: { expression: "=A1*3", cacheState: "imported" } });
      expect(reopened.getCell("sheet-1", 1, 1)).toMatchObject({ value: { type: "blank" }, formula: { cacheState: "unavailable" } });
      expect(reopened.snapshot().metadata.workbook.calculationMode).toBe("manual");
      expect(session.serialize()).toEqual(before);
      expect(session.history().undoSteps).toBe(1);
    } finally { session.dispose(); }
  });

  it("round-trips edited advanced appearance and SpreadsheetML escape literals", async () => {
    const session = createDocumentSession(openDocument(createWorkbook({ documentId: "native-appearance" }), { kind: "workbook" }));
    try {
      const at = { row: 0, column: 0 }, next = { row: 0, column: 1 };
      commit(session, [
        { op: "setWorkbookCells", sheetId: "sheet-1", cells: [
          { ...at, source: { kind: "value", value: { type: "text", value: " _x0041_\r\n😀禰󠄀 " } } },
          { ...next, source: { kind: "value", value: { type: "text", value: "竖排" } } },
        ] },
        { op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: at, to: at }],
          properties: { "--office-cell-text-rotation": "135", "--office-cell-border-diagonal-kind": "thin",
            "--office-cell-border-diagonal-up": "true", "--office-cell-border-diagonal-color": "#003344",
            "--office-cell-fill-kind": "gradient", "--office-cell-gradient-kind": "linear", "--office-cell-gradient-stops": "2",
            "--office-cell-gradient-degree": "35", "--office-cell-gradient-stop-0-position": "0",
            "--office-cell-gradient-stop-0-color": "#224466", "--office-cell-gradient-stop-1-position": "1", "--office-cell-gradient-stop-1-color": "#aabbcc" } },
        { op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: next, to: next }],
          properties: { "--office-cell-text-rotation": "255", "--office-cell-fill-kind": "darkTrellis", "--office-cell-fill-color": "#ffaa00" } },
      ]);
      const { bytes } = await exportXlsx(session);
      const output = await open(bytes), snapshot = output.snapshot();
      const css = new SourceStyleResolver(snapshot.styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS);
      const first = output.getCell("sheet-1", 0, 0)!;
      expect(first.value).toEqual({ type: "text", value: " _x0041_\r\n😀禰󠄀 " });
      expect(css.specifiedClass(first.cssClass!)).toMatchObject({
        "--office-cell-text-rotation": "135", "--office-cell-border-diagonal-kind": "thin", "--office-cell-border-diagonal-up": "true",
        "--office-cell-border-diagonal-color": "#003344", "--office-cell-fill-kind": "gradient", "--office-cell-gradient-degree": "35",
        "--office-cell-gradient-stop-1-color": "#aabbcc",
      });
      expect(css.specifiedClass(output.getCell("sheet-1", 0, 1)!.cssClass!)).toMatchObject({
        "--office-cell-text-rotation": "255", "--office-cell-fill-kind": "darkTrellis",
      });
    } finally { session.dispose(); }
  });

  it("retains an authored worksheet text box line height when reopening XLSX", async () => {
    const session = createDocumentSession(openDocument(createWorkbook({ documentId: "drawing-line-height" }), { kind: "workbook" }));
    try {
      commit(session, [{
        op: "insertWorkbookObject", sheetId: "sheet-1",
        anchor: { kind: "absolute", xPt: 20, yPt: 20, widthPt: 240, heightPt: 100 },
        object: { kind: "text", html: '<p style="font-size:18pt;line-height:1.8">Editable text</p>' },
      }]);
      const { bytes } = await exportXlsx(session);
      expect(part(bytes, "xl/drawings/drawing1.xml")).toContain('<a:lnSpc><a:spcPct val="180000"/></a:lnSpc>');
      const reopened = await open(bytes);
      expect(reopened.snapshot().styles.map(style => style.css).join("\n")).toMatch(/line-height:\s*1\.8\b/);
    } finally { session.dispose(); }
  });

  it("exports a representative cell and reports an occupied cell without replacing its comment", async () => {
    const session = createDocumentSession(openDocument(createWorkbook({ documentId: "range-comments" }), { kind: "workbook" }));
    try {
      const a = { row: 1, column: 1 }, b = { row: 2, column: 3 };
      commit(session, [
        { op: "addWorkbookComment", format: "modern", author: { name: "审核人" }, date: "2026-11-20T07:00:00Z", html: "<p>范围😀</p>",
          anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: a, to: b }, { from: { row: 4, column: 0 }, to: { row: 4, column: 1 } }] } },
        { op: "addWorkbookComment", format: "classic", html: "<p>保留在源中</p>",
          anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: a, to: a }] } },
      ]);
      const [first, second] = session.snapshot().annotations;
      const result = await exportXlsx(session);
      expect(result.report.exportedAnnotationIds).toEqual([first!.annotationId]);
      expect(result.report.unexportedAnnotations).toEqual([{ annotationId: second!.annotationId, reason: "representative_collision" }]);
      expect(result.report.losses.filter(loss => loss.annotationId)).toEqual([expect.objectContaining({ annotationId: first!.annotationId, kind: "anchor" })]);
      const model = await open(result.bytes), [thread] = model.snapshot().annotations;
      expect(thread?.anchor).toEqual({ kind: "cells", sheetId: "sheet-1", ranges: [{ from: a, to: a }] });
      expect(model.getNode(thread!.bodyNodeId!)?.text).toContain("[doffi ranges: B2:D3, A5:B5]");
      expect(session.snapshot().annotations).toHaveLength(2);
    } finally { session.dispose(); }
  });

  it("retains relative defined-name semantics when rebasing the native origin to A1", async () => {
    const session = createDocumentSession(openDocument(createWorkbook({ documentId: "relative-name" }), { kind: "workbook" }));
    try {
      const sheetId = "sheet-1", origin = { sheetId, row: 2, column: 2 };
      commit(session, [{ op: "setWorkbookDefinedName", name: {
        nameId: "relative-name", sheetId, name: "RelativeTarget", expression: "=A1", origin: { row: 2, column: 2 },
      } }]);
      const expected = session.resolveReference("=RelativeTarget", origin);
      const { bytes } = await exportXlsx(session);
      const reopened = await open(bytes);
      expect(reopened.resolveReference("=RelativeTarget", origin)).toEqual(expected);
      expect(expected).toEqual({ kind: "ranges", areas: [{ sheetId, from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] });
    } finally { session.dispose(); }
  });

  it("writes the active filter and its current visibility into the worksheet", async () => {
    const session = createDocumentSession(openDocument(createWorkbook({ documentId: "native-filter" }), { kind: "workbook" }));
    try {
      const sheetId = "sheet-1";
      commit(session, [
        { op: "setWorkbookCells", sheetId, cells: [
          { row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "Amount" } } },
          value(1, 0, 4), value(2, 0, 9), value(3, 0, 2),
        ] },
        { op: "updateWorkbookSheet", sheetId, changes: { rowDimensions: [{ from: 3, to: 3, hidden: true }] } },
        { op: "setWorkbookAutoFilter", sheetId, filter: {
          range: { from: { row: 0, column: 0 }, to: { row: 3, column: 0 } },
          columns: [{ column: 0, hiddenButton: false, showButton: true,
            criterion: { kind: "custom", and: false, tests: [{ operator: "greaterThan", value: "5" }] } }],
        } },
      ]);
      const { bytes } = await exportXlsx(session), xml = part(bytes, "xl/worksheets/sheet1.xml");
      expect(xml).toContain('<row r="2" hidden="1">');
      expect(xml).toContain('<row r="3">');
      expect(xml).toContain('<row r="4" hidden="1">');
      expect(xml.match(/<doffi:filterHidden\/>/g)).toHaveLength(1);
      const reopened = await open(bytes);
      expect(reopened.snapshot().metadata.workbook.sheets[0]!.autoFilter?.columns[0]!.criterion)
        .toEqual(session.snapshot().metadata.workbook.sheets[0]!.autoFilter!.columns[0]!.criterion);
      expect(reopened.snapshot().metadata.workbook.sheets[0]!.rowDimensions).toEqual([{ from: 3, to: 3, hidden: true }]);
      expect(reopened.getCell(sheetId, 1, 0)?.value).toEqual({ type: "number", value: 4 });
    } finally { session.dispose(); }
  });

  it("exports current font choices and produces valid native enums for CSS alignment and borders", async () => {
    const session = createDocumentSession(await open(source("09-classic-notes")));
    try {
      const at = { row: 0, column: 0 }, sheetId = "sheet-1";
      commit(session, [{ op: "formatWorkbookRanges", sheetId, ranges: [{ from: at, to: at }], properties: {
        "--office-font-ascii": '"Courier New"', "--office-font-h-ansi": '"Courier New"', "--office-font-east-asia": '"SimHei"',
      } }, { op: "updateWorkbookSheet", sheetId, changes: { defaultColumnWidthUnit: "baseCharacters", defaultColumnWidth: 8.5 } }]);
      const { bytes, report } = await exportXlsx(session);
      const model = await open(bytes), snapshot = model.snapshot();
      expect(snapshot.metadata.workbook.sheets[0]!.defaultColumnWidth).toBe(9);
      const resolver = new SourceStyleResolver(snapshot.styles.map(sheet => sheet.css), DEFAULT_DOCUMENT_READ_LIMITS);
      const properties = resolver.specifiedClass(model.getCell(sheetId, 0, 0)!.cssClass!);
      expect(properties["--office-font-ascii"]).toBe('"Courier New"');
      expect(report.losses.map(loss => loss.message)).toEqual(expect.arrayContaining([
        expect.stringContaining("script-specific typefaces"), expect.stringContaining("base column width"),
      ]));
    } finally { session.dispose(); }
    const authored = createDocumentSession(openDocument(createWorkbook({ documentId: "css-appearance" }), { kind: "workbook" }));
    try {
      const at = { row: 0, column: 0 }, sheetId = "sheet-1";
      commit(authored, [{ op: "setWorkbookCells", sheetId, cells: [value(0, 0, 1)] },
        { op: "formatWorkbookRanges", sheetId, ranges: [{ from: at, to: at }], properties: {
          "text-align": "start", "border-bottom-style": "groove", "border-bottom-color": "#123456",
        } }]);
      const { bytes, report } = await exportXlsx(authored);
      expect(part(bytes, "xl/styles.xml")).not.toMatch(/horizontal="start"|style="groove"/);
      expect(report.losses.map(loss => loss.message)).toContainEqual(expect.stringContaining("groove"));
      const reopened = await open(bytes);
      const resolver = new SourceStyleResolver(reopened.snapshot().styles.map(sheet => sheet.css), DEFAULT_DOCUMENT_READ_LIMITS);
      expect(resolver.specifiedClass(reopened.getCell(sheetId, 0, 0)!.cssClass!)["--office-cell-border-bottom-kind"]).toBe("thin");
    } finally { authored.dispose(); }
  });

  it("writes differential built-in formats with the required native format code", async () => {
    const session = createDocumentSession(await open(source("06-rules-validation")));
    try {
      const sheet = session.snapshot().metadata.workbook.sheets[0]!;
      const rule = sheet.conditionalRules!.find(rule => rule.kind === "cellIs")!;
      commit(session, [{ op: "setWorkbookNumberFormat", format: { formatId: "percentage", builtinId: 10 } },
        { op: "formatWorkbookConditionalRule", sheetId: sheet.sheetId, ruleId: rule.ruleId, properties: {}, numberFormatId: "percentage" }]);
      const { bytes } = await exportXlsx(session);
      const reopened = (await open(bytes)).snapshot().metadata.workbook;
      const formatId = reopened.sheets[0]!.conditionalRules!.find(rule => rule.kind === "cellIs")!.numberFormatId;
      expect(reopened.numberFormats.find(format => format.formatId === formatId)).toMatchObject({ code: "0.00%" });
    } finally { session.dispose(); }
  });

  it.each([
    ["vertical", "middle"], ["reading-order", "sideways"],
    ["text-rotation", "220"], ["fill-kind", "stripe"],
  ])("rejects a resolved invalid native %s before writing XLSX bytes", async (property, invalid) => {
    const session = createDocumentSession(openDocument(createWorkbook({ documentId: "invalid-resolved-style" }), { kind: "workbook" }));
    try {
      const at = { row: 0, column: 0 }, sheetId = "sheet-1";
      commit(session, [{ op: "setWorkbookCells", sheetId, cells: [value(0, 0, 1)] },
        { op: "formatWorkbookRanges", sheetId, ranges: [{ from: at, to: at }], properties: {
          [`--office-cell-${property}`]: `var(--absent, ${invalid})`,
        } }]);
      await expect(exportXlsx(session)).rejects.toMatchObject({ code: "UnsupportedFeature" });
    } finally { session.dispose(); }
  });

  it("cancels without returning bytes, guards edits and bounds the archive", async () => {
    const model = await open(source("04-formulas-names"));
    for (const phase of ["calculate", "worksheet", "package"] as const) {
      const signal = new AbortController();
      await expect(exportXlsx(model, { signal: signal.signal, onProgress(progress) {
        if (progress.phase === phase) signal.abort();
      } })).rejects.toMatchObject({ code: "OperationCancelled" });
    }
    await expect(exportXlsx(model, { nativeLimits: { maxZipBytes: 100 } })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    await expect(exportXlsx(model, { nativeLimits: { maxFiles: 2 } })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    const session = createDocumentSession(model);
    try {
      await expect(exportXlsx(session, { onProgress(progress) {
        if (progress.phase === "worksheet" && progress.completed === 0)
          commit(session, [{ op: "setWorkbookCells", sheetId: model.snapshot().metadata.workbook.sheets[0]!.sheetId, cells: [value(7, 7, 8)] }]);
      } })).rejects.toMatchObject({ code: "RevisionConflict" });
    } finally { session.dispose(); }
  });

  for (const scale of [false, true]) it.skipIf(scale && dailyCi)("reopens the independent and producer XLSX corpus" + (scale ? " @scale" : ""), async () => {
    for (const directory of ["examples/xlsx", "examples/xlsx/producer"]) for (const name of readdirSync(directory).filter(name => name.endsWith(".xlsx") && name.startsWith("15-") === scale)) {
      const input = await open(readFileSync(directory + "/" + name)), original = input.snapshot();
      const { bytes, report } = await exportXlsx(input);
      const reopened = await open(bytes), snapshot = reopened.snapshot();
      const book = snapshot.metadata.workbook, expected = original.metadata.workbook;
      expect(book.sheets.map(sheet => [sheet.name, sheet.visibility, sheet.frozen, sheet.merges]), name)
        .toEqual(expected.sheets.map(sheet => [sheet.name, sheet.visibility, sheet.frozen, sheet.merges]));
      expect(book.sheets.map(sheet => [sheet.tables?.length, sheet.conditionalRules?.length, sheet.dataValidations?.length]), name)
        .toEqual(expected.sheets.map(sheet => [sheet.tables?.length, sheet.conditionalRules?.length, sheet.dataValidations?.length]));
      expect(book.objects?.map(item => item.kind), name).toEqual(expected.objects?.map(item => item.kind));
      expect(book.charts?.map(item => item.kind), name).toEqual(expected.charts?.map(item => item.kind));
      expect(snapshot.annotations, name).toHaveLength(original.annotations.length);
      expect(report.unexportedAnnotations, name).toEqual([]);
      for (const [index, sheet] of expected.sheets.entries()) {
        const current = iterateWorkbookRows(reopened, book.sheets[index]!.sheetId);
        const compare = (cells: readonly WorkbookCell[]) => cells.map(cell => [cell.row, cell.column, cell.formula?.expression ?? cell.value]);
        for (const original of iterateWorkbookRows(input, sheet.sheetId)) {
          const row = current.next();
          expect(row.done, name).toBe(false);
          if (!row.done) expect(compare(row.value), name).toEqual(compare(original));
        }
        expect(current.next().done, name).toBe(true);
      }
      if (scale) {
        expect(report.formulaCaches).toEqual({ written: 20_000, omitted: 0 });
        expect(reopened.getCell(book.sheets[0]!.sheetId, 9999, 19)?.value).toEqual({ type: "number", value: 3_400_152 });
      }
      if (name.startsWith("07-")) expect(part(bytes, "xl/charts/chart1.xml")).toMatch(/<c:(?:numRef|numLit)>/);
    }
  }, scale ? 240_000 : 90_000);
});
