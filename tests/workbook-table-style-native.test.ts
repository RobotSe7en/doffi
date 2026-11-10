import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, DEFAULT_DOCUMENT_READ_LIMITS, openDocument, SourceStyleResolver } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

for (const region of ["wholeTable", "headerRow"] as const) test(`independent native table style ${region} edits its current regional source and reopens the package`, async () => {
  const model = openDocument(await importXlsx(readFileSync("packages/office/tests/fixtures/workbook-table-style.xlsx"), { documentId: "table-style-native" }), { kind: "workbook" }), session = createDocumentSession(model);
  try {
    const before = session.snapshot(), style = before.metadata.workbook.tableStyles![0]!, element = style.elements.find(element => element.region === region)!;
    const oldClass = element.style!.cssClass, original = new SourceStyleResolver(before.styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set()).classStyle(oldClass);
    session.commit(session.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: before.manifest.revision,
      operations: [{ op: "formatWorkbookTableStyleElement", styleId: style.styleId, region: element.region, properties: { "background-color": "#33bb88", "--office-font-east-asia": '"Noto Serif CJK SC"' } }] }));
    const after = session.snapshot(), next = after.metadata.workbook.tableStyles!.find(value => value.styleId === style.styleId)!.elements.find(value => value.region === element.region)!;
    const resolver = new SourceStyleResolver(after.styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set()); expect(next.style!.cssClass).not.toBe(oldClass); expect(resolver.classStyle(next.style!.cssClass)).toMatchObject({ "background-color": "#33bb88", "--office-cell-fill-color": "#33bb88" }); expect(resolver.classStyle(oldClass)).toEqual(original);
    expect(after.manifest.chunks).toEqual(before.manifest.chunks); expect(after.metadata.workbook.sheets).toEqual(before.metadata.workbook.sheets); expect(after.metadata.workbook.defaultTableStyle).toEqual(before.metadata.workbook.defaultTableStyle);
    const reopened = openDocument(readDocumentZip(writeDocumentZip(session.serialize())), { kind: "workbook" }); expect(reopened.snapshot().metadata).toEqual(after.metadata); expect(reopened.snapshot().styles).toEqual(after.styles);
    expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.snapshot().metadata).toEqual(before.metadata); session.redo(); expect(session.snapshot().metadata).toEqual(after.metadata);
  } finally { session.dispose(); }
});

for (const name of ["resave", "edit"]) test(`pinned Calc ${name} keeps its documented custom-style loss during current input and package reopen`, async () => {
  const model = openDocument(await importXlsx(readFileSync(`packages/office/tests/fixtures/workbook-table-style-${name}-calc.xlsx`), { documentId: "table-style-producer" }), { kind: "workbook" }), session = createDocumentSession(model);
  try {
    const before = session.snapshot(); expect(before.metadata.workbook.tableStyles ?? []).toEqual([]); expect(before.metadata.workbook.defaultTableStyle).toBeUndefined();
    expect(before.metadata.workbook.sheets[0]!.tables![0]!.style).toBeUndefined(); expect(session.getCell("sheet-1", 1, 1)?.value).toEqual({ type: "number", value: name === "edit" ? 9 : 2 });
    session.commit(session.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: before.manifest.revision,
      operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 1, column: 1, source: { kind: "value", value: { type: "number", value: 11 } } }] }] }));
    expect(session.snapshot().metadata).toEqual(before.metadata); const reopened = openDocument(readDocumentZip(writeDocumentZip(session.serialize())), { kind: "workbook" });
    expect(reopened.getCell("sheet-1", 1, 1)?.value).toEqual({ type: "number", value: 11 }); expect(reopened.snapshot().metadata).toEqual(before.metadata);
    expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.getCell("sheet-1", 1, 1)?.value).toEqual({ type: "number", value: name === "edit" ? 9 : 2 });
  } finally { session.dispose(); }
});
