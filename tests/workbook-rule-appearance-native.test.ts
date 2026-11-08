import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, DEFAULT_DOCUMENT_READ_LIMITS, openDocument, SourceStyleResolver } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

for (const index of [0, 1, 2, 3]) test(`native conditional definition ${index + 1} paints independently and reopens its package with one-step history`, async () => {
  const model = openDocument(await importXlsx(readFileSync("examples/xlsx/06-rules-validation.xlsx"), { documentId: "rule-appearance-native" }), { kind: "workbook" }), session = createDocumentSession(model);
  try {
    const before = session.snapshot(), source = before.metadata.workbook.sheets[0]!.conditionalRules![index]!, oldClasses = new SourceStyleResolver(before.styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set()),
      original = source.kind === "colorScale" ? source.stops.at(-1)!.cssClass : source.cssClass!, expected = oldClasses.classStyle(original);
    const operation = { op: "formatWorkbookConditionalRule" as const, sheetId: "sheet-1", ruleId: source.ruleId, ...(source.kind === "colorScale" ? { stopIndex: source.stops.length - 1 } : {}), properties: { "background-color": "#33bb88" } };
    const manifest = session.snapshot().manifest;
    session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [operation] }));
    const after = session.snapshot(), next = after.metadata.workbook.sheets[0]!.conditionalRules![index]!, cssClass = next.kind === "colorScale" ? next.stops.at(-1)!.cssClass : next.cssClass!, resolver = new SourceStyleResolver(after.styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS, new Set());
    expect(cssClass).not.toBe(original); expect(resolver.classStyle(cssClass)["background-color"]).toBe("#33bb88"); expect(resolver.classStyle(original)).toEqual(expected);
    expect(after.manifest.chunks).toEqual(before.manifest.chunks); expect(after.metadata.workbook.sheets[0]!.dataValidations).toEqual(before.metadata.workbook.sheets[0]!.dataValidations);
    expect(after.metadata.workbook.sheets[0]!.conditionalRules!.filter((_, at) => at !== index)).toEqual(before.metadata.workbook.sheets[0]!.conditionalRules!.filter((_, at) => at !== index));
    expect(openDocument(readDocumentZip(writeDocumentZip(session.serialize())), { kind: "workbook" }).snapshot()).toEqual(after);
    const saved = session.serialize(); expect(session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: after.manifest.revision, operations: [operation] })).changed).toBe(false); expect(session.serialize()).toEqual(saved);
    session.undo(); expect(session.snapshot().styles).toEqual(before.styles); expect(session.snapshot().metadata).toEqual(before.metadata); session.redo(); expect(session.snapshot().styles).toEqual(after.styles);
  } finally { session.dispose(); }
});
