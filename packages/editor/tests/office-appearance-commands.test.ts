import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentSelection, DocumentOperation } from "@doffi/core";
import { exportXlsx, importXlsx } from "@doffi/office";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { createEditorCommands } from "../src/commands.js";
import { registerWorkbookAppearanceCommands } from "../src/workbook/appearance-commands.js";

for (const id of ["numberGeneral", "numberNumber", "numberCurrency", "numberAccounting", "numberShortDate", "numberLongDate", "numberTime", "numberPercent", "numberFraction", "numberScientific", "numberText", "numberComma"]) {
  test(`Office ${id} preserves content, exports the chosen format and restores its source in one undo`, async () => {
    const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 12.3 } }]));
    const before = session.serialize(), manifest = session.snapshot().manifest;
    const selected: DocumentSelection = { documentId: manifest.documentId, baseRevision: manifest.revision, anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 1, column: 0 } }] } };
    const commands = createEditorCommands(() => true);
    registerWorkbookAppearanceCommands(commands, { snapshot: () => session.snapshot(), selection: () => selected, style: () => ({}), enabled: () => true,
      async run(operations: DocumentOperation[]) { session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: session.snapshot().manifest.revision, operations })); } });
    await commands.executeCommand(id);
    expect(session.history().undoSteps).toBe(1); expect(session.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 12.3 });
    const formats = session.snapshot().metadata.workbook.numberFormats;
    const code = formats.find(format => format.formatId === session.getCell("sheet-1", 0, 0)?.formatId);
    expect(code).toHaveProperty("code");
    const reopened = await importXlsx((await exportXlsx(openDocument(session.serialize(), { kind: "workbook" }))).bytes, { documentId: "reopened-ribbon-format" });
    const nativeBuiltins: Record<string, number> = { numberGeneral: 0, numberNumber: 2, numberTime: 21, numberPercent: 10, numberScientific: 11, numberFraction: 12, numberText: 49, numberComma: 4 };
    expect(openDocument(reopened, { kind: "workbook" }).snapshot().metadata.workbook.numberFormats.some(format => "code" in format && "code" in code! && format.code === code.code || "builtinId" in format && format.builtinId === nativeBuiltins[id])).toBe(true);
    session.undo(); expect(session.serialize()).toEqual(new Map([...before].map(([path, value]) => path === "manifest.json" ? [path, session.serialize().get(path)!] : [path, value]))); session.dispose();
  });
}
