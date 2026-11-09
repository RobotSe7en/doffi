import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";

const evidence = JSON.parse(readFileSync("docs/evidence/excel-web-name-semantics/observations.json", "utf8")) as {
  captures: { input: string; sourceSha256: string; cells: { sheet: string; cell: string; cache: number; observed: number | string }[] }[];
};
function coordinate(address: string) {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(address)!;
  let column = 0; for (const letter of letters!) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(digits) - 1, column: column - 1 };
}
for (const capture of evidence.captures) test(`current calculation matches independently observed Excel web values: ${capture.input}`, async () => {
  const bytes = readFileSync("docs/evidence/excel-web-name-semantics/" + capture.input);
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(capture.sourceSha256);
  const session = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "excel-web-names" }), { kind: "workbook" }));
  try {
    const before = session.serialize(); session.calculateSync();
    for (const cell of capture.cells) {
      const sheetId = session.snapshot().metadata.workbook.sheets.find(sheet => sheet.name === cell.sheet)!.sheetId;
      const { row, column } = coordinate(cell.cell);
      expect(session.getCell(sheetId, row, column)?.value, `${cell.sheet}!${cell.cell} imported cache`).toEqual({ type: "number", value: cell.cache });
      const type = typeof cell.observed === "number" ? "number" : cell.observed.startsWith("#") ? "error" : "text";
      expect(session.getCalculatedCell(sheetId, row, column), `${cell.sheet}!${cell.cell}`).toMatchObject({ state: "valid", value: { type, value: cell.observed } });
    }
    expect(session.serialize()).toEqual(before);
    const reopened = openDocument(readDocumentZip(await writeDocumentZip(before)), { kind: "workbook" });
    expect(reopened.snapshot()).toEqual(session.snapshot());
  } finally { session.dispose(); }
});

function apply(session: WorkbookSession, operations: DocumentOperation[]) {
  const manifest = session.snapshot().manifest;
  session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations }));
}
test("native names retain their bound identities through rename, current-value invalidation, conflicts and undo", async () => {
  const session = createDocumentSession(openDocument(await importXlsx(readFileSync("docs/evidence/excel-web-inputs/name-semantics.xlsx"), { documentId: "excel-name-edits" }), { kind: "workbook" }));
  try {
    const before = session.snapshot(), [data, other, third] = before.metadata.workbook.sheets;
    session.calculateSync(); expect(session.calculateSync()).toMatchObject({ calculated: 0, reused: 45 });
    const definition = before.metadata.workbook.names!.find(name => name.name === "ProbeShadow" && name.sheetId === undefined)!;
    apply(session, [{ op: "setWorkbookDefinedName", name: { ...definition, name: "GlobalShadow" } }]); session.calculateSync();
    const names = session.snapshot().metadata.workbook.names!;
    for (const name of names.filter(name => ["ProbePlainAlias", "ProbeLocalPlain"].includes(name.name))) expect(name.expression).toBe("=GlobalShadow");
    expect(names.find(name => name.name === "ProbeBangAlias")!.expression).toBe("=!ProbeShadow");
    expect(session.getCell(third!.sheetId, 9, 2)?.formula?.expression).toBe("=GlobalShadow");
    expect(session.getCell(data!.sheetId, 9, 2)?.formula?.expression).toBe("=ProbeShadow");
    for (const sheet of before.metadata.workbook.sheets) expect(session.getCalculatedCell(sheet.sheetId, 10, 2)).toMatchObject({ state: "valid", value: { type: "number", value: 30 } });
    const stale = session.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: 1,
      operations: [{ op: "setWorkbookCells", sheetId: data!.sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 9999 } } }] }] });
    apply(session, [{ op: "setWorkbookCells", sheetId: data!.sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 21 } } }] }]);
    expect(() => session.commit(stale)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
    const calculation = session.calculateSync(); expect(calculation.calculated).toBeGreaterThan(0); expect(calculation.reused).toBeGreaterThan(0);
    for (const sheet of [data!, other!, third!]) {
      expect(session.getCalculatedCell(sheet.sheetId, 4, 2)).toMatchObject({ state: "valid", value: { type: "number", value: 21 } });
      expect(session.getCalculatedCell(sheet.sheetId, 13, 2)).toMatchObject({ state: "valid", value: { type: "number", value: 1021 } });
    }
    expect(session.getCalculatedCell(other!.sheetId, 5, 2)).toMatchObject({ value: { type: "number", value: 101 } });
    expect(session.history().undoSteps).toBe(2); session.undo(); session.undo();
    expect(session.snapshot().metadata).toEqual(before.metadata); expect(session.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { session.dispose(); }
});
