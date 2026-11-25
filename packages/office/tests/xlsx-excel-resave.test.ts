import { readFileSync } from "node:fs";
import { createDocumentSession, openDocument } from "@doffi/core";
import { exportXlsx, importXlsx } from "@doffi/office";
import { expect, test } from "vitest";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";

const original = readFileSync("examples/xlsx/04-formulas-names.xlsx");
const spreadsheet = "http://schemas.microsoft.com/office/spreadsheetml/";
const mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const revision = spreadsheet + "2014/revision";
const uuid = "{00000000-0001-0000-0000-000000000000}";
const workbookPart = "xl/workbook.xml", sheetPart = "xl/worksheets/sheet1.xml", stylePart = "xl/styles.xml";
const pathHint = `<mc:AlternateContent><mc:Choice Requires="x15"><ac:absPath xmlns:ac="${spreadsheet}2010/11/ac" url="file:///local/save/path/"/></mc:Choice></mc:AlternateContent>`;

function resaveMetadata() {
  const files = new Map(readDocumentZip(original)), encoder = new TextEncoder(), decoder = new TextDecoder();
  const update = (part: string, edit: (xml: string) => string) => files.set(part, encoder.encode(edit(decoder.decode(files.get(part)!))));
  update(workbookPart, xml => xml
    .replace("<workbook ", `<workbook xmlns:mc="${mc}" mc:Ignorable="x15 xr" xmlns:x15="${spreadsheet}2010/11/main" xmlns:xr="${revision}" `)
    .replace("<bookViews>", `${pathHint}<xr:revisionPtr revIDLastSave="0" documentId="local-test"/><bookViews>`)
    .replace("</workbook>", `<extLst><ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:calc="${spreadsheet}2018/calcfeatures"><calc:calcFeatures><calc:feature name="microsoft.com:LET_WF"/></calc:calcFeatures></ext></extLst></workbook>`));
  update(sheetPart, xml => xml
    .replace("<worksheet ", `<worksheet xmlns:mc="${mc}" mc:Ignorable="xr" xmlns:xr="${revision}" xr:uid="${uuid}" `)
    .replace("<pageMargins", '<phoneticPr fontId="1" type="noConversion"/><pageMargins'));
  update(stylePart, xml => xml
    .replace("<styleSheet ", `<styleSheet xmlns:mc="${mc}" mc:Ignorable="xr" xmlns:xr="${revision}" `)
    .replace("</styleSheet>", `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="${spreadsheet}2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="${spreadsheet}2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>`));
  return { files, update };
}

const open = async (bytes: Uint8Array) => openDocument(await importXlsx(bytes, { documentId: "excel-resave" }), { kind: "workbook" });

test("desktop Excel save metadata leaves the imported workbook and formula source unchanged", async () => {
  const { files } = resaveMetadata(), expected = (await open(original)).serialize();
  const model = await open(writeOpcZip(files));
  expect(model.serialize()).toEqual(expected);
  expect(model.getCell("sheet-1", 1, 1)).toMatchObject({ formula: { expression: "=A2*2" }, value: { type: "number", value: 4 } });
  const result = await exportXlsx(model);
  expect((await open(result.bytes)).getCell("sheet-1", 1, 1)?.formula?.expression).toBe("=A2*2");
});

test.each([
  [workbookPart, 'mc:Ignorable="x15 xr"', 'mc:Ignorable="missing"', "InvalidDocument"],
  [workbookPart, '</mc:Choice>', '</mc:Choice><mc:Fallback><sheetProtection/></mc:Fallback>', "UnsupportedFeature"],
  [workbookPart, 'revIDLastSave="0"', 'revIDLastSave="1"', "UnsupportedFeature"],
  [sheetPart, 'xr:uid=', 'xr:unmodeled=', "UnsupportedFeature"],
  [stylePart, '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>', '<x14:slicerStyles><x14:slicerStyle/></x14:slicerStyles>', "UnsupportedFeature"],
])("rejects unsupported Excel %s payload containing %s", async (part, find, replacement, code) => {
  const { files, update } = resaveMetadata();
  update(part, xml => xml.replace(find, replacement));
  await expect(open(writeOpcZip(files))).rejects.toMatchObject({ code });
});

test.each([
  ["row", { row: 2, column: 3 }, "B2:D3", 5, 4],
  ["column", { row: 3, column: 2 }, "B2:C4", 6, 3],
] as const)("matches desktop Excel when deleting the leading merged %s", async (axis, to, address, formulaRow, formulaColumn) => {
  const { files, update } = resaveMetadata();
  update(sheetPart, xml => xml.replace(/<sheetData>.*?<\/sheetData>/s, '<sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>MERGE-TOP-LEFT</t></is></c></row><row r="7"><c r="E7"><f>IF(ISREF(B2),1,0)</f><v>1</v></c></row><row r="8"><c r="E8"><f>COUNTA(B2:D4)</f><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="B2:D4"/></mergeCells>'));
  const model = await open(writeOpcZip(files)), session = createDocumentSession(model);
  try {
    const { documentId, revision } = model.snapshot().manifest, sheetId = "sheet-1";
    session.commit(session.prepare({ protocolVersion: 1, documentId, baseRevision: revision, operations: [
      { op: "editWorkbookAxis", sheetId, axis, action: "delete", index: 1, count: 1 },
    ] }));
    const reopened = await open((await exportXlsx(session)).bytes);
    expect(reopened.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([{ from: { row: 1, column: 1 }, to }]);
    expect(reopened.getCell(sheetId, 1, 1)?.value?.type ?? "blank").toBe("blank");
    expect(reopened.getCell(sheetId, formulaRow, formulaColumn)?.formula?.expression).toBe("=IF(ISREF(#REF!),1,0)");
    expect(reopened.getCell(sheetId, formulaRow + 1, formulaColumn)?.formula?.expression).toBe(`=COUNTA(${address})`);
    session.undo();
    expect(session.getCell(sheetId, 1, 1)?.value).toEqual({ type: "text", value: "MERGE-TOP-LEFT" });
  } finally {
    session.dispose();
  }
});
