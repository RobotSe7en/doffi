import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { openDocument } from "@doffi/core";
import { createDocumentConverter } from "@doffi/node";
import { importXlsx, readDocumentZip } from "@doffi/office";

const [directory, browserExecutable] = process.argv.slice(2);
const archive = await readFile(join(directory, "edited.zip"));
const xlsx = await readFile(join(directory, "edited.xlsx"));
const packaged = openDocument(readDocumentZip(archive), { kind: "workbook" });
const native = openDocument(await importXlsx(xlsx, { documentId: "packed-native" }), { kind: "workbook" });
assert.equal(packaged.snapshot().manifest.documentId, "packed-workbook");
assert.equal(packaged.snapshot().manifest.revision, 4);
assert.equal(native.snapshot().manifest.revision, 0);
for (const model of [packaged, native]) {
  const { metadata, annotations } = model.snapshot();
  assert.equal(metadata.workbook.sheets[0].name, "备注");
  const { sheetId } = metadata.workbook.sheets[0];
  assert.equal(model.getCell(sheetId, 4, 0).value.value, 37);
  assert.equal(model.getCell(sheetId, 4, 1).formula.expression, "=A5*3");
  assert.equal(annotations.length, 1);
  assert.equal(annotations[0].anchor.sheetId, sheetId);
  assert.deepEqual(annotations[0].anchor.ranges, [{ from: { row: 1, column: 1 }, to: { row: 1, column: 1 } }]);
  assert.equal(model.getNode(annotations[0].bodyNodeId).text, "请确认：预算😀禰󠄀");
}
assert.equal(native.getCell(native.snapshot().metadata.workbook.sheets[0].sheetId, 4, 1).value.value, 111);

const area = { from: { row: 4, column: 0 }, to: { row: 4, column: 1 } };
const workbook = { sheets: [{ sheetId: packaged.snapshot().metadata.workbook.sheets[0].sheetId, areas: [area] }] };
const converter = createDocumentConverter({ browserExecutable });
try {
  const pdf = await converter.renderWorkbookPdf(packaged, workbook);
  const png = await converter.renderWorkbookPage(packaged, { ...workbook, pageIndex: 0 });
  assert.equal(pdf.report.kind, "workbook");
  assert.equal(pdf.report.revision, 4);
  assert.equal(pdf.report.formulaValues, "current");
  assert.equal(pdf.report.formulaCells.valid, 1);
  assert.equal(pdf.report.pages.length, 1);
  assert.deepEqual(png.report, pdf.report);
  await writeFile(join(directory, "edited.pdf"), pdf.bytes);
  await writeFile(join(directory, "edited.png"), png.bytes);
  await writeFile(join(directory, "print-report.json"), JSON.stringify(pdf.report, null, 2) + "\n");
  await assert.rejects(converter.renderWorkbookPage(packaged, { ...workbook, pageIndex: 1 }), { code: "InvalidRequest" });
  await assert.rejects(converter.renderWorkbookPdf(packaged, { signal: AbortSignal.abort() }), { code: "Cancelled" });
} finally {
  await converter.close();
}
await assert.rejects(converter.renderWorkbookPdf(packaged), { code: "InvalidOperation" });
