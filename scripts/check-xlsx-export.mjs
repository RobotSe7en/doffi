import assert from "node:assert/strict";
import { mkdir, readFile, readdir, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { createDocumentSession, openDocument } from "../packages/core/dist/index.js";
import { exportXlsx, importXlsx } from "../packages/office/dist/index.js";

const [phase, path] = process.argv.slice(2);
assert(["export", "reopen"].includes(phase) && path, "Select export or reopen and a directory.");
const directory = resolve(path);
const open = async (file, documentId) =>
  openDocument(await importXlsx(await readFile(file), { documentId }), { kind: "workbook" });
const cell = (row, column, value) => ({ row, column, source: { kind: "value", value: { type: "number", value } } });

if (phase === "export") {
  await mkdir(directory, { recursive: true });
  const inputs = ["examples/xlsx", "examples/xlsx/producer"];
  for (const [producer, input] of inputs.entries()) {
    for (const name of (await readdir(input)).filter(name => name.endsWith(".xlsx")).sort()) {
      if (process.env.DOFFI_CI_MODE === "daily" && name.startsWith("15-")) continue;
      const documentId = (producer ? "producer-" : "") + name;
      const document = await open(join(input, name), documentId);
      const edit = !producer && ["05-native-table.xlsx", "07-three-charts.xlsx", "09-classic-notes.xlsx"].includes(name);
      const session = edit ? createDocumentSession(document) : undefined;
      try {
        if (session) {
          const snapshot = session.snapshot(), sheetId = snapshot.metadata.workbook.sheets[0].sheetId;
          const operations = name === "09-classic-notes.xlsx"
            ? [{ op: "setWorkbookCommentBody", annotationId: snapshot.annotations[0].annotationId, html: "<p>SDK source change 禰󠄀</p>" }]
            : [{ op: "setWorkbookCells", sheetId, cells: name === "05-native-table.xlsx"
              ? [cell(1, 1, 6), { row: 0, column: 7, source: { kind: "formula", expression: "=B2*2" } }]
              : [cell(1, 1, 39)] }];
          session.commit(session.prepare({ protocolVersion: 1, documentId, baseRevision: 0, operations }));
        }
        const result = await exportXlsx(session ?? document);
        const output = openDocument(await importXlsx(result.bytes, { documentId: "reopen-" + documentId }), { kind: "workbook" });
        assert.deepEqual(output.snapshot().metadata.workbook.sheets.map(sheet => sheet.name), document.snapshot().metadata.workbook.sheets.map(sheet => sheet.name));
        assert.equal(result.report.revision, edit ? 1 : 0);
        assert.equal(result.report.unexportedAnnotations.length, 0);
        if (edit && name === "09-classic-notes.xlsx")
          assert.equal(output.getNode(output.snapshot().annotations[0].bodyNodeId).text, "SDK source change 禰󠄀");
        if (edit && name === "07-three-charts.xlsx") {
          const chart = output.snapshot().metadata.workbook.charts[0];
          assert(chart.series.some(item => item.values.cache?.points.some(point => point.value === 39)));
        }
        if (name.startsWith("15-")) {
          assert.deepEqual(result.report.formulaCaches, { written: 20_000, omitted: 0 });
          assert.equal(output.getCell("sheet-1", 9999, 19).value.value, 3_400_152);
        }
        await writeFile(join(directory, documentId), result.bytes);
      } finally {
        session?.dispose();
      }
    }
  }
} else {
  const table = await open(join(directory, "05-native-table.xlsx"), "native-table");
  const sheet = table.snapshot().metadata.workbook.sheets[0];
  assert.equal(table.getCell(sheet.sheetId, 1, 1).value.value, 7);
  assert.equal(table.getCell(sheet.sheetId, 0, 7).formula.expression, "=B2*4");
  assert.equal(table.getCell(sheet.sheetId, 0, 7).value.value, 28);
  assert.equal(sheet.tables[0].name, "SalesTable");
  assert.equal(sheet.tables[0].range.to.row, 6);
  const charts = await open(join(directory, "07-three-charts.xlsx"), "native-charts");
  const chart = charts.snapshot().metadata.workbook.charts[0];
  assert.equal(charts.getNode(chart.titleNodeId).text, "Calc title edit");
  assert.deepEqual(charts.snapshot().metadata.workbook.charts.map(chart => chart.kind), ["column", "line", "pie"]);
  const notes = await open(join(directory, "09-classic-notes.xlsx"), "native-notes");
  assert.equal(notes.getNode(notes.snapshot().annotations[0].bodyNodeId).text, "Calc note edit 禰󠄀");
  console.log("SDK reopened the actual Calc value, formula, table range, chart title and note edits.");
}
