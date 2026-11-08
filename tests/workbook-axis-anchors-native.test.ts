import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

const evidence = JSON.parse(readFileSync("docs/evidence/workbook-axis-drawings.json", "utf8")) as { sources: { name: string; path: string; sha256: string; cases: { name: string; producerSha256: string }[] }[] };
const digest = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex");
function apply(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations })); }
for (const source of evidence.sources) {
  for (const record of source.cases) test("actual fixed-font " + source.name + " " + record.name + " producer reopens independently with its recorded geometry", async () => {
    expect(digest(readFileSync(source.path))).toBe(source.sha256); const bytes = readFileSync("packages/office/tests/fixtures/workbook-axis-drawings-" + source.name + "-" + record.name + "-calc.xlsx"); expect(digest(bytes)).toBe(record.producerSha256);
    const model = openDocument(await importXlsx(bytes, { documentId: "native-axis-anchors-producer" }), { kind: "workbook" });
    expect(model.snapshot().metadata.workbook.sheets).not.toHaveLength(0); expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(model.snapshot());
  });
  for (const [direction, action, index, count] of [["row", "insert", 1, 2], ["column", "insert", 1, 2], ["row", "delete", 0, 1], ["column", "delete", 0, 1]] as const) test(source.name + " native " + direction + " " + action + " keeps source identities, references and both package/undo boundaries", async () => {
    const s = createDocumentSession(openDocument(await importXlsx(readFileSync(source.path), { documentId: "native-axis-anchors-source" }), { kind: "workbook" }));
    try {
      const before = s.snapshot(), sheetId = before.metadata.workbook.sheets[0]!.sheetId, op = { op: "editWorkbookAxis" as const, sheetId, axis: direction, action, index, count };
      apply(s, [op]); const after = s.snapshot(), objects = after.metadata.workbook.objects ?? [];
      expect(objects.map(object => object.nodeId)).toEqual(before.metadata.workbook.objects?.map(object => object.nodeId) ?? []); expect(after.content).toEqual(before.content); expect(after.styles).toEqual(before.styles); expect(after.manifest.assets).toEqual(before.manifest.assets);
      if (source.name === "07-three-charts") {
        const positions = direction === "row" ? action === "insert" ? [8, 8, 22] : [5, 5, 19] : action === "insert" ? [0, 7, 0] : [0, 4, 0];
        expect(objects.map(object => object.anchor.kind === "one-cell" ? object.anchor.from[direction] : -1)).toEqual(positions);
        for (const [i, chart] of (after.metadata.workbook.charts ?? []).entries()) expect(chart.series.map(series => series.seriesId)).toEqual(before.metadata.workbook.charts![i]!.series.map(series => series.seriesId));
      } else if (source.name === "08-images-shapes") {
        expect(objects[0]!.anchor).toMatchObject(direction === "row" ? { kind: "one-cell", from: { row: action === "insert" ? 4 : 1 } } : { kind: "one-cell", from: { column: action === "insert" ? 3 : 0 } });
        expect(objects[1]!.anchor).toMatchObject(direction === "row" ? { kind: "two-cell", from: { row: action === "insert" ? 4 : 1 }, to: { row: action === "insert" ? 11 : 8 } } : { kind: "two-cell", from: { column: action === "insert" ? 6 : 3 }, to: { column: action === "insert" ? 10 : 7 } });
        expect(objects[2]).toEqual(before.metadata.workbook.objects![2]);
      } else {
        const note = after.metadata.workbook.comments![0]!; if (note.format !== "classic") throw Error("Missing independent classic note.");
        expect(note.box!.anchor).toMatchObject(direction === "row" ? { from: { row: action === "insert" ? 3 : 0 }, to: { row: action === "insert" ? 7 : 4 } } : { from: { column: action === "insert" ? 4 : 1 }, to: { column: action === "insert" ? 7 : 4 } });
        expect(after.annotations[0]).toMatchObject({ annotationId: before.annotations[0]!.annotationId, bodyNodeId: before.annotations[0]!.bodyNodeId, status: before.annotations[0]!.status });
      }
      expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(after); expect(s.history().undoSteps).toBe(1);
      s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().content).toEqual(before.content);
    } finally { s.dispose(); }
  });
}
