import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

test.each(["examples/xlsx/08-images-shapes.xlsx", "examples/xlsx/producer/08-images-shapes.xlsx"])("native %s keeps original runs, font sources and resources through precise object input and package reopening", async path => {
  const session = createDocumentSession(openDocument(await importXlsx(readFileSync(path), { documentId: "object-text-native" }), { kind: "workbook" }));
  try {
    const before = session.snapshot(), files = session.serialize(), object = before.metadata.workbook.objects!.find(object => object.kind === "geometry")!, original = session.getNode(object.nodeId)!, paragraphId = original.childNodeIds[0]!, paragraph = session.getNode(paragraphId)!;
    const runs = new Map(paragraph.childNodeIds.map(id => [id, session.getNode(id)!.html])), prefix = "新增禰󠄀😀";
    const operations: DocumentOperation[] = [
      { op: "replaceWorkbookObjectText", nodeId: object.nodeId, paragraphId, from: 0, to: 0, expectedText: "", text: prefix },
      { op: "formatWorkbookObjectText", nodeId: object.nodeId, paragraphId, from: 0, to: prefix.length, expectedText: prefix, properties: { color: "#a02030", "--office-font-east-asia": '"Noto Sans CJK SC"', "font-size": "18pt" } },
      { op: "setWorkbookObjectParagraphStyle", nodeId: object.nodeId, paragraphId, properties: { "text-align": "center" } },
      { op: "insertWorkbookObjectParagraph", nodeId: object.nodeId, html: "<p>原生来源后的新段落</p>" },
      { op: "updateWorkbookObject", nodeId: object.nodeId, changes: { text: { wrap: false, autoFit: "normal" } } },
    ];
    session.commit(session.prepare({ protocolVersion: 1, documentId: "object-text-native", baseRevision: 0, operations }));
    expect(session.getNode(paragraphId)!.text).toBe(prefix + paragraph.text); for (const [id] of runs) expect(session.getNode(id)).toBeDefined();
    expect(session.getNode(object.nodeId)!.childNodeIds[0]).toBe(paragraphId); expect(session.getNode(object.nodeId)!.html).toContain("--office-text-source:direct");
    for (const other of before.metadata.workbook.objects!) if (other.nodeId !== object.nodeId) expect(session.getNode(other.nodeId)!.html).toBe(openDocument(files, { kind: "workbook" }).getNode(other.nodeId)!.html);
    expect(session.snapshot().manifest.assets).toEqual(before.manifest.assets); for (const asset of before.manifest.assets) expect(session.serialize().get(asset.path)).toEqual(files.get(asset.path));
    expect(session.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(session.snapshot().styles).toEqual(before.styles);
    const reopened = openDocument(await readDocumentZip(writeDocumentZip(session.serialize())), { kind: "workbook" }); expect(reopened.snapshot()).toEqual(session.snapshot()); expect(reopened.getNode(paragraphId)!.text).toBe(prefix + paragraph.text);
    session.undo(); expect(session.snapshot().metadata).toEqual(before.metadata); expect(session.snapshot().content).toEqual(before.content); session.redo(); expect(session.getNode(paragraphId)!.text).toBe(prefix + paragraph.text);
  } finally { session.dispose(); }
});
