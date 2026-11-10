import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, DEFAULT_DOCUMENT_READ_LIMITS, openDocument, SourceStyleResolver } from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { exportDocx, exportPptx, importDocx, importPptx, readDocumentZip } from "@doffi/office";

for (const kind of ["word", "presentation"] as const) test(`${kind} native runs clear underline, retain strike and reopen with the same effective paint`, () => {
  const input = kind === "word" ? importDocx(readFileSync("examples/docx/01-text-numbering.docx"), { documentId: "native-reset" }) : importPptx(readFileSync("examples/pptx/01-rich-text.pptx"), { documentId: "native-reset" });
  const base = openDocument(input); if (base.kind === "workbook") throw Error("Expected content model");
  const s = createDocumentSession(base);
  try {
    const paragraph = [...s.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(match => s.getNode(match[1]!)!).find(node => node.tagName === "p")!, m = s.snapshot().manifest;
    s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [
      { op: "setStyle", nodeId: paragraph.nodeId, property: "text-decoration-line", value: "underline line-through" },
      { op: "insertFragment", nodeId: paragraph.nodeId, position: "end", html: '<span style="text-decoration-line:none;--office-font-decoration-reset:underline">ResetMarkerGlyph</span>' },
    ] }));
    const before = s.serialize(), html = s.snapshot().html, result = kind === "word" ? exportDocx(s) : exportPptx(s), files = readDocumentZip(result.bytes), xml = new TextDecoder().decode(files.get(kind === "word" ? "word/document.xml" : "ppt/slides/slide1.xml")!);
    const run = [...xml.matchAll(kind === "word" ? /<w:r>[\s\S]*?<\/w:r>/g : /<a:r>[\s\S]*?<\/a:r>/g)].find(match => match[0].includes("ResetMarkerGlyph"))![0];
    if (kind === "word") { expect(run).toContain('<w:u w:val="none"'); expect(run).toMatch(/<w:strike(?: w:val="1")?\/>/); }
    else { expect(run).toContain('u="none"'); expect(run).toContain('strike="sngStrike"'); }
    expect(s.serialize()).toEqual(before);
    const reopened = openDocument(kind === "word" ? importDocx(result.bytes, { documentId: "reopened" }) : importPptx(result.bytes, { documentId: "reopened" })); if (reopened.kind === "workbook") throw Error("Expected content model");
    const target = [...reopened.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(match => reopened.getNode(match[1]!)!).find(node => node.tagName === "span" && node.text === "ResetMarkerGlyph")!;
    function projection(model: ContentDocumentModel, id: string): SourceStyleNode {
      const node = model.getNode(id)!; return { tag: node.tagName, attrs: node.attributes, parent: node.parentNodeId ? projection(model, node.parentNodeId) : undefined };
    }
    const css = new SourceStyleResolver(reopened.snapshot().styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS).get(projection(reopened, target.nodeId));
    expect(css["text-decoration-line"]!.split(/\s+/)).toContain("line-through"); expect(css["text-decoration-line"]!.split(/\s+/)).not.toContain("underline");
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().html).toBe(base.snapshot().html); s.redo(); expect(s.snapshot().html).toBe(html);
  } finally { s.dispose(); }
});
