import { openContentFixture as openDocument } from "../../content-fixture.js";
import { readFileSync } from "node:fs";
import { createDocumentSession } from "@doffi/core";
import type { DocumentPackage, PageSetup } from "@doffi/core";
import { importDocx } from "@doffi/office";

export const smallSetup: PageSetup = { unit: "pt", width: 300, height: 240, margins: { top: 24, right: 24, bottom: 24, left: 24 } };
export function pagedFiles(body: string): Map<string, Uint8Array> {
  const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path,
    new Uint8Array(readFileSync(new URL("../../../examples/report/" + path, import.meta.url))),
  ]));
  const encoder = new TextEncoder(), decoder = new TextDecoder();
  files.set("content.html", encoder.encode(decoder.decode(files.get("content.html")).replace(/<article[\s\S]*<\/article>/,
    '<article data-node-id="root" data-office-kind="document">' + body + "</article>")));
  const metadata = JSON.parse(decoder.decode(files.get("metadata.json"))); metadata.pageSetup = smallSetup;
  files.set("metadata.json", encoder.encode(JSON.stringify(metadata))); files.set("annotations.json", encoder.encode("[]"));
  files.set("styles.css", encoder.encode('article{font-family:"Liberation Serif",serif;font-size:12pt}p{line-height:16pt;margin:0}td,th{line-height:16pt;padding:0}table{border-collapse:collapse;width:240pt}tr{break-inside:avoid}'));
  return files;
}
export function longTable(): Map<string, Uint8Array> {
  return pagedFiles('<table data-node-id="table"><tr data-node-id="header" data-word-repeat-header="true"><th data-node-id="header-cell">REPEATED HEADER</th></tr>'
    + Array.from({ length: 32 }, (_, index) => '<tr data-node-id="row-' + index + '"><td data-node-id="cell-' + index + '">UNIQUE-' + index + "</td></tr>").join("") + "</table>");
}
export function fieldsFiles(): DocumentPackage {
  const files = new Map(importDocx(readFileSync(new URL("../../../examples/docx/07-fields.docx", import.meta.url)), { documentId: "fields-layout" }));
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json"))); metadata.pageSetup = { ...smallSetup, height: 400 };
  for (const section of metadata.sections) section.pageSetup = metadata.pageSetup;
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const session = createDocumentSession(openDocument(files)), target = session.snapshot().metadata.sections![0]!.rootNodeId;
  session.commit(session.prepare({ protocolVersion: 1, documentId: "fields-layout", baseRevision: 0, operations: [
    { op: "insertFragment", nodeId: target, position: "start", html: "<h1>Alpha</h1>" },
    { op: "insertFragment", nodeId: target, position: "end", html: '<h2 style="break-before:page">Beta</h2>' },
  ] }));
  const result = session.serialize(); session.dispose(); return result;
}
