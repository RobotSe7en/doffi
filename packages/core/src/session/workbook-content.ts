import type { WorkbookManifest } from "../contracts/document.js";
import type { DocumentReadLimits } from "../runtime.js";
import { DocumentEdit } from "../document/edit.js";
import { parseHtml } from "../document/html.js";
import type { HtmlSourceDescriptor } from "../document/html.js";
import { decodeText, encode, resolveReference } from "../document/package.js";
import { escapeWorkbookText } from "./workbook-cells.js";

/** Lazily opens only the bounded rich-content part, never the cell grid. */
export class WorkbookContentEdit {
  private editor?: DocumentEdit;
  constructor(private readonly manifest: WorkbookManifest, private readonly files: Map<string, Uint8Array>,
    private readonly limits: DocumentReadLimits, private readonly allocate: () => string) {}
  private descriptor(): HtmlSourceDescriptor {
    return { kind: "workbook-content", entrypoint: this.manifest.content!, styles: this.manifest.styles, assets: this.manifest.assets };
  }
  edit(): DocumentEdit {
    if (this.editor) return this.editor;
    if (!this.manifest.content) {
      // A root-level path keeps stylesheet/resource resolution independent of
      // the input package's nested paths. It is still an opaque package path.
      let path: string; do { path = "workbook-" + this.allocate() + ".html"; } while (this.files.has(path));
      this.manifest.content = path;
      this.files.set(path, encode('<!doctype html><html><head>' + this.manifest.styles.map(path => '<link rel="stylesheet" href="' + escapeWorkbookText(path.split("/").map(encodeURIComponent).join("/")) + '">').join("")
        + '</head><body><article data-node-id="' + this.allocate() + '" data-office-kind="workbook-content"></article></body></html>'));
    }
    const descriptor = this.descriptor(), source = decodeText(this.files.get(descriptor.entrypoint)!);
    const html = parseHtml(source, descriptor, this.files, this.limits, { nodes: 0 });
    return this.editor = new DocumentEdit({ files: this.files, snapshot: { html: html.html, manifest: descriptor, metadata: {} }, html }, this.limits, this.allocate);
  }
  parse() { return parseHtml(this.edit().html(), this.descriptor(), this.files, this.limits, { nodes: 0 }); }
  root(): string { const html = this.parse(); return html.nodeIds().find(id => html.peekNode(id)?.parentNodeId === null)!; }
  relative(path: string): string {
    this.edit();
    const directory = this.manifest.content!.split("/").slice(0, -1), target = path.split("/");
    while (directory.length && directory[0] === target[0]) { directory.shift(); target.shift(); }
    const relative = [...directory.map(() => ".."), ...target.map(encodeURIComponent)].join("/");
    if (resolveReference(relative, this.manifest.content!) !== path) throw new Error("Invalid workbook resource path.");
    return relative;
  }
  flush(): void { if (this.editor) this.files.set(this.manifest.content!, encode(this.editor.html())); }
}
