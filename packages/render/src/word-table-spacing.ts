import { SourceStyleResolver, DEFAULT_DOCUMENT_READ_LIMITS } from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { ownViewStyle } from "./view-nodes.js";
/** Native spacing is authoritative over an earlier CSS projection. Nested
 * tables resolve their own spacing instead of inheriting a browser variable. */
export function composeWordTableSpacing(document: Document, model: ContentDocumentModel): void {
  const snapshot = model.snapshot(),
    tables = new Map(snapshot.metadata.word?.tables.map((table) => [table.nodeId, table] as const));
  if (!tables.size) return;
  const css = new SourceStyleResolver(
      snapshot.styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    project = (id: string): SourceStyleNode => {
      const node = model.getNode(id)!;
      return {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
    };
  for (const table of document.querySelectorAll<HTMLTableElement>("table")) {
    const id = table.dataset.nodeId;
    if (!id || !tables.has(id)) continue;
    const computed = css.get(project(id)),
      spacing = computed["--word-table-cell-spacing"];
    ownViewStyle(table, "border-spacing");
    table.style.borderSpacing = spacing ?? computed["border-spacing"] ?? "0pt";
    if (computed["table-layout"] !== "fixed") {
      const info = tables.get(id)!;
      for (const [index, column] of [
        ...table.querySelectorAll<HTMLTableColElement>(":scope > colgroup > col"),
      ].entries())
        column.style.width = (info.grid[index]! / info.grid.reduce((a, b) => a + b, 0)) * 100 + "%";
      if (!computed.width || computed.width === "auto") {
        ownViewStyle(table, "width", "max-width");
        table.style.width = info.grid.reduce((a, b) => a + b, 0) + "pt";
        table.style.maxWidth = "100%";
      }
    }
    if (spacing !== undefined) {
      ownViewStyle(table, "border-collapse");
      table.style.borderCollapse =
        parseFloat(document.defaultView!.getComputedStyle(table).borderSpacing) > 0
          ? "separate"
          : "collapse";
    }
  }
}
