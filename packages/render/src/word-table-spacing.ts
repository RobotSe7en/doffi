import { SourceStyleResolver, DEFAULT_DOCUMENT_READ_LIMITS } from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { ownViewStyle } from "./view-nodes.js";

// Word expects the native preferred table width to fit a signed short in twips.
const maxNativeWidthPt = 32767 / 20;

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
    const id = table.dataset.nodeId,
      info = id ? tables.get(id) : undefined;
    if (!id || !info) continue;
    const computed = css.get(project(id)),
      spacing = computed["--word-table-cell-spacing"],
      fixed = computed["table-layout"] === "fixed",
      width = computed.width,
      gridWidth = info.grid.reduce((a, b) => a + b, 0);
    ownViewStyle(table, "border-spacing");
    table.style.borderSpacing = spacing ?? computed["border-spacing"] ?? "0pt";
    if (!fixed) {
      for (const [index, column] of [
        ...table.querySelectorAll<HTMLTableColElement>(":scope > colgroup > col"),
      ].entries())
        column.style.width = (info.grid[index]! / gridWidth) * 100 + "%";
    }
    if (
      (width?.endsWith("pt") && parseFloat(width) > maxNativeWidthPt) ||
      (!fixed && (!width || width === "auto"))
    ) {
      ownViewStyle(table, "width");
      table.style.width = gridWidth + "pt";
      if (!fixed) {
        ownViewStyle(table, "max-width");
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
