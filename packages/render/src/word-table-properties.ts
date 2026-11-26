import {
  DEFAULT_DOCUMENT_READ_LIMITS, SourceStyleResolver, wordEffectiveTableProperties, wordRowExceptionKey,
} from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { ownViewStyle } from "./view-nodes.js";

/** Resolve from the current source, including structural edits to the first
 * row. These properties are projections, never authoritative table edits. */
export function wordTablePropertyResolver(model: ContentDocumentModel): (table: HTMLTableElement) => Record<string, string> {
  const snapshot = model.snapshot(),
    resolver = new SourceStyleResolver(snapshot.styles.map(sheet => sheet.css), DEFAULT_DOCUMENT_READ_LIMITS),
    project = (id: string): SourceStyleNode => {
      const node = model.getNode(id)!;
      return { tag: node.tagName, attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined };
    };
  return table => {
    const node = project(table.dataset.nodeId!), css = { ...resolver.get(node) };
    if (resolver.specified(node)["--word-table-indent-origin"] === undefined)
      delete css["--word-table-indent-origin"];
    return wordEffectiveTableProperties(css,
      table.rows[0]?.dataset.nodeId ? resolver.specified(project(table.rows[0].dataset.nodeId)) : {},
      snapshot.metadata.word?.compatibilityMode);
  };
}

function firstRowHasWidth(table: HTMLTableElement): boolean {
  const row = table.rows[0];
  return !!row && !!table.ownerDocument.defaultView!.getComputedStyle(row)
    .getPropertyValue(wordRowExceptionKey("width")).trim();
}

export function composeWordTableProperties(document: Document, model: ContentDocumentModel): void {
  const snapshot = model.snapshot(), resolve = wordTablePropertyResolver(model),
    ids = new Set(snapshot.metadata.word?.tables.map(table => table.nodeId));
  for (const table of document.querySelectorAll<HTMLTableElement>("table[data-node-id]")) {
    if (!ids.has(table.dataset.nodeId!)) continue;
    const properties = resolve(table);
    for (const key of ["width", "table-layout", "margin-left", "margin-right", "--word-table-indent-origin"]) {
      if (properties[key] === undefined) continue;
      ownViewStyle(table, key);
      table.style.setProperty(key, properties[key]!);
    }
    if (!firstRowHasWidth(table)) continue;
    // Native preferred widths rescale the saved grid, including fixed tables.
    const info = snapshot.metadata.word!.tables.find(info => info.nodeId === table.dataset.nodeId)!;
    const total = info.grid.reduce((a, b) => a + b, 0);
    for (const [index, col] of [...table.querySelectorAll<HTMLTableColElement>(":scope > colgroup > col")].entries()) {
      ownViewStyle(col, "width");
      col.style.width = (info.grid[index]! / total) * 100 + "%";
    }
  }
}
