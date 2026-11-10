import { defaultTreeAdapter as adapter, html as namespaces } from "parse5";
import type { DefaultTreeAdapterTypes as Tree } from "parse5";
import type { WorkbookCell } from "../runtime.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import { DocumentError } from "../errors.js";
import { serializeHtmlTree } from "../document/html-parser.js";
import { isTextBoundary } from "../document/text.js";
import { CHUNK_COLUMNS, CHUNK_ROWS, parseWorkbookCellElement } from "../workbook/chunk-index.js";

type Format = Extract<WorkbookOperation, { op: "formatWorkbookText" }>;
export function formatWorkbookCellText(cell: WorkbookCell | undefined, operation: Format, format?: (root: Tree.Element) => void, cssClass?: string): string {
  if (!cell || cell.formula || cell.value.type !== "text") throw new DocumentError("InvalidOperation", "Partial font formatting requires a text constant.");
  const { from, to, expectedText } = operation;
  if (from >= to || !isTextBoundary(cell.text, from) || !isTextBoundary(cell.text, to) || cell.text.slice(from, to) !== expectedText)
    throw new DocumentError("InvalidOperation", "Text formatting requires an exact nonempty UTF-16 range preserving surrogate pairs and variation sequences.");
  if (!format) return cell.html;
  const root = parseWorkbookCellElement(cell.html), before = serializeHtmlTree(root), attributes = [...root.attrs];
  if (cssClass !== undefined) root.attrs = [...root.attrs.filter(attribute => attribute.name !== "class"), { name: "class", value: cssClass }];
  const html = adapter.createElement("html", namespaces.NS.HTML, []), body = adapter.createElement("body", namespaces.NS.HTML, []);
  const table = adapter.createElement("table", namespaces.NS.HTML, [{ name: "data-sheet-id", value: cell.sheetId }, { name: "data-row", value: String(Math.floor(cell.row / CHUNK_ROWS) * CHUNK_ROWS) }, { name: "data-column", value: String(Math.floor(cell.column / CHUNK_COLUMNS) * CHUNK_COLUMNS) }]);
  const tbody = adapter.createElement("tbody", namespaces.NS.HTML, []), row = adapter.createElement("tr", namespaces.NS.HTML, [{ name: "data-row", value: String(cell.row) }]);
  adapter.appendChild(html, body); adapter.appendChild(body, table); adapter.appendChild(table, tbody); adapter.appendChild(tbody, row); adapter.detachNode(root); adapter.appendChild(row, root);
  format(root); root.attrs = attributes;
  const result = serializeHtmlTree(root);
  if (result.length > 2 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "Rich cell formatting exceeds its source budget.");
  return result === before ? cell.html : result;
}
