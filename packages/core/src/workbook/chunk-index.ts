import type { DefaultTreeAdapterTypes } from "parse5";
import { defaultTreeAdapter, html as htmlNamespace } from "parse5";
import { parseHtmlTree, parseHtmlFragment } from "../document/html-parser.js";
import { decodeText } from "../document/package.js";
import { checkLimit, invalid } from "../document/validation.js";
import type { FormulaExpression, FormulaTemplatePool } from "./formula.js";
import type { WorkbookFormulaSource, WorkbookValue } from "../contracts/workbook.js";
import type { WorkbookChunkTextCache } from "./chunk-text-cache.js";

// Private R3 storage baseline. The workbook reader will additionally validate
// formula syntax, styles, resources and cross-chunk references before opening.
export const CHUNK_ROWS = 256;
export const CHUNK_COLUMNS = 32;
export const MAX_SHEET_ROWS = 1_048_576;
export const MAX_SHEET_COLUMNS = 16_384;
export interface ChunkPosition { readonly sheetId: string; readonly row: number; readonly column: number }
export interface ChunkWindow { readonly row: number; readonly column: number; readonly rows: number; readonly columns: number }
export interface IndexedCellSource { readonly row: number; readonly column: number; readonly html: string }
export interface ChunkValidation {
  /** Internal callers already own detached immutable bytes. */
  readonly takeOwnership?: boolean;
  readonly textCache?: WorkbookChunkTextCache;
  readonly formulas?: FormulaTemplatePool;
  readonly cssClasses?: ReadonlySet<string>;
  readonly cellStyleClasses?: ReadonlySet<string>;
  readonly numberFormats?: ReadonlySet<string>;
  readonly maxNodes?: number;
  readonly maxDepth?: number;
}
export interface ChunkIndex {
  readonly position: ChunkPosition;
  readonly cellCount: number;
  readonly formulaCount: number;
  readonly sourceBytes: number;
  readonly indexBytes: number;
  readonly formulaTemplates: readonly FormulaExpression[];
  formulaLocations(): Iterable<{ readonly row: number; readonly column: number; readonly expression: FormulaExpression }>;
  /** Window coordinates are absolute, zero based. Returned HTML is source data. */
  readWindow(window: ChunkWindow): readonly IndexedCellSource[];
  getFormula(row: number, column: number): FormulaExpression | undefined;
  /** Always returns a copy. Immutable source can be shared by revisions internally. */
  serialize(): Uint8Array;
}

type Element = DefaultTreeAdapterTypes.Element;
type Node = DefaultTreeAdapterTypes.Node;
const tags = new Set(["html", "head", "body", "table", "tbody", "tr", "td", "span", "br"]);
const errors = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"]);
const numeric = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
const cellAttributes = ["data-column", "data-value-type", "data-value", "data-formula", "data-cache-state", "data-format-id", "class"] as const;
const cellAttributeNames = new Set<string>(cellAttributes);

function attributes(element: Element, allowed: readonly string[]): Record<string, string> {
  const result: Record<string, string> = Object.create(null) as Record<string, string>;
  for (const attribute of element.attrs) {
    if (attribute.namespace || !allowed.includes(attribute.name)) invalid("Unsupported workbook chunk attribute.");
    result[attribute.name] = attribute.value;
  }
  return result;
}
function elements(node: Node): Element[] {
  if (!("childNodes" in node)) return [];
  const result: Element[] = [];
  for (const child of node.childNodes) {
    if ("tagName" in child) result.push(child);
    else if (child.nodeName === "#documentType") continue;
    else if (child.nodeName !== "#text" || (child as DefaultTreeAdapterTypes.TextNode).value.trim()) invalid("Unexpected workbook chunk content.");
  }
  return result;
}
function only(node: Node, names: readonly string[]): Element[] {
  const children = elements(node);
  if (children.length !== names.length || children.some((child, i) => child.tagName !== names[i])) invalid("Invalid workbook chunk structure.");
  return children;
}
function integer(value: string | undefined): number {
  if (value === undefined || !/^(?:0|[1-9]\d*)$/.test(value)) invalid("Invalid workbook coordinate.");
  const number = Number(value);
  if (!Number.isSafeInteger(number)) invalid("Invalid workbook coordinate.");
  return number;
}
function validateClass(value: string | undefined, classes?: ReadonlySet<string>): void {
  if (value !== undefined && classes) for (const name of value.split(/[\t\n\f\r ]+/).filter(Boolean)) if (!classes.has(name)) invalid("A workbook CSS class is not defined.");
}
function textContent(node: Node, classes?: ReadonlySet<string>): string {
  if (node.nodeName === "#text") return (node as DefaultTreeAdapterTypes.TextNode).value;
  if (!("tagName" in node)) invalid("Unsupported workbook cell content.");
  if (node.tagName === "br") { attributes(node, []); return "\n"; }
  if (node.tagName !== "span" && node.tagName !== "td") invalid("Unsupported workbook rich text.");
  if (node.tagName === "span") validateClass(attributes(node, ["class"])["class"], classes);
  return node.childNodes.map(child => textContent(child, classes)).join("");
}
function validateCell(cell: Element, validation?: ChunkValidation): { column: number; formula?: string } {
  const attrs = attributes(cell, cellAttributes);
  const column = integer(attrs["data-column"]), type = attrs["data-value-type"], value = attrs["data-value"];
  validateClass(attrs["class"], validation?.cssClasses);
  if (validation?.cellStyleClasses?.size && attrs["class"] !== undefined) {
    const bindings = new Set(attrs["class"].split(/[\t\n\f\r ]+/).filter(name => validation.cellStyleClasses!.has(name)));
    if (bindings.size > 1) invalid("A workbook cell cannot select competing named style bindings.");
  }
  if (attrs["data-format-id"] !== undefined && validation?.numberFormats && !validation.numberFormats.has(attrs["data-format-id"])) invalid("A workbook number format is not defined.");
  const text = textContent(cell, validation?.cssClasses);
  checkLimit(text.length, 32_767);
  if (type === "number") { if (value === undefined || !numeric.test(value) || !Number.isFinite(Number(value))) invalid("Invalid workbook number."); }
  else if (type === "boolean") { if (value !== "true" && value !== "false") invalid("Invalid workbook boolean."); }
  else if (type === "error") { if (!errors.has(value ?? "")) invalid("Invalid workbook error."); }
  else if (type === "text" || type === "blank") {
    if (value !== undefined || (type === "blank" && text !== "")) invalid("Invalid workbook text or blank value.");
  } else invalid("Unsupported workbook value type.");
  const formula = attrs["data-formula"];
  if (formula !== undefined) {
    if (!formula.startsWith("=") || formula.length < 2) invalid("Invalid workbook formula source.");
    checkLimit(formula.length, 8_192);
    if (!["imported", "valid", "stale", "unavailable"].includes(attrs["data-cache-state"] ?? "")) invalid("Formula cache state is required.");
  } else if (attrs["data-cache-state"] !== undefined) invalid("Only formulas carry a cache state.");
  return { column, ...(formula !== undefined ? { formula } : {}) };
}

/** Validates one bounded source block, retaining bytes and compact offsets only. */
export function indexWorkbookChunk(input: Uint8Array, position: ChunkPosition, maxBytes = 8 * 1024 * 1024, validation?: ChunkValidation): ChunkIndex {
  const formulas = validation?.formulas;
  if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0) invalid("Invalid workbook chunk budget.");
  checkLimit(input.byteLength, maxBytes);
  if (!position.sheetId || !Number.isInteger(position.row) || !Number.isInteger(position.column)
    || position.row < 0 || position.row >= MAX_SHEET_ROWS || position.row % CHUNK_ROWS
    || position.column < 0 || position.column >= MAX_SHEET_COLUMNS || position.column % CHUNK_COLUMNS) invalid("Invalid workbook chunk position.");
  const bytes = validation?.takeOwnership ? input : new Uint8Array(input), location = Object.freeze({ ...position });
  const source = decodeText(bytes);
  const tree = parseHtmlTree(source, { maxFiles: 1, maxPackageBytes: maxBytes, maxFileBytes: maxBytes, maxNodes: validation?.maxNodes ?? 100_000, maxDepth: Math.min(16, validation?.maxDepth ?? 16), maxCssNodes: 1 }, tags);
  const [html] = only(tree, ["html"]), [head, body] = only(html!, ["head", "body"]);
  attributes(html!, []); attributes(head!, []); attributes(body!, []); only(head!, []);
  const [table] = only(body!, ["table"]), tableAttrs = attributes(table!, ["data-sheet-id", "data-row", "data-column"]);
  if (tableAttrs["data-sheet-id"] !== location.sheetId || integer(tableAttrs["data-row"]) !== location.row || integer(tableAttrs["data-column"]) !== location.column) invalid("Workbook chunk identity does not match its index.");
  const [tbody] = only(table!, ["tbody"]); attributes(tbody!, []);
  const coordinates: number[] = [], offsets: number[] = [], ends: number[] = [], formulaCoordinates: number[] = [], templateIds: number[] = [];
  let previousRow = -1, formulaCount = 0;
  for (const row of elements(tbody!)) {
    if (row.tagName !== "tr") invalid("Expected a workbook row.");
    const rowNumber = integer(attributes(row, ["data-row"])["data-row"]);
    if (rowNumber <= previousRow || rowNumber < location.row || rowNumber >= location.row + CHUNK_ROWS) invalid("Workbook rows must be ordered within their chunk.");
    previousRow = rowNumber;
    let previousColumn = -1;
    for (const cell of elements(row)) {
      if (cell.tagName !== "td") invalid("Expected a workbook cell.");
      const { column, formula } = validateCell(cell, validation);
      if (column <= previousColumn || column < location.column || column >= location.column + CHUNK_COLUMNS) invalid("Workbook cells must be ordered within their chunk.");
      previousColumn = column;
      const span = cell.sourceCodeLocation;
      if (!span?.startTag || !span.endTag) invalid("Workbook cells require explicit source boundaries.");
      coordinates.push((rowNumber - location.row) * CHUNK_COLUMNS + column - location.column);
      offsets.push(span.startOffset); ends.push(span.endOffset);
      if (formula !== undefined) {
        formulaCount++;
        if (formulas) {
          formulaCoordinates.push(coordinates[coordinates.length - 1]!);
          templateIds.push(formulas.intern(formula, { row: rowNumber, column }));
        }
      }
    }
  }
  const coords = Uint16Array.from(coordinates), starts = Uint32Array.from(offsets), stops = Uint32Array.from(ends);
  // Do not close over the temporary tree, decoded source or per-cell objects.
  return createIndex(bytes, location, coords, starts, stops, formulaCount, Uint16Array.from(formulaCoordinates), Uint32Array.from(templateIds), formulas, validation?.textCache);
}

function plainCell(source: string): { attrs: Record<string, string>; text: string } | undefined {
  if (!source.startsWith("<td") || !/[\t\n\f >]/.test(source[3] ?? "") || !source.endsWith("</td>") || /[&\r\u0000]/.test(source)) return undefined;
  const end = source.indexOf(">"), opening = source.slice(3, end), text = source.slice(end + 1, -5);
  if (end < 0 || opening.includes("<") || text.includes("<")) return undefined;
  const attrs = Object.create(null) as Record<string, string>, pattern = /[\t\n\f ]+([a-z][a-z0-9-]*)="([^"<>]*)"/gy; let cursor = 0;
  while (cursor < opening.length) {
    pattern.lastIndex = cursor; const match = pattern.exec(opening);
    if (!match) { if (!/^[\t\n\f ]*$/.test(opening.slice(cursor))) return undefined; break; }
    if (!cellAttributeNames.has(match[1]!) || Object.hasOwn(attrs, match[1]!)) return undefined;
    attrs[match[1]!] = match[2]!; cursor = pattern.lastIndex;
  }
  return { attrs, text };
}

/** Shared inert parser for bounded source slices and rich-input extraction. */
export function parseWorkbookCellElement(source: string): Element {
  const context = defaultTreeAdapter.createElement("tr", htmlNamespace.NS.HTML, []);
  const fragment = parseHtmlFragment(source, context, { maxFiles: 1, maxPackageBytes: 8 * 1024 * 1024, maxFileBytes: 8 * 1024 * 1024, maxNodes: 100_000, maxDepth: 16, maxCssNodes: 1 }, tags);
  const [cell] = only(fragment, ["td"]); return cell!;
}
/** Only used on already validated, bounded cell slices. Canonical entity-free
 * plain cells need no second HTML parse. Opening validates every source chunk. */
export function decodeWorkbookCell(source: string): { value: WorkbookValue; formula?: WorkbookFormulaSource; text: string; cssClass?: string; formatId?: string } {
  let plain = plainCell(source);
  if (!plain) {
    const cell = parseWorkbookCellElement(source);
    plain = { attrs: Object.fromEntries(cell.attrs.map(a => [a.name, a.value])), text: textContent(cell) };
  }
  const { attrs, text } = plain;
  const type = attrs["data-value-type"], raw = attrs["data-value"];
  const value: WorkbookValue = type === "number" ? { type, value: Number(raw) } : type === "boolean" ? { type, value: raw === "true" }
    : type === "error" ? { type, value: raw as Extract<WorkbookValue, { type: "error" }>["value"] } : type === "blank" ? { type } : { type: "text", value: text };
  return { value, text, ...(attrs["class"] !== undefined ? { cssClass: attrs["class"] } : {}), ...(attrs["data-format-id"] !== undefined ? { formatId: attrs["data-format-id"] } : {}),
    ...(attrs["data-formula"] !== undefined ? { formula: { expression: attrs["data-formula"], cacheState: attrs["data-cache-state"] as WorkbookFormulaSource["cacheState"] } } : {}) };
}

function createIndex(bytes: Uint8Array, position: ChunkPosition, coords: Uint16Array, starts: Uint32Array, stops: Uint32Array, formulaCount: number, formulaCoords: Uint16Array, templateIds: Uint32Array, formulas: FormulaTemplatePool | undefined, textCache: WorkbookChunkTextCache | undefined): ChunkIndex {
  const expressions = Array.from(templateIds, id => formulas?.get(id));
  const formulaTemplates = Object.freeze([...new Set(expressions.filter((item): item is FormulaExpression => item !== undefined))]);
  return Object.freeze({
    position, cellCount: coords.length, formulaCount, sourceBytes: bytes.byteLength,
    indexBytes: coords.byteLength + starts.byteLength + stops.byteLength + formulaCoords.byteLength + templateIds.byteLength,
    formulaTemplates,
    *formulaLocations() {
      for (let index = 0; index < formulaCoords.length; index++) {
        const coordinate = formulaCoords[index]!, expression = expressions[index];
        if (expression) yield { row: position.row + Math.floor(coordinate / CHUNK_COLUMNS), column: position.column + coordinate % CHUNK_COLUMNS, expression };
      }
    },
    serialize: () => new Uint8Array(bytes),
    getFormula(row: number, column: number): FormulaExpression | undefined {
      if (!Number.isInteger(row) || !Number.isInteger(column) || row < position.row || row >= position.row + CHUNK_ROWS || column < position.column || column >= position.column + CHUNK_COLUMNS) return undefined;
      const coordinate = (row - position.row) * CHUNK_COLUMNS + column - position.column;
      let low = 0, high = formulaCoords.length;
      while (low < high) { const middle = (low + high) >>> 1; if (formulaCoords[middle]! < coordinate) low = middle + 1; else high = middle; }
      return formulaCoords[low] === coordinate ? expressions[low] : undefined;
    },
    readWindow(window: ChunkWindow): readonly IndexedCellSource[] {
      if (![window.row, window.column, window.rows, window.columns].every(Number.isSafeInteger)
        || window.row < 0 || window.column < 0 || window.rows <= 0 || window.columns <= 0
        || window.row + window.rows > MAX_SHEET_ROWS || window.column + window.columns > MAX_SHEET_COLUMNS) invalid("Invalid workbook window.");
      const firstRow = Math.max(window.row, position.row), lastRow = Math.min(window.row + window.rows, position.row + CHUNK_ROWS);
      const firstColumn = Math.max(window.column, position.column), lastColumn = Math.min(window.column + window.columns, position.column + CHUNK_COLUMNS);
      if (firstRow >= lastRow || firstColumn >= lastColumn) return [];
      const source = textCache?.read(bytes) ?? decodeText(bytes), result: IndexedCellSource[] = [];
      let low = 0, high = coords.length;
      const first = (firstRow - position.row) * CHUNK_COLUMNS;
      while (low < high) { const middle = (low + high) >>> 1; if (coords[middle]! < first) low = middle + 1; else high = middle; }
      for (let i = low; i < coords.length; i++) {
        const coordinate = coords[i]!, row = position.row + Math.floor(coordinate / CHUNK_COLUMNS), column = position.column + coordinate % CHUNK_COLUMNS;
        if (row >= lastRow) break;
        if (column >= firstColumn && column < lastColumn) result.push(Object.freeze({ row, column, html: source.slice(starts[i], stops[i]) }));
      }
      return Object.freeze(result);
    },
  });
}
