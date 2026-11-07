import type { WorkbookCellFormatInput, WorkbookCellInput } from "../contracts/workbook-operations.js";
import type { WorkbookManifest } from "../contracts/document.js";
import type { WorkbookChunkDescriptor, WorkbookRange } from "../contracts/workbook.js";
import type { DocumentReadLimits, WorkbookCell, WorkbookModel } from "../runtime.js";
import { assetDigest } from "../document/assets.js";
import { decodeText, encode } from "../document/package.js";
import { checkLimit, invalid } from "../document/validation.js";
import { CHUNK_COLUMNS, CHUNK_ROWS, decodeWorkbookCell, indexWorkbookChunk, parseWorkbookCellElement } from "../workbook/chunk-index.js";
import { parseFormula } from "../workbook/formula.js";
import { checkRange } from "../workbook/ranges.js";
import { workbookSource } from "../workbook/source.js";

function checkWorkbookUnicode(value: string): void {
  for (const character of value) {
    const point = character.codePointAt(0)!;
    if (point >= 0xd800 && point <= 0xdfff || point === 0) invalid("Workbook text must contain well-formed Unicode without null characters.");
  }
}
export function escapeWorkbookText(value: string): string {
  checkWorkbookUnicode(value);
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\r/g, "&#13;");
}
export function rangeContains(range: WorkbookRange, row: number, column: number): boolean {
  return row >= range.from.row && row <= range.to.row && column >= range.from.column && column <= range.to.column;
}
export function rangesIntersect(a: WorkbookRange, b: WorkbookRange): boolean {
  return a.from.row <= b.to.row && a.to.row >= b.from.row && a.from.column <= b.to.column && a.to.column >= b.from.column;
}
interface EditedChunk { descriptor: WorkbookChunkDescriptor; cells: Map<number, string>; pending: Set<number> }
const coordinate = (row: number, column: number): number => row * 16_384 + column;

/** A preparation owns only touched, bounded cell blocks. Other bytes stay shared. */
export class WorkbookCellEdit {
  private readonly chunks = new Map<string, EditedChunk>();
  private readonly originals: ReadonlyMap<string, import("../workbook/chunk-index.js").ChunkIndex>;
  constructor(private readonly base: WorkbookModel, readonly manifest: WorkbookManifest,
    readonly files: Map<string, Uint8Array>, private readonly limits: DocumentReadLimits, private readonly allocate: () => string) {
    this.originals = workbookSource(base)!.chunks;
  }
  private key(sheetId: string, row: number, column: number): string { return JSON.stringify([sheetId, Math.floor(row / CHUNK_ROWS), Math.floor(column / CHUNK_COLUMNS)]); }
  private chunk(sheetId: string, row: number, column: number): EditedChunk {
    const key = this.key(sheetId, row, column), known = this.chunks.get(key); if (known) return known;
    const position = { sheetId, row: Math.floor(row / CHUNK_ROWS) * CHUNK_ROWS, column: Math.floor(column / CHUNK_COLUMNS) * CHUNK_COLUMNS };
    let descriptor = this.manifest.chunks.find(chunk => chunk.sheetId === sheetId && chunk.row === position.row && chunk.column === position.column);
    if (!descriptor) {
      let path: string; do { path = "content/" + this.allocate() + ".html"; } while (this.files.has(path));
      // File quotas apply to the final candidate, independent of operation
      // order. Intermediate descriptors are bounded by the transaction and
      // node budget; flush plus package validation checks final file counts.
      checkLimit(this.chunks.size + 1, this.limits.maxNodes);
      descriptor = { ...position, path, sha256: "0".repeat(64) }; this.manifest.chunks.push(descriptor);
    }
    const index = this.originals.get(descriptor.path), cells = new Map<number, string>();
    for (const cell of index?.readWindow({ ...position, rows: CHUNK_ROWS, columns: CHUNK_COLUMNS }) ?? []) cells.set(coordinate(cell.row, cell.column), cell.html);
    const result = { descriptor, cells, pending: new Set<number>() }; this.chunks.set(key, result); return result;
  }
  get(sheetId: string, row: number, column: number): WorkbookCell | undefined {
    const chunk = this.chunks.get(this.key(sheetId, row, column));
    if (!chunk) return this.base.getCell(sheetId, row, column);
    const html = chunk.cells.get(coordinate(row, column)); return html === undefined ? undefined : { sheetId, row, column, html, ...decodeWorkbookCell(html) };
  }
  *entries(sheetId?: string, formulasOnly = false): Generator<WorkbookCell> {
    for (const descriptor of [...this.manifest.chunks]) {
      if (sheetId !== undefined && descriptor.sheetId !== sheetId) continue;
      const edited = this.chunks.get(this.key(descriptor.sheetId, descriptor.row, descriptor.column));
      if (edited) {
        for (const [key, html] of [...edited.cells]) {
          const cell = { sheetId: descriptor.sheetId, row: Math.floor(key / 16_384), column: key % 16_384, html, ...decodeWorkbookCell(html) };
          if (!formulasOnly || cell.formula) yield cell;
        }
      } else if (formulasOnly) {
        for (const location of this.originals.get(descriptor.path)?.formulaLocations() ?? []) yield this.base.getCell(descriptor.sheetId, location.row, location.column)!;
      } else for (const cell of this.originals.get(descriptor.path)?.readWindow({ ...descriptor, rows: CHUNK_ROWS, columns: CHUNK_COLUMNS }) ?? []) yield { sheetId: descriptor.sheetId, ...cell, ...decodeWorkbookCell(cell.html) };
    }
  }
  removeSheet(sheetId: string): void {
    for (const descriptor of this.manifest.chunks.filter(chunk => chunk.sheetId === sheetId)) {
      const edited = this.chunks.get(this.key(sheetId, descriptor.row, descriptor.column));
      if (edited) for (const key of edited.pending) this.validatePending(edited, key);
      this.files.delete(descriptor.path); this.chunks.delete(this.key(sheetId, descriptor.row, descriptor.column));
    }
    this.manifest.chunks = this.manifest.chunks.filter(chunk => chunk.sheetId !== sheetId);
  }
  set(sheetId: string, input: WorkbookCellInput): void {
    const previous = this.get(sheetId, input.row, input.column);
    const cssClass = input.cssClass === undefined ? previous?.cssClass : input.cssClass;
    const formatId = input.formatId === undefined ? previous?.formatId : input.formatId;
    const sameStyle = (cssClass ?? undefined) === previous?.cssClass && (formatId ?? undefined) === previous?.formatId;
    if (!previous && !cssClass && !formatId && input.source.kind === "value" && input.source.value.type === "blank") return;
    if (previous && sameStyle && (input.source.kind === "value" && !previous.formula && JSON.stringify(input.source.value) === JSON.stringify(previous.value)
      || input.source.kind === "formula" && previous.formula?.expression === input.source.expression)) return;
    const attrs: Record<string, string> = { "data-column": String(input.column) };
    let html: string;
    if (input.source.kind === "formula") {
      parseFormula(input.source.expression, input);
      attrs["data-value-type"] = "blank"; attrs["data-formula"] = input.source.expression; attrs["data-cache-state"] = "stale"; html = "";
    } else if (input.source.kind === "richText") { checkWorkbookUnicode(input.source.html); attrs["data-value-type"] = "text"; html = input.source.html; }
    else {
      const value = input.source.value; attrs["data-value-type"] = value.type;
      if (value.type !== "blank" && value.type !== "text") attrs["data-value"] = String(value.value);
      html = value.type === "blank" ? "" : escapeWorkbookText(value.type === "boolean" ? value.value ? "TRUE" : "FALSE" : String(value.value)).replace(/\n/g, "<br>");
    }
    if (cssClass) attrs.class = cssClass; if (formatId) attrs["data-format-id"] = formatId;
    const source = "<td" + Object.entries(attrs).map(([name, value]) => " " + name + '="' + escapeWorkbookText(value) + '"').join("") + ">" + html + "</td>";
    this.store(sheetId, input.row, input.column, source, attrs["data-value-type"] === "blank" && !attrs["data-formula"] && !cssClass && !formatId);
  }
  /** Attributes alone change; formula caches and the exact rich-text body stay
   * source bytes. A format-only edit never passes through typed value input. */
  format(sheetId: string, input: WorkbookCellFormatInput): void {
    const previous = this.get(sheetId, input.row, input.column);
    const cssClass = input.cssClass === undefined ? previous?.cssClass : input.cssClass;
    const formatId = input.formatId === undefined ? previous?.formatId : input.formatId;
    if ((cssClass ?? undefined) === previous?.cssClass && (formatId ?? undefined) === previous?.formatId) return;
    if (!previous) { this.set(sheetId, { row: input.row, column: input.column, source: { kind: "value", value: { type: "blank" } }, ...(cssClass ? { cssClass } : {}), ...(formatId ? { formatId } : {}) }); return; }
    const root = parseWorkbookCellElement(previous.html), attrs = Object.fromEntries(root.attrs.map(attr => [attr.name, attr.value]));
    if (cssClass) attrs.class = cssClass; else delete attrs.class;
    if (formatId) attrs["data-format-id"] = formatId; else delete attrs["data-format-id"];
    const tail = (html: string): string => {
      let quote = "";
      for (let index = 0; index < html.length; index++) {
        const character = html[index]!;
        if (quote) { if (character === quote) quote = ""; }
        else if (character === '"' || character === "'") quote = character;
        else if (character === ">") return html.slice(index + 1);
      }
      return invalid("A workbook cell has no complete opening tag.");
    };
    const body = tail(previous.html);
    let source = "<td" + Object.entries(attrs).map(([name, value]) => " " + name + '="' + escapeWorkbookText(value) + '"').join("") + ">" + body;
    const original = this.base.getCell(sheetId, input.row, input.column);
    if (original && (cssClass ?? undefined) === original.cssClass && (formatId ?? undefined) === original.formatId && body === tail(original.html)) {
      const originalAttrs = Object.fromEntries(parseWorkbookCellElement(original.html).attrs.map(attr => [attr.name, attr.value]));
      if (Object.keys(attrs).length === Object.keys(originalAttrs).length && Object.entries(attrs).every(([key, value]) => originalAttrs[key] === value)) source = original.html;
    }
    this.store(sheetId, input.row, input.column, source, previous.value.type === "blank" && !previous.formula && !cssClass && !formatId && root.childNodes.length === 0);
  }
  private store(sheetId: string, row: number, column: number, source: string, empty: boolean): void {
    const chunk = this.chunk(sheetId, row, column), key = coordinate(row, column);
    this.validatePending(chunk, key);
    if (empty) chunk.cells.delete(key);
    else { chunk.cells.set(key, source); chunk.pending.add(key); }
  }
  /** Final blocks are validated together by reopen. Inputs that disappear
   * before that boundary still receive the same bounded single-cell check. */
  private validatePending(chunk: EditedChunk, key: number): void {
    if (!chunk.pending.has(key)) return;
    const source = chunk.cells.get(key)!;
    const probe = this.serialize(chunk.descriptor, new Map([[key, source]]));
    indexWorkbookChunk(probe, chunk.descriptor, this.limits.maxFileBytes, { maxNodes: this.limits.maxNodes, maxDepth: this.limits.maxDepth });
    chunk.pending.delete(key);
  }
  *clearSteps(sheetId: string, ranges: readonly WorkbookRange[], mode: "contents" | "all"): Generator<{ completed: number; total: number }, void> {
    for (const range of ranges) checkRange(range);
    const descriptors = this.manifest.chunks.filter(descriptor => descriptor.sheetId === sheetId && ranges.some(range => rangesIntersect(range, { from: descriptor, to: { row: descriptor.row + CHUNK_ROWS - 1, column: descriptor.column + CHUNK_COLUMNS - 1 } })));
    const total = descriptors.reduce((count, descriptor) => count + (this.chunks.get(this.key(sheetId, descriptor.row, descriptor.column))?.cells.size ?? this.originals.get(descriptor.path)?.cellCount ?? 0), 0);
    let completed = 0;
    for (const descriptor of descriptors) {
      const chunk = this.chunk(sheetId, descriptor.row, descriptor.column);
      for (const key of [...chunk.cells.keys()]) {
        const row = Math.floor(key / 16_384), column = key % 16_384;
        if (ranges.some(range => rangeContains(range, row, column))) {
          if (mode === "all") { this.validatePending(chunk, key); chunk.cells.delete(key); }
          else this.set(sheetId, { row, column, source: { kind: "value", value: { type: "blank" } } });
        }
        if (++completed % 256 === 0) yield { completed, total };
      }
    }
    if (completed % 256 !== 0 || completed === 0) yield { completed, total };
  }
  private serialize(descriptor: WorkbookChunkDescriptor, cells: ReadonlyMap<number, string>): Uint8Array {
    let html = '<!doctype html><html><head></head><body><table data-sheet-id="' + escapeWorkbookText(descriptor.sheetId) + '" data-row="' + descriptor.row + '" data-column="' + descriptor.column + '"><tbody>';
    let previous = -1;
    for (const [key, cell] of [...cells].sort((a, b) => a[0] - b[0])) {
      const row = Math.floor(key / 16_384);
      if (row !== previous) { if (previous >= 0) html += "</tr>"; html += '<tr data-row="' + row + '">'; previous = row; }
      html += cell; checkLimit(html.length, this.limits.maxFileBytes);
    }
    if (previous >= 0) html += "</tr>";
    const bytes = encode(html + "</tbody></table></body></html>"); checkLimit(bytes.byteLength, this.limits.maxFileBytes); return bytes;
  }
  *flushSteps(): Generator<{ completed: number; total: number }, void> {
    let completed = 0; const total = this.chunks.size;
    for (const chunk of this.chunks.values()) {
      if (!chunk.cells.size) { this.files.delete(chunk.descriptor.path); this.manifest.chunks = this.manifest.chunks.filter(item => item.path !== chunk.descriptor.path); yield { completed: ++completed, total }; continue; }
      const bytes = this.serialize(chunk.descriptor, chunk.cells), previous = this.files.get(chunk.descriptor.path);
      // Preserve original bytes when the final cell set is unchanged, including source syntax.
      const original = this.originals.get(chunk.descriptor.path);
      const originalCells = original?.readWindow({ ...chunk.descriptor, rows: CHUNK_ROWS, columns: CHUNK_COLUMNS });
      if (!(previous && originalCells?.length === chunk.cells.size && originalCells.every(cell => chunk.cells.get(coordinate(cell.row, cell.column)) === cell.html))
        && !(previous && decodeText(previous) === decodeText(bytes))) { this.files.set(chunk.descriptor.path, bytes); chunk.descriptor.sha256 = assetDigest(bytes); }
      yield { completed: ++completed, total };
    }
  }
}
