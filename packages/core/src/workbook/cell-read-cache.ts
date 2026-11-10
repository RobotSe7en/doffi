import type { WorkbookCell } from "../runtime.js";
import { freeze } from "../document/validation.js";
import { createUtf8TextCopy } from "../document/package.js";
import { decodeWorkbookCell } from "./chunk-index.js";
import type { IndexedCellSource } from "./chunk-index.js";

/** A single immutable model owns this bounded derived LRU. Indexed source
 * bytes remain authoritative; no parse tree or implicit blank is retained. */
interface Entry { readonly key: string; readonly cell: WorkbookCell; readonly units: number; previous?: Entry | undefined; next?: Entry | undefined }
export class WorkbookCellReadCache {
  private readonly entries = new Map<string, Entry>();
  private first: Entry | undefined;
  private last: Entry | undefined;
  private units = 0;
  private readonly copyText = createUtf8TextCopy();
  constructor(private readonly maxEntries: number, private readonly maxUnits = 8 * 1024 * 1024) {}
  clear(): void { this.entries.clear(); this.first = undefined; this.last = undefined; this.units = 0; }
  private detach(entry: Entry): void {
    if (entry.previous) entry.previous.next = entry.next; else this.first = entry.next;
    if (entry.next) entry.next.previous = entry.previous; else this.last = entry.previous;
    entry.previous = undefined; entry.next = undefined;
  }
  private append(entry: Entry): void {
    entry.previous = this.last;
    if (this.last) this.last.next = entry; else this.first = entry;
    this.last = entry;
  }
  private lookup(key: string): WorkbookCell | undefined {
    const known = this.entries.get(key); if (!known) return undefined;
    if (known !== this.last) { this.detach(known); this.append(known); }
    return known.cell;
  }
  get(sheetId: string, row: number, column: number): WorkbookCell | undefined { return this.lookup(JSON.stringify([sheetId, row, column])); }
  read(sheetId: string, source: IndexedCellSource): WorkbookCell {
    const key = JSON.stringify([sheetId, source.row, source.column]), known = this.lookup(key);
    if (known) return known;
    // A substring may retain the entire decoded chunk in some engines. A UTF-8
    // copy detaches this already validated source slice before retaining it;
    // the decoder then also derives text/attributes from that small string.
    const html = this.copyText(source.html);
    const cell = freeze({ sheetId, row: source.row, column: source.column, html, ...decodeWorkbookCell(html) });
    const units = key.length + sheetId.length + cell.html.length + cell.text.length + (cell.value.type === "text" || cell.value.type === "error" ? cell.value.value.length : 0)
      + (cell.formula?.expression.length ?? 0) + (cell.cssClass?.length ?? 0) + (cell.formatId?.length ?? 0);
    if (units <= this.maxUnits && this.maxEntries > 0) {
      while (this.entries.size && (this.entries.size >= this.maxEntries || this.units + units > this.maxUnits)) {
        const previous = this.first!; this.detach(previous); this.entries.delete(previous.key); this.units -= previous.units;
      }
      const entry: Entry = { key, cell, units }; this.entries.set(key, entry); this.append(entry); this.units += units;
    }
    return cell;
  }
}
