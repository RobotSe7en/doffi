import { decodeText } from "../document/package.js";

/** Model-owned decoded source text, bounded independently of immutable bytes
 * and decoded cells. No parser nodes or metadata become authoritative here. */
export class WorkbookChunkTextCache {
  private readonly entries = new Map<Uint8Array, string>();
  private units = 0;
  constructor(private readonly maxEntries = 4, private readonly maxUnits = 8 * 1024 * 1024) {}
  get retainedCodeUnits(): number { return this.units; }
  get retainedChunks(): number { return this.entries.size; }
  clear(): void { this.entries.clear(); this.units = 0; }
  read(bytes: Uint8Array): string {
    const known = this.entries.get(bytes);
    if (known !== undefined) { this.entries.delete(bytes); this.entries.set(bytes, known); return known; }
    const source = decodeText(bytes);
    if (source.length <= this.maxUnits && this.maxEntries > 0) {
      while (this.entries.size && (this.entries.size >= this.maxEntries || this.units + source.length > this.maxUnits)) {
        const first = this.entries.keys().next().value!, previous = this.entries.get(first)!; this.entries.delete(first); this.units -= previous.length;
      }
      this.entries.set(bytes, source); this.units += source.length;
    }
    return source;
  }
}
