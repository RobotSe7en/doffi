import { DocumentError } from "@doffi/core";

/** One renderer owner retains bounded local-font measurements. Font-set
 * identity/descriptors invalidate them; no document text or DOM is shared. */
export class FontProbeCache {
  private readonly entries = new Map<string, number | boolean>();
  private readonly faces = new WeakMap<FontFace, number>();
  private faceId = 0;
  private fonts: FontFaceSet | undefined;
  private context: CanvasRenderingContext2D | undefined;
  private state: string | undefined;
  private epoch = 0;
  private units = 0;
  constructor(private readonly maxEntries: number = 4096, private readonly maxUnits: number = 4 * 1024 * 1024) {
    if (![maxEntries, maxUnits].every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Font probes require positive finite cache budgets.");
  }
  watch(fonts: FontFaceSet): { readonly ready: boolean; readonly epoch: number } {
    const state = JSON.stringify([...fonts].map(face => {
      let id = this.faces.get(face); if (id === undefined) { id = ++this.faceId; this.faces.set(face, id); }
      return [id, face.family, face.style, face.weight, face.stretch, face.unicodeRange, face.featureSettings, face.variationSettings,
        face.ascentOverride, face.descentOverride, face.lineGapOverride, face.display, face.status];
    }));
    const ready = fonts.status !== "loading";
    // fonts.ready can force source-table layout; this descriptor check does not.
    if (fonts !== this.fonts) this.context = undefined;
    if (fonts !== this.fonts || state !== this.state || !ready) { this.entries.clear(); this.units = 0; this.epoch++; }
    this.fonts = fonts; this.state = state;
    return { ready, epoch: this.epoch };
  }
  private read<Value extends number | boolean>(key: string, create: () => Value): Value {
    const known = this.entries.get(key);
    if (known !== undefined) { this.entries.delete(key); this.entries.set(key, known); return known as Value; }
    const value = create();
    if (key.length <= 4096 && key.length <= this.maxUnits) {
      while (this.entries.size && (this.entries.size >= this.maxEntries || this.units + key.length > this.maxUnits)) {
        const first = this.entries.keys().next().value!; this.entries.delete(first); this.units -= first.length;
      }
      this.entries.set(key, value); this.units += key.length;
    }
    return value;
  }
  available(family: string, create: () => boolean): boolean { return this.read("a:" + family, create); }
  glyph(key: string, create: () => number): number { return this.read("g:" + key, create); }
  canvas(create: () => CanvasRenderingContext2D): CanvasRenderingContext2D {
    return this.context ??= create();
  }
  clear(): void { this.entries.clear(); this.units = 0; this.context = undefined; this.fonts = undefined; this.state = undefined; this.epoch++; }
}
