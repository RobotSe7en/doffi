import { DocumentError } from "../errors.js";

export interface AxisOverride { readonly from: number; readonly to: number; readonly sizePt: number }
interface Segment extends AxisOverride { readonly startPt: number; readonly endPt: number }
export interface AxisCell { readonly index: number; readonly startPt: number; readonly sizePt: number }
const invalid = (message: string): never => { throw new DocumentError("InvalidOperation", message); };

/** Prefix sums over sparse dimension runs. Hidden cells have zero extent but
 * retain their source coordinates; neither lookup nor scrolling expands them. */
export class AxisDimensions {
  private readonly segments: Segment[] = [];
  readonly extentPt: number;
  constructor(readonly count: number, defaultSizePt: number, overrides: readonly AxisOverride[]) {
    if (!Number.isSafeInteger(count) || count < 1 || !Number.isFinite(defaultSizePt) || defaultSizePt < 0) invalid("Invalid worksheet axis dimensions.");
    let next = 0, extent = 0;
    const append = (from: number, to: number, sizePt: number): void => {
      if (to < from) return;
      const startPt = extent; extent += (to - from + 1) * sizePt;
      if (!Number.isFinite(extent)) invalid("Worksheet axis extent exceeds its numeric range.");
      const previous = this.segments.at(-1);
      if (previous?.sizePt === sizePt) this.segments[this.segments.length - 1] = { ...previous, to, endPt: extent };
      else this.segments.push({ from, to, sizePt, startPt, endPt: extent });
    };
    for (const run of overrides) {
      if (!Number.isSafeInteger(run.from) || !Number.isSafeInteger(run.to) || run.from < next || run.to < run.from || run.to >= count || !Number.isFinite(run.sizePt) || run.sizePt < 0) invalid("Worksheet dimension overrides must be finite, ordered and nonoverlapping.");
      append(next, run.from - 1, defaultSizePt); append(run.from, run.to, run.sizePt); next = run.to + 1;
    }
    append(next, count - 1, defaultSizePt); this.extentPt = extent;
  }
  get segmentCount(): number { return this.segments.length; }
  private atIndex(index: number): Segment {
    if (!Number.isSafeInteger(index) || index < 0 || index >= this.count) invalid("Worksheet axis index is outside its bounds.");
    let low = 0, high = this.segments.length;
    while (low < high) { const middle = (low + high) >>> 1; if (this.segments[middle]!.to < index) low = middle + 1; else high = middle; }
    return this.segments[low]!;
  }
  private atPosition(point: number): number {
    let low = 0, high = this.segments.length;
    while (low < high) { const middle = (low + high) >>> 1; if (this.segments[middle]!.endPt <= point) low = middle + 1; else high = middle; }
    return low;
  }
  /** The trailing boundary at count is valid, including an entirely hidden axis. */
  position(index: number): number {
    if (index === this.count) return this.extentPt;
    const run = this.atIndex(index); return run.startPt + (index - run.from) * run.sizePt;
  }
  size(index: number): number { return this.atIndex(index).sizePt; }
  /** Nearest displayed neighbor in source order, skipping whole hidden runs. */
  neighbor(index: number, step: -1 | 1): number | undefined {
    this.atIndex(index);
    if (step !== -1 && step !== 1) invalid("Worksheet neighbor direction must be -1 or 1.");
    let next = index + step;
    while (next >= 0 && next < this.count) {
      const run = this.atIndex(next);
      if (run.sizePt > 0 && run.endPt > run.startPt) return next;
      next = step === -1 ? run.from - 1 : run.to + 1;
    }
    return undefined;
  }
  private indexIn(run: Segment, point: number): number {
    let index = Math.min(run.to, Math.max(run.from, run.from + Math.floor((point - run.startPt) / run.sizePt)));
    // Compare the canonical positions at an exact edge, avoiding an off-by-one
    // when division of fractional row heights rounds just below an integer.
    if (index > run.from && point < run.startPt + (index - run.from) * run.sizePt) index--;
    if (index < run.to && point >= run.startPt + (index + 1 - run.from) * run.sizePt) index++;
    return index;
  }
  hit(point: number): number | undefined {
    if (!Number.isFinite(point)) invalid("Worksheet axis coordinates must be finite.");
    if (point < 0 || point >= this.extentPt) return undefined;
    const run = this.segments[this.atPosition(point)]!;
    return this.indexIn(run, point);
  }
  /** Returns only positive-extent cells intersecting [fromPt, toPt). A budget
   * failure is explicit; a truncated viewport must not masquerade as complete. */
  visible(fromPt: number, toPt: number, maxCells: number): readonly AxisCell[] {
    if (!Number.isFinite(fromPt) || !Number.isFinite(toPt) || toPt < fromPt || !Number.isSafeInteger(maxCells) || maxCells <= 0) invalid("Invalid worksheet axis viewport or cell budget.");
    const from = Math.max(0, fromPt), to = Math.min(this.extentPt, toPt), cells: AxisCell[] = [];
    if (from >= to) return cells;
    for (let i = this.atPosition(from); i < this.segments.length; i++) {
      const run = this.segments[i]!; if (run.startPt >= to) break;
      if (!run.sizePt || run.endPt <= from || run.endPt <= run.startPt) continue;
      for (let index = this.indexIn(run, Math.max(from, run.startPt)); index <= run.to; index++) {
        const startPt = run.startPt + (index - run.from) * run.sizePt; if (startPt >= to) break;
        if (cells.length === maxCells) throw new DocumentError("ResourceLimitExceeded", "The worksheet viewport exceeds its visible-cell budget.");
        cells.push({ index, startPt, sizePt: run.sizePt });
      }
    }
    return cells;
  }
}
