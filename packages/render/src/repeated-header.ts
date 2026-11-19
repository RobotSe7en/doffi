/** A derived flow may use a different first and continuation header. The
 * alternate remains measured at the same width, outside the authoritative flow. */
const continuations = new WeakMap<HTMLTableRowElement, HTMLTableRowElement>();
export function setContinuationHeader(first: HTMLTableRowElement, next: HTMLTableRowElement): void {
  continuations.set(first, next);
}
export function continuationHeader(first: HTMLTableRowElement): HTMLTableRowElement {
  return continuations.get(first) ?? first;
}
