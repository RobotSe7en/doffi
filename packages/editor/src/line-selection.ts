import type { PageSetup } from '@doffi/core';
import type { DocumentRenderReport } from '@doffi/render';

const blocks = 'p,h1,h2,h3,h4,h5,h6,li,td,th';
const selectionCursor = "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='20' height='23' viewBox='0 0 20 23'%3E%3Cpath d='M18 2L18 17L14 13L11 20L8 18L11 12L5 12Z' fill='white' stroke='%23212832' stroke-width='1.2'/%3E%3C/svg%3E\") 18 2, default";
/** Word's left selection bar uses rendered lines, including soft wraps. It
 * changes the native selection only; input/history retain their existing owner. */
export function bindWordLineSelection(frame: Document, root: HTMLElement, options: { signal: AbortSignal; enabled(): boolean; ignored(node: Node): boolean; layout(): DocumentRenderReport | undefined; setup(): Readonly<PageSetup>; changed(): void }): void {
  const cursorStyle = frame.createElement('style'); cursorStyle.textContent = `html[data-doffi-line-cursor],html[data-doffi-line-cursor] *{cursor:${selectionCursor}!important}`; frame.head.append(cursorStyle);
  const clearCursor = (): void => { frame.documentElement.removeAttribute('data-doffi-line-cursor'); };
  frame.documentElement.addEventListener('mouseleave', clearCursor, { signal: options.signal });
  let dragging: Range | undefined, pressed = false, scheduled: number | undefined, latest: { x: number; y: number } | undefined;
  function margin(x: number, y: number, extending = false): number | undefined {
    if (!options.enabled()) return;
    const box = root.getBoundingClientRect(), report = options.layout(), page = report?.layout === 'pages' ? report.pages.find(page => y >= box.top + page.rect.y * 4 / 3 && y < box.top + (page.rect.y + page.rect.height) * 4 / 3) : undefined;
    if (report?.layout === 'pages' && (!page || page.blank)) return;
    const setup = page?.setup ?? options.setup(), left = box.left + (page?.rect.x ?? 0) * 4 / 3, content = left + setup.margins.left * 4 / 3;
    if (!extending && (x < left || x >= content - 3)) return;
    return content;
  }
  function line(x: number, y: number, extending = false): { range: Range; block: HTMLElement } | undefined {
    const content = margin(x, y, extending); if (content === undefined) return;
    const caret = frame.caretPositionFromPoint?.(content + 2, y), native = caret ? undefined : frame.caretRangeFromPoint?.(content + 2, y);
    const node = caret?.offsetNode ?? native?.startContainer, offset = caret?.offset ?? native?.startOffset;
    if (!node || offset === undefined || !root.contains(node) || options.ignored(node)) return;
    const element = node.nodeType === 1 ? node as HTMLElement : node.parentElement, block = element?.closest<HTMLElement>(blocks);
    if (!block || !root.contains(block) || block.closest('td,th')) return;
    const all = frame.createRange(); all.selectNodeContents(block);
    const rects = [...all.getClientRects()].filter(rect => rect.width && rect.height);
    const band = rects.find(rect => y >= rect.top && y < rect.bottom) ?? rects.find(rect => y >= rect.top - 3 && y < rect.bottom + 3);
    if (!band) return;
    const walker = frame.createTreeWalker(block, 4), result = frame.createRange(); let text: Node | null, found = false;
    while ((text = walker.nextNode())) {
      if (!text.textContent || options.ignored(text)) continue;
      const current = text, offsets = [0]; for (const character of current.textContent!) offsets.push(offsets.at(-1)! + character.length);
      const rect = (index: number): DOMRect | undefined => { const range = frame.createRange(); range.setStart(current, offsets[index]!); range.setEnd(current, offsets[index + 1]!); return [...range.getClientRects()].find(rect => rect.height > 0); };
      const lower = (after: boolean): number => { let low = 0, high = offsets.length - 1; while (low < high) { const middle = (low + high) >>> 1, current = rect(middle); const before = current && (after ? current.top < band.bottom - 1 : current.bottom <= band.top + 1); if (before) low = middle + 1; else high = middle; } return low; };
      const from = lower(false), to = lower(true); if (from >= to) continue;
      if (!found) { result.setStart(current, offsets[from]!); found = true; } result.setEnd(current, offsets[to]!);
    }
    return found ? { range: result, block } : undefined;
  }
  function select(range: Range): void {
    frame.defaultView!.focus(); if (root.isContentEditable) root.focus({ preventScroll: true });
    const selected = frame.getSelection()!; selected.removeAllRanges(); selected.addRange(range); options.changed();
  }
  frame.addEventListener('mousedown', event => {
    if (event.button !== 0) return;
    const selected = line(event.clientX, event.clientY); if (!selected) return;
    event.preventDefault(); event.stopImmediatePropagation(); pressed = true;
    if (event.detail >= 3 || event.ctrlKey || event.metaKey) selected.range.selectNodeContents(root);
    else if (event.detail === 2) selected.range.selectNodeContents(selected.block);
    if (event.shiftKey) {
      const native = frame.getSelection(); if (native?.anchorNode && root.contains(native.anchorNode)) {
        const anchor = frame.createRange(); anchor.setStart(native.anchorNode, native.anchorOffset); anchor.collapse(true);
        if (selected.range.compareBoundaryPoints(0, anchor) < 0) selected.range.setEnd(anchor.startContainer, anchor.startOffset); else selected.range.setStart(anchor.startContainer, anchor.startOffset);
      }
    }
    dragging = selected.range.cloneRange(); select(selected.range);
  }, { signal: options.signal });
  function extend(): void {
    if (!dragging || !pressed || !latest) return;
    const selected = line(latest.x, latest.y, true); if (!selected) return;
    const range = dragging.cloneRange();
    if (selected.range.compareBoundaryPoints(0, dragging) < 0) range.setStart(selected.range.startContainer, selected.range.startOffset);
    else range.setEnd(selected.range.endContainer, selected.range.endOffset);
    select(range);
  }
  frame.addEventListener('mousemove', event => {
    const inside = margin(event.clientX, event.clientY) !== undefined;
    frame.documentElement.toggleAttribute('data-doffi-line-cursor', inside);
    if (!dragging || !pressed || !event.buttons) return;
    event.preventDefault(); latest = { x: event.clientX, y: event.clientY };
    if (scheduled === undefined) scheduled = frame.defaultView!.requestAnimationFrame(() => { scheduled = undefined; extend(); });
  }, { signal: options.signal });
  frame.addEventListener('mouseup', () => { extend(); if (pressed) options.changed(); pressed = false; dragging = undefined; latest = undefined; }, { signal: options.signal });
  frame.addEventListener('dblclick', event => { if (margin(event.clientX, event.clientY) !== undefined) event.preventDefault(); }, { signal: options.signal });
  frame.addEventListener('click', event => { if (line(event.clientX, event.clientY)) event.preventDefault(); }, { signal: options.signal });
  options.signal.addEventListener('abort', () => { clearCursor(); cursorStyle.remove(); if (scheduled !== undefined) frame.defaultView!.cancelAnimationFrame(scheduled); }, { once: true });
}
