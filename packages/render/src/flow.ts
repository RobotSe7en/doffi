import{wordTableRepeatHeader}from'./word-table-layout.js';
import { DocumentError } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { isMeasuredViewContent, markViewNode, ownViewAttribute, ownViewStyle, viewNodeRole } from "./view-nodes.js";
import{wordRowVisibleRect}from'./word-table-layout.js';

export interface PageFlow {
  reset(): void;
  layout(element: HTMLElement, forcePage?: boolean, parity?: number): void;
  readonly count: number;
  readonly blanks: ReadonlySet<number>;
  pageAt(element: HTMLElement): number;
  headerHeight(table: HTMLTableElement): number;
  project?():void;
}
export interface PageFlowConstraints{reserved(index:number):number;minimumPage(node:Node):number|undefined}
/** Portable single-DOM fragmentation. Empty, owned spacers carry page margins;
 * source paragraphs and cells keep their identity even when a text node splits. */
export function pageFlow(document: Document, flow: HTMLElement, area: number, pitch: number, limits: Readonly<RenderLimits>, pageStride=1,constraints?:PageFlowConstraints): PageFlow {
  const view = document.defaultView!, gaps: HTMLElement[] = [], blanks = new Set<number>();
  const excluded = (node: Node): boolean => {
    for (let item: Node | null = node; item && item !== flow; item = item.parentNode)
      if (viewNodeRole(item) === "remove" && !isMeasuredViewContent(item)) return true;
    return false;
  };
  const canonicalSpans = new Map<HTMLTableCellElement, number>();
  let last = 0, operations = 0;
  const origin = (): number => flow.getBoundingClientRect().top;
  const page = (y: number): number => Math.max(0, Math.floor((y - origin() + 0.25) / pitch));
  const start = (index: number): number => origin() + index * pitch;
  const available=(index:number):number=>Math.max(Math.min(24,area),area-(constraints?.reserved(index)??0));
  const ensure = (index: number): void => {
    if (index >= limits.maxPages) throw new DocumentError("ResourceLimitExceeded", "Document exceeds the render page budget.");
    last = Math.max(last, index);
    if (++operations > limits.maxNodes + limits.maxPages * 32) throw new DocumentError("ResourceLimitExceeded", "Pagination exceeds the measurement budget.");
  };
  const topPage = (element: HTMLElement): number => page(element.getBoundingClientRect().top);
  const nextPhysical = (index:number):number => (Math.floor(index/pageStride)+1)*pageStride;
  function spacer(parent: Node, before: Node | null, target: number): HTMLElement {
    const gap = markViewNode(document.createElement("span")); gap.className = "sdk-page-gap";
    gap.contentEditable = "false"; gap.setAttribute("aria-hidden", "true");
    gap.style.cssText = "display:block;height:0;margin:0;padding:0;border:0;line-height:0;pointer-events:none;user-select:none";
    parent.insertBefore(gap, before); gaps.push(gap);
    gap.style.height = Math.max(0, target - gap.getBoundingClientRect().top) + "px";
    return gap;
  }
  function before(element: HTMLElement, targetPage: number, inset = 0): void {
    ensure(targetPage); spacer(element.parentNode!, element, start(targetPage) + inset);
  }
  function headerRows(table: HTMLTableElement): HTMLTableRowElement[] {
    const rows = [...table.rows].filter((row) => viewNodeRole(row) !== "remove"&&view.getComputedStyle(row).display!=="none"&&view.getComputedStyle(row).visibility!=="collapse"), result: HTMLTableRowElement[] = [];
    for (const row of rows) { if (!wordTableRepeatHeader(row)) break; result.push(row); }
    return result;
  }
  const headerHeight = (table: HTMLTableElement): number => headerRows(table).reduce((sum, row) => sum + row.getBoundingClientRect().height, 0);
  function rowGap(table: HTMLTableElement, row: HTMLTableRowElement, targetPage: number, inset: number): void {
    ensure(targetPage);
    const rows = [...table.rows].filter((item) => viewNodeRole(item) !== "remove"&&view.getComputedStyle(item).display!=="none"&&view.getComputedStyle(item).visibility!=="collapse"), ordinal = rows.indexOf(row);
    for (const [index, previous] of rows.entries()) for (const cell of previous.cells) {
      const span = canonicalSpans.get(cell) ?? cell.rowSpan;
      canonicalSpans.set(cell, span);
      if (span && index < ordinal && index + span > ordinal) { ownViewAttribute(cell, "rowspan"); cell.rowSpan++; }
    }
    const gap = markViewNode(document.createElement("tr")); gap.className = "sdk-page-gap"; gap.setAttribute("aria-hidden", "true"); gap.contentEditable = "false";
    const cell = document.createElement("td"); cell.colSpan = Math.max(1, ...rows.map((item) => [...item.cells].reduce((sum, value) => sum + value.colSpan, 0)));
    cell.style.cssText = "padding:0;border:0;line-height:0;font-size:0"; gap.append(cell);
    row.before(gap); gaps.push(gap);
    const height = Math.max(0, start(targetPage) + inset - gap.getBoundingClientRect().top);
    gap.style.height = height + "px"; cell.style.height = height + "px";
  }
  interface Line { readonly node: Text | HTMLElement; readonly rect: DOMRect; readonly offset?: number;readonly targetPage?:number }
  function visualLines(element: HTMLElement): Line[] {
    const result: Line[] = [], walker = document.createTreeWalker(element, 5);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const owner = node.nodeType === 1 ? node as HTMLElement : node.parentElement!;
      if(owner.closest('[data-word-drawing]')&&(!owner.hasAttribute('data-word-drawing')||node.nodeType===3))continue;
      if(owner.hasAttribute('data-word-drawing')&&view.getComputedStyle(owner.closest<HTMLElement>('[data-sdk-word-object]')??owner).position==='absolute')continue;
      if (excluded(owner) || view.getComputedStyle(owner).visibility === "hidden") continue;
      if (node.nodeType !== 3 && !owner.matches("img,[data-word-tab],[data-word-drawing],[data-word-note-reference],[data-word-equation]")) continue;
      const range = document.createRange(); range.selectNodeContents(node);
      for (const rect of node.nodeType === 3 ? range.getClientRects() : owner.getClientRects()) {
        if(!wordRowVisibleRect(owner,rect))continue;
        if (rect.width < 0.01 || rect.height < 0.1) continue;
        const previous = result.at(-1);
        if (previous && Math.min(previous.rect.bottom, rect.bottom) - Math.max(previous.rect.top, rect.top) > Math.min(previous.rect.height, rect.height) / 2) continue;
        let lo = 0, hi = node.nodeType === 3 ? (node as Text).length : 0;
        while (lo < hi) {
          const mid = Math.floor((lo + hi) / 2), probe = document.createRange();
          probe.setStart(node, mid); probe.setEnd(node, Math.min((node as Text).length, mid + 1));
          if (probe.getBoundingClientRect().top < rect.top - 0.5) lo = mid + 1; else hi = mid;
        }
        if (node.nodeType === 3 && lo && /[\udc00-\udfff]/.test((node as Text).data[lo] ?? "")) lo--;
        result.push({ node: node as Text | HTMLElement, rect, ...(node.nodeType === 3 ? { offset: lo } : {}) });
      }
    }
    return result;
  }
  function overflow(element: HTMLElement, inset: number): Line | undefined {
    const walker = document.createTreeWalker(element, 5 /* ELEMENT | TEXT */);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const owner = node.nodeType === 1 ? node as HTMLElement : node.parentElement!;
      if(owner.closest('[data-word-drawing]')&&(!owner.hasAttribute('data-word-drawing')||node.nodeType===3))continue;
      if(owner.hasAttribute('data-word-drawing')&&view.getComputedStyle(owner.closest<HTMLElement>('[data-sdk-word-object]')??owner).position==='absolute')continue;
      if (excluded(owner) || view.getComputedStyle(owner).visibility === "hidden") continue;
      if (node.nodeType === 1) {
        if (owner.matches("br[data-word-break=page],br[data-word-break=column]")) {
          if (owner.style.display === "none") continue;
          return { node: owner, rect: owner.getBoundingClientRect() };
        }
        if (!owner.matches("img,[data-word-tab],[data-word-drawing],[data-word-note-reference],[data-word-equation]")) continue;
      }
      const range = document.createRange(); range.selectNodeContents(node);
      const rectangles = node.nodeType === 3 ? [...range.getClientRects()] : [...(node as Element).getClientRects()];
      for (const rect of rectangles) {
        if(!wordRowVisibleRect(owner,rect))continue;
        if (rect.height < 0.1 || rect.width < 0.01) continue;
        const current = page(rect.top + rect.height / 2), end = start(current) + available(current),minimum=constraints?.minimumPage(node);if(minimum!==undefined&&current<minimum)return{node:node as HTMLElement,rect,targetPage:minimum};
        // Glyph rectangles can extend beyond a line box by sub-pixel rounding.
        const height = Math.max(rect.height, parseFloat(view.getComputedStyle(owner).lineHeight) || rect.height);
        if (rect.top >= start(current) + inset - 5 && rect.bottom + Math.max(0, height - rect.height) / 2 <= end + 0.5) { ensure(current); continue; }
        if (height > area - inset + 0.5) throw new DocumentError("UnsupportedFeature", "A line or inline object exceeds the printable page height.");
        if (node.nodeType !== 3) return { node: node as HTMLElement, rect };
        const text = node as Text;
        let lo = 0, hi = text.length;
        // Locate the first character on this visual line without measuring every character.
        while (lo < hi) {
          const mid = Math.floor((lo + hi) / 2), probe = document.createRange();
          probe.setStart(text, mid); probe.setEnd(text, Math.min(text.length, mid + 1));
          const box = probe.getBoundingClientRect();
          if (box.top < rect.top - 0.5) lo = mid + 1; else hi = mid;
        }
        if (lo && /[\udc00-\udfff]/.test(text.data[lo] ?? "")) lo--;
        return { node: text, rect, offset: lo };
      }
    }
    return;
  }
  function paragraph(element: HTMLElement, inset: number): void {
    const style = view.getComputedStyle(element), box = element.getBoundingClientRect(), current = page(box.top);
    const capacity = available(current) - inset;
    if ((style.breakInside.startsWith("avoid") || style.pageBreakInside === "avoid") && box.height <= capacity && box.bottom > start(current) + available(current) + 0.5)
      before(element, current + 1, inset);
    for (let line = overflow(element, inset); line; line = overflow(element, inset)) {
      if (line.targetPage===undefined&&!(line.node.nodeType === 1 && (line.node as Element).matches("br[data-word-break]"))) {
        const lines = visualLines(element), ordinal = lines.findIndex((item) => Math.abs(item.rect.top - line!.rect.top) < 0.5);
        if (ordinal >= 0) {
          const fragmentPage = page(line.rect.top + line.rect.height / 2), fragmentStart = lines.findIndex((item) => page(item.rect.top + item.rect.height / 2) === fragmentPage);
          const orphans = Math.max(1, Number(style.orphans) || 2), widows = Math.max(1, Number(style.widows) || 2);
          let cut = ordinal;
          if (lines.length - cut < widows) cut = Math.max(fragmentStart, lines.length - widows);
          if (cut - fragmentStart < orphans && fragmentStart === 0 && element.getBoundingClientRect().top > start(fragmentPage) + inset + 0.5) {
            before(element, fragmentPage + 1, inset); continue;
          }
          if (cut > fragmentStart) line = lines[cut]!;
        }
      }
      const currentPage = page(line.rect.top + line.rect.height / 2);
      const hard = line.node.nodeType === 1 && (line.node as Element).matches("br[data-word-break]");
      const next = line.targetPage??(hard && (line.node as HTMLElement).dataset.wordBreak==='page' ? nextPhysical(currentPage) : line.rect.top < start(currentPage) + inset - 5 ? currentPage : currentPage + 1);
      ensure(next);
      if (hard) {
        const breaker = line.node as HTMLElement;
        // The empty block itself ends the preceding line. The authored br remains
        // in the source index and is restored by SourceProjection on every save.
        ownViewStyle(breaker, "display"); breaker.style.display = "none";
        spacer(breaker.parentNode!, breaker, start(next) + inset);
      } else {
        const node = line.node;
        const tail = node.nodeType === 3 && line.offset ? (node as Text).splitText(line.offset) : node;
        const gap = spacer(tail.parentNode!, tail, start(next) + inset);
        if (gap.getBoundingClientRect().height < 0.1) throw new DocumentError("UnsupportedFeature", "This inline structure cannot be fragmented safely.");
      }
    }
    const after = element.getBoundingClientRect(); ensure(page(Math.max(after.top, after.bottom - 0.5)));
  }
  function table(element: HTMLTableElement): void {
    const header = headerRows(element), height = headerHeight(element);
    if (height >= area) throw new DocumentError("UnsupportedFeature", "A repeating table header exceeds the printable page height.");
    const rows = [...element.rows].filter((row) => viewNodeRole(row) !== "remove"&&view.getComputedStyle(row).display!=="none"&&view.getComputedStyle(row).visibility!=="collapse");
    for (const row of rows) {
      let box = row.getBoundingClientRect(), current = page(box.top);
      if (!header.includes(row) && box.height <= available(current) - height && (box.bottom > start(current) + available(current) + 0.5 || box.top < start(current) + (current > topPage(element) ? height : 0) - 0.5)) {
        const next = box.top < start(current) + height - 0.5 ? current : current + 1;
        rowGap(element, row, next, height); box = row.getBoundingClientRect(); current = page(box.top);
      }
      if (box.height > available(current) - height) {
        for (const cell of row.cells) {
          if (view.getComputedStyle(cell).verticalAlign !== "top") throw new DocumentError("UnsupportedFeature", "Splitting a non-top-aligned table cell across pages is not supported.");
          const blocks = [...cell.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li")].filter((node) => !node.querySelector("p,h1,h2,h3,h4,h5,h6,li"));
          for (const block of blocks.length ? blocks : [cell]) paragraph(block, height);
        }
      }
      ensure(page(Math.max(box.top, row.getBoundingClientRect().bottom - 0.5)));
    }
  }
  function walk(element: HTMLElement): void {
    if (viewNodeRole(element) === "remove") return;
    const style = view.getComputedStyle(element), rect = element.getBoundingClientRect();
    if(style.display==="none"||style.visibility==="collapse")return;
    const force = ["page", "always", "left", "right"].includes(style.breakBefore) || style.pageBreakBefore === "always";
    if (force && (rect.top > start(page(rect.top)) + 0.5 || page(rect.top)%pageStride)) before(element, nextPhysical(page(rect.top)));
    if (style.breakAfter.startsWith("avoid") || style.pageBreakAfter === "avoid") {
      let height = element.getBoundingClientRect().height, next = element.nextElementSibling as HTMLElement | null;
      while (next && viewNodeRole(next) === "remove") next = next.nextElementSibling as HTMLElement | null;
      for (let count = 0; next && count < limits.maxNodes; count++) {
        const nextStyle = view.getComputedStyle(next), box = next.getBoundingClientRect();
        if (["page", "always"].includes(nextStyle.breakBefore)) break;
        const lines = visualLines(next), lineHeight = parseFloat(nextStyle.lineHeight) || lines[0]?.rect.height || box.height;
        const keep = nextStyle.breakAfter.startsWith("avoid") || nextStyle.pageBreakAfter === "avoid";
        height += keep ? box.height : Math.min(box.height, lineHeight * Math.max(1, Number(nextStyle.orphans) || 2));
        if (!keep) break;
        next = next.nextElementSibling as HTMLElement | null;
        while (next && viewNodeRole(next) === "remove") next = next.nextElementSibling as HTMLElement | null;
      }
      const box = element.getBoundingClientRect(), current = page(box.top);
      if (height <= available(current) && box.top > start(current) + 0.5 && box.top + height > start(current) + available(current) + 0.5) before(element, current + 1);
    }
    if (element.localName === "table") table(element as HTMLTableElement);
    else if (!element.querySelector("p,h1,h2,h3,h4,h5,h6,li,table,div,section") && (element.matches("p,h1,h2,h3,h4,h5,h6,li") || element.textContent?.trim() || element.querySelector("br,img"))) paragraph(element, 0);
    else if (element.localName === "img") {
      const box = element.getBoundingClientRect(); if (box.height > area) throw new DocumentError("UnsupportedFeature", "An object exceeds the printable page height.");
      if (box.bottom > start(page(box.top)) + area) before(element, page(box.top) + 1);
    } else {
      let inline: Node[] = [];
      const flush = (): void => {
        if (!inline.length) return;
        if (inline.some((node) => node.nodeType === 1 || node.textContent?.trim())) {
          const wrapper = markViewNode(document.createElement("span"), "unwrap");
          wrapper.className = "sdk-inline-flow"; wrapper.style.display = "block";
          element.insertBefore(wrapper, inline[0]!); wrapper.append(...inline); paragraph(wrapper, 0);
        }
        inline = [];
      };
      for (const child of [...element.childNodes]) {
        if (viewNodeRole(child) === "remove") continue;
        const node = child.nodeType === 1 ? child as HTMLElement : undefined;
        const display = node ? view.getComputedStyle(node).display : "inline";
        if (!node || display === "inline" || display === "inline-block") inline.push(child);
        else { flush(); walk(node); }
      }
      flush();
    }
    if (["page", "always"].includes(style.breakAfter) || style.pageBreakAfter === "always") {
      const box = element.getBoundingClientRect(), next = nextPhysical(page(box.bottom - 0.5)); ensure(next);
      spacer(element.parentNode!, element.nextSibling, start(next));
    }
  }
  return {
    reset() {
      for (const gap of gaps.splice(0)) gap.remove();
      for (const [cell, span] of canonicalSpans) cell.rowSpan = span;
      for (const breaker of flow.querySelectorAll<HTMLElement>('br[data-word-break]')) breaker.style.removeProperty("display");
      flow.style.minHeight = "0"; last = 0; operations = 0; blanks.clear();
    },
    layout(element, forcePage = false, parity) {
      let index = topPage(element);
      if (forcePage && (element.getBoundingClientRect().top > start(index) + 0.5 || index%pageStride)) { index=nextPhysical(index); before(element, index); }
      if (parity !== undefined && Math.floor(index/pageStride) % 2 !== parity) {for(let n=Math.floor(index/pageStride)*pageStride;n<nextPhysical(index);n++)blanks.add(n);index=nextPhysical(index);before(element,index);}
      if (element === flow) for (const child of [...flow.children]) walk(child as HTMLElement); else walk(element);
      flow.style.minHeight = last * pitch + area + "px";
    },
    get count() { return last + 1; }, get blanks() { return blanks; }, pageAt: topPage, headerHeight,
  };
}
