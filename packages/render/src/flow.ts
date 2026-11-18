import { wordRowVisibleRect, wordTableRepeatHeader } from "./word-table-layout.js";
import { continuationHeader } from "./repeated-header.js";
import { DocumentError } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import {
  isMeasuredViewContent,
  markViewNode,
  ownViewAttribute,
  ownViewStyle,
  viewNodeRole,
} from "./view-nodes.js";

export interface PageFlow {
  reset(): void;
  layout(element: HTMLElement, forcePage?: boolean, parity?: number): void;
  readonly count: number;
  readonly blanks: ReadonlySet<number>;
  pageAt(element: HTMLElement): number;
  headerHeight(table: HTMLTableElement): number;
  bottom(index: number): number;
  readonly endColumn?: number;
  project?(): void;
}
export interface PageFlowConstraints {
  firstPageInset?: number;
  minimumStart?: number;
  reserved(index: number): number;
  minimumPage(node: Node): number | undefined;
}
/** Portable single-DOM fragmentation. Empty, owned spacers carry page margins;
 * source paragraphs and cells keep their identity even when a text node splits. */
export function pageFlow(
  document: Document,
  flow: HTMLElement,
  area: number,
  pitch: number,
  limits: Readonly<RenderLimits>,
  pageStride = 1,
  constraints?: PageFlowConstraints,
): PageFlow {
  const view = document.defaultView!,
    gaps: HTMLElement[] = [],
    tableEdges: HTMLElement[] = [],
    tableGaps = new Map<HTMLTableElement, { row: HTMLTableRowElement; spanning: boolean }[]>(),
    tableFragmentBreaks = new Map<HTMLTableElement, Set<number>>(),
    alignedCells = new Map<
      HTMLTableCellElement,
      { content: HTMLElement; value: string; priority: string }
    >(),
    blanks = new Set<number>();
  const excluded = (node: Node): boolean => {
    for (let item: Node | null = node; item && item !== flow; item = item.parentNode)
      if (viewNodeRole(item) === "remove" && !isMeasuredViewContent(item)) return true;
    return false;
  };
  const canonicalSpans = new Map<HTMLTableCellElement, number>();
  let last = 0,
    operations = 0;
  const bottoms = new Map<number, number>();
  const origin = (): number => flow.getBoundingClientRect().top;
  const page = (y: number): number => Math.max(0, Math.floor((y - origin() + 0.25) / pitch));
  const start = (index: number): number => origin() + index * pitch;
  const available = (index: number): number =>
    Math.max(Math.min(24, area), area - (constraints?.reserved(index) ?? 0));
  const ensure = (index: number, bottom?: number): void => {
    if (bottom !== undefined)
      bottoms.set(index, Math.max(bottoms.get(index) ?? 0, bottom - start(index)));
    if (index >= limits.maxPages)
      throw new DocumentError("ResourceLimitExceeded", "Document exceeds the render page budget.");
    last = Math.max(last, index);
    if (++operations > limits.maxNodes + limits.maxPages * 32)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Pagination exceeds the measurement budget.",
      );
  };
  const topPage = (element: HTMLElement): number => page(element.getBoundingClientRect().top);
  const nextPhysical = (index: number): number => (Math.floor(index / pageStride) + 1) * pageStride;
  function spacer(parent: Node, before: Node | null, target: number): HTMLElement {
    const gap = markViewNode(document.createElement("span"));
    gap.className = "sdk-page-gap";
    gap.contentEditable = "false";
    gap.setAttribute("aria-hidden", "true");
    gap.style.cssText =
      "display:block;height:0;margin:0;padding:0;border:0;line-height:0;pointer-events:none;user-select:none";
    parent.insertBefore(gap, before);
    gaps.push(gap);
    gap.style.height = Math.max(0, target - gap.getBoundingClientRect().top) + "px";
    return gap;
  }
  function before(element: HTMLElement, targetPage: number, inset = 0): void {
    ensure(targetPage);
    spacer(element.parentNode!, element, start(targetPage) + inset);
  }
  function headerRows(table: HTMLTableElement): HTMLTableRowElement[] {
    const rows = [...table.rows].filter(
        (row) =>
          viewNodeRole(row) !== "remove" &&
          view.getComputedStyle(row).display !== "none" &&
          view.getComputedStyle(row).visibility !== "collapse",
      ),
      result: HTMLTableRowElement[] = [];
    for (const row of rows) {
      if (!wordTableRepeatHeader(row)) break;
      result.push(row);
    }
    return result;
  }
  const headerHeight = (table: HTMLTableElement): number =>
    headerRows(table).reduce(
      (sum, row) => sum + continuationHeader(row).getBoundingClientRect().height,
      0,
    );
  function rowGap(
    table: HTMLTableElement,
    rows: readonly HTMLTableRowElement[],
    ordinal: number,
    targetPage: number,
    inset: number,
  ): void {
    ensure(targetPage);
    const spanningCells: HTMLTableCellElement[] = [];
    for (let index = 0; index < ordinal; index++)
      for (const cell of rows[index]!.cells) {
        const span = canonicalSpans.get(cell) ?? cell.rowSpan;
        canonicalSpans.set(cell, span);
        if (!span || index + span > ordinal) spanningCells.push(cell);
      }
    const firstFragmentEnd = rows[ordinal]!.getBoundingClientRect().top;
    fixCellAlignment(spanningCells, firstFragmentEnd);
    for (const cell of spanningCells)
      if (cell.rowSpan) {
        ownViewAttribute(cell, "rowspan");
        cell.rowSpan++;
      }
    const gap = markViewNode(document.createElement("tr"));
    gap.className = "sdk-page-gap";
    gap.setAttribute("aria-hidden", "true");
    gap.contentEditable = "false";
    const cell = document.createElement("td");
    cell.colSpan = Math.max(
      1,
      ...rows.map((item) => [...item.cells].reduce((sum, value) => sum + value.colSpan, 0)),
    );
    cell.style.cssText = "padding:0;border:0;line-height:0;font-size:0";
    gap.append(cell);
    rows[ordinal]!.before(gap);
    gaps.push(gap);
    const spanning = spanningCells.length > 0;
    if (!spanning && view.getComputedStyle(table).borderCollapse === "collapse")
      cell.style.borderTop = view.getComputedStyle(table).borderBottom;
    if (!spanning || spanningCells.every((cell) =>
      visualLines(cell).every(({ rect }) => rect.bottom <= firstFragmentEnd + 0.5),
    )) {
      const tableRows = tableGaps.get(table) ?? [];
      tableRows.push({ row: gap, spanning });
      tableGaps.set(table, tableRows);
    }
    const height = Math.max(0, start(targetPage) + inset - gap.getBoundingClientRect().top);
    gap.style.height = height + "px";
    cell.style.height = height + "px";
  }
  function clipTableGaps(): void {
    for (const [table, rows] of tableGaps) {
      const box = table.getBoundingClientRect(),
        parts: string[] = [];
      // A single table stays in flow. Its spacer rows occupy space but should
      // never extend the table's perimeter or background into that space.
      let top = -1;
      const append = (bottom: number): void => {
        if (bottom > top)
          parts.push(`M -1 ${top} h ${box.width + 2} v ${bottom - top} h ${-box.width - 2} Z`);
      };
      const border = view.getComputedStyle(table).borderBottom;
      for (const { row, spanning } of rows) {
        const gap = row.getBoundingClientRect(),
          closingBorder = (parseFloat(row.cells[0]!.style.borderTopWidth) || 0) / 2;
        append(gap.top - box.top + closingBorder);
        top = gap.bottom - box.top - 0.5;
        if (spanning && parseFloat(border) > 0)
          paintTableEdge("bottom", box, gap.top - parseFloat(border), border);
      }
      append(box.height + 1);
      ownViewStyle(table, "clip-path");
      table.style.clipPath = `path("${parts.join(" ")}")`;
    }
  }
  function markTableBreak(table: HTMLTableElement, next: number): void {
    const breaks = tableFragmentBreaks.get(table) ?? new Set<number>();
    breaks.add(next);
    tableFragmentBreaks.set(table, breaks);
  }
  function finishRow(table: HTMLTableElement, top: number, bottom: number): void {
    const end = page(Math.max(top, bottom - 0.5));
    ensure(end, bottom);
    const first = page(top);
    if (first === end || pitch <= area || table.closest(".sdk-word-floating-table")) return;
    for (let next = first + 1; next <= end; next++) markTableBreak(table, next);
  }
  function paintTableEdge(
    side: "top" | "bottom", rect: DOMRect, y: number, border: string,
  ): void {
    if (document.getElementsByTagName("*").length >= limits.maxNodes)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Table fragment borders exceed the rendering node budget.",
      );
    const bounds = flow.getBoundingClientRect(),
      edge = markViewNode(document.createElement("div"));
    edge.className = "sdk-table-fragment-edge";
    edge.dataset.sdkTableEdge = side;
    edge.setAttribute("aria-hidden", "true");
    edge.contentEditable = "false";
    edge.style.cssText =
      `position:absolute;pointer-events:none;left:${rect.left - bounds.left}px;top:${y - bounds.top}px;width:${rect.width}px;height:0;border-top:${border}`;
    flow.append(edge);
    tableEdges.push(edge);
  }
  function paintSplitTableEdges(): void {
    for (const [table, breaks] of tableFragmentBreaks) {
      const rect = table.getBoundingClientRect(),
        style = view.getComputedStyle(table),
        topWidth = parseFloat(style.borderTopWidth) || 0,
        bottomWidth = parseFloat(style.borderBottomWidth) || 0,
        repeatedHeader = headerRows(table).length > 0;
      for (const next of breaks) {
        const previous = next - 1;
        // The shared page clip ends at the printable bottom. Draw the stroke
        // inward so its full width survives the clip and print projection.
        if (bottomWidth > 0 && !blanks.has(previous))
          paintTableEdge("bottom", rect, start(previous) + available(previous) - bottomWidth, style.borderBottom);
        if (topWidth > 0 && !repeatedHeader && !blanks.has(next) && rect.bottom > start(next))
          paintTableEdge("top", rect, start(next), style.borderTop);
      }
    }
  }
  interface Line {
    readonly node: Text | HTMLElement;
    readonly rect: DOMRect;
    readonly offset?: number;
    readonly targetPage?: number;
  }
  const scaledPiece = (owner: HTMLElement): HTMLElement | null => {
    const piece = owner.closest<HTMLElement>(".sdk-word-run-width");
    return piece && viewNodeRole(piece) === "unwrap" ? piece : null;
  };
  function visualLines(element: HTMLElement): Line[] {
    const result: Line[] = [],
      walker = document.createTreeWalker(element, 5);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const owner = node.nodeType === 1 ? (node as HTMLElement) : node.parentElement!;
      const scaled = scaledPiece(owner);
      if (scaled && scaled !== node) continue;
      if (
        owner.closest("[data-word-drawing]") &&
        (!owner.hasAttribute("data-word-drawing") || node.nodeType === 3)
      )
        continue;
      if (
        owner.hasAttribute("data-word-drawing") &&
        (() => {
          const css = view.getComputedStyle(
            owner.closest<HTMLElement>("[data-sdk-word-object]") ?? owner,
          );
          return css.position === "absolute" || css.cssFloat !== "none";
        })()
      )
        continue;
      if (excluded(owner) || view.getComputedStyle(owner).visibility === "hidden") continue;
      if (
        node.nodeType !== 3 && !scaled &&
        !owner.matches(
          "img,[data-word-tab],[data-word-drawing],[data-word-note-reference],[data-word-equation]",
        )
      )
        continue;
      const range = document.createRange();
      range.selectNodeContents(scaled?.firstChild?.firstChild ?? node);
      for (const rect of node.nodeType === 3 || scaled ? range.getClientRects() : owner.getClientRects()) {
        if (!wordRowVisibleRect(owner, rect)) continue;
        if (rect.width < 0.01 || rect.height < 0.1) continue;
        const previous = result.at(-1);
        if (
          previous &&
          Math.min(previous.rect.bottom, rect.bottom) - Math.max(previous.rect.top, rect.top) >
            Math.min(previous.rect.height, rect.height) / 2
        )
          continue;
        let lo = 0,
          hi = node.nodeType === 3 ? (node as Text).length : 0;
        while (lo < hi) {
          const mid = Math.floor((lo + hi) / 2),
            probe = document.createRange();
          probe.setStart(node, mid);
          probe.setEnd(node, Math.min((node as Text).length, mid + 1));
          if (probe.getBoundingClientRect().top < rect.top - 0.5) lo = mid + 1;
          else hi = mid;
        }
        if (node.nodeType === 3 && lo && /[\udc00-\udfff]/.test((node as Text).data[lo] ?? ""))
          lo--;
        result.push({
          node: node as Text | HTMLElement,
          rect,
          ...(node.nodeType === 3 ? { offset: lo } : {}),
        });
      }
    }
    return result;
  }
  function overflow(
    element: HTMLElement,
    inset: number,
    exactHeight: number,
    fragmentGaps: readonly HTMLElement[],
  ): Line | undefined {
    const fragmentStarts: number[] = [];
    if (exactHeight) {
      const css = view.getComputedStyle(element);
      fragmentStarts.push(
        element.getBoundingClientRect().top + parseFloat(css.paddingTop) + parseFloat(css.borderTopWidth),
        ...fragmentGaps.map((gap) => gap.getBoundingClientRect().bottom),
      );
      fragmentStarts.reverse();
    }
    const walker = document.createTreeWalker(element, 5 /* ELEMENT | TEXT */);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const owner = node.nodeType === 1 ? (node as HTMLElement) : node.parentElement!;
      const scaled = scaledPiece(owner);
      if (scaled && scaled !== node) continue;
      if (
        owner.closest("[data-word-drawing]") &&
        (!owner.hasAttribute("data-word-drawing") || node.nodeType === 3)
      )
        continue;
      if (
        owner.hasAttribute("data-word-drawing") &&
        view.getComputedStyle(owner.closest<HTMLElement>("[data-sdk-word-object]") ?? owner)
          .position === "absolute"
      )
        continue;
      if (excluded(owner) || view.getComputedStyle(owner).visibility === "hidden") continue;
      if (node.nodeType === 1) {
        if (owner.matches("br[data-word-break=page],br[data-word-break=column]")) {
          if (owner.style.display === "none") continue;
          return { node: owner, rect: owner.getBoundingClientRect() };
        }
        if (!scaled &&
          !owner.matches(
            "img,[data-word-tab],[data-word-drawing],[data-word-note-reference],[data-word-equation]",
          )
        )
          continue;
      }
      const range = document.createRange();
      range.selectNodeContents(scaled?.firstChild?.firstChild ?? node);
      const rectangles =
        node.nodeType === 3 || scaled
          ? [...range.getClientRects()]
          : [...(node as Element).getClientRects()].map((rect) => {
              // A floating frame rotates around its centre. Its paint bounds are
              // not a text line above the paragraph: paginate its layout box.
              if (
                !owner.hasAttribute("data-word-drawing") ||
                view.getComputedStyle(owner).cssFloat === "none"
              )
                return rect;
              const width = owner.offsetWidth,
                height = owner.offsetHeight;
              return new DOMRect(
                rect.left + (rect.width - width) / 2,
                rect.top + (rect.height - height) / 2,
                width,
                height,
              );
            });
      for (const rect of rectangles) {
        if (!wordRowVisibleRect(owner, rect)) continue;
        if (rect.height < 0.1 || rect.width < 0.01) continue;
        const middle = rect.top + rect.height / 2,
          origin = fragmentStarts.find((top) => top <= middle) ?? fragmentStarts.at(-1),
          // For exact spacing, the page needs room for the entire line box.
          // Glyph ink can finish before the bottom of the paragraph strut.
          lineTop = origin === undefined
            ? rect.top
            : origin + Math.max(0, Math.floor((middle - origin) / exactHeight)) * exactHeight,
          current = page(origin === undefined ? middle : lineTop + exactHeight / 2),
          end = start(current) + available(current),
          minimum = constraints?.minimumPage(node);
        if (minimum !== undefined && current < minimum)
          return { node: node as HTMLElement, rect, targetPage: minimum };
        // Glyph rectangles can extend beyond a line box by sub-pixel rounding.
        const height = exactHeight || Math.max(
          rect.height,
          parseFloat(view.getComputedStyle(owner).lineHeight) || rect.height,
        ),
          lineBottom = origin === undefined
            ? rect.bottom + Math.max(0, height - rect.height) / 2
            : lineTop + exactHeight;
        if (
          lineTop >= start(current) + inset - 5 &&
          lineBottom <= end + 0.5
        ) {
          ensure(current, lineBottom);
          continue;
        }
        if (height > area - inset + 0.5)
          throw new DocumentError(
            "UnsupportedFeature",
            "A line or inline object exceeds the printable page height.",
          );
        if (node.nodeType !== 3) return { node: node as HTMLElement, rect };
        const text = node as Text;
        let lo = 0,
          hi = text.length;
        // Locate the first character on this visual line without measuring every character.
        while (lo < hi) {
          const mid = Math.floor((lo + hi) / 2),
            probe = document.createRange();
          probe.setStart(text, mid);
          probe.setEnd(text, Math.min(text.length, mid + 1));
          const box = probe.getBoundingClientRect();
          if (box.top < rect.top - 0.5) lo = mid + 1;
          else hi = mid;
        }
        if (lo && /[\udc00-\udfff]/.test(text.data[lo] ?? "")) lo--;
        return { node: text, rect, offset: lo };
      }
    }
    return;
  }
  function paragraph(element: HTMLElement, inset: number, splitRow = false): void {
    const style = view.getComputedStyle(element),
      box = element.getBoundingClientRect(),
      current = page(box.top),
      exactHeight =
        style.writingMode === "horizontal-tb" && !element.closest("td,th") &&
        (element.dataset.wordLineRule ?? style.getPropertyValue("--word-line-rule").trim()) === "exact" &&
        !element.querySelector("img,[data-word-drawing],[data-word-native-object],[data-word-equation]")
          ? parseFloat(style.lineHeight) || 0
          : 0,
      fragmentGaps: HTMLElement[] = [];
    if (
      !splitRow &&
      (style.breakInside.startsWith("avoid") || style.pageBreakInside === "avoid") &&
      box.height <= area - inset &&
      box.bottom > start(current) + available(current) + 0.5
    ) {
      let target = current + 1;
      while (target < limits.maxPages && box.height > available(target) - inset) target++;
      before(element, target, inset);
    }
    for (
      let line = overflow(element, inset, exactHeight, fragmentGaps);
      line;
      line = overflow(element, inset, exactHeight, fragmentGaps)
    ) {
      if (
        !splitRow &&
        line.targetPage === undefined &&
        !(line.node.nodeType === 1 && (line.node as Element).matches("br[data-word-break]"))
      ) {
        const lines = visualLines(element),
          ordinal = lines.findIndex((item) => Math.abs(item.rect.top - line!.rect.top) < 0.5);
        if (ordinal >= 0) {
          const fragmentPage = page(line.rect.top + line.rect.height / 2),
            fragmentStart = lines.findIndex(
              (item) => page(item.rect.top + item.rect.height / 2) === fragmentPage,
            );
          const orphans = Math.max(1, Number(style.orphans) || 2),
            widows = Math.max(1, Number(style.widows) || 2);
          let cut = ordinal;
          if (lines.length - cut < widows) cut = Math.max(fragmentStart, lines.length - widows);
          if (
            cut - fragmentStart < orphans &&
            fragmentStart === 0 &&
            element.getBoundingClientRect().top > start(fragmentPage) + inset + 0.5
          ) {
            before(element, fragmentPage + 1, inset);
            continue;
          }
          if (cut > fragmentStart) line = lines[cut]!;
        }
      }
      const currentPage = page(line.rect.top + line.rect.height / 2);
      const hard =
        line.node.nodeType === 1 && (line.node as Element).matches("br[data-word-break]");
      const next =
        line.targetPage ??
        (hard && (line.node as HTMLElement).dataset.wordBreak === "page"
          ? nextPhysical(currentPage)
          : line.rect.top < start(currentPage) + inset - 5
            ? currentPage
            : currentPage + 1);
      ensure(next);
      if (hard) {
        const breaker = line.node as HTMLElement;
        // The empty block itself ends the preceding line. The authored br remains
        // in the source index and is restored by SourceProjection on every save.
        ownViewStyle(breaker, "display");
        breaker.style.display = "none";
        fragmentGaps.push(spacer(breaker.parentNode!, breaker, start(next) + inset));
      } else {
        const node = line.node;
        const tail =
          node.nodeType === 3 && line.offset ? (node as Text).splitText(line.offset) : node;
        const target = start(next) + inset,
          gap = spacer(tail.parentNode!, tail, target),
          gapHeight = gap.getBoundingClientRect().height;
        fragmentGaps.push(gap);
        if (gapHeight < 0.1)
          throw new DocumentError(
            "UnsupportedFeature",
            "This inline structure cannot be fragmented safely.",
          );
        if (inset) {
          // Negative font leading can lift the first glyph above the block
          // spacer and into the repeated header. Reserve the measured overhang.
          const range = document.createRange();
          range.selectNodeContents(tail);
          const first = (tail.nodeType === 3 ? range : (tail as HTMLElement)).getClientRects()[0];
          if (first && first.top < target - 0.25)
            gap.style.height = gapHeight + target - first.top + "px";
        }
      }
    }
    const after = element.getBoundingClientRect();
    ensure(
      page(Math.max(after.top, after.bottom - 0.5)),
      after.bottom + Math.max(0, parseFloat(style.marginBottom) || 0),
    );
  }
  function cellBlocks(cell: HTMLTableCellElement): HTMLElement[] {
    const blocks = [...cell.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li")].filter(
      (node) => !node.querySelector("p,h1,h2,h3,h4,h5,h6,li"),
    );
    return blocks.length ? blocks : [cell];
  }
  function canStartRow(blocks: readonly HTMLElement[][], end: number): boolean {
    return blocks.some(([block]) => {
      if (!block) return false;
      const style = view.getComputedStyle(block),
        box = block.getBoundingClientRect();
      if (box.height <= 0.1 || box.top >= end - 0.5) return false;
      if (box.bottom <= end + 0.5) return true;
      const line = visualLines(block)[0];
      if (!line) return false;
      const descent = Math.max(0, (parseFloat(style.lineHeight) || 0) - line.rect.height) / 2;
      return line.rect.bottom + descent <= end + 0.5;
    });
  }
  function fixCellAlignment(cells: HTMLTableCellElement[], fragmentEnd?: number): void {
    const aligned = cells
      .filter((cell) => !alignedCells.has(cell))
      .map((cell) => {
        const alignment = view.getComputedStyle(cell).verticalAlign,
          inset = fragmentEnd === undefined || alignment === "top"
            ? 0
            : Math.max(0, cell.getBoundingClientRect().bottom - fragmentEnd) *
              (alignment === "middle" ? 0.5 : 1);
        return { cell, alignment, inset };
      })
      .filter(({ alignment }) => alignment === "middle" || alignment === "bottom")
      .map(({ cell, inset }) => {
        const content = markViewNode(document.createElement("div"), "unwrap");
        content.style.display = "flow-root";
        content.append(...cell.childNodes);
        cell.append(content);
        alignedCells.set(cell, {
          content,
          value: cell.style.getPropertyValue("vertical-align"),
          priority: cell.style.getPropertyPriority("vertical-align"),
        });
        return { cell, content, inset };
      });
    // Freeze the unbroken alignment before synthetic page spacers enlarge the
    // row. Otherwise browsers recenter short cells into the page gutter.
    const tops = aligned.map(({ content }) => content.getBoundingClientRect().top);
    for (const { cell } of aligned) {
      ownViewStyle(cell, "vertical-align");
      cell.style.setProperty("vertical-align", "top", "important");
    }
    const offsets = aligned.map(
      ({ content, inset }, index) => tops[index]! - content.getBoundingClientRect().top - inset,
    );
    aligned.forEach(({ content }, index) => {
      if (offsets[index]! > 0.1) content.style.paddingTop = offsets[index]! + "px";
    });
  }
  function table(element: HTMLTableElement): void {
    const header = headerRows(element),
      height = headerHeight(element),
      tableStyle = view.getComputedStyle(element),
      collapsedStroke = tableStyle.borderCollapse === "collapse" &&
        (parseFloat(tableStyle.borderBottomWidth) > 0 ||
          parseFloat(tableStyle.getPropertyValue("--word-border-inside-h")) > 0);
    if (height >= area)
      throw new DocumentError(
        "UnsupportedFeature",
        "A repeating table header exceeds the printable page height.",
      );
    const rows = [...element.rows].filter(
      (row) =>
        viewNodeRole(row) !== "remove" &&
        view.getComputedStyle(row).display !== "none" &&
        view.getComputedStyle(row).visibility !== "collapse",
    );
    let mergeEnd = 0;
    for (const [ordinal, row] of rows.entries()) {
      let box = row.getBoundingClientRect(),
        current = page(box.top);
      const inMergedRow = mergeEnd > ordinal;
      for (const cell of row.cells)
        mergeEnd = Math.max(mergeEnd, ordinal + (cell.rowSpan || rows.length - ordinal));
      const repeated = header.includes(row),
        inset = current > topPage(element) ? height : 0,
        end = start(current) + available(current),
        style = view.getComputedStyle(row),
        canSplit = !repeated && !inMergedRow && mergeEnd <= ordinal + 1 &&
          !style.breakInside.startsWith("avoid") &&
          style.pageBreakInside !== "avoid" &&
          style.getPropertyValue("--word-row-height-rule").trim() !== "exact" &&
          !row.querySelector("table") &&
          !element.closest(".sdk-word-floating-table"),
        pastEdge = box.bottom > end + 0.5,
        // A browser can round a collapsed border outward by up to one CSS
        // pixel. Keep the row only if its contents are inside the page.
        strokeClipped = collapsedStroke && pastEdge && box.bottom <= end + 1.5 &&
            [...row.cells].every((cell) =>
              visualLines(cell).every(({ rect }) => rect.bottom <= end + 0.5),
          ),
        overflows = pastEdge && !strokeClipped;
      const blocks = box.top >= start(current) + inset - 0.5 && overflows && canSplit
        ? [...row.cells].map(cellBlocks)
        : undefined;
      if (blocks && canStartRow(blocks, end)) {
        fixCellAlignment([...row.cells]);
        // Word splits breakable rows at the available cell line, even when a
        // cell paragraph requests widow/orphan control or keep-lines.
        for (const cell of blocks)
          for (const block of cell) paragraph(block, height, true);
        finishRow(element, box.top, row.getBoundingClientRect().bottom);
        continue;
      }
      if (
        !repeated &&
        box.height <= available(current) - height &&
        (overflows || box.top < start(current) + inset - 0.5)
      ) {
        const next = box.top < start(current) + height - 0.5 ? current : current + 1;
        if (ordinal === 0) before(element, next);
        else rowGap(element, rows, ordinal, next, height);
        box = row.getBoundingClientRect();
        current = page(box.top);
      }
      if (box.height <= area - height && box.height > available(current) - height) {
        let target = current + 1;
        while (target < limits.maxPages && box.height > available(target) - height) target++;
        if (ordinal === 0) before(element, target);
        else rowGap(element, rows, ordinal, target, height);
        box = row.getBoundingClientRect();
        current = page(box.top);
      }
      if (box.height > available(current) - height) {
        for (const cell of row.cells) {
          if (view.getComputedStyle(cell).verticalAlign !== "top")
            throw new DocumentError(
              "UnsupportedFeature",
              "Splitting a non-top-aligned table cell across pages is not supported.",
            );
          for (const block of cellBlocks(cell)) paragraph(block, height, canSplit);
        }
      }
      const bottom = row.getBoundingClientRect().bottom;
      if (strokeClipped && bottom > end + 0.5 && bottom <= end + 1.5)
        markTableBreak(element, current + 1);
      finishRow(element, box.top, bottom);
    }
  }
  function walk(element: HTMLElement): void {
    if (viewNodeRole(element) === "remove" && !isMeasuredViewContent(element)) return;
    if (
      element.classList.contains("sdk-word-floating-table") &&
      element.style.position === "absolute"
    )
      return;
    const style = view.getComputedStyle(element),
      rect = element.getBoundingClientRect();
    if (style.display === "none" || style.visibility === "collapse") return;
    const force =
      ["page", "always", "left", "right"].includes(style.breakBefore) ||
      style.pageBreakBefore === "always";
    if (
      force &&
      (rect.top > start(page(rect.top)) + 0.5 ||
        page(rect.top) % pageStride ||
        (page(rect.top) < pageStride && (constraints?.firstPageInset ?? 0) > 0))
    )
      before(element, nextPhysical(page(rect.top)));
    if (style.breakAfter.startsWith("avoid") || style.pageBreakAfter === "avoid") {
      let height = element.getBoundingClientRect().height,
        next = element.nextElementSibling as HTMLElement | null;
      while (next && viewNodeRole(next) === "remove")
        next = next.nextElementSibling as HTMLElement | null;
      for (let count = 0; next && count < limits.maxNodes; count++) {
        const nextStyle = view.getComputedStyle(next),
          box = next.getBoundingClientRect();
        if (["page", "always"].includes(nextStyle.breakBefore)) break;
        const lines = visualLines(next),
          lineHeight = parseFloat(nextStyle.lineHeight) || lines[0]?.rect.height || box.height;
        const keep =
          nextStyle.breakAfter.startsWith("avoid") || nextStyle.pageBreakAfter === "avoid";
        height += keep
          ? box.height
          : Math.min(box.height, lineHeight * Math.max(1, Number(nextStyle.orphans) || 2));
        if (!keep) break;
        next = next.nextElementSibling as HTMLElement | null;
        while (next && viewNodeRole(next) === "remove")
          next = next.nextElementSibling as HTMLElement | null;
      }
      const box = element.getBoundingClientRect(),
        current = page(box.top);
      if (
        height <= available(current) &&
        box.top > start(current) + 0.5 &&
        box.top + height > start(current) + available(current) + 0.5
      )
        before(element, current + 1);
    }
    if (element.localName === "table") table(element as HTMLTableElement);
    else if (
      !element.querySelector("p,h1,h2,h3,h4,h5,h6,li,table,div,section") &&
      (element.matches("p,h1,h2,h3,h4,h5,h6,li") ||
        element.textContent?.trim() ||
        element.querySelector("br,img"))
    )
      paragraph(element, 0);
    else if (element.localName === "img") {
      const box = element.getBoundingClientRect();
      if (box.height > area)
        throw new DocumentError(
          "UnsupportedFeature",
          "An object exceeds the printable page height.",
        );
      if (box.bottom > start(page(box.top)) + area) before(element, page(box.top) + 1);
      const after = element.getBoundingClientRect();
      ensure(page(Math.max(after.top, after.bottom - 0.5)), after.bottom);
    } else {
      let inline: Node[] = [];
      const flush = (): void => {
        if (!inline.length) return;
        if (inline.some((node) => node.nodeType === 1 || node.textContent?.trim())) {
          const wrapper = markViewNode(document.createElement("span"), "unwrap");
          wrapper.className = "sdk-inline-flow";
          wrapper.style.display = "block";
          element.insertBefore(wrapper, inline[0]!);
          wrapper.append(...inline);
          paragraph(wrapper, 0);
        }
        inline = [];
      };
      for (const child of [...element.childNodes]) {
        if (viewNodeRole(child) === "remove" && !isMeasuredViewContent(child)) continue;
        const node = child.nodeType === 1 ? (child as HTMLElement) : undefined;
        const display = node ? view.getComputedStyle(node).display : "inline";
        if (!node || display === "inline" || display === "inline-block") inline.push(child);
        else {
          flush();
          walk(node);
        }
      }
      flush();
    }
    if (["page", "always"].includes(style.breakAfter) || style.pageBreakAfter === "always") {
      const box = element.getBoundingClientRect(),
        next = nextPhysical(page(box.bottom - 0.5));
      ensure(next);
      spacer(element.parentNode!, element.nextSibling, start(next));
    }
  }
  return {
    reset() {
      for (const gap of gaps.splice(0)) gap.remove();
      for (const edge of tableEdges.splice(0)) edge.remove();
      tableFragmentBreaks.clear();
      for (const [cell, span] of canonicalSpans) cell.rowSpan = span;
      canonicalSpans.clear();
      for (const [cell, { content, value, priority }] of alignedCells) {
        content.replaceWith(...content.childNodes);
        cell.style.setProperty("vertical-align", value, priority);
      }
      alignedCells.clear();
      for (const table of tableGaps.keys()) table.style.removeProperty("clip-path");
      tableGaps.clear();
      for (const breaker of flow.querySelectorAll<HTMLElement>("br[data-word-break]"))
        breaker.style.removeProperty("display");
      flow.style.minHeight = "0";
      last = 0;
      operations = 0;
      blanks.clear();
      bottoms.clear();
    },
    layout(element, forcePage = false, parity) {
      let index = topPage(element);
      if (index < (constraints?.minimumStart ?? 0)) {
        index = constraints!.minimumStart!;
        before(element, index);
      }
      if (
        forcePage &&
        (element.getBoundingClientRect().top > start(index) + 0.5 || index % pageStride)
      ) {
        index = nextPhysical(index);
        before(element, index);
      }
      if (parity !== undefined && Math.floor(index / pageStride) % 2 !== parity) {
        for (let n = Math.floor(index / pageStride) * pageStride; n < nextPhysical(index); n++)
          blanks.add(n);
        index = nextPhysical(index);
        before(element, index);
      }
      if (element === flow) for (const child of [...flow.children]) walk(child as HTMLElement);
      else walk(element);
      flow.style.minHeight = last * pitch + area + "px";
    },
    get count() {
      return last + 1;
    },
    get blanks() {
      return blanks;
    },
    pageAt: topPage,
    headerHeight,
    bottom: (index) => bottoms.get(index) ?? 0,
    project() {
      if (pitch > area) {
        clipTableGaps();
        paintSplitTableEdges();
      }
    },
  };
}
