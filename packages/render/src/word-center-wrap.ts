import {
  DEFAULT_DOCUMENT_READ_LIMITS,
  DocumentError,
  SourceStyleResolver,
  isDocumentTextBoundary,
  resolveSourceCssLength,
} from "@doffi/core";
import type { ContentDocumentModel, DeepReadonly, SourceStyleNode, WordDrawing } from "@doffi/core";
import type { RenderLimits, RenderedPage } from "./contracts.js";
import {
  hideFromViewSelection,
  isMeasuredViewContent,
  markViewNode,
  measureViewContent,
  ownViewStyle,
  viewNodeRole,
} from "./view-nodes.js";

export function needsWordDrawingLineExclusion(drawing: DeepReadonly<WordDrawing>): boolean {
  const layout = drawing.layout;
  return (
    layout.mode === "floating" &&
    (layout.alignH === "center" || (!layout.alignH && drawing.geometry.x > 0)) &&
    !layout.behindText &&
    (layout.wrap ?? "square") === "square" &&
    (layout.wrapSide ?? "bothSides") === "bothSides"
  );
}
export function needsWordTableLineExclusion(css: Readonly<Record<string, string>>): boolean {
  const align = css["--word-table-position-x-align"];
  return (
    css["--word-table-floating"] === "1" &&
    (align === "center" ||
      ((!align || align === "none") &&
        resolveSourceCssLength(css["--word-table-position-x"] ?? "0pt", {
          font: 12,
          percent: 468,
        }) > 0.05))
  );
}
interface Obstacle {
  key: string;
  left: number;
  right: number;
  top: number;
  bottom: number;
}
/** Inline exclusions leave one logical paragraph and the original inline owners
 * intact. They split only UTF-16-safe Text boundaries and carry no source text.
 * The browser still measures fonts, bidi runs, line breaks and native input. */
export interface WordCenterWrap {
  readonly active: boolean;
  readonly changed: boolean;
  reset(): void;
  measure(pages: readonly RenderedPage[], root: HTMLElement): void;
}
export function wordCenterWrap(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): WordCenterWrap {
  const snapshot = model.snapshot(),
    roots = new Map(snapshot.metadata.word?.stories.map((story) => [story.rootNodeId, story.kind]));
  const inMain = (id: string): boolean => {
    for (
      let node = model.getNode(id);
      node;
      node = node.parentNodeId ? model.getNode(node.parentNodeId) : undefined
    ) {
      const kind = roots.get(node.nodeId);
      if (kind) return kind === "main";
    }
    return false;
  };
  const drawings = new Map(
    snapshot.metadata.word?.drawings
      ?.filter((item) => needsWordDrawingLineExclusion(item) && inMain(item.nodeId))
      .map((item) => [item.nodeId, item]),
  );
  const resolver = new SourceStyleResolver(
    snapshot.styles.map((sheet) => sheet.css),
    DEFAULT_DOCUMENT_READ_LIMITS,
  );
  const project = (id: string): SourceStyleNode => {
    const node = model.getNode(id)!;
    return {
      tag: node.tagName,
      attrs: node.attributes,
      parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
    };
  };
  const tables = new Map(
    snapshot.metadata.word?.tables.flatMap((info) => {
      const css = resolver.get(project(info.nodeId));
      return needsWordTableLineExclusion(css) && inMain(info.nodeId)
        ? [[info.nodeId, css] as const]
        : [];
    }),
  );
  const active = drawings.size > 0 || tables.size > 0,
    view = document.defaultView!;
  const spacers = new Map<
      HTMLElement,
      { key: string; height: number; baseline: number; rtl: boolean }
    >(),
    paragraphStyles = new Map<
      HTMLElement,
      { value: string; priority: string; align: string; alignPriority: string; right: boolean }
    >(),
    aligned = new Set<HTMLElement>(),
    alignmentFillers = new Set<HTMLElement>(),
    longWords = new Set<HTMLElement>(),
    softBreaks = new Set<HTMLElement>();
  const canvas = document.createElement("canvas"),
    context = canvas.getContext("2d")!;
  let changed = false;
  const reset = () => {
    const parents = new Set<Node>();
    for (const spacer of spacers.keys()) {
      if (spacer.parentNode) parents.add(spacer.parentNode);
      spacer.remove();
    }
    spacers.clear();
    for (const filler of alignmentFillers) {
      if (filler.parentNode) parents.add(filler.parentNode);
      filler.remove();
    }
    alignmentFillers.clear();
    for (const br of softBreaks) br.remove();
    softBreaks.clear();
    for (const wrapper of longWords)
      if (wrapper.parentNode) {
        parents.add(wrapper.parentNode);
        wrapper.replaceWith(...wrapper.childNodes);
      }
    longWords.clear();
    aligned.clear();
    for (const parent of parents) parent.normalize();
    for (const [p, old] of paragraphStyles) {
      if (old.value) p.style.setProperty("overflow-wrap", old.value, old.priority);
      else p.style.removeProperty("overflow-wrap");
      if (old.right) {
        if (old.align) p.style.setProperty("text-align", old.align, old.alignPriority);
        else p.style.removeProperty("text-align");
      }
    }
    paragraphStyles.clear();
    changed = false;
  };
  return {
    active,
    get changed(): boolean {
      return changed;
    },
    reset,
    measure(pages: readonly RenderedPage[], root: HTMLElement): void {
      changed = false;
      if (!active) return;
      const obstacles: Obstacle[] = [],
        origin = root.getBoundingClientRect();
      const add = (
        node: HTMLElement,
        id: string,
        distances: { left: number; right: number; top: number; bottom: number },
      ) => {
        if (!node.getClientRects().length || view.getComputedStyle(node).visibility === "hidden")
          return;
        const rect = node.getBoundingClientRect();
        const page = pages.find(
          (page) =>
            !page.blank &&
            rect.top >= origin.top + page.rect.y / 0.75 - 0.5 &&
            rect.top < origin.top + (page.rect.y + page.rect.height) / 0.75,
        );
        if (!page) return;
        obstacles.push({
          key: id + ":" + page.pageId,
          left: rect.left - distances.left / 0.75,
          right: rect.right + distances.right / 0.75,
          top: rect.top - distances.top / 0.75,
          bottom: rect.bottom + distances.bottom / 0.75,
        });
      };
      for (const frame of document.querySelectorAll<HTMLElement>(
        ".sdk-word-floating-layer [data-sdk-word-object]",
      )) {
        const drawing = drawings.get(frame.dataset.sdkWordObject!);
        if (drawing)
          add(
            frame,
            drawing.nodeId,
            drawing.layout.distances ?? { left: 0, right: 0, top: 0, bottom: 0 },
          );
      }
      for (const table of document.querySelectorAll<HTMLTableElement>(
        ".sdk-word-floating-table > table[data-node-id]",
      )) {
        const css = tables.get(table.dataset.nodeId!);
        if (!css) continue;
        const distance = (side: string) =>
          resolveSourceCssLength(css["--word-table-position-" + side + "-distance"] ?? "0pt", {
            font: 12,
            percent: 468,
          });
        add(table, table.dataset.nodeId!, {
          left: distance("left"),
          right: distance("right"),
          top: distance("top"),
          bottom: distance("bottom"),
        });
      }
      const byKey = new Map(obstacles.map((item) => [item.key, item]));
      let work = 0,
        nodes = document.querySelectorAll("*").length;
      const reserve = () => {
        if (++work > limits.maxNodes * 32)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Centered wrapping exceeds its measurement budget.",
          );
      };
      for (const [spacer, entry] of spacers) {
        const box = spacer.getBoundingClientRect(),
          obstacle = byKey.get(entry.key);
        if (
          !spacer.isConnected ||
          !obstacle ||
          box.top - entry.baseline + entry.height <= obstacle.top + 0.25 ||
          box.top - entry.baseline >= obstacle.bottom - 0.25
        ) {
          spacer.remove();
          spacers.delete(spacer);
          changed = true;
          continue;
        }
        if (spacer.style.display === "block") continue;
        const width = entry.rtl ? box.right - obstacle.left : obstacle.right - box.left;
        if (width <= 0) {
          spacer.remove();
          spacers.delete(spacer);
          changed = true;
          continue;
        }
      }
      const charRect = (text: Text, from: number, to: number): DOMRect => {
        reserve();
        const range = document.createRange();
        range.setStart(text, from);
        range.setEnd(text, to);
        return range.getBoundingClientRect();
      };
      const overlaps = (
        a: { left: number; right: number; top: number; bottom: number },
        b: Obstacle,
      ) =>
        a.right > b.left + 0.25 &&
        a.left < b.right - 0.25 &&
        a.bottom > b.top + 0.25 &&
        a.top < b.bottom - 0.25;
      const find = ():
        | {
            text: Text;
            at: number;
            x: number;
            line: DOMRect;
            p: HTMLElement;
            obstacle: Obstacle;
            rtl: boolean;
          }
        | undefined => {
        for (const p of root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")) {
          reserve();
          if (
            p.closest(".sdk-word-floating-table,[data-sdk-word-object],.sdk-paper") ||
            p.closest("[data-office-story]")?.getAttribute("data-office-story") !== "main"
          )
            continue;
          const bounds = p.getBoundingClientRect(),
            relevant = obstacles.filter((o) => overlaps(bounds, o));
          if (!relevant.length) continue;
          const css = view.getComputedStyle(p),
            rtl = css.direction === "rtl";
          if (!paragraphStyles.has(p)) {
            paragraphStyles.set(p, {
              value: p.style.getPropertyValue("overflow-wrap"),
              priority: p.style.getPropertyPriority("overflow-wrap"),
              align: p.style.getPropertyValue("text-align"),
              alignPriority: p.style.getPropertyPriority("text-align"),
              right: css.direction !== "rtl" && css.textAlign === "right",
            });
            ownViewStyle(p, "overflow-wrap");
            p.style.overflowWrap = "anywhere";
            if (css.direction !== "rtl" && css.textAlign === "right") {
              ownViewStyle(p, "text-align");
              p.style.textAlign = "left";
              changed = true;
            }
          }
          const walker = document.createTreeWalker(p, 4);
          for (let current = walker.nextNode(); current; current = walker.nextNode()) {
            reserve();
            const text = current as Text;
            if (
              !text.length ||
              text.parentElement?.closest("[data-sdk-word-object],[data-word-native-object]")
            )
              continue;
            let excluded = false;
            for (
              let node: Node | null = text.parentNode;
              node && node !== p;
              node = node.parentNode
            )
              if (viewNodeRole(node) === "remove" && !isMeasuredViewContent(node)) {
                excluded = true;
                break;
              }
            if (excluded || view.getComputedStyle(text.parentElement!).visibility !== "visible")
              continue;
            const range = document.createRange();
            range.selectNodeContents(text);
            for (const line of range.getClientRects())
              for (const obstacle of relevant) {
                if (!overlaps(line, obstacle)) continue;
                let low = 0,
                  high = text.length;
                while (low < high) {
                  const mid = Math.floor((low + high) / 2),
                    rect = charRect(text, mid, Math.min(text.length, mid + 1));
                  if (rect.bottom <= line.top + 0.5) low = mid + 1;
                  else high = mid;
                }
                while (low > 0 && !isDocumentTextBoundary(text.data, low)) low--;
                const first = low;
                for (let at = first; at < text.length; ) {
                  let end = at + 1;
                  while (end < text.length && !isDocumentTextBoundary(text.data, end)) end++;
                  const glyph = charRect(text, at, end);
                  if (glyph.top >= line.bottom - 0.5) break;
                  if (overlaps(glyph, obstacle)) {
                    // Keep an ordinary Latin word together when it fits a side.
                    if (/[A-Za-z0-9]/.test(text.data[at] ?? "")) {
                      let start = at,
                        end = at + 1;
                      while (start > first && /[A-Za-z0-9]/.test(text.data[start - 1] ?? "")) {
                        reserve();
                        start--;
                      }
                      while (end < text.length && /[A-Za-z0-9]/.test(text.data[end] ?? "")) {
                        reserve();
                        end++;
                      }
                      const font = view.getComputedStyle(text.parentElement!);
                      context.font = font.font;
                      const wordWidth =
                        context.measureText(text.data.slice(start, end)).width +
                        (parseFloat(font.letterSpacing) || 0) * (end - start);
                      const available = Math.max(
                        obstacle.left - bounds.left,
                        bounds.right - obstacle.right,
                      );
                      if (wordWidth <= available) at = start;
                      else if (!text.parentElement!.closest(".sdk-word-unbroken-run")) {
                        if (++nodes > limits.maxNodes)
                          throw new DocumentError(
                            "ResourceLimitExceeded",
                            "Long-word wrapping exceeds its render node budget.",
                          );
                        const word = start ? text.splitText(start) : text;
                        if (end - start < word.length) word.splitText(end - start);
                        const wrapper = markViewNode(document.createElement("span"), "unwrap");
                        wrapper.className = "sdk-word-unbroken-run";
                        wrapper.style.cssText = "word-break:break-all;overflow-wrap:anywhere";
                        word.before(wrapper);
                        wrapper.append(word);
                        const widths = [...new Set(word.data)].map(
                          (character) =>
                            context.measureText(character).width +
                            (parseFloat(font.letterSpacing) || 0),
                        );
                        const sides = [
                          obstacle.left - bounds.left,
                          bounds.right - obstacle.right,
                        ].filter((width) => width > 1);
                        const chunk = Math.max(
                          1,
                          Math.min(
                            16,
                            Math.floor(Math.min(...sides) / Math.max(0.1, ...widths)) - 1,
                          ),
                        );
                        if (nodes + Math.ceil(word.length / chunk) > limits.maxNodes)
                          throw new DocumentError(
                            "ResourceLimitExceeded",
                            "Long-word soft breaks exceed the render node budget.",
                          );
                        const fragment = document.createDocumentFragment();
                        for (let index = 0; index < word.length; index += chunk) {
                          fragment.append(
                            document.createTextNode(word.data.slice(index, index + chunk)),
                          );
                          if (index + chunk < word.length) {
                            const br = markViewNode(document.createElement("wbr"));
                            softBreaks.add(br);
                            fragment.append(br);
                            nodes++;
                          }
                        }
                        word.replaceWith(fragment);
                        longWords.add(wrapper);
                        changed = true;
                        return find();
                      }
                    }
                    while (at > 0 && !isDocumentTextBoundary(text.data, at)) at--;
                    const boundary =
                      at === first
                        ? charRect(text, at, Math.min(text.length, at + 1))
                        : charRect(text, at - 1, at);
                    const x = rtl
                      ? at === first
                        ? boundary.right
                        : boundary.left
                      : at === first
                        ? boundary.left
                        : boundary.right;
                    return { text, at, x, line, p, obstacle, rtl };
                  }
                  at = end;
                }
              }
          }
        }
        return undefined;
      };
      for (let hit = find(); hit; hit = find()) {
        if (++nodes > limits.maxNodes)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Centered wrapping exceeds its render node budget.",
          );
        const spacer = markViewNode(document.createElement("span"));
        hideFromViewSelection(spacer);
        measureViewContent(spacer);
        spacer.className = "sdk-word-center-exclusion";
        spacer.contentEditable = "false";
        spacer.setAttribute("aria-hidden", "true");
        const width = hit.rtl ? hit.x - hit.obstacle.left : hit.obstacle.right - hit.x;
        const previous =
          hit.at === 0 && hit.text.previousSibling?.nodeType === 1
            ? (hit.text.previousSibling as HTMLElement)
            : undefined;
        const paragraphBox = hit.p.getBoundingClientRect();
        const breakOnly = hit.rtl
          ? hit.obstacle.left <= paragraphBox.left + 0.5 && hit.x < paragraphBox.right - 2
          : hit.obstacle.right >= paragraphBox.right - 0.5 && hit.x > paragraphBox.left + 2;
        const blocked =
          !breakOnly &&
          (width >= paragraphBox.width - 0.5 ||
            (previous && spacers.get(previous)?.key === hit.obstacle.key));
        spacer.style.cssText = breakOnly
          ? "display:block;height:0;line-height:0;font-size:0;width:0;padding:0;margin:0;border:0;pointer-events:none;user-select:none"
          : blocked
            ? `display:block;height:${Math.max(1, hit.obstacle.bottom - hit.line.top)}px;width:0;padding:0;margin:0;border:0;pointer-events:none;user-select:none`
            : `display:inline-block;width:${Math.max(1, width)}px;height:0;padding:0;margin:0;border:0;font-size:0;line-height:0;vertical-align:baseline;pointer-events:none;user-select:none`;
        const tail = hit.at ? hit.text.splitText(hit.at) : hit.text;
        tail.before(spacer);
        spacers.set(spacer, {
          key: hit.obstacle.key,
          height: hit.line.height,
          baseline: Math.min(
            hit.line.height,
            Math.max(0, spacer.getBoundingClientRect().top - hit.line.top),
          ),
          rtl: hit.rtl,
        });
        changed = true;
      }
      for (const [p, old] of paragraphStyles) {
        if (!old.right || aligned.has(p)) continue;
        const gaps = [...spacers].filter(([spacer]) => spacer.closest("p,h1,h2,h3,h4,h5,h6") === p);
        for (const [gap, entry] of gaps) {
          const obstacle = byKey.get(entry.key),
            box = gap.getBoundingClientRect();
          if (!obstacle || gap.style.display === "block") continue;
          const top = box.top - entry.baseline,
            bottom = top + entry.height,
            walker = document.createTreeWalker(p, 4);
          let first: { text: Text; at: number } | undefined,
            left = -Infinity,
            right = -Infinity;
          for (let current = walker.nextNode(); current; current = walker.nextNode()) {
            reserve();
            const text = current as Text;
            if (
              !text.length ||
              text.parentElement?.closest("[data-sdk-word-object],[data-word-native-object]") ||
              view.getComputedStyle(text.parentElement!).visibility !== "visible"
            )
              continue;
            const range = document.createRange();
            range.selectNodeContents(text);
            for (const rect of range.getClientRects()) {
              if (rect.bottom <= top + 0.5 || rect.top >= bottom - 0.5 || rect.width < 0.1)
                continue;
              if (!first) {
                let low = 0,
                  high = text.length;
                while (low < high) {
                  const mid = Math.floor((low + high) / 2);
                  if (charRect(text, mid, Math.min(text.length, mid + 1)).bottom <= top + 0.5)
                    low = mid + 1;
                  else high = mid;
                }
                while (low > 0 && !isDocumentTextBoundary(text.data, low)) low--;
                first = { text, at: low };
              }
              if (rect.right <= obstacle.left + 0.5) left = Math.max(left, rect.right);
              if (rect.left >= obstacle.right - 0.5) right = Math.max(right, rect.right);
            }
          }
          const shift = Number.isFinite(left) ? Math.max(0, obstacle.left - left) : 0,
            paragraph = view.getComputedStyle(p),
            end =
              p.getBoundingClientRect().right -
              (parseFloat(paragraph.paddingRight) || 0) -
              (parseFloat(paragraph.borderRightWidth) || 0),
            trailing = Number.isFinite(right) ? Math.max(0, end - right) : 0;
          if (first && shift > 0.1) {
            if (++nodes > limits.maxNodes)
              throw new DocumentError(
                "ResourceLimitExceeded",
                "Centered alignment exceeds its render node budget.",
              );
            const filler = markViewNode(document.createElement("span"));
            hideFromViewSelection(filler);
            measureViewContent(filler);
            filler.className = "sdk-word-center-alignment";
            filler.contentEditable = "false";
            filler.setAttribute("aria-hidden", "true");
            filler.style.cssText = `display:inline-block;width:${shift}px;height:0;margin:0;padding:0;border:0;font-size:0;line-height:0;vertical-align:baseline;pointer-events:none;user-select:none`;
            const tail = first.at ? first.text.splitText(first.at) : first.text;
            tail.before(filler);
            alignmentFillers.add(filler);
          }
          gap.style.width = Math.max(0, box.width - shift + trailing - 0.05) + "px";
        }
        if (old.align) p.style.setProperty("text-align", old.align, old.alignPriority);
        else p.style.removeProperty("text-align");
        aligned.add(p);
      }
    },
  };
}
