import { floatingStory } from "./word-floating-placement.js";
import { columnRegion } from "./column-regions.js";
import { wordDrawingLayoutInCell } from "./word-drawing-position.js";
import { wordWrapIntervals, wordWrapPoints } from "./word-wrap-geometry.js";
import type { WrapPoint } from "./word-wrap-geometry.js";
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
  return layout.mode === "floating" && !layout.behindText && (layout.wrap ?? "square") !== "none";
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
  layout?: DeepReadonly<WordDrawing>["layout"];
  points?: WrapPoint[] | undefined;
  cell?: string | undefined;
  story?: HTMLElement | undefined;
}
/** Inline exclusions leave one logical paragraph and the original inline owners
 * intact. They split only UTF-16-safe Text boundaries and carry no source text.
 * The browser still measures fonts, bidi runs, line breaks and native input. */
export interface WordCenterWrap {
  readonly active: boolean;
  readonly changed: boolean;
  readonly storiesChanged: boolean;
  reset(): void;
  measure(pages: readonly RenderedPage[], root: HTMLElement, storiesOnly?: boolean): void;
}
export function wordCenterWrap(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): WordCenterWrap {
  const snapshot = model.snapshot(),
    roots = new Map(snapshot.metadata.word?.stories.map((story) => [story.rootNodeId, story.kind]));
  const inFlow = (id: string): boolean => {
    for (
      let node = model.getNode(id);
      node;
      node = node.parentNodeId ? model.getNode(node.parentNodeId) : undefined
    ) {
      const kind = roots.get(node.nodeId);
      if (kind) return kind === "main" || kind === "header" || kind === "footer";
    }
    return false;
  };
  const drawings = new Map(
    snapshot.metadata.word?.drawings
      ?.filter((item) => needsWordDrawingLineExclusion(item) && inFlow(item.nodeId))
      .map((item) => [item.nodeId, item]),
  );
  let resolver: SourceStyleResolver | undefined;
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
      resolver ??= new SourceStyleResolver(snapshot.styles.map((sheet) => sheet.css), DEFAULT_DOCUMENT_READ_LIMITS);
      const css = resolver.get(project(info.nodeId));
      return css["--word-table-floating"] === "1" && inFlow(info.nodeId)
        ? [[info.nodeId, css] as const]
        : [];
    }),
  );
  const active = drawings.size > 0 || tables.size > 0,
    view = document.defaultView!;
  const spacers = new Map<
      HTMLElement,
      {
        key: string;
        height: number;
        baseline: number;
        rtl: boolean;
        obstacle: Obstacle;
        baseWidth: number;
      }
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
  let changed = false,
    storiesChanged = false;
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
    changed = storiesChanged = false;
  };
  return {
    active,
    get changed(): boolean {
      return changed;
    },
    get storiesChanged() {
      return storiesChanged;
    },
    reset,
    measure(pages, root, storiesOnly = false): void {
      changed = storiesChanged = false;
      const markChange = (story?: HTMLElement): void => {
        if (story) storiesChanged = true;
        else changed = true;
      };
      if (!active) return;
      let work = 0,
        nodes = document.querySelectorAll("*").length;
      const reserve = () => {
        if (++work > limits.maxNodes * 32)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Word wrapping exceeds its measurement budget.",
          );
      };
      const obstacles: Obstacle[] = [],
        origin = root.getBoundingClientRect();
      const add = (
        node: HTMLElement,
        id: string,
        distances: { left: number; right: number; top: number; bottom: number },
        drawing?: DeepReadonly<WordDrawing>,
        anchor?: HTMLElement | null,
      ) => {
        if (!node.getClientRects().length || view.getComputedStyle(node).visibility === "hidden")
          return;
        const rect = node.getBoundingClientRect(),
          pageId = node.closest<HTMLElement>(".sdk-word-floating-page,.sdk-paper")?.dataset.pageId,
          story = floatingStory(anchor);
        if (storiesOnly && !story) return;
        const page = pages.find(
          (page) =>
            !page.blank &&
            (pageId
              ? page.pageId === pageId
              : rect.top >= origin.top + page.rect.y / 0.75 - 0.5 &&
                rect.top < origin.top + (page.rect.y + page.rect.height) / 0.75),
        );
        if (!page) return;
        const points =
          drawing && (drawing.layout.wrap === "tight" || drawing.layout.wrap === "through")
            ? wordWrapPoints(
                drawing.layout.wrapPolygon ?? [
                  { x: 0, y: 0 },
                  { x: 21600, y: 0 },
                  { x: 21600, y: 21600 },
                  { x: 0, y: 21600 },
                ],
                {
                  left: origin.left + Number(node.dataset.sdkWordLayoutLeft),
                  top: origin.top + Number(node.dataset.sdkWordLayoutTop),
                  width: Number(node.dataset.sdkWordWidth ?? drawing.geometry.width) / 0.75,
                  height: Number(node.dataset.sdkWordHeight ?? drawing.geometry.height) / 0.75,
                },
                drawing.geometry,
                reserve,
              )
            : undefined;
        const paintBounds = points
          ? {
              left: points.reduce((v, p) => Math.min(v, p.x), Infinity),
              right: points.reduce((v, p) => Math.max(v, p.x), -Infinity),
              top: points.reduce((v, p) => Math.min(v, p.y), Infinity),
              bottom: points.reduce((v, p) => Math.max(v, p.y), -Infinity),
            }
          : rect;
        const extent = drawing?.layout.effectExtent,
          bounds = { left: paintBounds.left - (extent?.left ?? 0) / .75,
            right: paintBounds.right + (extent?.right ?? 0) / .75,
            top: paintBounds.top - (extent?.top ?? 0) / .75,
            bottom: paintBounds.bottom + (extent?.bottom ?? 0) / .75 };
        obstacles.push({
          key: id + ":" + page.pageId,
          story,
          ...(drawing
            ? {
                layout: drawing.layout,
                points,
                cell:
                  node.dataset.sdkWordWrapCell ??
                  (wordDrawingLayoutInCell(drawing.layout, snapshot.metadata.word?.compatibilityMode)
                    ? anchor?.closest<HTMLElement>("td,th")?.dataset.nodeId
                    : undefined),
              }
            : { cell: anchor?.closest<HTMLElement>("td,th")?.dataset.nodeId }),
          left: bounds.left - distances.left / 0.75,
          right: bounds.right + distances.right / 0.75,
          top: Math.max(bounds.top - distances.top / 0.75, origin.top + page.rect.y / 0.75),
          bottom: Math.min(
            bounds.bottom + distances.bottom / 0.75,
            origin.top + (page.rect.y + page.rect.height) / 0.75,
          ),
        });
      };
      for (const frame of document.querySelectorAll<HTMLElement>(
        ".sdk-word-floating-layer [data-sdk-word-object], .sdk-paper [data-sdk-word-story-anchor]",
      )) {
        const drawing = drawings.get(frame.dataset.sdkWordObject!);
        const paper = frame.closest(".sdk-paper"),
          anchor = paper?.querySelector<HTMLElement>(
            '[data-sdk-word-anchor="' + CSS.escape(frame.dataset.sdkWordObject!) + '"]',
          );
        if (drawing && (!paper || anchor))
          add(frame, drawing.nodeId, { left: 0, right: 0, top: 0, bottom: 0 }, drawing, anchor);
      }
      for (const table of document.querySelectorAll<HTMLTableElement>(
        ".sdk-word-floating-table > table[data-node-id]",
      )) {
        const id = table.dataset.nodeId!,
          css = tables.get(id),
          box = table.parentElement!;
        if (!css || box.style.position !== "absolute" || table.closest(".sdk-repeat-header"))
          continue;
        const anchor =
          table
            .closest(".sdk-paper")
            ?.querySelector<HTMLElement>('[data-sdk-word-table-anchor="' + CSS.escape(id) + '"]') ??
          (box.previousElementSibling as HTMLElement | null);
        const distance = (side: string) =>
          resolveSourceCssLength(css["--word-table-position-" + side + "-distance"] ?? "0pt", {
            font: 12,
            percent: 468,
          });
        const distances = {
          left: distance("left"),
          right: distance("right"),
          top: distance("top"),
          bottom: distance("bottom"),
        };
        add(table, id, distances, undefined, anchor);
      }
      const unalign = (p: HTMLElement | null | undefined): void => {
        if (!p || !aligned.has(p)) return;
        for (const [spacer, entry] of spacers)
          if (spacer.style.display !== "block" && spacer.closest("p,h1,h2,h3,h4,h5,h6") === p)
            spacer.style.width = entry.baseWidth + "px";
        ownViewStyle(p, "text-align");
        p.style.textAlign = "left";
        aligned.delete(p);
      };
      const byKey = new Map(obstacles.map((item) => [item.key, item]));
      const resolveTrackedObstacle = (key: string): Obstacle | undefined => {
        const parts = key.split("\u001f");
        let combined: Obstacle | undefined;
        for (const part of parts) {
          const raw = byKey.get(part);
          if (!raw) return undefined;
          const expanded: Obstacle = {
            ...raw,
            top: raw.top - (raw.layout?.distances?.top ?? 0) / 0.75,
            bottom: raw.bottom + (raw.layout?.distances?.bottom ?? 0) / 0.75,
          };
          if (!combined) {
            combined = expanded;
          } else {
            combined = {
              ...combined,
              left: Math.min(combined.left, expanded.left),
              right: Math.max(combined.right, expanded.right),
              top: Math.min(combined.top, expanded.top),
              bottom: Math.max(combined.bottom, expanded.bottom),
              key,
            };
          }
        }
        return combined;
      };
      for (const [spacer, entry] of spacers) {
        if (storiesOnly && !entry.obstacle.story) continue;
        const connected = spacer.isConnected,
          owner = spacer.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6"),
          box = spacer.getBoundingClientRect(),
          obstacle = resolveTrackedObstacle(entry.key);
        if (
          !spacer.isConnected ||
          !obstacle ||
          box.top - entry.baseline + entry.height <= obstacle.top + 0.25 ||
          box.top - entry.baseline >= obstacle.bottom - 0.25
        ) {
          unalign(owner);
          spacer.remove();
          spacers.delete(spacer);
          if (connected) markChange(entry.obstacle.story);
          continue;
        }
        if (spacer.style.display === "block") continue;
        const width = entry.rtl ? box.right - obstacle.left : obstacle.right - box.left;
        if (width <= 0) {
          unalign(owner);
          spacer.remove();
          spacers.delete(spacer);
          markChange(entry.obstacle.story);
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
            bounds: { left: number; right: number };
            obstacle: Obstacle;
            rtl: boolean;
          }
        | undefined => {
        const paragraphs = obstacles.some((obstacle) => obstacle.story)
          ? document.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")
          : root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6");
        for (const p of paragraphs) {
          reserve();
          const story = floatingStory(p),
            insideFloatingTable = Boolean(p.closest(".sdk-word-floating-table")),
            cell = p.closest<HTMLElement>("td,th")?.dataset.nodeId;
          if (
            p.closest("[data-sdk-word-object]") ||
            (!story &&
              p.closest("[data-office-story]")?.getAttribute("data-office-story") !== "main")
          )
            continue;
          const paragraphBox = p.getBoundingClientRect(),
            fragmented = columnRegion(p, paragraphBox, p)?.fragmented,
            relevant = obstacles.filter(
              (o) =>
                o.story === story &&
                (o.cell ? cell === o.cell : !insideFloatingTable) &&
                (fragmented ||
                  (paragraphBox.bottom > o.top - (o.layout?.distances?.top ?? 0) / 0.75 &&
                    paragraphBox.top < o.bottom + (o.layout?.distances?.bottom ?? 0) / 0.75)),
            );
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
              markChange(story);
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
            for (const line of range.getClientRects()) {
              const region = columnRegion(p, line, p),
                fragment = region?.fragment ?? region,
                bounds = fragment ? { left: fragment.left, right: fragment.right } : paragraphBox;
              const bandObstacles: Obstacle[] = [];
              for (const original of relevant) {
                reserve();
                if (
                  line.bottom <=
                    original.top - (original.layout?.distances?.top ?? 0) / 0.75 + 0.25 ||
                  line.top >=
                    original.bottom + (original.layout?.distances?.bottom ?? 0) / 0.75 - 0.25 ||
                  (region &&
                    (original.right + (original.layout?.distances?.right ?? 0) / 0.75 <=
                      bounds.left ||
                      original.left - (original.layout?.distances?.left ?? 0) / 0.75 >=
                        bounds.right))
                )
                  continue;
                for (const span of original.layout
                  ? wordWrapIntervals(
                      original,
                      original.points,
                      original.layout,
                      line,
                      bounds,
                      rtl,
                      reserve,
                    )
                  : [original]) {
                  bandObstacles.push({
                    ...original,
                    ...span,
                    top: original.top - (original.layout?.distances?.top ?? 0) / 0.75,
                    bottom: original.bottom + (original.layout?.distances?.bottom ?? 0) / 0.75,
                  });
                }
              }
              if (!bandObstacles.length) continue;
              // Merge overlapping/adjacent concurrent floating obstacles on the same line band
              // and process them in reading order so multi-object wraps remain stable.
              bandObstacles.sort((a, b) => (rtl ? b.right - a.right : a.left - b.left));
              const mergedObstacles: Obstacle[] = [];
              for (const item of bandObstacles) {
                const prev = mergedObstacles.at(-1);
                if (
                  prev &&
                  prev.layout &&
                  item.layout &&
                  prev.layout.wrapSide === item.layout.wrapSide &&
                  (rtl ? item.right >= prev.left - 2 : item.left <= prev.right + 2)
                ) {
                  prev.left = Math.min(prev.left, item.left);
                  prev.right = Math.max(prev.right, item.right);
                  prev.top = Math.min(prev.top, item.top);
                  prev.bottom = Math.max(prev.bottom, item.bottom);
                  prev.key = prev.key + "\u001f" + item.key;
                } else {
                  mergedObstacles.push({ ...item });
                }
              }
              for (const obstacle of mergedObstacles) {
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
                      if (aligned.has(p)) {
                        unalign(p);
                        markChange(story);
                        return find();
                      }
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
                          markChange(story);
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
                      return { text, at, x, line, p, bounds, obstacle, rtl };
                    }
                    at = end;
                  }
              }
            }
          }
        }
        return undefined;
      };
      for (let hit = find(); hit; hit = find()) {
        const width = hit.rtl ? hit.x - hit.obstacle.left : hit.obstacle.right - hit.x;
        const previous =
          hit.at === 0 && hit.text.previousSibling?.nodeType === 1
            ? (hit.text.previousSibling as HTMLElement)
            : undefined;
        const paragraphBox = hit.bounds,
          widthInFlow = paragraphBox.right - paragraphBox.left;
        const previousEntry = previous ? spacers.get(previous) : undefined;
        // Glyph bounds can overlap their adjacent advance by a pixel (notably
        // in WebKit's CJK runs). Extend an existing gap by that residual instead
        // of treating a rounding/ink overlap as a wholly blocked text line.
        if (
          previous &&
          previousEntry?.key === hit.obstacle.key &&
          previous.style.display !== "block" &&
          Math.abs(previous.getBoundingClientRect().top - previousEntry.baseline - hit.line.top) <
            0.5
        ) {
          const next = (parseFloat(previous.style.width) || 0) + Math.max(1, width);
          if (next < widthInFlow - 0.5) {
            previous.style.width = next + "px";
            previousEntry.baseWidth = next;
            markChange(hit.obstacle.story);
            continue;
          }
        }
        const breakOnly = hit.rtl
          ? hit.obstacle.left <= paragraphBox.left + 0.5 && hit.x < paragraphBox.right - 2
          : hit.obstacle.right >= paragraphBox.right - 0.5 && hit.x > paragraphBox.left + 2;
        const blocked =
          !breakOnly &&
          (width >= widthInFlow - 0.5 ||
            (previous &&
              spacers.get(previous)?.key === hit.obstacle.key &&
              Math.abs(
                previous.getBoundingClientRect().top -
                  spacers.get(previous)!.baseline -
                  hit.line.top,
              ) < 0.5));
        if (++nodes > limits.maxNodes)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Word wrapping exceeds its render node budget.",
          );
        const spacer = markViewNode(document.createElement("span"));
        hideFromViewSelection(spacer);
        measureViewContent(spacer);
        spacer.className = "sdk-word-center-exclusion";
        spacer.contentEditable = "false";
        spacer.setAttribute("aria-hidden", "true");
        spacer.style.cssText = breakOnly
          ? "display:block;height:0;line-height:0;font-size:0;width:0;padding:0;margin:0;border:0;pointer-events:none;user-select:none"
          : blocked
            ? `display:block;height:${Math.max(1, hit.obstacle.bottom - hit.line.top)}px;width:0;padding:0;margin:0;border:0;pointer-events:none;user-select:none`
            : `display:inline-block;width:${Math.max(1, width)}px;height:0;padding:0;margin:0;border:0;font-size:0;line-height:0;vertical-align:baseline;pointer-events:none;user-select:none`;
        const tail = hit.at ? hit.text.splitText(hit.at) : hit.text;
        tail.before(spacer);
        if (blocked) {
          // A block inserted after an inline anchor can start below the old
          // line box. Measure that actual start so the old line is not charged
          // again as extra clearance below a full-width floating object.
          spacer.style.height = Math.max(1, hit.obstacle.bottom - spacer.getBoundingClientRect().top) + "px";
        }
        spacers.set(spacer, {
          key: hit.obstacle.key,
          height: hit.line.height,
          baseline: Math.min(
            hit.line.height,
            Math.max(0, spacer.getBoundingClientRect().top - hit.line.top),
          ),
          rtl: hit.rtl,
          obstacle: hit.obstacle,
          baseWidth: Math.max(1, width),
        });
        markChange(hit.obstacle.story);
      }
      for (const [p, old] of paragraphStyles) {
        if (!p.isConnected) {
          paragraphStyles.delete(p);
          aligned.delete(p);
          continue;
        }
        if (!old.right || aligned.has(p)) continue;
        const pStyle = view.getComputedStyle(p),
          pRect = p.getBoundingClientRect(),
          pBounds = {
            left:
              pRect.left +
              (parseFloat(pStyle.paddingLeft) || 0) +
              (parseFloat(pStyle.borderLeftWidth) || 0),
            right:
              pRect.right -
              (parseFloat(pStyle.paddingRight) || 0) -
              (parseFloat(pStyle.borderRightWidth) || 0),
          };
        const gaps = [...spacers].filter(
          ([spacer, entry]) =>
            spacer.style.display !== "block" &&
            Boolean(resolveTrackedObstacle(entry.key)) &&
            spacer.closest("p,h1,h2,h3,h4,h5,h6") === p,
        );
        // Measure every wrapped line before mutating spacer widths so adjusting
        // one line cannot shift or misattribute later lines in the paragraph.
        const adjustments: { gap: HTMLElement; width: number }[] = [];
        for (const [gap, entry] of gaps) {
          const obstacle = entry.obstacle,
            box = gap.getBoundingClientRect(),
            top = box.top - entry.baseline,
            bottom = top + entry.height,
            region = columnRegion(
              p,
              new DOMRect(box.left, top, box.right - box.left, bottom - top),
              p,
            ),
            fragment = region?.fragment ?? region,
            bounds = fragment ? { left: fragment.left, right: fragment.right } : pBounds,
            walker = document.createTreeWalker(p, 4);
          let rightEdge: number | undefined;
          for (let current = walker.nextNode(); current; current = walker.nextNode()) {
            reserve();
            const text = current as Text;
            if (
              !text.length ||
              !(gap.compareDocumentPosition(text) & 4) ||
              text.parentElement?.closest("[data-sdk-word-object],[data-word-native-object]") ||
              view.getComputedStyle(text.parentElement!).visibility !== "visible"
            )
              continue;
            const range = document.createRange();
            range.selectNodeContents(text);
            for (const rect of range.getClientRects()) {
              const mid = (rect.top + rect.bottom) / 2;
              if (mid < top - 0.5 || mid > bottom + 0.5 || rect.width < 0.1) continue;
              if (rect.left >= box.right - 1) {
                let low = 0,
                  high = text.length;
                while (low < high) {
                  const midIndex = Math.floor((low + high) / 2),
                    r = charRect(text, midIndex, Math.min(text.length, midIndex + 1));
                  if ((r.top + r.bottom) / 2 < top - 0.5) low = midIndex + 1;
                  else high = midIndex;
                }
                let endIndex = low;
                while (endIndex < text.length) {
                  let step = endIndex + 1;
                  while (step < text.length && !isDocumentTextBoundary(text.data, step)) step++;
                  const r = charRect(text, endIndex, step);
                  if ((r.top + r.bottom) / 2 > bottom + 0.5) break;
                  endIndex = step;
                }
                // Ignore trailing collapsible ASCII spaces that disappear under text-align:right.
                let trimmed = endIndex;
                while (trimmed > low && /[ \t]/.test(text.data[trimmed - 1]!)) trimmed--;
                if (trimmed > low) {
                  const edge = trimmed === endIndex ? rect.right : charRect(text, low, trimmed).right;
                  rightEdge = Math.max(rightEdge ?? -Infinity, edge);
                }
              }
            }
          }
          const leftWidth = Math.max(0, box.left - bounds.left),
            rightWidth = rightEdge !== undefined ? Math.max(0, rightEdge - box.right) : 0,
            minGap = Math.max(0, obstacle.right - obstacle.left),
            maxFit = Math.max(0, bounds.right - bounds.left - leftWidth - rightWidth - 0.25),
            target =
              leftWidth > 0.5
                ? Math.min(maxFit, Math.max(minGap, bounds.right - rightWidth - obstacle.left + 0.5))
                : maxFit;
          adjustments.push({ gap, width: Math.max(0, target) });
        }
        for (const { gap, width } of adjustments) gap.style.width = width + "px";
        if (old.align) p.style.setProperty("text-align", old.align, old.alignPriority);
        else p.style.removeProperty("text-align");
        aligned.add(p);
      }
    },
  };
}
