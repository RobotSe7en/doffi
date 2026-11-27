import { DocumentError } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { markViewNode, viewNodeRole } from "./view-nodes.js";

const scaleProperty = "--word-run-native-w";
const paragraphs = "p,h1,h2,h3,h4,h5,h6,li";

/** Shape whole breakable pieces at their original font size. The outer box
 * reserves the scaled advance, and the actual source text remains selectable. */
export function wordRunWidthLayout(document: Document, limits: Readonly<RenderLimits>) {
  const view = document.defaultView!,
    segmenter = new Intl.Segmenter("und", { granularity: "grapheme" }),
    overlays: HTMLElement[] = [];
  return {
    reset(): void {
      // Repeated stories and floating text clone the view tree, including
      // ownership marks. Reset those projections when fonts/layout change too.
      const owned = new Set([...overlays.splice(0), ...document.querySelectorAll<HTMLElement>(".sdk-word-run-width,.sdk-word-run-width-text,.sdk-word-kinsoku-boundary")]);
      for (const node of owned) {
        if (viewNodeRole(node) === "unwrap") node.replaceWith(...node.childNodes);
        else if (viewNodeRole(node) === "remove") node.remove();
      }
    },
    measure(root: HTMLElement): void {
      const range = document.createRange();
      let count = document.getElementsByTagName("*").length, work = 0;
      const add = (role: "remove" | "unwrap"): HTMLElement => {
        if (++count > limits.maxNodes)
          throw new DocumentError("ResourceLimitExceeded", "Scaled Word runs exceed the view node budget.");
        const node = markViewNode(document.createElement("span"), role);
        overlays.push(node);
        return node;
      };
      const textNodes = (node: Node): Text[] => {
        const walker = document.createTreeWalker(node, 4), nodes: Text[] = [];
        for (let item = walker.nextNode(); item; item = walker.nextNode()) nodes.push(item as Text);
        return nodes;
      };
      for (const paragraph of root.querySelectorAll<HTMLElement>(paragraphs)) {
        if (paragraph.querySelector(paragraphs + ",.sdk-word-run-width") || !paragraph.getClientRects().length) continue;
        const sources = textNodes(paragraph),
          styles = new Map<HTMLElement, CSSStyleDeclaration>(),
          runs = sources.map((text) => {
            const host = text.parentElement!;
            let css = styles.get(host);
            if (!css) {
              css = view.getComputedStyle(host);
              styles.set(host, css);
            }
            const scale = Number(css.getPropertyValue(scaleProperty)) / 100;
            return scale >= 0.01 && scale <= 6 && scale !== 1 &&
              css.visibility === "visible" &&
              viewNodeRole(host) !== "remove" &&
              (!host.closest("[data-word-drawing],[data-word-native-object],[data-word-equation],.sdk-page-gap") || !!host.closest(".sdk-word-drawing-text"))
              ? scale : 1;
          });
        if (!runs.some((scale) => scale !== 1)) continue;

        const css = view.getComputedStyle(paragraph),
          vertical = css.writingMode.startsWith("vertical"),
          copy = paragraph.cloneNode(true) as HTMLElement;
        copy.removeAttribute("id");
        copy.style.cssText += ";position:absolute!important;visibility:hidden!important;pointer-events:none;"
          + "inset:0 auto auto -10000px!important;inline-size:1px!important;max-inline-size:none!important;"
          + "min-inline-size:0!important;block-size:auto!important;text-indent:0!important;"
          + "text-align:start!important;text-align-last:start!important;margin:0!important;";
        const fragments: { index: number; from: number; to: number; join: boolean }[] = [];
        let last: DOMRect | undefined;
        paragraph.after(copy);
        try {
          for (const [index, text] of textNodes(copy).entries()) {
            for (const { index: offset, segment } of segmenter.segment(text.data)) {
              if (++work > limits.maxNodes * 32)
                throw new DocumentError("ResourceLimitExceeded", "Scaled Word runs exceed the layout work budget.");
              const end = offset + segment.length;
              range.setStart(text, offset);
              range.setEnd(text, end);
              const rect = range.getClientRects()[0];
              if (!rect || (!rect.height && !rect.width)) continue;
              const joined = last !== undefined && (vertical
                ? Math.min(last.right, rect.right) - Math.max(last.left, rect.left) > 0.5
                : Math.min(last.bottom, rect.bottom) - Math.max(last.top, rect.top) > 0.5);
              const fragment = fragments.at(-1);
              if (joined && fragment?.index === index && fragment.to === offset) fragment.to = end;
              else fragments.push({ index, from: offset, to: end, join: joined });
              last = /[\r\n]/.test(segment) ? undefined : rect;
            }
          }
        } finally {
          copy.remove();
        }

        const boxes: { box: HTMLElement; text: HTMLElement; scale: number }[] = [];
        for (let i = fragments.length - 1; i >= 0; i--) {
          const { index, from, to, join } = fragments[i]!,
            scale = runs[index]!,
            previous = fragments[i - 1],
            joinScaled = join && previous && (scale !== 1 || runs[previous.index] !== 1);
          if (scale === 1 && !joinScaled) continue;
          const node = sources[index]!;
          if (to < node.length) node.splitText(to);
          const piece = from ? node.splitText(from) : node;
          let target: Node = piece;
          if (scale !== 1) {
            const box = add("unwrap"), text = add("unwrap");
            box.className = "sdk-word-run-width";
            text.className = "sdk-word-run-width-text";
            box.style.cssText =
              "display:inline-block;vertical-align:baseline;white-space:pre;text-indent:0";
            text.style.cssText =
              "display:inline-block;white-space:pre;inline-size:max-content;text-indent:0;text-align:start";
            // Text decorations do not propagate through atomic inline boxes.
            for (let owner = piece.parentElement; owner; owner = owner.parentElement) {
              const decoration = view.getComputedStyle(owner);
              if (decoration.textDecorationLine !== "none") {
                text.style.textDecoration = decoration.textDecoration;
                break;
              }
              if (owner === paragraph) break;
            }
            piece.before(box);
            text.append(piece);
            box.append(text);
            boxes.push({ box, text, scale });
            target = box;
          }
          if (joinScaled) {
            const boundary = add("remove");
            boundary.setAttribute("aria-hidden", "true");
            boundary.className = "sdk-word-kinsoku-boundary";
            boundary.contentEditable = "false";
            boundary.style.cssText = "font-size:0;line-height:0;letter-spacing:0";
            boundary.textContent = "\u2060";
            target.parentNode!.insertBefore(boundary, target);
          }
        }
        const widths = boxes.map(({ text }) => parseFloat(view.getComputedStyle(text).inlineSize));
        boxes.forEach(({ box, text, scale }, i) => {
          box.style.inlineSize = widths[i]! * scale + "px";
          text.style.transform = vertical ? `scaleY(${scale})` : `scaleX(${scale})`;
          text.style.transformOrigin = vertical ? "center top" : css.direction === "rtl" ? "right center" : "left center";
        });
      }
    },
  };
}

export type WordRunWidthLayout = ReturnType<typeof wordRunWidthLayout>;
