import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WordMetadata } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { isMeasuredViewContent, markViewNode, viewNodeRole } from "./view-nodes.js";

const paragraphSelector = "p,h1,h2,h3,h4,h5,h6";
const atomSelector =
  "br,img,[data-word-tab],[data-word-drawing],[data-word-native-object]:not([data-word-native-boundary]),[data-word-note-reference],[data-word-equation]";
const eastAsianText =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\u3000-\u303f\uff01-\uffef]/u;
const fullWidth = /[\u3000-\u303f\ufe10-\ufe6f\uff01-\uff60]/u;
const kana = /[\p{Script=Hiragana}\p{Script=Katakana}]/u;

// ECMA-376 §17.3.1.16 and MS-OI29500 §2.1.52.
const lineRules = {
  hans: {
    before:
      "!%),.:;>?]}¢¨°·ˇˉ―‖’”…‰′″›℃∶、。〃〉》」』】〕〗〞︶︺︾﹀﹄﹚﹜﹞！＂％＇），．：；？］｀｜｝～￠",
    after: "$([{£¥·‘“〈《「『【〔〖〝﹙﹛﹝＄（．［｛￡￥",
    hang: "!%),.:;>?]}¢°·ˇ’”‰′″℃∶、。〃〉》」』〗〕〞﹚﹜﹞！＂％＇），．：；？］｝￠",
  },
  hant: {
    before:
      "!),.:;?]}¢·–—’”•‥…‧′╴、。〉》」』】〕〞︰︱︳︴︶︸︺︼︾﹀﹂﹄﹏﹐﹑﹒﹔﹕﹖﹗﹚﹜﹞！），．：；？］｜｝､",
    after: "([{£¥‘“‵〈《「『【〔〝︵︷︹︻︽︿﹁﹃﹙﹛﹝（｛",
    hang: "!),.:;?]}’”′、。〉》」』〗〕〞﹚﹜﹞！），．：；？］｝",
  },
  ja: {
    before: "!%),.:;?]}¢°’”‰′″℃、。々〉》」』】〕゛゜ゝゞ・ヽヾ！％），．：；？］｝｡｣､･ﾞﾟ￠",
    after: "$([\\{£¥‘“〈《「『【〔＄（［｛｢￡￥",
    hang: ",.’”、。」』〗），．］｝｡､",
  },
  ko: {
    before: "!%),.:;?]}¢°’”′″℃〉》」』】〕！％），．：；？］｝￠",
    after: "$([\\{£¥‘“〈《「『【〔＄（［｛￡￥￦",
    hang: "!%),.:;?]}¢°’”′″℃〉》」』〗〕！％），．：；？］｝￠",
  },
};
const strictKana = "ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶーｧｨｩｪｫｬｭｮｯｰ";
const languageAliases = new Map([
  ["zh", "zh-cn"],
  ["zh-hans", "zh-cn"],
  ["zh-hant", "zh-tw"],
  ["ja", "ja-jp"],
  ["ko", "ko-kr"],
]);
type Locale = keyof typeof lineRules;
function locale(language: string): Locale | undefined {
  if (/^zh(?:-|$)/.test(language))
    return /(?:^|-)(?:hant|tw|hk|mo)(?:-|$)/.test(language) ? "hant" : "hans";
  if (/^ja(?:-|$)/.test(language)) return "ja";
  if (/^ko(?:-|$)/.test(language)) return "ko";
  return undefined;
}
function languageKey(value: string): string {
  const tag = value
    .trim()
    .replace(/^["']|["']$/g, "")
    .toLowerCase();
  return languageAliases.get(tag) ?? tag;
}
interface Glyph {
  node: Text;
  offset: number;
  text: string;
  language: string;
  grid: boolean;
  fontSize: number;
  prior: Glyph | undefined;
  next?: Glyph;
  guard?: "join" | "break";
}
interface Punctuation {
  glyph: Glyph;
  span: HTMLElement;
  hanging: boolean;
  leading: boolean;
  start: number;
  end: number;
}

function fitPunctuation(
  items: Punctuation[],
  writingMode: string,
  range: Range,
  step: () => void,
): void {
  const vertical = writingMode.startsWith("vertical"),
    progress = (rect: DOMRect) =>
      vertical ? (writingMode === "vertical-rl" ? -rect.right : rect.left) : rect.top,
    extent = (rect: DOMRect) => (vertical ? rect.height : rect.width),
    sameLine = (a: DOMRect | undefined, b: DOMRect | undefined): boolean => {
      if (!a || !b) return false;
      const [start, end, size] = vertical
        ? (["left", "right", "width"] as const)
        : (["top", "bottom", "height"] as const);
      return (
        Math.min(a[end], b[end]) - Math.max(a[start], b[start]) > Math.min(a[size], b[size]) / 2
      );
    };
  const measure = () => {
    const measured = new Map<Glyph, DOMRect | undefined>();
    return (glyph: Glyph | undefined): DOMRect | undefined => {
      if (!glyph) return;
      if (measured.has(glyph)) return measured.get(glyph);
      step();
      range.setStart(glyph.node, glyph.offset);
      range.setEnd(glyph.node, glyph.offset + glyph.text.length);
      const rect = range.getBoundingClientRect(),
        visible = extent(rect) > 0.01 ? rect : undefined;
      measured.set(glyph, visible);
      return visible;
    };
  };
  const place = (item: Punctuation, start: number, end: number): void => {
    item.span.style.marginInlineStart = start ? -start + "px" : "";
    item.span.style.marginInlineEnd = end ? -end + "px" : "";
  };
  for (const item of items) place(item, item.start, item.end);

  const leading = new Map<Punctuation, number>();
  for (let pass = 0; pass < 3; pass++) {
    const read = measure(),
      changes: [Punctuation, number][] = [];
    for (const item of items) {
      if (!item.leading) continue;
      const rect = read(item.glyph),
        prior = read(item.glyph.prior);
      const trim =
        rect && !sameLine(rect, prior)
          ? Math.max(item.start, (extent(rect) - item.end) / 2)
          : item.start;
      if (trim !== (leading.get(item) ?? item.start)) changes.push([item, trim]);
    }
    if (!changes.length) break;
    for (const [item, trim] of changes) {
      leading.set(item, trim);
      place(item, trim, item.end);
    }
  }

  // A zero-advance inline box paints the original glyph past the line edge.
  // Keep it only when it joins the preceding line at its end.
  const before = measure(),
    probes: { item: Punctuation; rect: DOMRect }[] = [];
  for (const item of items) {
    if (!item.hanging) continue;
    const glyph = item.glyph,
      rect = before(glyph),
      prior = before(glyph.prior);
    if (!rect || !prior) continue;
    if (
      sameLine(rect, prior) &&
      sameLine(rect, before(glyph.next)) &&
      (!glyph.prior?.prior || sameLine(prior, before(glyph.prior.prior)))
    )
      continue;
    probes.push({ item, rect });
  }
  const hang = (item: Punctuation, enabled: boolean): void => {
    item.span.style.display = enabled ? "inline-block" : "";
    item.span.style.inlineSize = enabled ? "0px" : "";
    item.span.style.whiteSpace = enabled ? "nowrap" : "";
    place(item, enabled ? 0 : (leading.get(item) ?? item.start), enabled ? 0 : item.end);
  };
  for (const { item } of probes) hang(item, true);
  const after = measure();
  const rejected = probes.filter(({ item, rect }) => {
    const glyph = item.glyph,
      moved = after(glyph);
    return (
      !moved ||
      progress(moved) >= progress(rect) - 0.5 ||
      !sameLine(moved, after(glyph.prior)) ||
      sameLine(moved, after(glyph.next))
    );
  });
  for (const { item } of rejected) hang(item, false);
  if (leading.size || probes.length !== rejected.length) {
    const read = measure();
    const invalid = items.filter((item) => {
      const rect = read(item.glyph);
      return (
        (item.span.style.inlineSize === "0px" &&
          (!sameLine(rect, read(item.glyph.prior)) || sameLine(rect, read(item.glyph.next)))) ||
        (leading.get(item) !== undefined && sameLine(rect, read(item.glyph.prior)))
      );
    });
    for (const item of invalid) {
      leading.delete(item);
      hang(item, false);
    }
  }
}

/** Layout-only additions carry no source identity and can be rebuilt at a new width. */
export function wordPunctuationLayout(
  document: Document,
  word: DeepReadonly<WordMetadata> | undefined,
  grids: ReadonlySet<string>,
  limits: Readonly<RenderLimits>,
): { reset(): void; measure(root: HTMLElement, fixedGrid?: boolean): void } {
  const options = word?.typography,
    view = document.defaultView!,
    range = document.createRange(),
    segmenter = new Intl.Segmenter("und", { granularity: "grapheme" }),
    overlays: HTMLElement[] = [];
  const before = options?.noLineBreaksBefore,
    after = options?.noLineBreaksAfter,
    beforeKey = before && languageKey(before.language),
    afterKey = after && languageKey(after.language),
    beforeSet = new Set(before?.characters),
    afterSet = new Set(after?.characters);
  const spacing = options?.characterSpacingControl;
  function reset(): void {
    for (const element of overlays.splice(0)) {
      if (viewNodeRole(element) === "unwrap") element.replaceWith(...element.childNodes);
      else element.remove();
    }
  }
  function measure(root: HTMLElement, fixedGrid?: boolean): void {
    if (!word) return;
    let work = 0,
      nodes = document.getElementsByTagName("*").length;
    const step = () => {
      if (++work > limits.maxNodes * 32)
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Word punctuation exceeds the render work budget.",
        );
    };
    const overlay = (tag: "span" | "wbr", role: "unwrap" | "remove") => {
      if (++nodes > limits.maxNodes)
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Word punctuation exceeds the render node budget.",
        );
      const element = markViewNode(document.createElement(tag), role);
      overlays.push(element);
      return element;
    };
    for (const paragraph of root.querySelectorAll<HTMLElement>(paragraphSelector)) {
      if (
        paragraph.closest(
          ".sdk-auxiliary,[data-word-drawing],[data-word-native-object]:not([data-word-native-boundary])",
        ) ||
        !paragraph.textContent
      )
        continue;
      const css = view.getComputedStyle(paragraph);
      if (
        !paragraph.getClientRects().length ||
        css.whiteSpace === "nowrap" ||
        css.whiteSpace === "pre"
      )
        continue;
      const top = css.getPropertyValue("--word-top-line-punctuation").trim() === "1",
        overflow = css.getPropertyValue("--word-overflow-punctuation").trim() !== "0",
        kinsoku = css.getPropertyValue("--word-kinsoku").trim() !== "0",
        custom =
          kinsoku &&
          (before !== undefined ||
            after !== undefined ||
            options?.strictFirstAndLastChars !== undefined),
        compress = spacing && spacing !== "doNotCompress";
      if (!top && !overflow && !custom && !compress) continue;
      const raw = paragraph.textContent;
      if (
        !eastAsianText.test(raw) &&
        !custom &&
        !paragraph.querySelector("[lang]") &&
        !locale(languageKey(css.getPropertyValue("--word-language-eastAsia") || paragraph.lang))
      )
        continue;
      const grid =
        (fixedGrid ??
          grids.has(
            paragraph.closest("section[data-node-id]")?.getAttribute("data-node-id") ?? "",
          )) &&
        css.getPropertyValue("--word-paragraph-snap-to-grid").trim() !== "0";
      const glyphs: Glyph[] = [];
      let prior: Glyph | undefined,
        inferredLanguage = "";
      const visit = (node: Node): void => {
        step();
        if (node.nodeType === 1) {
          const element = node as HTMLElement;
          if (element.matches(atomSelector)) {
            prior = undefined;
            return;
          }
          if (viewNodeRole(element) === "remove" && !isMeasuredViewContent(element)) return;
          const style = view.getComputedStyle(element);
          if (style.display === "none" || style.visibility === "hidden") return;
        }
        if (node.nodeType !== 3) {
          for (const child of node.childNodes) visit(child);
          return;
        }
        const text = node as Text,
          parent = text.parentElement!;
        if (!text.length || parent.closest("[data-sdk-word-autospace]")) return;
        const style = view.getComputedStyle(parent),
          lang = languageKey(
            style.getPropertyValue("--word-language-eastAsia") ||
              parent.closest("[lang]")?.getAttribute("lang") ||
              "",
          ),
          snapped = grid && style.getPropertyValue("--word-run-snap-to-grid").trim() !== "0",
          fontSize = parseFloat(style.fontSize);
        if (inferredLanguage !== "ja-jp" && kana.test(text.data)) inferredLanguage = "ja-jp";
        else if (inferredLanguage !== "ja-jp" && /\p{Script=Hangul}/u.test(text.data))
          inferredLanguage = "ko-kr";
        else if (!inferredLanguage && eastAsianText.test(text.data)) inferredLanguage = "zh-cn";
        for (const part of segmenter.segment(text.data)) {
          step();
          if (/^[\r\n]/.test(part.segment) && /^(?:pre|break-spaces)/.test(style.whiteSpace)) {
            prior = undefined;
            continue;
          }
          const glyph: Glyph = {
            node: text,
            offset: part.index,
            text: part.segment,
            language: lang,
            grid: snapped,
            fontSize,
            prior,
          };
          if (prior) prior.next = glyph;
          glyphs.push(glyph);
          prior = glyph;
        }
      };
      visit(paragraph);
      if (!glyphs.length) continue;
      for (const glyph of glyphs) glyph.language ||= inferredLanguage;
      for (const glyph of glyphs) {
        const prev = glyph.prior;
        if (!custom || !prev) continue;
        const currentLocale = locale(glyph.language),
          prevLocale = locale(prev.language),
          first = [...glyph.text][0]!,
          last = [...prev.text][0]!,
          nativeBefore = currentLocale ? lineRules[currentLocale].before.includes(first) : false,
          nativeAfter = prevLocale ? lineRules[prevLocale].after.includes(last) : false;
        const useBefore =
          currentLocale !== undefined &&
          glyph.language === beforeKey &&
          !(currentLocale === "ja" && options?.strictFirstAndLastChars !== undefined);
        const useAfter =
          prevLocale !== undefined &&
          prev.language === afterKey &&
          !(prevLocale === "ja" && options?.strictFirstAndLastChars !== undefined);
        const strict = currentLocale === "ja" ? options?.strictFirstAndLastChars : undefined;
        if (
          (useBefore ? beforeSet.has(first) : nativeBefore) ||
          (useAfter ? afterSet.has(last) : nativeAfter) ||
          (strict && strictKana.includes(first))
        )
          glyph.guard = "join";
        else if (
          (useBefore && nativeBefore) ||
          (useAfter && nativeAfter) ||
          (strict === false && strictKana.includes(first))
        )
          glyph.guard = "break";
      }
      const punctuation: Punctuation[] = [];
      const pieces = new Map<Text, { node: Text; offset: number }[]>();
      const split = (text: Text, offset: number): Text => {
        const node = text.splitText(offset),
          list = pieces.get(text) ?? [];
        if (!pieces.has(text)) pieces.set(text, list);
        list.push({ node, offset });
        return node;
      };
      // Splitting from the end keeps offsets within each original text valid.
      for (let index = glyphs.length - 1; index >= 0; index--) {
        const glyph = glyphs[index]!;
        const script = locale(glyph.language),
          base = [...glyph.text][0]!,
          hanging = Boolean(
            script &&
              overflow &&
              !(glyph.grid && word.doNotWrapTextWithPunct) &&
              glyph.prior &&
              lineRules[script].hang.includes(base),
          ),
          mark = fullWidth.test(base) && /[\p{P}\p{Zs}]/u.test(base),
          leading = Boolean(script && top && /\p{P}/u.test(base)),
          trim = Boolean(script && compress && mark && !glyph.grid),
          proportional =
            script === "ja" &&
            spacing === "compressPunctuationAndJapaneseKana" &&
            kana.test(base) &&
            !glyph.grid;
        if (!hanging && !leading && !trim && !proportional && !glyph.guard) continue;
        const span =
          hanging || leading || trim || proportional ? overlay("span", "unwrap") : undefined;
        let target: Node;
        if (span) {
          span.className = "sdk-word-punctuation";
          const end = glyph.offset + glyph.text.length;
          if (end < glyph.node.length) split(glyph.node, end);
          const text = glyph.offset ? split(glyph.node, glyph.offset) : glyph.node;
          text.before(span);
          span.append(text);
          const opening = /[\p{Ps}\p{Pi}]/u.test(base);
          const item = { glyph, span, hanging, leading, start: 0, end: 0 };
          if (trim || proportional) {
            // Full-width punctuation has up to half a cell of trim space.
            // Kana use the font's proportional metrics where supplied.
            if (trim) {
              if (opening) item.start = glyph.fontSize / 2;
              else item.end = glyph.fontSize / 2;
            } else span.style.fontVariantEastAsian = "proportional-width";
          }
          punctuation.push(item);
          target = span;
        } else target = glyph.offset ? split(glyph.node, glyph.offset) : glyph.node;
        if (glyph.guard) {
          const guard = overlay(glyph.guard === "join" ? "span" : "wbr", "remove");
          guard.className = "sdk-word-kinsoku-boundary";
          if (glyph.guard === "join") {
            guard.contentEditable = "false";
            guard.setAttribute("aria-hidden", "true");
            guard.style.cssText = "font-size:0;line-height:0;letter-spacing:0";
            guard.textContent = "\u2060";
          }
          target.parentNode!.insertBefore(guard, target);
        }
      }
      if (!punctuation.length) continue;
      // A split glyph's text may now be in a sibling of its original node.
      for (const list of pieces.values()) list.reverse();
      const cursors = new Map<Text, number>();
      for (const glyph of glyphs) {
        const list = pieces.get(glyph.node);
        if (!list) continue;
        let index = cursors.get(glyph.node) ?? -1;
        while (list[index + 1] && list[index + 1]!.offset <= glyph.offset) index++;
        cursors.set(glyph.node, index);
        const piece = list[index];
        if (!piece) continue;
        glyph.node = piece.node;
        glyph.offset -= piece.offset;
      }
      fitPunctuation(punctuation.reverse(), css.writingMode, range, step);
    }
  }
  return { reset, measure };
}
