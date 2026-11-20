import { markViewNode } from "../view-nodes.js";

// ECMA-376 §17.3.1.16 / §21.1.2.2.7 hanging closing punctuation set.
const HANGING_PUNCT = new Set([..."!%),.:;>?]}¢°·ˇ’”‰′″℃∶、。〃〉》」』〗〕〞﹚﹜﹞！＂％＇），．：；？］｝￠"]);

interface CharPos {
  readonly node: Text;
  readonly offset: number;
  readonly ch: string;
}

function collectChars(paragraph: HTMLElement): CharPos[] {
  const doc = paragraph.ownerDocument, view = doc.defaultView;
  if (!view) return [];
  const chars: CharPos[] = [];
  const walker = doc.createTreeWalker(paragraph, view.NodeFilter.SHOW_TEXT);
  for (let current = walker.nextNode(); current; current = walker.nextNode()) {
    const text = current as Text;
    if (!text.data || text.parentElement?.closest(".sdk-field-display,.sdk-ppt-marker")) continue;
    let offset = 0;
    for (const ch of text.data) {
      chars.push({ node: text, offset, ch });
      offset += ch.length;
    }
  }
  return chars;
}

function sameLine(a: DOMRect, b: DOMRect): boolean {
  if (a.height <= 0.1 || b.height <= 0.1) return false;
  return Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top) > Math.min(a.height, b.height) / 2;
}

function rectOf(range: Range, pos: CharPos): DOMRect {
  range.setStart(pos.node, pos.offset);
  range.setEnd(pos.node, pos.offset + pos.ch.length);
  return range.getBoundingClientRect();
}

/**
 * Implements DrawingML `a:pPr/@hangingPunct="1"` for browsers without native
 * CSS `hanging-punctuation: allow-end`. Every wrapper uses `unwrap` view-node
 * metadata so source DOM text and UTF-16 projections remain lossless.
 */
export function applyPresentationHangingPunctuation(content: HTMLElement): boolean {
  const doc = content.ownerDocument, view = doc.defaultView;
  let changed = false;
  const parents = new Set<Node>();
  for (const existing of content.querySelectorAll<HTMLElement>("[data-sdk-hanging-punct]")) {
    if (existing.parentNode) parents.add(existing.parentNode);
    existing.replaceWith(...existing.childNodes);
    changed = true;
  }
  for (const existing of content.querySelectorAll<HTMLElement>("[data-sdk-hanging-break]")) {
    if (existing.parentNode) parents.add(existing.parentNode);
    existing.remove();
    changed = true;
  }
  for (const parent of parents) parent.normalize();
  if (!view || view.CSS?.supports?.("hanging-punctuation", "allow-end")) return changed;
  if (content.clientWidth <= 0 || content.clientHeight <= 0) return changed;

  const contentStyle = view.getComputedStyle(content);
  const padRight = parseFloat(contentStyle.paddingRight) || 0;
  const innerRight = content.getBoundingClientRect().right - padRight;
  const range = doc.createRange();

  for (const child of content.children) {
    if (!(child instanceof view.HTMLElement) || child.tagName !== "P") continue;
    const pStyle = view.getComputedStyle(child);
    if (pStyle.getPropertyValue("--office-hanging-punctuation").trim() !== "allow-end") continue;
    if (pStyle.whiteSpace === "pre" || pStyle.whiteSpace === "nowrap") continue;
    if (!["left", "start", "justify"].includes(pStyle.textAlign)) continue;
    const lineEstimate = parseFloat(pStyle.lineHeight) || (parseFloat(pStyle.fontSize) || 16) * 1.2;
    if (child.offsetHeight <= lineEstimate * 1.35) continue;
    const textContent = child.textContent ?? "";
    let hasCandidate = false;
    for (const ch of textContent) if (HANGING_PUNCT.has(ch)) { hasCandidate = true; break; }
    if (!hasCandidate) continue;

    let chars = collectChars(child);
    let checks = 0;
    for (let k = 2; k < chars.length && checks < 256; k++) {
      const punct = chars[k]!;
      if (!HANGING_PUNCT.has(punct.ch)) continue;
      const pulled = chars[k - 1]!, prior = chars[k - 2]!;
      if (/\s/.test(pulled.ch) || HANGING_PUNCT.has(pulled.ch)) continue;
      checks++;
      const rP = rectOf(range, punct), rC = rectOf(range, pulled), rB = rectOf(range, prior);
      if (rP.width <= 0.1 || rC.width <= 0.1 || rB.width <= 0.1) continue;
      if (!sameLine(rC, rP) || sameLine(rB, rC) || rC.top <= rB.top + 1) continue;
      if (rB.right + rC.width > innerRight + 0.5) continue;
      if (rB.right + rC.width + rP.width <= innerRight + 0.5) continue;

      const owner = punct.node.parentElement;
      if (!owner) continue;
      // ECMA-376 §21.1.2.3.9: omitted a:rPr/@lang defaults to en-US. Runs with
      // explicit East Asian lang hang closing punctuation at the line end;
      // default/Latin runs allow a break before closing CJK punctuation instead
      // of pulling the preceding ideograph onto the next line.
      const langHost = owner.closest<HTMLElement>("[data-node-id]");
      const explicitLang = langHost?.getAttribute("lang")
        ?? langHost?.closest<HTMLElement>("p[data-node-id]")?.getAttribute("lang")
        ?? "";
      const eastAsianLang = /^(?:zh|ja|ko)\b/i.test(explicitLang);

      if (!eastAsianLang) {
        const target = punct.offset > 0 ? punct.node.splitText(punct.offset) : punct.node;
        const wbr = markViewNode(doc.createElement("wbr"), "remove");
        wbr.dataset.sdkHangingBreak = "true";
        target.before(wbr);
        const updated = collectChars(child);
        const afterC = updated[k - 1] ? rectOf(range, updated[k - 1]!) : undefined;
        const afterP = updated[k] ? rectOf(range, updated[k]!) : undefined;
        const ok = !!afterC && !!afterP && sameLine(rB, afterC) && !sameLine(afterC, afterP);
        if (!ok) {
          wbr.remove();
          owner.normalize();
          chars = collectChars(child);
        } else {
          changed = true;
          chars = updated;
        }
        continue;
      }

      const overflow = Math.min(
        rP.width,
        Math.max(0.5, Math.ceil((rB.right + rC.width + rP.width - innerRight + 0.75) * 100) / 100),
      );
      const target = punct.offset > 0 ? punct.node.splitText(punct.offset) : punct.node;
      if (target.data.length > punct.ch.length) target.splitText(punct.ch.length);
      const span = markViewNode(doc.createElement("span"), "unwrap");
      span.dataset.sdkHangingPunct = "true";
      span.style.display = "inline-block";
      span.style.marginInlineEnd = -overflow + "px";
      span.style.whiteSpace = "nowrap";
      span.style.textIndent = "0";
      span.style.textAlign = "start";
      if (owner.style.backgroundColor) span.style.backgroundColor = owner.style.backgroundColor;
      target.replaceWith(span);
      span.append(target);

      const updated = collectChars(child);
      const afterC = updated[k - 1] ? rectOf(range, updated[k - 1]!) : undefined;
      const afterP = updated[k] ? rectOf(range, updated[k]!) : undefined;
      const afterNext = updated[k + 1] ? rectOf(range, updated[k + 1]!) : undefined;
      const ok = !!afterC && !!afterP && sameLine(rB, afterC) && sameLine(afterC, afterP)
        && (!afterNext || afterNext.width <= 0.1 || !sameLine(afterP, afterNext));
      if (!ok) {
        span.replaceWith(...span.childNodes);
        owner.normalize();
        chars = collectChars(child);
      } else {
        changed = true;
        chars = updated;
      }
    }
  }
  return changed;
}

/** Right-edge horizontal overhang caused solely by hanging punctuation spans. */
export function hangingPunctuationOverhang(content: HTMLElement): number {
  const spans = content.querySelectorAll<HTMLElement>("[data-sdk-hanging-punct]");
  if (!spans.length) return 0;
  const right = content.getBoundingClientRect().right;
  let overhang = 0;
  for (const span of spans) overhang = Math.max(overhang, span.getBoundingClientRect().right - right);
  return overhang;
}
