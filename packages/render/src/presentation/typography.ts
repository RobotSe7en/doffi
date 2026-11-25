import { property } from "./appearance.js";
import { ownViewStyle } from "../view-nodes.js";

interface NormalizedMetrics {
  readonly asc: number;
  readonly desc: number;
}

const FONT_PROBE_SIZE = 100;

/**
 * View-only DrawingML typography refinements applied after font resolution and
 * auto-fit sizing settle:
 * 1. Proportional first-line baseline and bottom-anchored negative half-leading
 *    alignment matching PowerPoint's line-box geometry.
 * 2. Cross-platform consistent synthetic bold for single-weight loaded web fonts.
 */
export class PresentationTypography {
  private context: CanvasRenderingContext2D | null | undefined;
  private readonly metrics = new Map<string, NormalizedMetrics | null>();

  /** DrawingML measures a nonempty paragraph from its actual runs. An inherited
   * larger paragraph font must not add a CSS strut to every small-text line. */
  prepareParagraph(paragraph: HTMLElement): void {
    const doc = paragraph.ownerDocument, view = doc.defaultView!;
    const walker = doc.createTreeWalker(paragraph, view.NodeFilter.SHOW_TEXT);
    let largest = 0;
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (!node.textContent) continue;
      const size = parseFloat(view.getComputedStyle(node.parentElement!).fontSize);
      if (Number.isFinite(size)) largest = Math.max(largest, size);
    }
    const inherited = parseFloat(view.getComputedStyle(paragraph).fontSize);
    if (!largest || Math.abs(largest - inherited) < 0.01) return;
    // Freeze effective run sizes before changing their ancestor, including
    // relative sizes and nested formatting. Projection restores authored CSS.
    const runs = [...paragraph.querySelectorAll<HTMLElement>("*")].map(node =>
      ({ node, size: view.getComputedStyle(node).fontSize }));
    for (const { node, size } of runs) { ownViewStyle(node, "font-size"); node.style.fontSize = size; }
    ownViewStyle(paragraph, "font-size"); paragraph.style.fontSize = largest + "px";
  }

  private measureMetrics(document: Document, style: CSSStyleDeclaration): NormalizedMetrics | null {
    const key = `${style.fontStyle}|${style.fontWeight}|${style.fontFamily}`;
    if (this.metrics.has(key)) return this.metrics.get(key)!;
    if (this.context === undefined) {
      const canvas = document.createElement("canvas");
      this.context = canvas.getContext("2d");
    }
    const ctx = this.context;
    if (!ctx) {
      this.metrics.set(key, null);
      return null;
    }
    ctx.font = `${style.fontStyle} ${style.fontWeight} ${FONT_PROBE_SIZE}px ${style.fontFamily}`;
    const measured = ctx.measureText("汉Hg");
    const asc = measured.fontBoundingBoxAscent / FONT_PROBE_SIZE;
    const desc = measured.fontBoundingBoxDescent / FONT_PROBE_SIZE;
    const result = Number.isFinite(asc) && Number.isFinite(desc) && asc > 0 && desc >= 0 && asc + desc <= 3
      ? { asc, desc }
      : null;
    this.metrics.set(key, result);
    return result;
  }

  alignTextBox(content: HTMLElement, anchor: string, rotated: boolean): void {
    content.style.marginTop = "";
    if (rotated || (anchor !== "t" && anchor !== "b")) return;
    const doc = content.ownerDocument, view = doc.defaultView;
    if (!view) return;
    const paragraphs = [...content.children].filter(
      (child): child is HTMLElement => child instanceof view.HTMLElement && child.tagName === "P",
    );
    if (!paragraphs.length) return;
    const targetP = anchor === "t" ? paragraphs[0]! : paragraphs[paragraphs.length - 1]!;
    const pStyle = view.getComputedStyle(targetP);
    if (anchor === "t" && property(pStyle, "--office-line-spacing-rule") === "exact") return;

    let bestStyle: CSSStyleDeclaration | undefined;
    let bestSize = 0;
    let bestEastAsia = false;
    for (const span of targetP.querySelectorAll<HTMLElement>("[data-office-script]")) {
      if (!span.textContent?.trim()) continue;
      const s = view.getComputedStyle(span);
      const sz = parseFloat(s.fontSize);
      if (!Number.isFinite(sz) || sz <= 0) continue;
      const isEastAsia = span.dataset.officeScript === "east-asia";
      if (sz > bestSize + 0.05 || (Math.abs(sz - bestSize) <= 0.05 && isEastAsia && !bestEastAsia)) {
        bestStyle = s;
        bestSize = sz;
        bestEastAsia = isEastAsia;
      }
    }
    if (!bestStyle || bestSize <= 0) return;
    const linePx = parseFloat(bestStyle.lineHeight);
    if (!Number.isFinite(linePx) || linePx <= 0) return;
    const metrics = this.measureMetrics(doc, bestStyle);
    if (!metrics) return;

    const L = linePx / bestSize;
    const { asc, desc } = metrics;
    let deltaEm = 0;
    if (anchor === "t") {
      const cssBaseline = 0.5 * L + 0.5 * (asc - desc);
      const targetBaseline = Math.max(0.75 * L, 0.75 * (asc + desc), 1.2 - desc);
      deltaEm = targetBaseline - cssBaseline;
    } else {
      const halfLeading = (L - (asc + desc)) / 2;
      if (halfLeading < 0) deltaEm = halfLeading;
    }
    const deltaPx = Math.max(-bestSize * 0.25, Math.min(bestSize * 0.25, deltaEm * bestSize));
    if (Math.abs(deltaPx) >= 0.25) content.style.marginTop = Math.round(deltaPx * 100) / 100 + "px";
  }

  applySyntheticBold(root: HTMLElement): void {
    const doc = root.ownerDocument, view = doc.defaultView;
    if (!view) return;
    const buckets = new Map<string, { regular: boolean; bold: boolean }>();
    for (const face of doc.fonts) {
      if (face.status !== "loaded") continue;
      const family = face.family.replace(/^['"]|['"]$/g, "").trim().toLowerCase();
      if (!family) continue;
      const bucket = buckets.get(family) ?? { regular: false, bold: false };
      const weights = String(face.weight ?? "400")
        .trim()
        .toLowerCase()
        .split(/\s+/)
        .map((token) => token === "bold" ? 700 : token === "normal" ? 400 : Number(token))
        .filter(Number.isFinite);
      const maxW = weights.length ? Math.max(...weights) : 400;
      const minW = weights.length ? Math.min(...weights) : 400;
      if (maxW >= 600) bucket.bold = true;
      if (minW < 600) bucket.regular = true;
      buckets.set(family, bucket);
    }
    const singleWeight = new Set<string>();
    for (const [family, bucket] of buckets) {
      if (bucket.regular && !bucket.bold) singleWeight.add(family);
    }
    for (const span of root.querySelectorAll<HTMLElement>("[data-office-script]")) {
      const style = view.getComputedStyle(span);
      const weight = Number(style.fontWeight);
      const primary = style.fontFamily.split(",")[0]?.trim().replace(/^['"]|['"]$/g, "").toLowerCase() ?? "";
      const active = Number.isFinite(weight) && weight >= 600 && singleWeight.has(primary)
        && (style.textShadow === "none" || span.dataset.sdkSyntheticBold === "true");
      if (active) {
        span.style.setProperty("font-synthesis-weight", "none");
        span.style.textShadow = "0.028em 0 0 currentColor";
        span.dataset.sdkSyntheticBold = "true";
      } else if (span.dataset.sdkSyntheticBold === "true") {
        span.style.removeProperty("font-synthesis-weight");
        span.style.textShadow = "";
        delete span.dataset.sdkSyntheticBold;
      }
    }
  }
}
