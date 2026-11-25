import { DocumentError } from "@doffi/core";
import { markViewNode, ownViewStyle } from "../view-nodes.js";
import { property, resolvedColor } from "./appearance.js";
import { resolvePreviewDefaultFont } from "../fonts.js";

/** Native group scaling changes object rectangles; it does not mirror glyphs
 * or rewrite their point sizes. Resolve the cumulative linear view transform. */
export function textTransform(node: HTMLElement, rect: { x: number; y: number; width: number; height: number }): string {
  const view = node.ownerDocument.defaultView!; let matrix = new view.DOMMatrix(), flipH = false;
  for (let current: HTMLElement | null = node; current && current.dataset.officeKind !== "slide"; current = current.parentElement) {
    const transform = view.getComputedStyle(current).transform;
    if (transform !== "none") matrix = new view.DOMMatrix(transform).multiply(matrix);
    if (current.dataset.officeFlipH === "true") flipH = !flipH;
  }
  const scaleX = Math.hypot(matrix.a, matrix.b), scaleY = Math.hypot(matrix.c, matrix.d);
  if (!scaleX || !scaleY) throw new DocumentError("InvalidDocument", "A text instance has a singular group transform.");
  const angle = Math.atan2(matrix.b * (flipH ? -1 : 1), matrix.a * (flipH ? -1 : 1));
  const rotation = new view.DOMMatrix([Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0]);
  matrix.e = 0; matrix.f = 0;
  const local = matrix.inverse().multiply(rotation), width = rect.width * scaleX, height = rect.height * scaleY;
  return `left:${rect.x + (rect.width - width) / 2}pt;top:${rect.y + (rect.height - height) / 2}pt;width:${width}pt;height:${height}pt;transform-origin:50% 50%;transform:matrix(${local.a},${local.b},${local.c},${local.d},0,0)`;
}
function numberLabel(value: number, format: string): string {
  let label: string;
  if (format.startsWith("alpha")) {
    label = ""; for (let n = value; n > 0; n = Math.floor((n - 1) / 26)) label = String.fromCharCode(65 + (n - 1) % 26) + label;
    if (format.startsWith("alphaLc")) label = label.toLowerCase();
  } else if (format.startsWith("roman")) {
    if (value > 1_000_000) throw new DocumentError("ResourceLimitExceeded", "A presentation list label exceeds its budget.");
    label = ""; let remaining = value;
    for (const [amount, token] of [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]] as const) while (remaining >= amount) { label += token; remaining -= amount; }
    if (format.startsWith("romanLc")) label = label.toLowerCase();
  } else label = String(value);
  return format.endsWith("ParenBoth") ? "(" + label + ")" : format.endsWith("ParenR") ? label + ")" : format.endsWith("Plain") ? label : label + ".";
}
export function decorateLists(root: HTMLElement, account: (n: number) => void): void {
  for (const box of root.querySelectorAll<HTMLElement>(".sdk-ppt-text,.sdk-ppt-rich-story")) {
    decorateDrawingLists(box, account);
  }
}

/** Lists shared by independently anchored DrawingML text bodies. */
export function decorateDrawingLists(box: HTMLElement, account: (n: number) => void): void {
    const document = box.ownerDocument;
    const counters = new Map<string, number>();
    const bulletFonts = new Map<string, string>();
    for (const paragraph of box.querySelectorAll<HTMLElement>("p")) {
      const s = document.defaultView!.getComputedStyle(paragraph), kind = paragraph.hasAttribute("data-office-list-none") ? "none" : paragraph.dataset.officeBullet ? "bullet" : paragraph.dataset.officeNumbering ? "number" : property(s, "--office-list-kind");
      if (kind !== "bullet" && kind !== "number") { counters.clear(); continue; }
      account(2); // Marker element and its generated label text.
      const marker = markViewNode(document.createElement("span")); marker.className = "sdk-ppt-marker";
      const character = paragraph.dataset.officeBullet ?? property(s, "--office-list-character", "•").replace(/^"|"$/g, "");
      const format = paragraph.dataset.officeNumbering ?? property(s, "--office-list-numbering", "arabicPeriod"), key = (paragraph.dataset.officeLevel ?? "0") + ":" + format;
      const value = counters.get(key) ?? Number(paragraph.dataset.officeNumberingStart ?? property(s, "--office-list-start", "1")); counters.set(key, value + 1);
      marker.textContent = kind === "bullet" ? character : numberLabel(value, format);
      const sourceIndent = parseFloat(s.textIndent), sourceMargin = parseFloat(s.marginLeft);
      const indent = Number.isFinite(sourceIndent) ? sourceIndent : -24, margin = Number.isFinite(sourceMargin) ? sourceMargin : 24;
      const sourceFont = property(s, "--office-bullet-font", s.fontFamily);
      let bulletFont = sourceFont;
      if (character === "▪") {
        // macOS PowerPoint uses its color-emoji fallback for this custom square
        // even with an Arial bullet declaration. Prefer that installed face;
        // other platforms keep their available source font. Native paint and
        // character declarations remain unchanged in the document package.
        bulletFont = bulletFonts.get(sourceFont) ?? resolvePreviewDefaultFont(document, '"Apple Color Emoji",' + sourceFont);
        bulletFonts.set(sourceFont, bulletFont);
      }
      marker.style.cssText = `position:absolute;pointer-events:none;white-space:pre;left:${indent}px;width:${Math.abs(indent)}px;font-family:${bulletFont};font-size:${property(s, "--office-bullet-size", "1em")};color:${resolvedColor(document, s, "--office-bullet", s.color)}`;
      ownViewStyle(paragraph, "position", "text-indent", "margin-left"); paragraph.style.position = "relative"; paragraph.style.textIndent = "0"; paragraph.style.marginLeft = margin + "px"; paragraph.prepend(marker);
    }
}
