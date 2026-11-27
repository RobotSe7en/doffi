import { WORD_ART_PAGE_BORDER_STYLES, wordColorLuminance } from "@doffi/core";
import type { DeepReadonly, WordPageBorders, WordMetadata } from "@doffi/core";
import type { RenderedPage } from "./contracts.js";

const ART_STYLES = new Set<string>(WORD_ART_PAGE_BORDER_STYLES);

function artGlyphSvg(style: string, color: string): string {
  switch (style) {
    case "apples":
      return `<circle cx="10" cy="11" r="6.5" fill="#d63031"/><path d="M10 4.5 C10 2.5 12 2 13 2.5" stroke="#5b3a29" stroke-width="1.5" fill="none"/><ellipse cx="12.5" cy="4.2" rx="2.5" ry="1.3" fill="#2ecc71"/>`;
    case "stars":
      return `<polygon points="10,2 12.4,7.4 18.2,7.8 13.7,11.6 15.1,17.4 10,14.2 4.9,17.4 6.3,11.6 1.8,7.8 7.6,7.4" fill="#f1c40f" stroke="#d4ac0d" stroke-width="0.8"/>`;
    case "starsBlack":
      return `<polygon points="10,2 12.4,7.4 18.2,7.8 13.7,11.6 15.1,17.4 10,14.2 4.9,17.4 6.3,11.6 1.8,7.8 7.6,7.4" fill="${color}"/>`;
    case "hearts":
      return `<path d="M10 16.5 C4 12.5 2.5 9.2 2.5 6.8 C2.5 4.4 4.3 3 6.5 3 C8.1 3 9.4 4 10 5.3 C10.6 4 11.9 3 13.5 3 C15.7 3 17.5 4.4 17.5 6.8 C17.5 9.2 16 12.5 10 16.5 Z" fill="#e74c3c"/>`;
    case "trees":
    case "christmasTree":
      return `<polygon points="10,2 16,10 13,10 17,15.5 3,15.5 7,10 4,10" fill="#27ae60"/><rect x="8.8" y="15.5" width="2.4" height="3" fill="#795548"/>`;
    case "checkered":
      return `<rect x="0" y="0" width="10" height="10" fill="${color}"/><rect x="10" y="10" width="10" height="10" fill="${color}"/>`;
    case "classicalWave":
      return `<path d="M0 10 Q5 3 10 10 T20 10" fill="none" stroke="${color}" stroke-width="2.2"/><path d="M0 15 Q5 8 10 15 T20 15" fill="none" stroke="${color}" stroke-width="1.2"/>`;
    case "diamondsGray":
      return `<polygon points="10,2 17,10 10,18 3,10" fill="#95a5a6" stroke="${color}" stroke-width="1"/>`;
    case "doubleDiamonds":
      return `<polygon points="10,2 17,10 10,18 3,10" fill="none" stroke="${color}" stroke-width="1.5"/><polygon points="10,5.5 14,10 10,14.5 6,10" fill="${color}"/>`;
    case "flowersBlockPrint":
      return `<circle cx="10" cy="5.5" r="3" fill="${color}"/><circle cx="14.5" cy="10" r="3" fill="${color}"/><circle cx="10" cy="14.5" r="3" fill="${color}"/><circle cx="5.5" cy="10" r="3" fill="${color}"/><circle cx="10" cy="10" r="2.2" fill="#f1c40f"/>`;
    case "basicBlackSquares":
      return `<rect x="4" y="4" width="12" height="12" fill="${color}"/>`;
    case "basicBlackDots":
      return `<circle cx="10" cy="10" r="5" fill="${color}"/>`;
    case "mapleLeaf":
      return `<path d="M10 2 L12 7 L17 6 L14 10 L17 13 L12 13 L11 18 L9 18 L8 13 L3 13 L6 10 L3 6 L8 7 Z" fill="#d35400"/>`;
    case "palmsColor":
      return `<path d="M10 18 L10 9" stroke="#8d6e63" stroke-width="2"/><path d="M10 9 Q4 6 2 10 M10 9 Q16 6 18 10 M10 9 Q6 3 4 4 M10 9 Q14 3 16 4" stroke="#2ecc71" stroke-width="2" fill="none"/>`;
    case "confetti":
      return `<rect x="2" y="3" width="4" height="4" fill="#e74c3c" transform="rotate(20 4 5)"/><circle cx="14" cy="6" r="2.5" fill="#3498db"/><polygon points="7,12 11,17 4,17" fill="#f1c40f"/><rect x="13" y="13" width="4" height="3" fill="#2ecc71"/>`;
    case "wave":
      return `<path d="M0 10 Q5 4 10 10 T20 10" fill="none" stroke="${color}" stroke-width="2.4"/>`;
    case "doubleWave":
      return `<path d="M0 7 Q5 2 10 7 T20 7 M0 13 Q5 8 10 13 T20 13" fill="none" stroke="${color}" stroke-width="1.8"/>`;
    default:
      return `<rect x="3" y="3" width="14" height="14" fill="${color}"/>`;
  }
}

/** Page decoration stays outside the editable source and shares physical-page
 * geometry with headers, mirrored margins and print projections. */
export function composeWordPageBorders(
  document: Document,
  article: HTMLElement,
  paper: HTMLElement,
  front: HTMLElement,
  page: RenderedPage,
  borders: DeepReadonly<WordPageBorders> | undefined,
  first: boolean,
  word: DeepReadonly<WordMetadata> | undefined,
): number {
  if (
    !borders ||
    (borders.display === "firstPage" && !first) ||
    (borders.display === "notFirstPage" && first)
  )
    return 0;
  const theme = document.defaultView!.getComputedStyle(article),
    roles: Record<string, string> = {
      dark1: "dk1",
      light1: "lt1",
      dark2: "dk2",
      light2: "lt2",
      text1: "dk1",
      background1: "lt1",
      text2: "dk2",
      background2: "lt2",
      hyperlink: "hlink",
      followedHyperlink: "folHlink",
    };
  let count = 0;
  const distances = { ...page.setup.margins };
  if (borders.offsetFrom === "text")
    for (const story of paper.querySelectorAll<HTMLElement>(".sdk-page-story")) {
      const header = story.dataset.officeStory === "header",
        footer = story.dataset.officeStory === "footer";
      if (header && !word?.bordersDoNotSurroundHeader)
        distances.top = Math.min(distances.top, parseFloat(story.style.top) || 0);
      if (footer && !word?.bordersDoNotSurroundFooter)
        distances.bottom = Math.min(distances.bottom, parseFloat(story.style.bottom) || 0);
    }
  const inset = (side: "top" | "left" | "bottom" | "right"): number => {
    const border = borders[side],
      space = border?.space ?? 0;
    return borders.offsetFrom === "page"
      ? space
      : Math.max(0, distances[side] - space - (border?.width ?? 0));
  };
  const top = inset("top"),
    left = inset("left"),
    bottom = inset("bottom"),
    right = inset("right");
  for (const side of ["top", "left", "bottom", "right"] as const) {
    const border = borders[side];
    if (!border || ["nil", "none"].includes(border.style) || !border.width) continue;
    const node = document.createElement("div");
    node.dataset.sdkPageBorder = side;
    node.dataset.sdkPageBorderStyle = border.style;
    node.dataset.pageId = page.pageId;
    node.setAttribute("aria-hidden", "true");
    const raw =
        border.themeColor && border.themeColor !== "none"
          ? theme
              .getPropertyValue("--office-theme-" + (roles[border.themeColor] ?? border.themeColor))
              .trim()
              .replace(/^#/, "")
          : undefined,
      color = wordColorLuminance(
        raw && /^[a-f\d]{6}$/i.test(raw) ? raw : border.color === "auto" ? "000000" : border.color,
        border.themeTint,
        border.themeShade,
      ),
      styles: Record<string, string> = {
        single: "solid",
        thick: "solid",
        double: "double",
        triple: "double",
        dotted: "dotted",
        dashed: "dashed",
        dashSmallGap: "dashed",
        dotDash: "dashed",
        dotDotDash: "dashed",
        threeDEmboss: "ridge",
        threeDEngrave: "groove",
      };
    const horizontal = side === "top" || side === "bottom";
    const isPattern = ART_STYLES.has(border.style) || border.style === "wave" || border.style === "doubleWave";
    if (isPattern) {
      const tile = Math.max(6, Math.min(31, border.width));
      const tilePx = tile * (96 / 72);
      const glyph = artGlyphSvg(border.style, "#" + color);
      const isDiscreteArt = ART_STYLES.has(border.style);
      const spanW = Math.max(tile, page.setup.width - left - right);
      const sideTile = (edge: "top" | "bottom"): number => {
        const b = borders[edge];
        if (!b || ["nil", "none"].includes(b.style) || !b.width) return 0;
        const bPat = ART_STYLES.has(b.style) || b.style === "wave" || b.style === "doubleWave";
        return bPat ? Math.max(6, Math.min(31, b.width)) : b.width;
      };
      const topInset = horizontal ? 0 : sideTile("top");
      const bottomInset = horizontal ? 0 : sideTile("bottom");
      const spanH = Math.max(tile, page.setup.height - top - bottom - topInset - bottomInset);
      const boxLeft = side === "right" ? page.setup.width - right - tile : left;
      const boxTop =
        side === "bottom"
          ? page.setup.height - bottom - tile
          : horizontal
            ? top
            : top + topInset;
      const countAlong = horizontal
        ? Math.max(1, Math.round(spanW / tile))
        : Math.max(1, Math.round(spanH / tile));
      const patW =
        isDiscreteArt && horizontal ? (spanW / countAlong) * (96 / 72) : tilePx;
      const patH =
        isDiscreteArt && !horizontal ? (spanH / countAlong) * (96 / 72) : tilePx;
      const drawW = Math.min(tilePx, patW);
      const drawH = Math.min(tilePx, patH);
      const offX = ((patW - drawW) / 2).toFixed(2);
      const offY = ((patH - drawH) / 2).toFixed(2);
      const patternId = `sdk-pb-${page.pageId.replace(/[^a-zA-Z0-9_-]/g, "")}-${side}`;
      node.style.cssText =
        `position:absolute;pointer-events:none;box-sizing:border-box;overflow:hidden;` +
        `left:${boxLeft}pt;` +
        `top:${boxTop}pt;` +
        `width:${horizontal ? spanW : tile}pt;` +
        `height:${horizontal ? tile : spanH}pt;`;
      const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      svg.setAttribute("width", "100%");
      svg.setAttribute("height", "100%");
      svg.setAttribute("aria-hidden", "true");
      svg.style.cssText = "display:block;width:100%;height:100%";
      svg.innerHTML =
        `<defs><pattern id="${patternId}" width="${patW.toFixed(2)}" height="${patH.toFixed(2)}" patternUnits="userSpaceOnUse">` +
        `<svg x="${offX}" y="${offY}" viewBox="0 0 20 20" width="${drawW.toFixed(2)}" height="${drawH.toFixed(2)}">${glyph}</svg>` +
        `</pattern></defs>` +
        `<rect width="100%" height="100%" fill="url(#${patternId})"/>`;
      node.append(svg);
    } else {
      node.style.cssText = `position:absolute;pointer-events:none;box-sizing:border-box;left:${left}pt;top:${top}pt;width:${page.setup.width - left - right}pt;height:${page.setup.height - top - bottom}pt;border-${side}:${border.width}pt ${styles[border.style] ?? "solid"} #${color}`;
      if (horizontal) {
        node.style.height = "0";
        if (side === "bottom") node.style.top = page.setup.height - bottom - border.width + "pt";
      } else {
        node.style.width = "0";
        if (side === "right") node.style.left = page.setup.width - right - border.width + "pt";
      }
    }
    if (border.shadow) node.style.filter = "drop-shadow(2pt 2pt 0 #0005)";
    if (border.frame) node.style.borderColor = "#" + color;
    if (borders.zOrder === "front") {
      node.style.left = page.rect.x + parseFloat(node.style.left) + "pt";
      node.style.top = page.rect.y + parseFloat(node.style.top) + "pt";
      front.append(node);
    } else paper.append(node);
    count++;
  }
  return count;
}
