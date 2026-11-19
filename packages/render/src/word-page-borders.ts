import { wordColorLuminance } from "@doffi/core";
import type { DeepReadonly, WordPageBorders, WordMetadata } from "@doffi/core";
import type { RenderedPage } from "./contracts.js";
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
        dotted: "dotted",
        dashed: "dashed",
        dashSmallGap: "dashed",
      };
    const horizontal = side === "top" || side === "bottom";
    node.style.cssText = `position:absolute;pointer-events:none;box-sizing:border-box;left:${left}pt;top:${top}pt;width:${page.setup.width - left - right}pt;height:${page.setup.height - top - bottom}pt;border-${side}:${border.width}pt ${styles[border.style]} #${color}`;
    if (border.shadow) node.style.filter = "drop-shadow(2pt 2pt 0 #0005)";
    if (border.frame) node.style.borderColor = "#" + color;
    if (horizontal) {
      node.style.height = "0";
      if (side === "bottom") node.style.top = page.setup.height - bottom - border.width + "pt";
    } else {
      node.style.width = "0";
      if (side === "right") node.style.left = page.setup.width - right - border.width + "pt";
    }
    if (borders.zOrder === "front") {
      node.style.left = page.rect.x + parseFloat(node.style.left) + "pt";
      node.style.top = page.rect.y + parseFloat(node.style.top) + "pt";
      front.append(node);
    } else paper.append(node);
    count++;
  }
  return count;
}
