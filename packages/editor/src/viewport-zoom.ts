import { DocumentError } from "@doffi/core";
import type { DocumentRenderReport } from "@doffi/render";

export type WordPageZoomMode = "singlePage" | "multiplePages" | "pageWidth";
export interface ZoomAnchor { readonly xFromCenter: number; readonly y: number }

export function computeWordFitZoom(
  stage: HTMLElement,
  pageWidthPt: number,
  pageHeightPt: number,
  mode: WordPageZoomMode,
): number {
  const pageW = Math.max(1, (pageWidthPt * 96) / 72),
    pageH = Math.max(1, (pageHeightPt * 96) / 72);
  const availW = Math.max(240, stage.clientWidth - 48),
    availH = Math.max(200, stage.clientHeight - 36);
  const clamp = (value: number) => Math.max(0.25, Math.min(2, Math.floor(value * 100) / 100));
  if (mode === "pageWidth") return clamp(availW / pageW);
  const single = Math.min(availW / pageW, availH / pageH);
  if (mode === "singlePage") return clamp(single);
  const multiple = Math.min((availW - 24) / (2 * pageW), availH / (2 * pageH + 24));
  return clamp(Math.min(single * 0.85, multiple));
}
/** Preserve the visible document point, including when a narrow page is centered
 * inside a wider viewport. Geometry reads stay bounded to these two elements. */
export function captureZoomAnchor(stage: HTMLElement, plane: HTMLElement, zoom: number): ZoomAnchor | undefined {
  if (!stage.clientWidth || !plane.offsetWidth || !plane.offsetHeight) return;
  const viewport = stage.getBoundingClientRect(), content = plane.getBoundingClientRect();
  return { xFromCenter: (viewport.left + stage.clientLeft + stage.clientWidth / 2 - content.left - plane.offsetWidth / 2) / zoom,
    y: (viewport.top + stage.clientTop + stage.clientHeight / 2 - content.top) / zoom };
}
export function restoreZoomAnchor(stage: HTMLElement, plane: HTMLElement, zoom: number, anchor: ZoomAnchor | undefined): void {
  if (!stage.clientWidth) return;
  if (!anchor) { stage.scrollLeft = Math.max(0, (plane.offsetWidth - stage.clientWidth) / 2); return; }
  const viewport = stage.getBoundingClientRect(), content = plane.getBoundingClientRect();
  stage.scrollLeft += content.left + plane.offsetWidth / 2 + anchor.xFromCenter * zoom - viewport.left - stage.clientLeft - stage.clientWidth / 2;
  stage.scrollTop += content.top + anchor.y * zoom - viewport.top - stage.clientTop - stage.clientHeight / 2;
}

export function scrollToWordPage(stage: HTMLElement, frame: HTMLIFrameElement, report: DocumentRenderReport, index: number, zoom: number): void {
  const page = report.layout === "pages" && Number.isInteger(index) && report.pages[index - 1];
  const paper = page && frame.contentDocument?.querySelector<HTMLElement>(`.sdk-paper[data-page-id="${CSS.escape(page.pageId)}"]`);
  if (!paper) throw new DocumentError("InvalidOperation", "页码超出当前页面视图。");
  const viewport = stage.getBoundingClientRect();
  stage.scrollTop += frame.getBoundingClientRect().top + paper.getBoundingClientRect().top * zoom
    - viewport.top - stage.clientTop - 12;
}
