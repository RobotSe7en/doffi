import { DocumentError } from "@doffi/core";
import type { DocumentRenderReport } from "@doffi/render";

export interface ZoomAnchor { readonly xFromCenter: number; readonly y: number }
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
