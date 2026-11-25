import { DocumentError } from "@doffi/core";
import type { RenderedPage, RenderLimits } from "./contracts.js";
import { safeStyleText } from "./preview-package.js";

function copyHead(document: Document, output: Document): void {
  output.head.replaceWith(output.importNode(document.head, true));
  const source = [...document.head.querySelectorAll("style")],
    target = [...output.head.querySelectorAll("style")];
  // CSSOM insertion does not update style.textContent. Materialize derived
  // inheritance rules when moving a view into the isolated print document.
  for (const [index, style] of source.entries())
    if (style.sheet)
      target[index]!.textContent = safeStyleText(
        [...style.sheet.cssRules].map((rule) => rule.cssText).join("\n"),
      );
}

/** A thumbnail uses one composed slide and the shared style context. */
export function presentationThumbnailSource(document: Document): (slide: HTMLElement) => string {
  const output = document.implementation.createHTMLDocument("");
  copyHead(document, output);
  output.head.querySelector("[data-sdk-slide-visibility]")?.remove();
  const style = output.createElement("style");
  style.textContent = "html,body{width:100%;height:100%;overflow:hidden;margin:0!important;padding:0!important}article,article>[data-office-kind=slide]{margin:0!important}.sdk-ppt-hidden-label{display:none!important}";
  output.head.append(style);
  const head = output.head.outerHTML;
  return slide => "<!doctype html><html>" + head + '<body><article data-office-kind="presentation">' + slide.outerHTML + "</article></body></html>";
}

/** Clip the measured canvas per page. Chromium retains searchable text and culls
 * content outside each page; the editable source and its flow remain unchanged. */
export function createPrintDocument(
  document: Document,
  root: HTMLElement,
  pages: readonly RenderedPage[],
  limits: Readonly<RenderLimits>,
): Document {
  if (root.dataset.officeKind === "presentation")
    return printPresentation(document, root, pages, limits);
  const article = root.closest<HTMLElement>('[data-office-kind="document"]')!;
  const nodes =
    article.querySelectorAll("*").length +
    1 +
    document.querySelectorAll(".sdk-paper *,.sdk-page-border-layer *").length;
  if (pages.length * nodes > limits.maxPrintNodes)
    throw new DocumentError("ResourceLimitExceeded", "Print projection exceeds its node budget.");
  const output = document.implementation.createHTMLDocument("");
  copyHead(document, output);
  output.body.replaceChildren();
  const style = output.createElement("style");
  style.textContent =
    "html,body{margin:0!important;padding:0!important;background:white!important}.sdk-print-page{position:relative;overflow:hidden;break-after:page;box-sizing:border-box}.sdk-print-page:last-child{break-after:auto}.sdk-page-number{display:none!important}";
  const surface = root.getBoundingClientRect(),
    origin = article.getBoundingClientRect();
  for (const page of pages) {
    style.textContent += `@page sdk-page-${page.index}{size:${page.setup.width}pt ${page.setup.height}pt;margin:0}`;
    const paper = output.createElement("div");
    paper.className = "sdk-print-page";
    paper.style.cssText = `width:${page.setup.width}pt;height:${page.setup.height}pt;page:sdk-page-${page.index}`;
    const background = document.querySelector<HTMLElement>(
      '.sdk-paper[data-page-id="' + CSS.escape(page.pageId) + '"]',
    );
    if (background) {
      const clone = output.importNode(background, true);
      clone.removeAttribute("aria-hidden");
      clone.style.cssText = `position:absolute;left:0;top:0;width:${page.setup.width}pt;height:${page.setup.height}pt;box-shadow:none;background:white`;
      paper.append(clone);
    }
    const content = output.importNode(article, true);
    content.style.cssText += `;position:absolute;left:${-(page.rect.x + ((surface.left - origin.left) * 72) / 96)}pt;top:${-(page.rect.y + ((surface.top - origin.top) * 72) / 96)}pt;margin:0;`;
    paper.append(content);
    for (const border of document.querySelectorAll<HTMLElement>(
      '.sdk-page-border-layer [data-page-id="' + CSS.escape(page.pageId) + '"]',
    )) {
      const clone = output.importNode(border, true);
      clone.style.left = parseFloat(border.style.left) - page.rect.x + "pt";
      clone.style.top = parseFloat(border.style.top) - page.rect.y + "pt";
      clone.style.zIndex = "3";
      paper.append(clone);
    }
    output.body.append(paper);
  }
  output.head.append(style);
  return output;
}

function printPresentation(
  document: Document,
  root: HTMLElement,
  pages: readonly RenderedPage[],
  limits: Readonly<RenderLimits>,
): Document {
  const slides = [...root.children].filter(
    (node): node is HTMLElement => node.getAttribute("data-office-kind") === "slide",
  );
  if (slides.length !== pages.length)
    throw new DocumentError("InvalidOperation", "The print projection has stale slide geometry.");
  const count = slides.reduce((sum, node) => sum + node.querySelectorAll("*").length + 2, 0);
  if (count > limits.maxPrintNodes)
    throw new DocumentError("ResourceLimitExceeded", "Print projection exceeds its node budget.");
  const output = document.implementation.createHTMLDocument("");
  copyHead(document, output);
  output.head.querySelector("[data-sdk-slide-visibility]")?.remove();
  const article = output.importNode(root, false);
  article.style.cssText = "margin:0;padding:0;width:auto";
  output.body.replaceChildren(article);
  const style = output.createElement("style");
  style.textContent =
    "html,body{margin:0!important;padding:0!important;background:white!important}[data-office-kind=slide]{margin:0!important;break-after:page}[data-office-kind=slide]:last-child{break-after:auto}.sdk-ppt-hidden-label{display:none!important}";
  for (const [index, slide] of slides.entries()) {
    const page = pages[index]!,
      clone = output.importNode(slide, true);
    clone.style.page = "sdk-ppt-" + index;
    style.textContent += `@page sdk-ppt-${index}{size:${page.setup.width}pt ${page.setup.height}pt;margin:0}`;
    article.append(clone);
  }
  output.head.append(style);
  return output;
}
