import { DocumentError } from "@doffi/core";
import type { SourceProjection } from "@doffi/render";
/** Paint formatting symbols around the original whitespace; copying, offsets
 * and serialization retain those same characters. */
export function showWordWhitespace(
  root: HTMLElement,
  projection: SourceProjection,
  maximumNodes: number,
): () => void {
  const document = root.ownerDocument,
    walker = document.createTreeWalker(root, 4),
    texts: Text[] = [];
  let count = document.getElementsByTagName("*").length;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (projection.ignored(node) || node.parentElement?.closest("[data-word-tab]")) continue;
    const text = node as Text;
    for (let at = 0; at < text.length; at++)
      if (text.data[at] === " " || text.data[at] === "\u00a0")
        if (++count > maximumNodes)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Formatting marks exceed the render node budget.",
          );
    texts.push(text);
  }
  const wrappers: HTMLElement[] = [];
  for (const text of texts) {
    if (!/[ \u00a0]/.test(text.data)) continue;
    const fragment = document.createDocumentFragment();
    for (const part of text.data.split(/([ \u00a0])/)) {
      if (!part) continue;
      if (part === " " || part === "\u00a0") {
        const span = document.createElement("span");
        projection.markWrapper(span);
        span.className = "sdk-space-mark";
        span.dataset.sdkSpaceMark = part === " " ? "·" : "°";
        span.textContent = part;
        fragment.append(span);
        wrappers.push(span);
      } else fragment.append(document.createTextNode(part));
    }
    text.replaceWith(fragment);
  }
  return () => {
    for (const wrapper of wrappers)
      if (wrapper.parentNode) wrapper.replaceWith(...wrapper.childNodes);
  };
}
