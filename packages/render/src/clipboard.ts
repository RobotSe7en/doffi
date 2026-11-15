import { SourceProjection } from "./projection.js";
import type{ContentDocumentModel}from'@doffi/core';
import{copyWordClipboardFormat}from'./word-clipboard-format.js';
import{copyWordClipboardImages}from'./word-clipboard-images.js';
export function sourceSelectionText(root: Node): string {
  interface Part { text: string; boundary: boolean }
  const visit = (node: Node): Part => {
    if (node.nodeType === 3) return { text: node.textContent ?? "", boundary: false };
    const element = node.nodeType === 1 ? node as Element : undefined;
    if (element?.localName === "br") return { text: "\n", boundary: false };
    if (element?.hasAttribute("data-word-tab")) return { text: "\t", boundary: false };
    if (element?.localName === "img") return { text: "\ufffc", boundary: false };
    const children = element?.localName === "tr" ? [...element.children].filter((child) => child.matches("td,th"))
      : element?.matches("table,thead,tbody,tfoot") ? [...element.children].filter((child) => child.matches("thead,tbody,tfoot,tr")) : [...node.childNodes];
    const parts = children.map(visit);
    if (element?.localName === "tr") return { text: parts.map((part) => part.boundary ? part.text.slice(0, -1) : part.text).join("\t") + "\n", boundary: true };
    let text = parts.map((part) => part.text).join(""), boundary = parts.at(-1)?.boundary ?? false;
    if (element?.matches("p,h1,h2,h3,h4,h5,h6")) { text += "\n"; boundary = true; }
    else if (element?.matches("div,section,article,li,td,th,table,ul,ol") && !boundary) { text += "\n"; boundary = true; }
    return { text, boundary };
  };
  const value = visit(root); return value.boundary ? value.text.slice(0, -1) : value.text;
}
export function stripClipboardIdentities(root: ParentNode): void {
  for (const node of root.querySelectorAll("[data-node-id],[id],[data-word-field],[data-word-field-end],[data-word-field-result]"))
    for (const name of ["data-node-id", "id", "data-word-field", "data-word-field-end", "data-word-field-result"]) node.removeAttribute(name);
}

/** Copy the authored selection, excluding layout spacers, font substitutions and
 * repeated stories. This never enables input or creates an editing session. */
export function copyPreviewSelection(event: ClipboardEvent, root: HTMLElement, projection: SourceProjection,model?:ContentDocumentModel): void {
  if (event.defaultPrevented || !event.clipboardData) return;
  const selection = root.ownerDocument.getSelection();
  if (!selection?.rangeCount) return;
  const range = selection.getRangeAt(0);
  if (range.collapsed || !root.contains(range.startContainer) || !root.contains(range.endContainer)) return;
  let fragment: Node = projection.serializeRange(range), ancestor = range.commonAncestorContainer;
  if (ancestor.nodeType !== 1) ancestor = ancestor.parentNode!;
  while (ancestor !== root && root.contains(ancestor)) {
    if ((ancestor as Element).matches("p,h1,h2,h3,h4,h5,h6,li,td,th,section,article,div,ul,ol")) break;
    const shell = projection.shell(ancestor as Element); if (shell) { shell.append(fragment); fragment = shell; }
    if ((ancestor as Element).matches("table")) break;
    ancestor = ancestor.parentNode!;
  }
  const holder = root.ownerDocument.createElement("div"); holder.append(fragment);
  event.preventDefault();if(model?.kind==='document'){copyWordClipboardImages(holder,model);copyWordClipboardFormat(holder,model,(ancestor as HTMLElement).dataset.nodeId??root.dataset.nodeId!);}
  stripClipboardIdentities(holder);
  event.preventDefault(); event.clipboardData.setData("text/html", holder.innerHTML);
  event.clipboardData.setData("text/plain", sourceSelectionText(holder));
}
