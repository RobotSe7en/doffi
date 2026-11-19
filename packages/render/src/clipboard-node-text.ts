export function clipboardNodeText(node: Node): string {
  if (node.nodeType === 3) return node.textContent ?? "";
  if (node.nodeType === 1) {
    const element = node as Element;
    if (element.localName === "img") return "\ufffc";
    if (element.localName === "br") return "\n";
  }
  return [...node.childNodes].map(clipboardNodeText).join("");
}
