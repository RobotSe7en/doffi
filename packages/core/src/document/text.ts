import type { DefaultTreeAdapterTypes as Tree } from "parse5";

/** UTF-16 offsets may not split a surrogate pair or detach a variation selector. */
export function isTextBoundary(text: string, offset: number): boolean {
  if (typeof text !== "string" || !Number.isSafeInteger(offset) || offset < 0 || offset > text.length) return false;
  const before = text.charCodeAt(offset - 1);
  const after = text.charCodeAt(offset);
  const code = text.codePointAt(offset);
  return !(before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff)
    && !(offset > 0 && code !== undefined && (code >= 0xfe00 && code <= 0xfe0f || code >= 0xe0100 && code <= 0xe01ef));
}

/** Positions in the same normalized text used by DocumentNode and DOM selections. */
export function textPositions(root: Tree.Node): Map<Tree.Node, { from: number; to: number }> {
  const positions = new Map<Tree.Node, { from: number; to: number }>();
  let offset = 0;
  const visit = (node: Tree.Node): void => {
    const from = offset;
    if (node.nodeName === "#text" && "value" in node) offset += node.value.length;
    else if ("tagName" in node && ["br", "img"].includes(node.tagName)) offset++;
    else if ("childNodes" in node) for (const child of node.childNodes) visit(child);
    positions.set(node, { from, to: offset });
  };
  visit(root);
  return positions;
}
