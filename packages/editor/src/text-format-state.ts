export interface TextFormatState {
  readonly enabled: boolean;
  readonly hasSelection?: boolean;
  readonly editingText?: boolean;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly fontSize?: number;
  readonly fontFamily?: string;
  readonly color?: string;
  readonly alignment?: string;
}
function pointSize(element: Element, style: CSSStyleDeclaration, root: HTMLElement): number {
  const measured = parseFloat(style.fontSize) * .75;
  // Firefox quantizes CSS pixels: 15.5pt can compute as 20.65px. Preserve an
  // explicit authored value when it agrees within one layout subpixel.
  for (let node: Element | null = element; node && root.contains(node); node = node.parentElement) {
    const value = (node as HTMLElement).style.fontSize; if (!value) continue;
    const match = /^(\d+(?:\.\d+)?)(pt|px)$/.exec(value);
    if (match) { const exact = Number(match[1]) * (match[2] === 'px' ? .75 : 1); if (Math.abs(measured - exact) < .02) return exact; }
    break;
  }
  return Math.round(measured * 100) / 100;
}
/** Read the current source projection; mixed selections have no single value. */
export function readTextFormat(root: HTMLElement, range: Range | undefined, typing: ReadonlyMap<string, string> = new Map(), ignored: (node: Node) => boolean = () => false): TextFormatState {
  const document = root.ownerDocument, view = document.defaultView!;
  if (!range) return { enabled: true };
  const elements: Element[] = [];
  if (range.collapsed) {
    let node: Node = range.startContainer;
    if (node.nodeType === 1) {
      node = node.childNodes[range.startOffset] ?? node.childNodes[Math.max(0, range.startOffset - 1)] ?? node;
      while (node.nodeType === 1 && node.firstChild) node = node.firstChild;
    }
    const element = node.nodeType === 1 ? node as Element : node.parentElement;
    if (element) elements.push(element);
  } else {
    const ancestor = range.commonAncestorContainer, walker = document.createTreeWalker(ancestor, 4);
    const include = (node: Node) => {
      if (!node.textContent || ignored(node) || !node.parentElement || !range.intersectsNode(node)) return;
      const boundary = document.createRange(); boundary.selectNodeContents(node);
      if (range.compareBoundaryPoints(3, boundary) >= 0 || range.compareBoundaryPoints(1, boundary) <= 0) return;
      elements.push(node.parentElement);
    };
    if (ancestor.nodeType === 3) include(ancestor); else { let node: Node | null; while ((node = walker.nextNode())) include(node); }
  }
  if (!elements.length) return { enabled: true };
  const styles = elements.map(element => view.getComputedStyle(element));
  const same = <Value>(values: Value[]): Value | undefined => values.every(value => value === values[0]) ? values[0] : undefined;
  const uniform = (read: (style: CSSStyleDeclaration) => string) => same(styles.map(read));
  const weight = range.collapsed && typing.has("font-weight") ? typing.get("font-weight") : uniform(style => Number(style.fontWeight) >= 600 || style.fontWeight === "bold" ? "bold" : "normal");
  const italic = range.collapsed && typing.has("font-style") ? typing.get("font-style") : uniform(style => style.fontStyle);
  const decoration = range.collapsed && typing.has("text-decoration") ? typing.get("text-decoration") : same(elements.map(element => {
    for (let node: Element | null = element; node && root.contains(node); node = node.parentElement) if (view.getComputedStyle(node).textDecorationLine.includes("underline")) return "underline";
    return "none";
  }));
  const size = range.collapsed && typing.has("font-size") ? parseFloat(typing.get("font-size")!) : same(elements.map((element, index) => pointSize(element, styles[index]!, root)));
  const fontFamily = range.collapsed && typing.has("font-family") ? typing.get("font-family") : same(elements.map((element, index) => {
    const style = styles[index]!, scope = element.closest<HTMLElement>("[data-word-script],[data-office-script]"), role = scope?.dataset.wordScript ?? scope?.dataset.officeScript;
    const wordRole = role === "eastAsia" || role === "east-asia" ? "eastAsia" : role === "cs" || role === "complex-script" ? "cs" : role === "h-ansi" || role === "hAnsi" ? "hAnsi" : "ascii";
    const officeRole = wordRole === "eastAsia" ? "east-asia" : wordRole === "cs" ? "complex-script" : wordRole === "hAnsi" ? "h-ansi" : "ascii";
    return style.getPropertyValue("--word-font-" + wordRole).trim() || style.getPropertyValue("--office-font-" + officeRole).trim() || style.fontFamily;
  }));
  const color = range.collapsed && typing.has("color") ? typing.get("color") : uniform(style => style.color), alignment = uniform(style => style.textAlign);
  return { enabled: true, hasSelection: !range.collapsed, ...(weight === undefined ? {} : { bold: weight === "bold" || Number(weight) >= 600 }), ...(italic === undefined ? {} : { italic: italic !== "normal" }), ...(decoration === undefined ? {} : { underline: decoration.includes("underline") }),
    ...(size === undefined ? {} : { fontSize: size }), ...(fontFamily === undefined ? {} : { fontFamily }), ...(color === undefined ? {} : { color }), ...(alignment === undefined ? {} : { alignment }) };
}
