export interface TextFormatState {
  readonly enabled: boolean;
  readonly hasSelection?: boolean;
  readonly editingText?: boolean;
  readonly editingInPlace?: boolean;
  readonly drawingKind?: string;
  readonly drawingCount?: number;
  readonly storyKind?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
  readonly doubleStrikethrough?: boolean;
  readonly characterBorder?: boolean;
  readonly inTable?: boolean;
  readonly themeColors?: readonly string[];
  readonly verticalAlign?: string;
  readonly backgroundColor?: string;
  readonly fontSize?: number;
  readonly fontFamily?: string;
  readonly color?: string;
  readonly alignment?: string;
  readonly paragraphStyleId?: string;
  readonly characterProperties?: Readonly<Record<string, string>>;
  readonly paragraphProperties?: Readonly<Record<string, string>>;
  readonly numberPreview?: string;
  readonly numberFormatCode?: string;
  readonly numberFormatBuiltinId?: number;
}
function pointSize(element: Element, style: CSSStyleDeclaration, root: HTMLElement): number {
  const measured = parseFloat(style.fontSize) * 0.75;
  // Firefox quantizes CSS pixels: 15.5pt can compute as 20.65px. Preserve an
  // explicit authored value when it agrees within one layout subpixel.
  for (let node: Element | null = element; node && root.contains(node); node = node.parentElement) {
    const value = (node as HTMLElement).style.fontSize;
    if (!value) continue;
    const match = /^(\d+(?:\.\d+)?)(pt|px)$/.exec(value);
    if (match) {
      const exact = Number(match[1]) * (match[2] === "px" ? 0.75 : 1);
      if (Math.abs(measured - exact) < 0.02) return exact;
    }
    break;
  }
  return Math.round(measured * 100) / 100;
}
/** Read the current source projection; mixed selections have no single value. */
export function readTextFormat(
  root: HTMLElement,
  range: Range | undefined,
  typing: ReadonlyMap<string, string> = new Map(),
  ignored: (node: Node) => boolean = () => false,
): TextFormatState {
  const document = root.ownerDocument,
    view = document.defaultView!;
  if (!range) return { enabled: true };
  const elements: Element[] = [];
  if (range.collapsed) {
    let node: Node = range.startContainer;
    if (node.nodeType === 1) {
      node =
        node.childNodes[range.startOffset] ??
        node.childNodes[Math.max(0, range.startOffset - 1)] ??
        node;
      while (node.nodeType === 1 && node.firstChild) node = node.firstChild;
    }
    const element = node.nodeType === 1 ? (node as Element) : node.parentElement;
    if (element) elements.push(element);
  } else {
    const ancestor = range.commonAncestorContainer,
      walker = document.createTreeWalker(ancestor, 4);
    const include = (node: Node) => {
      if (!node.textContent || ignored(node) || !node.parentElement || !range.intersectsNode(node))
        return;
      const boundary = document.createRange();
      boundary.selectNodeContents(node);
      if (
        range.compareBoundaryPoints(3, boundary) >= 0 ||
        range.compareBoundaryPoints(1, boundary) <= 0
      )
        return;
      elements.push(node.parentElement);
    };
    if (ancestor.nodeType === 3) include(ancestor);
    else {
      let node: Node | null;
      while ((node = walker.nextNode())) include(node);
    }
  }
  if (!elements.length) return { enabled: true };
  const styles = elements.map((element) => view.getComputedStyle(element));
  const same = <Value>(values: Value[]): Value | undefined =>
    values.every((value) => value === values[0]) ? values[0] : undefined;
  const uniform = (read: (style: CSSStyleDeclaration) => string) => same(styles.map(read));
  const weight =
    range.collapsed && typing.has("font-weight")
      ? typing.get("font-weight")
      : uniform((style) =>
          Number(style.fontWeight) >= 600 || style.fontWeight === "bold" ? "bold" : "normal",
        );
  const italic =
    range.collapsed && typing.has("font-style")
      ? typing.get("font-style")
      : uniform((style) => style.fontStyle);
  const decoration =
    range.collapsed && typing.has("text-decoration")
      ? typing.get("text-decoration")
      : same(
          elements.map((element) => {
            for (
              let node: Element | null = element;
              node && root.contains(node);
              node = node.parentElement
            )
              if (view.getComputedStyle(node).textDecorationLine.includes("underline"))
                return "underline";
            return "none";
          }),
        );
  const strike = same(
    elements.map((element) => {
      for (
        let node: Element | null = element;
        node && root.contains(node);
        node = node.parentElement
      )
        if (view.getComputedStyle(node).textDecorationLine.includes("line-through")) return true;
      return false;
    }),
  );
  const verticalAlign =
    range.collapsed && typing.has("vertical-align")
      ? typing.get("vertical-align")
      : uniform((style) => style.verticalAlign);
  const backgroundColor =
    range.collapsed && typing.has("background-color")
      ? typing.get("background-color")
      : uniform((style) => style.backgroundColor);
  const size =
    range.collapsed && typing.has("font-size")
      ? parseFloat(typing.get("font-size")!)
      : same(elements.map((element, index) => pointSize(element, styles[index]!, root)));
  const fontFamily =
    range.collapsed && typing.has("font-family")
      ? typing.get("font-family")
      : same(
          elements.map((element, index) => {
            const style = styles[index]!,
              scope = element.closest<HTMLElement>("[data-word-script],[data-office-script]"),
              role = scope?.dataset.wordScript ?? scope?.dataset.officeScript;
            const wordRole =
              role === "eastAsia" || role === "east-asia"
                ? "eastAsia"
                : role === "cs" || role === "complex-script"
                  ? "cs"
                  : role === "h-ansi" || role === "hAnsi"
                    ? "hAnsi"
                    : "ascii";
            const officeRole =
              wordRole === "eastAsia"
                ? "east-asia"
                : wordRole === "cs"
                  ? "complex-script"
                  : wordRole === "hAnsi"
                    ? "h-ansi"
                    : "ascii";
            return (
              style.getPropertyValue("--word-font-" + wordRole).trim() ||
              style.getPropertyValue("--office-font-" + officeRole).trim() ||
              style.fontFamily
            );
          }),
        );
  const characterProperties: Record<string, string> = {};
  for (const key of [
    "letter-spacing",
    "text-decoration-style",
    "text-decoration-color",
    "font-variant",
    "text-transform",
    "--word-run-native-w",
    "--word-dstrike",
    "--word-underline-style",
    "--word-run-native-vanish",
    "--word-run-native-position",
    "--word-run-native-em",
    "--word-run-border",
    "--word-opentype-ligatures",
    "--word-opentype-num-spacing",
    "--word-opentype-num-form",
    "--word-opentype-stylistic-set",
    "--word-opentype-contextual-alternates",
    "--word-font-ascii",
    "--word-font-hAnsi",
    "--word-font-eastAsia",
    "--word-font-cs",
    "--word-font-size-cs",
    "--word-bold-cs",
    "--word-italic-cs",
    "--word-kerning-threshold",
    "--word-run-snap-to-grid",
    "--word-language-val",
    "--word-language-eastAsia",
    "--word-language-bidi",
  ]) {
    const value =
      range.collapsed && typing.has(key)
        ? typing.get(key)
        : uniform(
            (style) =>
              style.getPropertyValue(key).trim() || (key === "--word-run-native-w" ? "100" : ""),
          );
    if (value) characterProperties[key] = value;
  }
  const doubleStrikethrough =
    range.collapsed && typing.has("--word-dstrike")
      ? typing.get("--word-dstrike") === "1"
      : same(styles.map((style) => style.getPropertyValue("--word-dstrike").trim() === "1"));
  const characterBorder =
    range.collapsed && typing.has("--word-run-border")
      ? Boolean(typing.get("--word-run-border") && typing.get("--word-run-border") !== "none")
      : same(
          styles.map((style) => {
            const border = style.getPropertyValue("--word-run-border").trim();
            return Boolean(border && border !== "none");
          }),
        );
  const inTable = elements.length > 0 && elements.every((element) => Boolean(element.closest("td,th")));
  const themeOwner = root.closest<HTMLElement>('[data-office-kind="document"]') ?? root;
  const rootStyle = view.getComputedStyle(themeOwner);
  const themeSlots = [
    "--office-theme-lt1",
    "--office-theme-dk1",
    "--office-theme-lt2",
    "--office-theme-dk2",
    "--office-theme-accent1",
    "--office-theme-accent2",
    "--office-theme-accent3",
    "--office-theme-accent4",
    "--office-theme-accent5",
    "--office-theme-accent6",
  ].map((slot) => rootStyle.getPropertyValue(slot).trim().toLowerCase());
  const themeColors = themeSlots.every((color) => /^#[0-9a-f]{6}$/i.test(color))
    ? themeSlots
    : undefined;
  const paragraphStyleId = same(
    elements.map(
      (element) => element.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li")?.dataset.wordStyle,
    ),
  );
  const paragraphs = [...new Set(elements.map(element => element.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li")).filter((element): element is HTMLElement => Boolean(element)))];
  const paragraphProperties: Record<string, string> = {};
  for (const key of ["margin-left", "margin-right", "text-indent", "margin-top", "margin-bottom", "line-height", "font-size", "text-align", "widows", "break-after", "break-inside", "break-before", "--word-line-rule", "--word-contextual-spacing", "--word-paragraph-snap-to-grid", "--word-kinsoku", "--word-word-wrap", "--word-overflow-punctuation", "--word-top-line-punctuation", "--word-auto-space-de", "--word-auto-space-dn"]) {
    const value = same(paragraphs.map(paragraph => view.getComputedStyle(paragraph).getPropertyValue(key).trim()));
    if (value) paragraphProperties[key] = value;
  }
  const color =
      range.collapsed && typing.has("color")
        ? typing.get("color")
        : uniform((style) => style.color),
    alignment = uniform((style) => style.textAlign);
  return {
    enabled: true,
    hasSelection: !range.collapsed,
    characterProperties,
    ...(paragraphs.length ? { paragraphProperties } : {}),
    ...(paragraphStyleId === undefined ? {} : { paragraphStyleId }),
    ...(strike === undefined ? {} : { strikethrough: strike }),
    ...(doubleStrikethrough === undefined ? {} : { doubleStrikethrough }),
    ...(characterBorder === undefined ? {} : { characterBorder }),
    ...(inTable ? { inTable: true } : {}),
    ...(themeColors ? { themeColors } : {}),
    ...(verticalAlign === undefined ? {} : { verticalAlign }),
    ...(backgroundColor === undefined ? {} : { backgroundColor }),
    ...(weight === undefined ? {} : { bold: weight === "bold" || Number(weight) >= 600 }),
    ...(italic === undefined ? {} : { italic: italic !== "normal" }),
    ...(decoration === undefined ? {} : { underline: decoration.includes("underline") }),
    ...(size === undefined ? {} : { fontSize: size }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(color === undefined ? {} : { color }),
    ...(alignment === undefined ? {} : { alignment }),
  };
}
