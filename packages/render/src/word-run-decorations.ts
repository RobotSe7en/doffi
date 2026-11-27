import { DocumentError } from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { hideFromViewSelection, markViewNode, ownViewAttribute, ownViewStyle } from "./view-nodes.js";

const triggerTokens = [
  "--word-dstrike",
  "--word-run-native-vanish",
  "--word-run-native-position",
  "--word-run-native-em",
  "--word-run-native-outline",
  "--word-run-native-shadow",
  "--word-run-native-emboss",
  "--word-run-native-imprint",
  "--word-text-glow",
  "--word-text-outline",
  "--word-text-shadow",
  "--word-text-reflection",
  "--word-text-fill-gradient",
  "--word-text-bevel",
  "--word-text-3d",
  "--word-ea-combine",
  "--word-ea-vert",
  "--word-fit-text",
  "--word-underline-style",
  "--word-run-border",
  "--word-opentype-",
  "--word-drop-cap",
  "--word-border-tl2br",
  "--word-border-tr2bl",
  "data-word-ruby",
] as const;

const heavyUnderlines = new Set([
  "thick",
  "dottedHeavy",
  "dashedHeavy",
  "dashLongHeavy",
  "dashDotHeavy",
  "dashDotDotHeavy",
  "wavyHeavy",
]);

const doubleStrikeGradient =
  "linear-gradient(to bottom, transparent 40%, currentColor 40%, currentColor 46%, transparent 46%, transparent 56%, currentColor 56%, currentColor 62%, transparent 62%)";

function openTypeFeatures(computed: CSSStyleDeclaration): string {
  const features: string[] = [];
  const ligatures = computed.getPropertyValue("--word-opentype-ligatures").trim();
  if (ligatures) {
    if (ligatures === "none") {
      features.push('"liga" 0', '"clig" 0', '"dlig" 0', '"hlig" 0');
    } else if (ligatures === "all") {
      features.push('"liga" 1', '"clig" 1', '"dlig" 1', '"hlig" 1');
    } else {
      const lower = ligatures.toLowerCase();
      features.push(
        `"liga" ${lower.includes("standard") ? 1 : 0}`,
        `"clig" ${lower.includes("contextual") ? 1 : 0}`,
        `"dlig" ${lower.includes("discretional") ? 1 : 0}`,
        `"hlig" ${lower.includes("historical") ? 1 : 0}`,
      );
    }
  }
  const numSpacing = computed.getPropertyValue("--word-opentype-num-spacing").trim();
  if (numSpacing === "proportional") features.push('"pnum" 1', '"tnum" 0');
  else if (numSpacing === "tabular") features.push('"tnum" 1', '"pnum" 0');
  const numForm = computed.getPropertyValue("--word-opentype-num-form").trim();
  if (numForm === "lining") features.push('"lnum" 1', '"onum" 0');
  else if (numForm === "oldStyle") features.push('"onum" 1', '"lnum" 0');
  const stylistic = computed.getPropertyValue("--word-opentype-stylistic-set").trim();
  if (stylistic && stylistic !== "none") {
    for (const part of stylistic.split(",")) {
      const id = Number(part.trim());
      if (Number.isInteger(id) && id >= 1 && id <= 20)
        features.push(`"ss${String(id).padStart(2, "0")}" 1`);
    }
  }
  const calt = computed.getPropertyValue("--word-opentype-contextual-alternates").trim();
  if (calt === "1" || calt === "0") features.push(`"calt" ${calt}`);
  return features.join(", ");
}

/** Project Word character-level visual properties, phonetic guides (ruby),
 * drop caps, enclosed characters, text effects, East Asian layout, and
 * diagonal table cell borders. */
export function composeWordRunDecorations(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): void {
  const snapshot = model.snapshot();
  if (!snapshot.metadata.word) return;
  const word = snapshot.metadata.word;
  const hasEnclosedField = word.fields.some((f) => /^\s*EQ\s+\\o/i.test(f.instruction));
  const hasFeature =
    hasEnclosedField ||
    triggerTokens.some(
      (token) =>
        snapshot.html.includes(token) || snapshot.styles.some((style) => style.css.includes(token)),
    );
  if (!hasFeature) return;
  const view = document.defaultView!;
  const skipSelector =
    "math,[data-word-native-object]:not([data-word-native-boundary]),[data-word-note-reference],[data-word-note-number],[data-word-note-separator],[data-word-equation],[data-word-tab],.sdk-number-marker,.sdk-story-label,.sdk-word-ruby-rt";
  const blockSelector = "p,h1,h2,h3,h4,h5,h6,td,th,div,section,article,li";

  let count = document.querySelectorAll("*").length;

  // 1. Drop Cap paragraphs (w:framePr w:dropCap="drop"|"margin")
  for (const para of document.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")) {
    const computed = view.getComputedStyle(para);
    const dropCap = computed.getPropertyValue("--word-drop-cap").trim();
    if (dropCap !== "drop" && dropCap !== "margin") continue;
    const lines = Math.max(1, Math.min(10, Number(computed.getPropertyValue("--word-drop-cap-lines").trim()) || 3));
    const space = computed.getPropertyValue("--word-drop-cap-space").trim() || "0pt";
    const baseSize = parseFloat(computed.fontSize) || 16;
    const lineH = parseFloat(computed.lineHeight) || baseSize * 1.3;
    const targetSize = Math.round(lines * lineH * 0.88 * 100) / 100;
    const marginSpace = space === "0pt" ? "4pt" : space;
    const visibleChars = [...(para.textContent ?? "").trim()].length;
    ownViewAttribute(para, "data-sdk-word-drop-cap");
    if (visibleChars <= 2) {
      // Native Word 1-character Drop Cap frame paragraph followed by a continuation paragraph.
      // Match the continuation paragraph's rendered line height so the cap spans the requested lines.
      const nextEl = para.nextElementSibling as HTMLElement | null;
      const nextComputed = nextEl ? view.getComputedStyle(nextEl) : null;
      const nextFont = nextComputed ? parseFloat(nextComputed.fontSize) || baseSize : baseSize;
      const rawNextLine = nextComputed ? parseFloat(nextComputed.lineHeight) || nextFont * 1.4 : lineH;
      // composeWordAutoLineSpacing scales auto-rule CJK paragraphs by the font's natural
      // line-spacing factor (~1.38x) later in the pipeline; account for it so the cap spans N lines.
      const effectiveLine =
        nextEl?.dataset.wordLineRule === "auto"
          ? Math.max(rawNextLine * 1.38, nextFont * 1.6)
          : rawNextLine;
      const targetBoxHeight = Math.round(lines * effectiveLine * 0.88 * 100) / 100;
      const finalSize = Math.round(targetBoxHeight * 0.88 * 100) / 100;
      ownViewStyle(
        para,
        "float",
        "font-size",
        "line-height",
        "min-height",
        "margin-right",
        "margin-left",
        "margin-top",
        "margin-bottom",
      );
      para.style.float = "left";
      para.style.fontSize = finalSize + "px";
      para.style.lineHeight = Math.round(targetBoxHeight * 0.96) + "px";
      para.style.minHeight = targetBoxHeight + "px";
      para.style.marginTop = "0";
      para.style.marginBottom = "0";
      para.style.marginRight = marginSpace;
      if (dropCap === "margin") para.style.marginLeft = -Math.round(finalSize * 1.1) + "px";
      para.dataset.sdkWordDropCap = dropCap;
    } else {
      // Full-paragraph Drop Cap authored in editor: wrap only the first grapheme in an unwrap view span
      const walker = document.createTreeWalker(para, 4);
      let firstText: Text | null = null;
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const txt = n as Text;
        if (txt.data.trim().length > 0 && !txt.parentElement?.closest(skipSelector)) {
          firstText = txt;
          break;
        }
      }
      if (firstText && ++count <= limits.maxNodes) {
        const leadingMatch = /^\s*/.exec(firstText.data);
        const startIdx = leadingMatch ? leadingMatch[0].length : 0;
        const firstChar = [...firstText.data.slice(startIdx)][0] ?? "";
        if (firstChar) {
          if (startIdx > 0) firstText = firstText.splitText(startIdx);
          if (firstText.data.length > firstChar.length) firstText.splitText(firstChar.length);
          const capSpan = markViewNode(document.createElement("span"), "unwrap");
          capSpan.className = "sdk-word-drop-cap";
          capSpan.style.cssText =
            `float:left;font-size:${targetSize}px;line-height:0.92;margin-right:${marginSpace};` +
            (dropCap === "margin" ? `margin-left:-${Math.round(targetSize * 1.1)}px;` : "");
          firstText.parentNode!.insertBefore(capSpan, firstText);
          capSpan.append(firstText);
          para.dataset.sdkWordDropCap = dropCap + "-inline";
        }
      }
    }
  }

  // 2. Phonetic Guide / Ruby ([data-word-ruby="true"])
  for (const rubyEl of document.querySelectorAll<HTMLElement>('[data-word-ruby="true"]')) {
    const rubyText = rubyEl.dataset.wordRubyText;
    if (!rubyText) continue;
    if (++count > limits.maxNodes)
      throw new DocumentError("ResourceLimitExceeded", "Ruby projection exceeds node budget.");
    const align = rubyEl.dataset.wordRubyAlign ?? "distributeSpace";
    const sizePt = Number(rubyEl.dataset.wordRubySize ?? "6") || 6;
    const raisePt = Number(rubyEl.dataset.wordRubyRaise ?? "11") || 11;
    ownViewStyle(rubyEl, "position", "display", "padding-top", "vertical-align");
    rubyEl.style.position = "relative";
    rubyEl.style.display = "inline-block";
    rubyEl.style.paddingTop = Math.max(sizePt + 1, Math.min(raisePt * 0.5, sizePt + 2.5)) + "pt";
    rubyEl.style.verticalAlign = "baseline";
    const rt = markViewNode(document.createElement("span"));
    hideFromViewSelection(rt);
    rt.className = "sdk-word-ruby-rt";
    rt.textContent = rubyText;
    const justify =
      align === "left"
        ? "left"
        : align === "right"
          ? "right"
          : align === "distributeLetter" || align === "distributeSpace"
            ? "center"
            : "center";
    rt.style.cssText =
      `position:absolute;left:0;right:0;top:1pt;font-size:${sizePt}pt;line-height:1;white-space:nowrap;text-align:${justify};pointer-events:none;user-select:none;` +
      (align === "distributeLetter" ? "letter-spacing:0.15em;" : "");
    rubyEl.prepend(rt);
  }

  // 3. Enclosed Characters (EQ \o\ac(○,字) fields)
  if (hasEnclosedField) {
    for (const field of word.fields) {
      const m = /^\s*EQ\s+\\o(?:\\ac)?\(\s*(?:\\s\\up\s*\d+\()?([^,)]+)\)?\s*,\s*([^)]+)\s*\)/i.exec(
        field.instruction,
      );
      if (!m) continue;
      const el = document.querySelector<HTMLElement>(`[data-node-id="${CSS.escape(field.nodeId)}"]`);
      if (!el) continue;
      const enclosure = m[1]!.trim();
      const char = m[2]!.trim();
      if (++count > limits.maxNodes)
        throw new DocumentError("ResourceLimitExceeded", "Enclosed character projection exceeds budget.");
      ownViewStyle(el, "background", "display", "position", "padding");
      el.style.background = "transparent";
      el.style.display = "inline-flex";
      el.style.alignItems = "center";
      el.style.justifyContent = "center";
      el.style.position = "relative";
      const badge = markViewNode(document.createElement("span"));
      hideFromViewSelection(badge);
      badge.className = "sdk-word-enclosed";
      const isSquare = enclosure === "□" || enclosure === "▢";
      const isDiamond = enclosure === "◇";
      const isTriangle = enclosure === "△";
      badge.style.cssText = isTriangle
        ? "display:inline-flex;align-items:center;justify-content:center;min-width:1.45em;height:1.45em;font-size:0.78em;line-height:1;position:relative;vertical-align:middle;"
        : `display:inline-flex;align-items:center;justify-content:center;min-width:1.35em;height:1.35em;padding:0 0.12em;border:0.75pt solid currentColor;border-radius:${isSquare || isDiamond ? "1pt" : "999px"};font-size:0.8em;line-height:1;vertical-align:middle;${isDiamond ? "transform:rotate(45deg);" : ""}`;
      if (isTriangle) {
        badge.innerHTML = `<svg viewBox="0 0 24 24" style="position:absolute;inset:0;width:100%;height:100%;pointer-events:none"><polygon points="12,2 22,21 2,21" fill="none" stroke="currentColor" stroke-width="1.5"/></svg><span style="position:relative;top:2px">${char.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</span>`;
      } else if (isDiamond) {
        badge.innerHTML = `<span style="transform:rotate(-45deg);display:inline-block">${char.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</span>`;
      } else {
        badge.textContent = char;
      }
      for (const child of [...el.childNodes]) {
        const hidden = markViewNode(document.createElement("span"), "unwrap");
        hideFromViewSelection(hidden);
        hidden.style.cssText = "position:absolute;visibility:hidden;pointer-events:none";
        hidden.append(child);
        el.append(hidden);
      }
      el.append(badge);
    }
  }

  // 4. Diagonal Table Cell Borders (--word-border-tl2br, --word-border-tr2bl)
  for (const cell of document.querySelectorAll<HTMLElement>("td,th")) {
    const computed = view.getComputedStyle(cell);
    const tl2br = computed.getPropertyValue("--word-border-tl2br").trim();
    const tr2bl = computed.getPropertyValue("--word-border-tr2bl").trim();
    const hasTl = Boolean(tl2br && !tl2br.includes("none") && !tl2br.startsWith("0"));
    const hasTr = Boolean(tr2bl && !tr2bl.includes("none") && !tr2bl.startsWith("0"));
    if (!hasTl && !hasTr) continue;
    if (++count + 3 > limits.maxNodes)
      throw new DocumentError("ResourceLimitExceeded", "Diagonal cell border exceeds node budget.");
    ownViewStyle(cell, "position");
    if (computed.position === "static") cell.style.position = "relative";
    const ns = "http://www.w3.org/2000/svg";
    const svg = markViewNode(document.createElementNS(ns, "svg") as unknown as HTMLElement);
    hideFromViewSelection(svg);
    svg.setAttribute("class", "sdk-word-diagonal-border");
    svg.setAttribute("preserveAspectRatio", "none");
    svg.setAttribute("viewBox", "0 0 100 100");
    svg.style.cssText =
      "position:absolute;inset:0;width:100%;height:100%;pointer-events:none;overflow:visible;z-index:0";
    const addLine = (spec: string, x1: string, y1: string, x2: string, y2: string) => {
      const m = /^(\S+)\s+(\S+)\s+(.+)$/.exec(spec);
      const widthPt = m ? parseFloat(m[1]!) || 0.5 : 0.5;
      const style = m?.[2] ?? "solid";
      const stroke = m?.[3] ?? "#000000";
      const line = document.createElementNS(ns, "line");
      line.setAttribute("x1", x1);
      line.setAttribute("y1", y1);
      line.setAttribute("x2", x2);
      line.setAttribute("y2", y2);
      line.setAttribute("stroke", stroke);
      line.setAttribute("stroke-width", String(Math.max(0.75, widthPt * 1.333)));
      line.setAttribute("vector-effect", "non-scaling-stroke");
      if (style === "dashed") line.setAttribute("stroke-dasharray", "4,2");
      else if (style === "dotted") line.setAttribute("stroke-dasharray", "1.5,1.5");
      svg.append(line);
    };
    if (hasTl) addLine(tl2br, "0", "0", "100", "100");
    if (hasTr) addLine(tr2bl, "100", "0", "0", "100");
    cell.prepend(svg);
  }

  // 5. Run-level decorations, text effects, East Asian layout, and Fit Text
  interface WorkItem {
    text: Text;
    vanish: boolean;
    dstrike: boolean;
    position: number;
    em: string;
    outline: boolean;
    shadow: boolean;
    emboss: boolean;
    imprint: boolean;
    textGlow: string;
    textOutline: string;
    textShadow: string;
    textReflection: boolean;
    textGradient: string;
    textBevel: string;
    text3d: string;
    eaCombine: string;
    eaVert: string;
    fitText: string;
    underlineStyle: string;
    hasUnderline: boolean;
    underlineColor: string;
    hasStrike: boolean;
    border: string;
    features: string;
    superSub: boolean;
  }

  const items: WorkItem[] = [];
  const walker = document.createTreeWalker(document.body, 4);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node as Text;
    const parent = text.parentElement;
    if (!text.data || !parent || parent.closest(skipSelector)) continue;
    const computed = view.getComputedStyle(parent);
    const vanish = computed.getPropertyValue("--word-run-native-vanish").trim() === "1";
    const dstrike = computed.getPropertyValue("--word-dstrike").trim() === "1";
    const rawPosition = computed.getPropertyValue("--word-run-native-position").trim();
    const position = /^-?\d+$/.test(rawPosition) ? Number(rawPosition) : 0;
    const em = computed.getPropertyValue("--word-run-native-em").trim();
    const outline = computed.getPropertyValue("--word-run-native-outline").trim() === "1";
    const shadow = computed.getPropertyValue("--word-run-native-shadow").trim() === "1";
    const emboss = computed.getPropertyValue("--word-run-native-emboss").trim() === "1";
    const imprint = computed.getPropertyValue("--word-run-native-imprint").trim() === "1";
    const textGlow = computed.getPropertyValue("--word-text-glow").trim();
    const textOutline = computed.getPropertyValue("--word-text-outline").trim();
    const textShadow = computed.getPropertyValue("--word-text-shadow").trim();
    const textReflection = computed.getPropertyValue("--word-text-reflection").trim() === "1";
    const textGradient = computed.getPropertyValue("--word-text-fill-gradient").trim();
    const textBevel = computed.getPropertyValue("--word-text-bevel").trim();
    const text3d = computed.getPropertyValue("--word-text-3d").trim();
    const eaCombine = computed.getPropertyValue("--word-ea-combine").trim();
    const eaVert = computed.getPropertyValue("--word-ea-vert").trim();
    const fitText = computed.getPropertyValue("--word-fit-text").trim();
    const underlineStyle = computed.getPropertyValue("--word-underline-style").trim();
    const border = computed.getPropertyValue("--word-run-border").trim();
    const features = openTypeFeatures(computed);
    const hasEm = em === "dot" || em === "comma" || em === "circle" || em === "underDot";
    const hasBorder = Boolean(border && border !== "none");
    const hasHeavyUnderline = heavyUnderlines.has(underlineStyle);
    const isWordsUnderline =
      underlineStyle === "words" && computed.textDecorationLine.includes("underline");
    const hasEffects =
      outline ||
      shadow ||
      emboss ||
      imprint ||
      Boolean(textGlow && textGlow !== "none") ||
      Boolean(textOutline && textOutline !== "none") ||
      Boolean(textShadow && textShadow !== "none") ||
      textReflection ||
      Boolean(textGradient && textGradient !== "none") ||
      Boolean(textBevel && textBevel !== "none") ||
      Boolean(text3d && text3d !== "none");
    const hasEa =
      Boolean(eaCombine && eaCombine !== "none") ||
      Boolean(eaVert && eaVert !== "0") ||
      Boolean(fitText && fitText !== "none");
    if (
      !vanish &&
      !dstrike &&
      position === 0 &&
      !hasEm &&
      !hasBorder &&
      !hasHeavyUnderline &&
      !isWordsUnderline &&
      !features &&
      !hasEffects &&
      !hasEa
    )
      continue;
    items.push({
      text,
      vanish,
      dstrike,
      position,
      em: hasEm ? em : "",
      outline,
      shadow,
      emboss,
      imprint,
      textGlow: textGlow !== "none" ? textGlow : "",
      textOutline: textOutline !== "none" ? textOutline : "",
      textShadow: textShadow !== "none" ? textShadow : "",
      textReflection,
      textGradient: textGradient !== "none" ? textGradient : "",
      textBevel: textBevel !== "none" ? textBevel : "",
      text3d: text3d !== "none" ? text3d : "",
      eaCombine: eaCombine !== "none" ? eaCombine : "",
      eaVert: eaVert !== "0" ? eaVert : "",
      fitText: fitText !== "none" ? fitText : "",
      underlineStyle,
      hasUnderline: computed.textDecorationLine.includes("underline"),
      underlineColor: computed.textDecorationColor || "currentColor",
      hasStrike: computed.textDecorationLine.includes("line-through"),
      border: hasBorder ? border : "",
      features,
      superSub: computed.verticalAlign === "super" || computed.verticalAlign === "sub",
    });
  }
  if (!items.length) return;

  const wrapText = (text: Text): HTMLElement => {
    const parent = text.parentElement!;
    if (
      !parent.matches(blockSelector) &&
      parent.childNodes.length === 1 &&
      parent.firstChild === text &&
      (parent.classList.contains("sdk-word-text-run") ||
        parent.hasAttribute("data-sdk-word-script-view"))
    )
      return parent;
    if (++count > limits.maxNodes)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Word run decoration projection exceeds the render node budget.",
      );
    const wrapper = markViewNode(document.createElement("span"), "unwrap");
    wrapper.className = "sdk-word-run-decoration";
    text.before(wrapper);
    wrapper.append(text);
    return wrapper;
  };

  for (const item of items) {
    if (!item.text.isConnected) continue;
    if (item.underlineStyle === "words" && item.hasUnderline) {
      const parent = item.text.parentElement!;
      if (parent.classList.contains("sdk-word-text-run") || parent.hasAttribute("data-sdk-word-script-view")) {
        ownViewStyle(parent, "text-decoration-line");
        parent.style.textDecorationLine = item.hasStrike ? "line-through" : "none";
      }
      const parts = item.text.data.split(/(\s+)/).filter(Boolean);
      const fragment = document.createDocumentFragment();
      for (const part of parts) {
        if (++count > limits.maxNodes)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Word underline projection exceeds the render node budget.",
          );
        const span = markViewNode(document.createElement("span"), "unwrap");
        span.className = "sdk-word-run-decoration";
        if (!/^\s+$/.test(part)) {
          span.style.textDecorationLine = item.hasStrike ? "underline line-through" : "underline";
          span.style.textDecorationStyle = "solid";
          span.style.textDecorationColor = item.underlineColor;
        } else if (item.hasStrike) {
          span.style.textDecorationLine = "line-through";
        }
        applyVisuals(span, item);
        span.append(document.createTextNode(part));
        fragment.append(span);
      }
      item.text.replaceWith(fragment);
      continue;
    }
    const target = wrapText(item.text);
    target.classList.add("sdk-word-run-decoration");
    applyVisuals(target, item);
  }
}

function applyVisuals(
  target: HTMLElement,
  item: {
    vanish: boolean;
    dstrike: boolean;
    position: number;
    em: string;
    outline: boolean;
    shadow: boolean;
    emboss: boolean;
    imprint: boolean;
    textGlow: string;
    textOutline: string;
    textShadow: string;
    textReflection: boolean;
    textGradient: string;
    textBevel: string;
    text3d: string;
    eaCombine: string;
    eaVert: string;
    fitText: string;
    underlineStyle: string;
    hasUnderline: boolean;
    border: string;
    features: string;
    superSub: boolean;
  },
): void {
  if (item.vanish) target.dataset.sdkWordHidden = "true";
  if (item.dstrike) {
    ownViewStyle(target, "background-image", "background-repeat");
    target.style.backgroundImage = doubleStrikeGradient;
    target.style.backgroundRepeat = "no-repeat";
  }
  if (item.position !== 0) {
    const pt = item.position / 2 + "pt";
    if (item.superSub) {
      ownViewStyle(target, "position", "bottom");
      target.style.position = "relative";
      target.style.bottom = pt;
    } else {
      ownViewStyle(target, "vertical-align");
      target.style.verticalAlign = pt;
    }
  }
  if (item.em) {
    ownViewStyle(target, "text-emphasis-style", "text-emphasis-position");
    if (item.em === "dot" || item.em === "underDot") {
      target.style.setProperty("-webkit-text-emphasis-style", "filled dot");
      target.style.setProperty("text-emphasis-style", "filled dot");
      target.style.setProperty("-webkit-text-emphasis-position", "under right");
      target.style.setProperty("text-emphasis-position", "under right");
    } else if (item.em === "comma") {
      target.style.setProperty("-webkit-text-emphasis-style", "filled sesame");
      target.style.setProperty("text-emphasis-style", "filled sesame");
      target.style.setProperty("-webkit-text-emphasis-position", "over right");
      target.style.setProperty("text-emphasis-position", "over right");
    } else if (item.em === "circle") {
      target.style.setProperty("-webkit-text-emphasis-style", "open circle");
      target.style.setProperty("text-emphasis-style", "open circle");
      target.style.setProperty("-webkit-text-emphasis-position", "over right");
      target.style.setProperty("text-emphasis-position", "over right");
    }
  }
  if (item.hasUnderline && heavyUnderlines.has(item.underlineStyle)) {
    ownViewStyle(target, "text-decoration-thickness");
    target.style.textDecorationThickness = "0.14em";
  }
  if (item.border) {
    ownViewStyle(target, "border", "box-decoration-break");
    target.style.border = item.border;
    target.style.setProperty("-webkit-box-decoration-break", "clone");
    target.style.boxDecorationBreak = "clone";
  }
  if (item.features) {
    ownViewStyle(target, "font-feature-settings");
    target.style.fontFeatureSettings = item.features;
  }
  let appliedGradient = false;
  if (item.textGradient) {
    const raw = item.textGradient.trim();
    const angleMatch = /^(-?\d+(?:\.\d+)?)deg\b/i.exec(raw);
    const angle = angleMatch ? `${angleMatch[1]}deg` : "90deg";
    const body = angleMatch ? raw.slice(angleMatch[0].length) : raw;
    const stops: string[] = [];
    const stopRegex = /(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\)|[a-zA-Z]+)\s*(\d+(?:\.\d+)?%)?/g;
    let m: RegExpExecArray | null;
    while ((m = stopRegex.exec(body)) !== null) {
      stops.push(m[2] ? `${m[1]} ${m[2]}` : m[1]!);
    }
    if (stops.length >= 2) {
      appliedGradient = true;
      ownViewStyle(target, "background-image", "-webkit-background-clip", "background-clip", "-webkit-text-fill-color");
      target.style.backgroundImage = `linear-gradient(${angle}, ${stops.join(", ")})`;
      target.style.setProperty("-webkit-background-clip", "text");
      target.style.backgroundClip = "text";
      target.style.setProperty("-webkit-text-fill-color", "transparent");
    }
  }
  if (appliedGradient) {
    // CSS text-shadow paints above background-clip:text when text-fill-color is transparent,
    // muting gradient fills. Use filter: drop-shadow(...) so effects composite behind the gradient.
    const dropFilters: string[] = [];
    if (item.textGlow) {
      const m = /^(\S+)\s+(.+)$/.exec(item.textGlow);
      if (m) dropFilters.push(`drop-shadow(0 0 ${m[1]} ${m[2]})`);
    }
    if (item.textShadow) {
      const m = /^(\S+)\s+(\S+)\s+(\S+)\s+(.+)$/.exec(item.textShadow);
      if (m) {
        const blur = m[1]!;
        const dist = parseFloat(m[2]!) || 2;
        const rad = ((Number(m[3]) || 45) * Math.PI) / 180;
        const dx = (Math.cos(rad) * dist).toFixed(2) + "pt";
        const dy = (Math.sin(rad) * dist).toFixed(2) + "pt";
        dropFilters.push(`drop-shadow(${dx} ${dy} ${blur} ${m[4]})`);
      }
    } else if (item.shadow) {
      dropFilters.push("drop-shadow(1pt 1.2pt 1.4pt rgba(0,0,0,0.35))");
    }
    if (item.text3d || item.textBevel) {
      dropFilters.push("drop-shadow(0.9pt 1.2pt 1.1pt rgba(60,30,10,0.36))");
    }
    if (dropFilters.length) {
      ownViewStyle(target, "filter");
      target.style.filter = dropFilters.join(" ");
    }
  } else {
    const shadows: string[] = [];
    if (item.shadow) shadows.push("1.2pt 1.2pt 1.5pt rgba(0,0,0,0.45)");
    if (item.emboss) shadows.push("-0.6pt -0.6pt 0 rgba(255,255,255,0.85), 0.75pt 0.75pt 0 rgba(0,0,0,0.38)");
    if (item.imprint) shadows.push("0.6pt 0.6pt 0 rgba(255,255,255,0.85), -0.75pt -0.75pt 0 rgba(0,0,0,0.38)");
    if (item.textGlow) {
      const m = /^(\S+)\s+(.+)$/.exec(item.textGlow);
      if (m) shadows.push(`0 0 ${m[1]} ${m[2]}, 0 0 ${m[1]} ${m[2]}`);
    }
    if (item.textShadow) {
      const m = /^(\S+)\s+(\S+)\s+(\S+)\s+(.+)$/.exec(item.textShadow);
      if (m) {
        const blur = m[1]!;
        const dist = parseFloat(m[2]!) || 2;
        const rad = ((Number(m[3]) || 45) * Math.PI) / 180;
        const dx = (Math.cos(rad) * dist).toFixed(2) + "pt";
        const dy = (Math.sin(rad) * dist).toFixed(2) + "pt";
        shadows.push(`${dx} ${dy} ${blur} ${m[4]}`);
      }
    }
    if (item.textBevel) {
      shadows.push(
        "-0.7pt -0.7pt 0.4pt rgba(255,255,255,0.75)",
        "0.9pt 0.9pt 0.6pt rgba(0,0,0,0.42)",
        "1.4pt 1.4pt 1.2pt rgba(0,0,0,0.22)",
      );
    }
    if (item.text3d) {
      shadows.push(
        "0.6pt 0.6pt 0 rgba(0,0,0,0.28)",
        "1.2pt 1.2pt 0 rgba(0,0,0,0.22)",
        "1.8pt 1.8pt 1.5pt rgba(0,0,0,0.18)",
      );
    }
    if (shadows.length) {
      ownViewStyle(target, "text-shadow");
      target.style.textShadow = shadows.join(", ");
    }
  }
  if (item.outline) {
    ownViewStyle(target, "-webkit-text-stroke", "-webkit-text-fill-color", "paint-order");
    target.style.setProperty("-webkit-text-stroke", "0.5pt currentColor");
    target.style.setProperty("-webkit-text-fill-color", "#ffffff");
    target.style.setProperty("paint-order", "stroke fill");
  } else if (item.textOutline) {
    const m = /^(\S+)\s+(.+)$/.exec(item.textOutline);
    if (m) {
      ownViewStyle(target, "-webkit-text-stroke");
      target.style.setProperty("-webkit-text-stroke", `${m[1]} ${m[2]}`);
    }
  }
  if (item.textReflection) {
    ownViewStyle(target, "display", "-webkit-box-reflect");
    target.style.display = "inline-block";
    target.style.setProperty(
      "-webkit-box-reflect",
      "below -3px linear-gradient(transparent 15%, rgba(0,0,0,0.38))",
    );
  }
  if (item.eaCombine) {
    ownViewStyle(target, "display", "font-size", "line-height", "vertical-align", "max-width");
    target.style.display = "inline-block";
    target.style.fontSize = "0.52em";
    target.style.lineHeight = "1.08";
    target.style.verticalAlign = "middle";
    const chars = Array.from(target.textContent ?? "").length;
    target.style.maxWidth = Math.max(1, Math.ceil(chars / 2)) + 0.5 + "em";
    target.dataset.sdkWordEaCombine = item.eaCombine;
    const brackets: Record<string, readonly [string, string]> = {
      round: ["(", ")"],
      square: ["[", "]"],
      angle: ["<", ">"],
      curly: ["{", "}"],
    };
    const pair = brackets[item.eaCombine];
    if (pair && target.ownerDocument) {
      const open = markViewNode(target.ownerDocument.createElement("span"));
      const close = markViewNode(target.ownerDocument.createElement("span"));
      hideFromViewSelection(open);
      hideFromViewSelection(close);
      open.className = "sdk-word-ea-bracket";
      close.className = "sdk-word-ea-bracket";
      open.textContent = pair[0];
      close.textContent = pair[1];
      target.before(open);
      target.after(close);
    }
  }
  if (item.eaVert) {
    ownViewStyle(target, "display", "writing-mode", "transform", "line-height", "vertical-align", "font-size");
    target.style.display = "inline-block";
    target.style.writingMode = "vertical-rl";
    target.style.transform = "rotate(180deg)";
    target.style.lineHeight = "1";
    target.style.verticalAlign = "middle";
    target.style.fontSize = item.eaVert === "compress" ? "0.74em" : "0.85em";
  }
  if (item.fitText) {
    const targetPt = parseFloat(item.fitText);
    if (Number.isFinite(targetPt) && targetPt > 0) {
      ownViewStyle(target, "display", "min-width", "text-align", "text-align-last", "text-justify");
      target.style.display = "inline-block";
      target.style.minWidth = targetPt + "pt";
      target.style.textAlign = "justify";
      target.style.textAlignLast = "justify";
      target.style.textJustify = "inter-character";
    }
  }
}
