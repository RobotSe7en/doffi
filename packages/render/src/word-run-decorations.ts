import { DocumentError } from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { markViewNode, ownViewStyle } from "./view-nodes.js";

const triggerTokens = [
  "--word-dstrike",
  "--word-run-native-vanish",
  "--word-run-native-position",
  "--word-run-native-em",
  "--word-underline-style",
  "--word-run-border",
  "--word-opentype-",
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

/** Project Word character-level visual properties that CSS does not directly
 * express on source elements: double strikethrough (independent of underline),
 * hidden text, raised/lowered baseline shift, emphasis marks, extended underline
 * variants, character borders, and OpenType font features. */
export function composeWordRunDecorations(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): void {
  const snapshot = model.snapshot();
  if (!snapshot.metadata.word) return;
  const hasFeature = triggerTokens.some(
    (token) =>
      snapshot.html.includes(token) || snapshot.styles.some((style) => style.css.includes(token)),
  );
  if (!hasFeature) return;
  const view = document.defaultView!;
  const skipSelector =
    "math,[data-word-native-object]:not([data-word-native-boundary]),[data-word-note-reference],[data-word-note-number],[data-word-note-separator],[data-word-equation],[data-word-tab],.sdk-number-marker,.sdk-story-label";
  const blockSelector = "p,h1,h2,h3,h4,h5,h6,td,th,div,section,article,li";

  interface WorkItem {
    text: Text;
    vanish: boolean;
    dstrike: boolean;
    position: number;
    em: string;
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
    const underlineStyle = computed.getPropertyValue("--word-underline-style").trim();
    const border = computed.getPropertyValue("--word-run-border").trim();
    const features = openTypeFeatures(computed);
    const hasEm = em === "dot" || em === "comma" || em === "circle" || em === "underDot";
    const hasBorder = Boolean(border && border !== "none");
    const hasHeavyUnderline = heavyUnderlines.has(underlineStyle);
    const isWordsUnderline =
      underlineStyle === "words" && computed.textDecorationLine.includes("underline");
    if (
      !vanish &&
      !dstrike &&
      position === 0 &&
      !hasEm &&
      !hasBorder &&
      !hasHeavyUnderline &&
      !isWordsUnderline &&
      !features
    )
      continue;
    items.push({
      text,
      vanish,
      dstrike,
      position,
      em: hasEm ? em : "",
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

  let count = document.querySelectorAll("*").length;
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
}
