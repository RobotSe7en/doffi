import { DocumentError, wordScriptRole } from "@doffi/core";
import type { ContentDocumentModel, WordScriptRole } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { ownViewAttribute, markViewNode, viewNodeRole } from "./view-nodes.js";
// General punctuation, symbols and ambiguous alphabetic blocks selected by
// the East Asian hint in the Word run-font table (MS-OI29500 §2.1.88).
const eastAsianHint =
  /[\u00a1\u00a4\u00a7-\u00a8\u00aa\u00ad\u00af-\u00b4\u00b6-\u00ba\u00bc-\u00bf\u00d7\u00f7\u02b0-\u04ff\u2000-\u27bf\ue000-\uf8ff\ufb00-\ufb1c]/u;
const chineseHint =
  /[\u00e0-\u00e1\u00e8-\u00ea\u00ec-\u00ed\u00f2-\u00f3\u00f9-\u00fa\u00fc\u0100-\u02af\u1e00-\u1eff]/u;
/** Font/language role wrappers belong to the view and retain exact source text. */
export function composeWordLanguages(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): void {
  if (!model.snapshot().metadata.word) return;
  for (const node of document.querySelectorAll<HTMLElement>("[data-sdk-word-script-view]"))
    if (viewNodeRole(node) === "unwrap") node.replaceWith(...node.childNodes);
  const walker = document.createTreeWalker(document.body, 4),
    texts: Text[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const parent = node.parentElement;
    if (
      parent?.closest('[data-office-kind="document"]') &&
      !parent.closest(
        '[data-word-drawing],[data-word-native-object],[data-word-equation],[data-word-note-reference],[data-word-content-control="checkBox"]',
      )
    )
      texts.push(node as Text);
  }
  let count = document.querySelectorAll("*").length;
  for (const text of texts) {
    const parent = text.parentElement!,
      css = document.defaultView!.getComputedStyle(parent);
    const preserveWords =
      css.getPropertyValue("--word-kinsoku").trim() === "0" &&
      css.getPropertyValue("--word-word-wrap").trim() !== "0";
    const hinted = css.getPropertyValue("--word-font-hint").trim() === "eastAsia",
      eastAsianLanguage = css.getPropertyValue("--word-language-eastAsia").trim(),
      chinese = /^["']?zh(?:[-"']|$)/i.test(eastAsianLanguage),
      eastAsianFamily = css.getPropertyValue("--word-font-eastAsia").trim(),
      // When both roles name the Chinese face, a substitute may expose
      // separate proportional and full-width forms of the same punctuation.
      sameChineseFamily =
        chinese && eastAsianFamily !== "" &&
        eastAsianFamily === css.getPropertyValue("--word-font-hAnsi").trim();
    if (
      !preserveWords &&
      !["ascii", "hAnsi", "eastAsia", "cs"].some((role) =>
        css.getPropertyValue("--word-font-" + role).trim(),
      ) &&
      !["val", "eastAsia", "bidi"].some((role) =>
        css.getPropertyValue("--word-language-" + role).trim(),
      )
    )
      continue;
    let role: WordScriptRole | undefined,
      fullWidth = false,
      part = "";
    const fragment = document.createDocumentFragment();
    const append = (): void => {
      if (!part || !role) return;
      if (++count > limits.maxNodes)
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Script roles exceed the render node budget.",
        );
      const span = markViewNode(document.createElement("span"), "unwrap");
      span.dataset.sdkWordScriptView = "";
      span.dataset.wordScript = role;
      span.style.fontFamily = "var(--word-font-" + role + ",var(--word-font-ascii,inherit))";
      if (fullWidth) span.style.fontVariantEastAsian = "full-width";
      if (role === "cs") {
        const size = css.getPropertyValue("--word-font-size-cs").trim(),
          bold = css.getPropertyValue("--word-bold-cs").trim(),
          italic = css.getPropertyValue("--word-italic-cs").trim();
        if (size) span.style.fontSize = size;
        if (bold) span.style.fontWeight = bold === "1" ? "700" : "400";
        if (italic) span.style.fontStyle = italic === "1" ? "italic" : "normal";
      }
      if (preserveWords && role !== "eastAsia") span.style.lineBreak = "auto";
      span.textContent = part;
      fragment.append(span);
      part = "";
    };
    for (const character of text.data) {
      const base = wordScriptRole(character, role),
        useHint =
          hinted && base === "hAnsi" &&
          (eastAsianHint.test(character) || (chinese && chineseHint.test(character))),
        next = useHint ? "eastAsia" : base,
        width =
          (useHint || (sameChineseFamily && base === "hAnsi")) &&
          /[\u2000-\u206f]/u.test(character) && /\p{P}/u.test(character);
      if (next !== role || width !== fullWidth) {
        append();
        role = next;
        fullWidth = width;
      }
      part += character;
    }
    append();
    text.replaceWith(fragment);
  }
  for (const node of document.querySelectorAll<HTMLElement>(
    "[data-node-id],[data-sdk-word-script-view]",
  )) {
    const css = document.defaultView!.getComputedStyle(node),
      role =
        node.dataset.wordScript === "eastAsia"
          ? "eastAsia"
          : node.dataset.wordScript === "cs"
            ? "bidi"
            : "val",
      raw =
        css.getPropertyValue("--word-language-" + role).trim() ||
        css.getPropertyValue("--word-language-val").trim();
    if (!raw) continue;
    let language = raw;
    if (/^"/.test(raw))
      try {
        language = JSON.parse(raw) as string;
      } catch {
        continue;
      }
    else if (/^'[^']*'$/.test(raw)) language = raw.slice(1, -1);
    if (!language || /[\u0000-\u0020\u007f]/.test(language)) continue;
    ownViewAttribute(node, "lang");
    node.lang = language;
  }
}
