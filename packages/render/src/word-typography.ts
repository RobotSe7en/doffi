import { DocumentError } from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { markViewNode, measureViewContent, ownViewStyle, viewNodeRole } from "./view-nodes.js";

type Kind = "eastAsia" | "letter" | "number" | "other";
const composed = new WeakMap<Document, { root: Element; snapshot: unknown; visibility: string }>();
// Grapheme bases classify spacing; combining marks and variation sequences stay
// with their base. Fullwidth letters/digits already have an East Asian advance.
function kind(text: string): Kind {
  const character = [...text][0]!,
    code = character.codePointAt(0)!;
  if (
    /\p{Script_Extensions=Han}|\p{Script_Extensions=Hiragana}|\p{Script_Extensions=Katakana}|\p{Script=Bopomofo}/u.test(
      character,
    ) &&
    !/[\p{P}\p{S}]/u.test(character)
  )
    return "eastAsia";
  if ((code >= 0xff01 && code <= 0xff60) || /\p{Script=Hangul}/u.test(character)) return "other";
  if (/\p{Nd}/u.test(character)) return "number";
  return /\p{L}/u.test(character) ? "letter" : "other";
}
/** Office paragraph line rules and deterministic inter-script spacing. The
 * removable, collapsible space is layout data: source UTF-16 and clipboard
 * content retain their exact bytes. CSS Text's 1/8ic is the preview pitch. */
export function composeWordTypography(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): void {
  const snapshot = model.snapshot();
  if (!snapshot.metadata.word) return;
  const view = document.defaultView!,
    segmenter = new Intl.Segmenter("und", { granularity: "grapheme" });
  const root = document.querySelector('[data-office-kind="document"]');
  if (!root) return;
  const visibility = [
    ...root.querySelectorAll<HTMLElement>(
      "ins,del,[data-sdk-word-inserted],[data-sdk-word-deleted]",
    ),
  ]
    .map((node) => (node.dataset.nodeId ?? "") + ":" + view.getComputedStyle(node).display)
    .join("|");
  const previous = composed.get(document);
  if (
    previous?.root === root &&
    previous.snapshot === snapshot &&
    previous.visibility === visibility
  )
    return;
  for (const space of root.querySelectorAll("[data-sdk-word-autospace]"))
    if (viewNodeRole(space) === "remove") space.remove();
  composed.set(document, { root, snapshot, visibility });
  const sections = new Map(
      snapshot.metadata.sections?.map((section) => [section.rootNodeId, section]),
    ),
    context = document.createElement("canvas").getContext("2d")!,
    pitches = new Map<string, number>();
  const pitch = (css: CSSStyleDeclaration): number => {
    const font = `${css.fontStyle} ${css.fontWeight} ${css.fontSize} ${css.fontFamily}`;
    let value = pitches.get(font);
    if (value === undefined) {
      context.font = font;
      value = context.measureText("水").width || parseFloat(css.fontSize);
      pitches.set(font, value);
    }
    return value;
  };
  let nodes = document.getElementsByTagName("*").length,
    work = 0;
  for (const paragraph of document.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")) {
    if (
      !paragraph.closest('[data-office-kind="document"]') ||
      paragraph.closest("[data-word-native-object],[data-word-drawing]")
    )
      continue;
    const css = view.getComputedStyle(paragraph),
      de = css.getPropertyValue("--word-auto-space-de").trim() !== "0",
      dn = css.getPropertyValue("--word-auto-space-dn").trim() !== "0";
    ownViewStyle(paragraph, "text-autospace", "line-break", "word-break");
    paragraph.style.setProperty("text-autospace", "no-autospace");
    const kinsoku = css.getPropertyValue("--word-kinsoku").trim() !== "0",
      wrap = css.getPropertyValue("--word-word-wrap").trim() !== "0";
    paragraph.style.lineBreak = kinsoku ? "strict" : "anywhere";
    paragraph.style.wordBreak = wrap ? "normal" : "break-all";
    // Fixed character snapping already supplies the inter-script grid gaps.
    const section = sections.get(
      paragraph.closest("section[data-node-id]")?.getAttribute("data-node-id") ?? "",
    );
    if (
      (!de && !dn) ||
      (section?.grid?.type === "snapToChars" &&
        css.getPropertyValue("--word-paragraph-snap-to-grid").trim() !== "0")
    )
      continue;
    let previous: Kind = "other",
      previousSize = 0;
    const visit = (node: Node): void => {
      if (++work > limits.maxNodes * 32)
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Word typography exceeds its render work budget.",
        );
      if (node.nodeType === 1) {
        const element = node as HTMLElement;
        if (
          viewNodeRole(node) === "remove" ||
          element.matches(
            "br,img,[data-word-tab],[data-word-drawing],[data-word-native-object],[data-word-note-reference],[data-word-equation]",
          )
        ) {
          previous = "other";
          return;
        }
        if (element.matches("[data-word-paragraph-mark]")) return;
        const computed = view.getComputedStyle(element);
        if (computed.display === "none" || computed.visibility === "hidden") return;
        if (
          ["inline-block", "inline-table", "block"].includes(computed.display) &&
          element !== paragraph
        ) {
          previous = "other";
          return;
        }
      }
      if (node.nodeType !== 3) {
        for (const child of [...node.childNodes]) visit(child);
        return;
      }
      const text = node as Text;
      if (!text.length) return;
      const computed = view.getComputedStyle(text.parentElement!),
        size = pitch(computed),
        boundaries: { at: number; size: number }[] = [];
      for (const cluster of segmenter.segment(text.data)) {
        if (++work > limits.maxNodes * 32)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Word typography exceeds its render work budget.",
          );
        const current = kind(cluster.segment),
          other = previous === "eastAsia" ? current : current === "eastAsia" ? previous : "other";
        if ((de && other === "letter") || (dn && other === "number"))
          boundaries.push({
            at: cluster.index,
            size: current === "eastAsia" ? size : previousSize,
          });
        previous = current;
        previousSize = size;
      }
      for (const boundary of boundaries.reverse()) {
        if (++nodes > limits.maxNodes)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Word typography exceeds its render node budget.",
          );
        const after = boundary.at ? text.splitText(boundary.at) : text,
          space = markViewNode(document.createElement("span"));
        measureViewContent(space);
        space.dataset.sdkWordAutospace = "";
        space.setAttribute("aria-hidden", "true");
        space.contentEditable = "false";
        space.style.cssText =
          "font-size:0;line-height:0;letter-spacing:0;white-space:normal;word-spacing:" +
          boundary.size / 8 +
          "px";
        space.textContent = " ";
        after.before(space);
      }
    };
    visit(paragraph);
  }
}
