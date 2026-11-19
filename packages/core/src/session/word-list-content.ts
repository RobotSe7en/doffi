import { DocumentError } from "../errors.js";
import { createOpaqueId } from "../identifiers.js";
import type { WordMetadata, WordNumberLevel } from "../contracts/word.js";
import type { ParsedHtml } from "../document/html.js";
import type { DocumentEdit } from "../document/edit.js";
import type { SourceStyleNode } from "../document/source-css.js";
import type { SourceStyleEdit } from "./source-styles.js";

/** HTML lists become the same numbered paragraphs as native Word import.
 * Container identities remain transparent source wrappers for exact review. */
export function normalizeWordListContent(
  word: WordMetadata,
  edit: DocumentEdit,
  source: ParsedHtml,
  inserted: ReadonlySet<string>,
  styles: SourceStyleEdit,
): void {
  const inside = (id: string): boolean => {
    for (let current: string | null = id; current; current = source.peekNode(current)!.parentNodeId)
      if (inserted.has(current)) return true;
    return false;
  };
  const listIds = source
    .nodeIds()
    .filter((id) => ["ol", "ul"].includes(source.peekNode(id)!.tagName) && inside(id));
  if (!listIds.length) return;
  const css = styles.resolver();
  const projected = new Map<string, SourceStyleNode>();
  const project = (id: string): SourceStyleNode => {
    let value = projected.get(id);
    if (!value) {
      const node = source.peekNode(id)!;
      value = {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
      projected.set(id, value);
    }
    return value;
  };
  const formatMap: Readonly<Record<string, WordNumberLevel["format"]>> = {
    decimal: "decimal",
    "lower-alpha": "lowerLetter",
    "upper-alpha": "upperLetter",
    "lower-roman": "lowerRoman",
    "upper-roman": "upperRoman",
    disc: "bullet",
    circle: "bullet",
    square: "bullet",
    none: "none",
  };
  const counter = (raw: string): number => {
    const value = Number(raw);
    if (!Number.isSafeInteger(value) || value < 0 || value > 2147483647)
      throw new DocumentError(
        "UnsupportedFeature",
        "A pasted list counter exceeds native Word numbering.",
      );
    return value;
  };
  for (const listId of listIds) {
    const list = source.peekNode(listId)!;
    if (list.attributes.reversed !== undefined)
      throw new DocumentError(
        "UnsupportedFeature",
        "Reversed HTML lists cannot be represented as editable Word numbering.",
      );
    let depth = 0;
    for (let parent = list.parentNodeId; parent; parent = source.peekNode(parent)!.parentNodeId)
      if (["ol", "ul"].includes(source.peekNode(parent)!.tagName) && inside(parent)) depth++;
    if (depth > 8)
      throw new DocumentError("UnsupportedFeature", "Word lists support at most nine levels.");
    const listStyle = css.get(project(listId));
    const rawFormat = listStyle["list-style-type"] ?? (list.tagName === "ol" ? "decimal" : "disc");
    const format = formatMap[rawFormat];
    if (!format)
      throw new DocumentError(
        "UnsupportedFeature",
        "The pasted HTML list marker has no native Word representation.",
      );
    let start = counter(list.attributes.start ?? "1");
    let definition: WordMetadata["numbering"][number] | undefined;
    let previousAppearance: string | undefined;
    for (const itemId of list.childNodeIds) {
      const item = source.peekNode(itemId)!;
      if (item.tagName !== "li") continue;
      if (list.tagName === "ol" && item.attributes.value !== undefined) {
        start = counter(item.attributes.value);
        definition = undefined;
      }
      const paragraphId = edit.sourceListItemParagraph(itemId);
      // Explicit native paragraph numbering is stronger than the surrounding
      // HTML list. Never continue a receiver's unrelated numbering instance.
      const existing = item.attributes["data-word-numbering"];
      const paragraph = source.peekNode(paragraphId);
      const disabled =
        item.attributes["data-word-numbering-off"] ??
        paragraph?.attributes["data-word-numbering-off"];
      if (!existing && !paragraph?.attributes["data-word-numbering"] && !disabled) {
        const appearance = Object.fromEntries(
          Object.entries(css.get(project(itemId))).filter(([key]) =>
            /^(?:font-|color$|letter-spacing$|word-spacing$|text-decoration|--word-(?:font|run|kerning|language|bold|italic))/.test(
              key,
            ),
          ),
        );
        const signature = JSON.stringify(appearance);
        if (previousAppearance !== signature) definition = undefined;
        previousAppearance = signature;
        if (!definition) {
          const cssClass = "word-list-content-" + createOpaqueId();
          styles.setClass(cssClass, appearance);
          definition = {
            numberingId: "word-list-content-" + createOpaqueId(),
            levels: [
              {
                level: depth,
                start,
                format,
                suffix: "tab",
                cssClass,
                text:
                  format === "none"
                    ? ""
                    : format === "bullet"
                      ? rawFormat === "square"
                        ? "▪"
                        : rawFormat === "circle"
                          ? "○"
                          : "•"
                      : "%" + (depth + 1) + ".",
              },
            ],
          };
          word.numbering.push(definition);
        }
        edit.attributes(paragraphId, {
          "data-word-numbering": definition.numberingId,
          "data-word-level": String(depth),
        });
        edit.sourceStyle(paragraphId, "margin-left", (depth + 1) * 24 + "pt");
        edit.sourceStyle(paragraphId, "text-indent", "-18pt");
      } else if (existing && paragraphId !== itemId) {
        edit.attributes(itemId, { "data-word-numbering": undefined, "data-word-level": undefined });
        edit.attributes(paragraphId, {
          "data-word-numbering": existing,
          "data-word-level": item.attributes["data-word-level"] ?? "0",
        });
      } else if (disabled && paragraphId !== itemId) {
        edit.attributes(itemId, { "data-word-numbering-off": undefined });
        edit.attributes(paragraphId, { "data-word-numbering-off": disabled });
      }
      start++;
    }
    edit.renameSource(listId, "div");
    edit.attributes(listId, { start: undefined, reversed: undefined, type: undefined });
  }
}
