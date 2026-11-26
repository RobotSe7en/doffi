import { SourceStyleResolver, DEFAULT_DOCUMENT_READ_LIMITS } from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { ownViewStyle, ownViewAttribute } from "./view-nodes.js";
import { wordTablePropertyResolver } from "./word-table-properties.js";

/** bidiVisual mirrors table geometry and logical sides, while paragraphs keep
 * their own writing direction and UTF-16 order. All changes are view-owned. */
export function composeWordTableBidi(document: Document, model: ContentDocumentModel): void {
  const word = model.snapshot().metadata.word;
  if (!word?.tables.length) return;
  const css = new SourceStyleResolver(
      model.snapshot().styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    view = document.defaultView!,
    project = (id: string): SourceStyleNode => {
      const node = model.getNode(id)!;
      return {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
    };
  const effective = wordTablePropertyResolver(model);
  const tables = [...document.querySelectorAll<HTMLTableElement>("table")].flatMap((table) => {
    const id = table.dataset.nodeId;
    if (!id || !word.tables.some((item) => item.nodeId === id)) return [];
    const source = project(id),
      bidi =
        css.specified(source)["--word-table-bidi"] !== undefined &&
        css.get(source)["--word-table-bidi"] === "1",
      computed = effective(table);
    return [
      {
        table,
        bidi,
        marginLeft: computed["margin-left"] ?? "0pt",
        marginRight: computed["margin-right"] ?? "0pt",
        paint: [table, ...table.querySelectorAll<HTMLElement>("td,th")]
          .filter((node) => node === table || node.closest("table") === table)
          .map((node) => {
            const computed = view.getComputedStyle(node),
              properties = ["padding", "border"].flatMap((prefix) =>
                prefix === "padding"
                  ? ["padding-left", "padding-right"]
                  : ["width", "style", "color"].flatMap((part) => [
                      "border-left-" + part,
                      "border-right-" + part,
                    ]),
              );
            return {
              node,
              direction: computed.direction,
              values: Object.fromEntries(
                properties.map((key) => [key, computed.getPropertyValue(key)]),
              ),
            };
          }),
      },
    ];
  });
  for (const { table, bidi, marginLeft, marginRight, paint } of tables) {
    ownViewStyle(table, "direction");
    table.style.direction = bidi ? "rtl" : "ltr";
    ownViewAttribute(table, "data-sdk-word-table-bidi");
    table.dataset.sdkWordTableBidi = String(bidi);
    if (!bidi) continue;
    ownViewStyle(table, "margin-left", "margin-right");
    table.style.marginLeft = marginLeft === "auto" ? marginRight : "auto";
    table.style.marginRight = marginLeft === "auto" ? "auto" : marginLeft;
    for (const { node, direction, values } of paint) {
      if (node !== table) {
        ownViewStyle(node, "direction");
        node.style.direction = direction;
        const content = [...node.children].find((child) =>
          child.classList.contains("sdk-word-row-spacing-content"),
        ) as HTMLElement | undefined;
        if (content) {
          const left = content.style.marginLeft;
          content.style.marginLeft = content.style.marginRight;
          content.style.marginRight = left;
        }
      }
      for (const [key, value] of Object.entries(values)) {
        const opposite = key.replace(/left|right/, (side) => (side === "left" ? "right" : "left"));
        ownViewStyle(node, opposite);
        node.style.setProperty(opposite, value);
      }
    }
  }
}
