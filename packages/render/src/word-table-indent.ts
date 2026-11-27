import type { ContentDocumentModel } from "@doffi/core";
import { ownViewStyle } from "./view-nodes.js";
import { wordTablePropertyResolver } from "./word-table-properties.js";

/** Pre-2013 Word measures native table indentation at the leading cell text.
 * Keep the native value in source; project it to the physical CSS border edge. */
export function composeWordTableIndent(document: Document, model: ContentDocumentModel): void {
  const snapshot = model.snapshot(), word = snapshot.metadata.word;
  if (!word || (word.compatibilityMode ?? 12) >= 15) return;
  const view = document.defaultView!;
  const resolve = wordTablePropertyResolver(model);
  for (const table of document.querySelectorAll<HTMLTableElement>("table[data-node-id]")) {
    // This is a table's native semantic, never an inherited instruction for
    // nested tables which may still own ordinary physical CSS margins.
    const source = resolve(table);
    if (source["--word-table-indent-origin"] !== "native") continue;
    if (source["margin-left"] === "auto" || source["margin-right"] === "auto") continue;
    const computed = view.getComputedStyle(table);
    if (computed.getPropertyValue("--word-table-indent-origin").trim() !== "native"
      || computed.getPropertyValue("--word-table-floating").trim() === "1") continue;
    const cell = table.rows[0]?.cells[0];
    if (!cell || cell.dataset.wordGridGap) continue;
    const rtl = table.dataset.sdkWordTableBidi === "true";
    const margin = rtl ? "margin-right" : "margin-left", padding = rtl ? "padding-right" : "padding-left";
    const raw = computed.getPropertyValue(margin);
    if (raw === "auto") continue;
    const indent = parseFloat(raw), inset = parseFloat(view.getComputedStyle(cell).getPropertyValue(padding));
    if (!Number.isFinite(indent) || !Number.isFinite(inset)) continue;
    ownViewStyle(table, margin);
    table.style.setProperty(margin, (indent - inset) + "px");
  }
}
