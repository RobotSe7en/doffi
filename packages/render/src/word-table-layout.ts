import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WordMetadata } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { markViewNode, ownViewStyle, ownViewAttribute, viewNodeRole } from "./view-nodes.js";
/** Exact row heights clip the view; all hidden cell content remains in source. */
export function composeWordTableLayout(
  document: Document,
  limits: Readonly<RenderLimits>,
  word?: DeepReadonly<WordMetadata>,
): void {
  const view = document.defaultView!;
  const tables = new Map(word?.tables.map((table) => [table.nodeId, table] as const));
  let count = document.querySelectorAll("*").length;
  for (const gap of document.querySelectorAll<HTMLElement>("[data-word-grid-gap]")) {
    const info = tables.get(gap.closest<HTMLTableElement>("table")?.dataset.nodeId ?? ""),
      span = Number(gap.getAttribute("colspan") ?? 1);
    if (info) {
      const columns =
        gap.dataset.wordGridGap === "before" ? info.grid.slice(0, span) : info.grid.slice(-span);
      ownViewStyle(gap, "width");
      gap.style.width = columns.reduce((a, b) => a + b, 0) + "pt";
    }
    ownViewAttribute(gap, "contenteditable");
    ownViewAttribute(gap, "aria-hidden");
    gap.contentEditable = "false";
    gap.setAttribute("aria-hidden", "true");
    for (const [name, value] of Object.entries({
      visibility: "hidden",
      "pointer-events": "none",
      padding: "0",
      border: "0",
      "background-color": "transparent",
    })) {
      ownViewStyle(gap, name);
      gap.style.setProperty(name, value, "important");
    }
  }
  for (const table of document.querySelectorAll<HTMLTableElement>("table")) {
    const rows = [...table.rows],
      heights = rows.map((row) => {
        const css = view.getComputedStyle(row),
          rule = css.getPropertyValue("--word-row-height-rule").trim();
        if (rule === "auto") {
          ownViewStyle(row, "height");
          row.style.height = "auto";
          return undefined;
        }
        if (rule !== "exact") return undefined;
        const match = /^([\d.]+)(pt|px|in|cm|mm|pc|em|rem)$/.exec(row.style.height);
        if (!match) return undefined;
        const factors: Record<string, number> = {
          pt: 4 / 3,
          px: 1,
          in: 96,
          cm: 96 / 2.54,
          mm: 96 / 25.4,
          pc: 16,
          em: parseFloat(css.fontSize),
          rem: parseFloat(view.getComputedStyle(document.documentElement).fontSize),
        };
        return Number(match[1]) * factors[match[2]!]!;
      });
    if (!heights.some((height) => height !== undefined)) continue;
    const targets = [0],
      exact = [0];
    rows.forEach((row, index) => {
      targets.push(targets[index]! + (heights[index] ?? row.getBoundingClientRect().height));
      exact.push(exact[index]! + (heights[index] !== undefined ? 1 : 0));
    });
    const specs: {
      cell: HTMLTableCellElement;
      row: number;
      span: number;
      target: number;
      padding: number;
      border: number;
      alignment: string;
    }[] = [];
    for (const [index, row] of rows.entries())
      for (const cell of row.cells) {
        const span = Math.min(cell.rowSpan || rows.length - index, rows.length - index);
        if (exact[index + span] === exact[index]) continue;
        const css = view.getComputedStyle(cell);
        specs.push({
          cell,
          row: index,
          span,
          target: targets[index + span]! - targets[index]!,
          padding: (parseFloat(css.paddingTop) || 0) + (parseFloat(css.paddingBottom) || 0),
          border: (parseFloat(css.borderTopWidth) || 0) + (parseFloat(css.borderBottomWidth) || 0),
          alignment: css.verticalAlign,
        });
      }
    const clips: { box: HTMLElement; row: number; span: number; target: number; height: number }[] =
      [];
    for (const spec of specs) {
      count += 2;
      if (count > limits.maxNodes)
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Fixed row clipping exceeds the rendering node budget.",
        );
      const box = markViewNode(document.createElement("div"), "unwrap"),
        content = markViewNode(document.createElement("div"), "unwrap"),
        height = Math.max(0, spec.target - spec.padding - spec.border);
      box.className = "sdk-word-row-clip";
      box.style.cssText = `height:${height}px;min-height:0;overflow:hidden;display:flex;flex-direction:column;justify-content:${spec.alignment === "middle" ? "center" : spec.alignment === "bottom" ? "flex-end" : "flex-start"}`;
      content.style.cssText = "display:flow-root;flex:none;min-height:0";
      content.append(...spec.cell.childNodes);
      box.append(content);
      spec.cell.append(box);
      clips.push({ box, row: spec.row, span: spec.span, target: spec.target, height });
    }
    // Collapsed borders may come from the table or adjacent cells and are
    // absent from this cell's computed border widths. Correct their actual
    // contribution with bounded, batched reads and writes.
    for (let pass = 0; pass < 2; pass++) {
      const measured = [0];
      rows.forEach((row, index) =>
        measured.push(measured[index]! + row.getBoundingClientRect().height),
      );
      let changed = false;
      for (const clip of clips) {
        const excess = measured[clip.row + clip.span]! - measured[clip.row]! - clip.target;
        if (excess <= 0.05) continue;
        const height = Math.max(0, clip.height - excess);
        if (clip.height - height <= 0.05) continue;
        clip.height = height;
        clip.box.style.height = height + "px";
        changed = true;
      }
      if (!changed) break;
    }
  }
}
/** Exclude clipped geometry from pagination and source hit projections. */
export function wordRowVisibleRect(owner: Element, rect: DOMRect): DOMRect | undefined {
  let left = rect.left,
    right = rect.right,
    top = rect.top,
    bottom = rect.bottom,
    clipped = false;
  for (let node: Element | null = owner; node; node = node.parentElement)
    if (node.classList.contains("sdk-word-row-clip") && viewNodeRole(node) === "unwrap") {
      clipped = true;
      const clip = node.getBoundingClientRect();
      left = Math.max(left, clip.left);
      right = Math.min(right, clip.right);
      top = Math.max(top, clip.top);
      bottom = Math.min(bottom, clip.bottom);
    }
  if (right - left < 0.01 || bottom - top < 0.1) return;
  return clipped ? new DOMRect(left, top, right - left, bottom - top) : rect;
}

/** An explicit row override wins over the table style's repeat-header flag. */
export function wordTableRepeatHeader(row: HTMLTableRowElement): boolean {
  const direct = row.dataset.wordRepeatHeader;
  if (direct !== undefined) return direct === "true";
  return (
    row.parentElement?.localName === "thead" ||
    row.ownerDocument
      .defaultView!.getComputedStyle(row)
      .getPropertyValue("--word-row-repeat-header")
      .trim() === "1"
  );
}
