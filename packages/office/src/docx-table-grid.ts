import { DocumentError, resolveSourceCssLength } from "@doffi/core";
import type { PageSetup } from "@doffi/core";
import type { DocxStyles } from "./ooxml/styles.js";
import { attr, child, check, elements, fail, number } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import { measurePoints } from "./ooxml/measure.js";
import { readWordRowGapCount } from "./docx-row-gaps.js";

/** Column widths are preferred geometry. Missing/zero native columns are
 * inferred from real cells and omissions, rather than preserving a whole table
 * as an opaque object. No text or auxiliary native source is synthesized. */
export function readEditableWordGrid(
  source: XmlNode,
  styles: DocxStyles,
  available: number,
  maxCells: number,
): number[] {
  const properties = child(source, "tblPr"),
    gridNode = child(source, "tblGrid");
  check(gridNode, ["gridCol", "tblGridChange"]);
  const declared = elements(gridNode, "gridCol").map((node) => {
    const width = measurePoints(attr(node, "w"), node, 20, 0);
    if (width < 0) fail(node, "A native table grid width cannot be negative.");
    return width;
  });
  const tableWidth = child(properties, "tblW"),
    preferred = tableWidth ? styles.width(tableWidth, "width").width : undefined;
  const requested =
    preferred && preferred !== "auto"
      ? resolveSourceCssLength(preferred, { font: 12, percent: available })
      : available;
  const target = requested > 0 ? requested : available;
  const rows = elements(source, "tr"),
    constraints: { start: number; span: number; width: number }[] = [];
  let columns = declared.length;
  const addWidth = (node: XmlNode | undefined, start: number, span: number) => {
    if (!node || !span) return;
    const css = styles.width(node, "width").width;
    if (css)
      constraints.push({
        start,
        span,
        width: resolveSourceCssLength(css, { font: 12, percent: target }),
      });
  };
  for (const row of rows) {
    const props = child(row, "trPr"),
      before = readWordRowGapCount(props, "before"),
      after = readWordRowGapCount(props, "after");
    let column = before;
    addWidth(child(props, "wBefore"), 0, before);
    for (const cell of elements(row, "tc")) {
      const props = child(cell, "tcPr"),
        span = number(attr(child(props, "gridSpan"), "val"), cell, 1);
      if (span < 1 || !Number.isSafeInteger(column + span))
        fail(cell, "A table cell span is invalid.");
      addWidth(child(props, "tcW"), column, span);
      column += span;
    }
    addWidth(child(props, "wAfter"), column, after);
    columns = Math.max(columns, column + after);
    if (columns > maxCells || columns * Math.max(1, rows.length) > maxCells)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Inferred table geometry exceeds its cell budget.",
      );
  }
  if (!columns) fail(source, "A table requires a nonempty cell grid.");
  if (declared.length === columns && declared.every((width) => width > 0)) return declared;
  const grid = Array.from({ length: columns }, (_, index) => declared[index] ?? 0),
    inferred = grid.map((width) => width === 0);
  for (const item of constraints) {
    let fixed = 0,
      missing = 0;
    for (let c = item.start; c < item.start + item.span; c++)
      if (inferred[c]) missing++;
      else fixed += grid[c]!;
    if (!missing) continue;
    const width = Math.max(0, item.width - fixed) / missing;
    for (let c = item.start; c < item.start + item.span; c++)
      if (inferred[c]) grid[c] = Math.max(grid[c]!, width);
  }
  const empty = grid.filter((width) => !width).length,
    sum = grid.reduce((a, b) => a + b, 0);
  const fallback = Math.max(1, (Math.max(target, sum) - sum) / (empty || 1));
  for (let c = 0; c < grid.length; c++) if (!grid[c]) grid[c] = fallback;
  // A fixed preferred table width distributes unresolved grid hints. The
  // resulting positive grid is editable and stable after native compilation.
  if (preferred && target > 0) {
    const total = grid.reduce((a, b) => a + b, 0);
    for (let c = 0; c < grid.length; c++) grid[c] = (grid[c]! * target) / total;
  }
  return grid;
}

export function wordSectionTableWidth(
  properties: XmlNode | undefined,
  fallback?: PageSetup,
): number {
  const size = child(properties, "pgSz"),
    margins = child(properties, "pgMar"),
    columns = child(properties, "cols");
  const width = size ? number(attr(size, "w"), size) / 20 : (fallback?.width ?? 612);
  const left = margins
    ? number(attr(margins, "left"), margins) / 20
    : (fallback?.margins.left ?? 72);
  const right = margins
    ? number(attr(margins, "right"), margins) / 20
    : (fallback?.margins.right ?? 72);
  const gutter = margins ? number(attr(margins, "gutter"), margins, 0) / 20 : 0;
  const explicit = elements(columns, "col");
  if (explicit.length) return number(attr(explicit[0], "w"), explicit[0]!) / 20;
  const count = columns ? number(attr(columns, "num"), columns, 1) : 1,
    gap = columns ? number(attr(columns, "space"), columns, 720) / 20 : 0;
  if (count < 1) fail(columns!, "A section must have a positive column count.");
  return Math.max(1, (width - left - right - gutter - gap * (count - 1)) / count);
}
