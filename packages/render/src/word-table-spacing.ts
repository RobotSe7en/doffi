import { SourceStyleResolver, DEFAULT_DOCUMENT_READ_LIMITS } from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { ownViewStyle } from "./view-nodes.js";
import { wordTablePropertyResolver } from "./word-table-properties.js";

// Word expects the native preferred table width to fit a signed short in twips.
const maxNativeWidthPt = 32767 / 20;
const numeric = (value: string): number => parseFloat(value) || 0;

/** In simple legacy AutoFit tables, automatic cells can enlarge a track
 * beyond the saved grid. Fit the preferred tracks towards their measured
 * minimum content widths when the table has no room to grow. */
function fitLegacyAutomaticColumns(
  table: HTMLTableElement,
  columns: readonly HTMLTableColElement[],
  grid: readonly number[],
  isAutomatic: (cell: HTMLTableCellElement) => boolean,
): void {
  const rows = [...table.rows];
  if (
    columns.length < 2 || columns.length !== grid.length ||
    rows.some((row) =>
      row.cells.length !== columns.length ||
      [...row.cells].some((cell) => cell.colSpan !== 1 || cell.rowSpan !== 1),
    ) ||
    table.querySelector("table,img,[data-word-drawing],[data-word-native-object],[data-word-grid-gap]")
  ) return;
  const cells = rows.flatMap((row) => [...row.cells].map((cell, column) => ({
    cell,
    column,
    automatic: isAutomatic(cell),
  })));
  if (!cells.some((entry) => entry.automatic) || !cells.some((entry) => !entry.automatic)) return;

  const document = table.ownerDocument,
    view = document.defaultView!,
    available = rows[0]!.getBoundingClientRect().width,
    preferred = grid.map((width) => width * 4 / 3),
    maximum = preferred.slice(),
    minimum = grid.map(() => 0);
  if (available <= 0) return;
  const probes = cells.map(({ cell, column, automatic }) => {
    const css = view.getComputedStyle(cell),
      padding = numeric(css.paddingLeft) + numeric(css.paddingRight),
      probe = document.createElement("div");
    probe.style.cssText =
      "position:absolute;visibility:hidden;pointer-events:none;margin:0;padding:0;border:0;width:max-content";
    probe.append(...[...cell.childNodes].map((node) => node.cloneNode(true)));
    return { cell, column, automatic, padding, probe };
  });
  try {
    for (const { cell, probe } of probes) cell.append(probe);
    for (const { automatic, column, padding, probe } of probes)
      if (automatic)
        maximum[column] = Math.max(maximum[column]!, probe.getBoundingClientRect().width + padding);
    if (!maximum.some((width, column) => width > preferred[column]! + 0.1)) return;
    for (const { probe } of probes) probe.style.width = "min-content";
    for (const { column, padding, probe } of probes) {
      minimum[column] = Math.max(minimum[column]!, probe.getBoundingClientRect().width + padding);
      maximum[column] = Math.max(maximum[column]!, minimum[column]!);
    }
  } finally {
    for (const { probe } of probes) probe.remove();
  }
  const totalMin = minimum.reduce((width, value) => width + value, 0),
    totalMax = maximum.reduce((width, value) => width + value, 0);
  if (totalMin >= available || totalMax <= available) return;
  const fraction = (available - totalMin) / (totalMax - totalMin);
  for (const [index, column] of columns.entries()) {
    const width = minimum[index]! + (maximum[index]! - minimum[index]!) * fraction;
    column.style.width = (width / available) * 100 + "%";
  }
}

/** Word's legacy grid is laid out between border centres. Recover the space
 * the browser charges for collapsed borders inside that same grid. */
function legacyCenteredGrid(
  table: HTMLTableElement,
  columns: readonly HTMLTableColElement[],
  grid: readonly number[],
  isAutomatic: (cell: HTMLTableCellElement) => boolean,
): void {
  const document = table.ownerDocument,
    view = document.defaultView!,
    tableBounds = table.getBoundingClientRect(),
    firstRow = table.rows[0];
  if (!columns.length || !firstRow?.cells.length) return;
  // Table rows span the rendered grid. Some engines expose empty or
  // partial geometry on the non-rendered <col> elements.
  const gridWidth = firstRow.getBoundingClientRect().width;
  if (!gridWidth) return;
  const leftCell = view.getComputedStyle(firstRow.cells[0]!),
    rightCell = view.getComputedStyle(firstRow.cells[firstRow.cells.length - 1]!),
    outsideBorder = Math.max(0, tableBounds.width - gridWidth),
    allowance = numeric(leftCell.paddingLeft) + numeric(rightCell.paddingRight) + outsideBorder,
    width = `calc(${grid.reduce((sum, column) => sum + column, 0)}pt + ${outsideBorder}px)`,
    cap = `calc(100% + ${allowance}px)`;
  ownViewStyle(table, "width", "max-width", "margin-left");
  table.style.width = width;
  table.style.maxWidth = cap;
  table.style.width = `min(${width}, ${cap})`;
  table.style.marginLeft = `calc((100% - min(${width}, ${cap})) / 2)`;
  fitLegacyAutomaticColumns(table, columns, grid, isAutomatic);

  const probes = [...table.rows].flatMap((row) =>
    [...row.cells].flatMap((cell) => {
      const style = view.getComputedStyle(cell),
        left = numeric(style.paddingLeft),
        right = numeric(style.paddingRight);
      return left || right ? [{ cell, left, right, probe: document.createElement("span") }] : [];
    }),
  );
  // Negative margins cancel existing padding for the measurement. This leaves
  // the actual border insets, including borders inherited from adjacent cells.
  for (const { cell, left, right, probe } of probes) {
    probe.style.cssText =
      `display:block;height:0;margin:0 ${-right}px 0 ${-left}px;padding:0;border:0;line-height:0`;
    cell.append(probe);
  }
  // Allow half a native twip for CSS layout rounding at a fullwidth break.
  const paddingInsideBorder = (padding: number, inset: number): number =>
    inset > 0 ? Math.max(0, padding - inset - 1 / 30) : padding;
  const adjusted = probes.map(({ cell, left, right, probe }) => {
    const bounds = cell.getBoundingClientRect(),
      inside = probe.getBoundingClientRect();
    return {
      left: paddingInsideBorder(left, inside.left - bounds.left),
      right: paddingInsideBorder(right, bounds.right - inside.right),
    };
  });
  probes.forEach(({ cell, probe, left, right }, index) => {
    probe.remove();
    const padding = adjusted[index]!;
    if (left - padding.left > 0.01) {
      ownViewStyle(cell, "padding-left");
      cell.style.paddingLeft = padding.left + "px";
    }
    if (right - padding.right > 0.01) {
      ownViewStyle(cell, "padding-right");
      cell.style.paddingRight = padding.right + "px";
    }
  });
}

/** Native spacing is authoritative over an earlier CSS projection. Nested
 * tables resolve their own spacing instead of inheriting a browser variable. */
export function composeWordTableSpacing(document: Document, model: ContentDocumentModel): void {
  const snapshot = model.snapshot(),
    tables = new Map(snapshot.metadata.word?.tables.map((table) => [table.nodeId, table] as const)),
    mode = snapshot.metadata.word?.compatibilityMode,
    legacy = mode === 11 || mode === 12 || mode === 14,
    view = document.defaultView!;
  if (!tables.size) return;
  const css = new SourceStyleResolver(
      snapshot.styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    project = (id: string): SourceStyleNode => {
      const node = model.getNode(id)!;
      return {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
    };
  const effective = wordTablePropertyResolver(model);
  for (const table of document.querySelectorAll<HTMLTableElement>("table")) {
    const id = table.dataset.nodeId,
      info = id ? tables.get(id) : undefined;
    if (!id || !info) continue;
    const computed = effective(table),
      spacing = computed["--word-table-cell-spacing"],
      fixed = computed["table-layout"] === "fixed",
      width = computed.width,
      centered = computed["margin-left"] === "auto" && computed["margin-right"] === "auto",
      gridWidth = info.grid.reduce((a, b) => a + b, 0);
    ownViewStyle(table, "border-spacing");
    table.style.borderSpacing = spacing ?? computed["border-spacing"] ?? "0pt";
    if (spacing !== undefined) {
      ownViewStyle(table, "border-collapse");
      table.style.borderCollapse =
        numeric(view.getComputedStyle(table).borderSpacing) > 0 ? "separate" : "collapse";
    }
    const columns = [...table.querySelectorAll<HTMLTableColElement>(":scope > colgroup > col")];
    if (fixed && centered && width?.endsWith("pt") && numeric(width) <= maxNativeWidthPt) {
      // The native center is unchanged when the grid extends into the margins.
      // CSS auto margins instead pin an overwide table to the leading margin.
      const measuredWidth = table.getBoundingClientRect().width;
      if (measuredWidth > 0) {
        ownViewStyle(table, "margin-left");
        table.style.marginLeft = `calc((100% - ${measuredWidth}px) / 2)`;
      }
    }
    if (!fixed) {
      for (const [index, column] of columns.entries())
        column.style.width = (info.grid[index]! / gridWidth) * 100 + "%";
    }
    if (
      (width?.endsWith("pt") && parseFloat(width) > maxNativeWidthPt) ||
      (!fixed && (!width || width === "auto"))
    ) {
      ownViewStyle(table, "width");
      table.style.width = gridWidth + "pt";
      if (!fixed) {
        ownViewStyle(table, "max-width");
        table.style.maxWidth = "100%";
        // WebKit can ignore max-width during auto table layout.
        table.style.width = `min(${gridWidth}pt, 100%)`;
      }
      if (
        legacy && !fixed && (!width || width === "auto") &&
        centered &&
        computed["--word-table-bidi"] !== "1" &&
        !table.parentElement?.closest("td,th") &&
        view.getComputedStyle(table).borderCollapse === "collapse"
      )
        legacyCenteredGrid(table, columns, info.grid, (cell) => {
          const id = cell.dataset.nodeId;
          if (!id || !model.getNode(id)) return false;
          const preferred = css.get(project(id)).width;
          return !preferred || preferred === "auto" || /^0(?:\.0+)?(?:pt|%)?$/.test(preferred);
        });
    }
  }
}
