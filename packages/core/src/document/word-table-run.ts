import type { DeepReadonly } from "../runtime.js";
import type { WordMetadata } from "../contracts/word.js";
import type { ParsedHtml } from "./html.js";
import type { SourceStyleResolver } from "./source-css.js";
import { DocumentError } from "../errors.js";
import { resolveWordTableStyle, wordTableRegions } from "../word-style-context.js";
/** Resolve one cell's current conditional run baseline. Intervals retain
 * vertical occupancy without expanding merged cells into a dense grid. */
export function wordTableRunDefaults(
  word: DeepReadonly<Pick<WordMetadata, "styles"> & Partial<Pick<WordMetadata, "tables">>>,
  html: Pick<ParsedHtml, "peekNode">,
  nodeId: string,
  resolver: SourceStyleResolver,
  maxWork: number,
): Record<string, string> {
  let cell: string | undefined,
    table = html.peekNode(nodeId);
  while (table) {
    if (!cell && ["td", "th"].includes(table.tagName)) cell = table.nodeId;
    if (table.tagName === "table") break;
    table = table.parentNodeId ? html.peekNode(table.parentNodeId) : undefined;
  }
  const info = word.tables?.find((info) => info.nodeId === table?.nodeId);
  if (!cell || !table || !info) return {};
  const rows = table.childNodeIds.flatMap((id) => {
      const node = html.peekNode(id)!;
      return node.tagName === "tr"
        ? [id]
        : node.childNodeIds.filter((id) => html.peekNode(id)!.tagName === "tr");
    }),
    styleId = table.attributes["data-word-style"],
    settings: Record<string, number> = {},
    seen = new Set<string>();
  let style = word.styles.find((style) => style.styleId === styleId);
  while (style) {
    if (seen.has(style.styleId))
      throw new DocumentError("InvalidDocument", "Word table style inheritance is cyclic.");
    seen.add(style.styleId);
    for (const key of ["tableLookMask", "rowBandSize", "columnBandSize"] as const)
      if (settings[key] === undefined && style[key] !== undefined) settings[key] = style[key]!;
    style = word.styles.find((item) => item.styleId === style!.basedOn);
  }
  let active: { from: number; to: number; until: number }[] = [],
    work = 0;
  for (const [rowIndex, row] of rows.entries()) {
    if ((work += active.length + 1) > maxWork)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Word table default resolution exceeds its work budget.",
      );
    active = active.filter((item) => item.until > rowIndex).sort((a, b) => a.from - b.from);
    let column = 0,
      index = 0;
    const added: typeof active = [];
    for (const id of html.peekNode(row)!.childNodeIds) {
      const source = html.peekNode(id)!;
      if (!["td", "th"].includes(source.tagName)) continue;
      if (++work > maxWork)
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Word table default resolution exceeds its work budget.",
        );
      while (index < active.length && active[index]!.from <= column) {
        column = Math.max(column, active[index++]!.to);
      }
      const span = Number(source.attributes.colspan ?? 1),
        height = Number(source.attributes.rowspan ?? 1) || rows.length - rowIndex;
      if (id === cell) {
        const regions = wordTableRegions(
          rowIndex,
          column,
          rows.length,
          info.grid.length,
          span,
          info.lookMask ?? settings.tableLookMask ?? 0,
          info.rowBandSize ?? settings.rowBandSize ?? 0,
          info.columnBandSize ?? settings.columnBandSize ?? 0,
        );
        return resolveWordTableStyle(
          word,
          (name) => resolver.specifiedClass(name),
          styleId,
          regions,
        ).r;
      }
      if (height > 1) added.push({ from: column, to: column + span, until: rowIndex + height });
      column += span;
    }
    active.push(...added);
  }
  return {};
}
