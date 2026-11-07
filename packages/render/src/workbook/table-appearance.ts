import { createWorkbookTableFormatResolver, DocumentError } from "@doffi/core";
import type { WorkbookCellFormat } from "@doffi/core";
import type { DeepReadonly, WorkbookCoordinate, WorkbookRange, WorkbookSheet, WorkbookSnapshot, WorkbookStyleApplication, WorkbookTable } from "@doffi/core";
import { AxisDimensions } from "./dimensions.js";
import { WorkbookMergeIndex } from "./projection.js";
import { WorkbookTableRegions, workbookTableRegionBorders } from "./table-regions.js";
import { workbookBuiltinTableStyle } from "./table-styles.js";
import type { WorkbookBuiltinTableColor, WorkbookBuiltinTableDifferential } from "./table-style-types.js";
import { WorkbookStyleLayers, workbookStyleDifferences } from "./style-layers.js";
import type { WorkbookStyleProperties } from "./style-layers.js";
import type { WorkbookSourceCell } from "./cell-source.js";

type Range = DeepReadonly<WorkbookRange>;
type Table = DeepReadonly<WorkbookTable>;
type Differential = { readonly cssClass: string; readonly numberFormatId?: string | undefined };
type Layer = { readonly range: Range; readonly style: WorkbookBuiltinTableDifferential | Differential; readonly borderOnly?: boolean; readonly apply?: DeepReadonly<WorkbookStyleApplication> };
const themeSlots = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
function color(target: Map<string, string>, prefix: string, value: WorkbookBuiltinTableColor): void {
  target.set(prefix + "-color", `var(--office-theme-${themeSlots[value.theme]})`); target.set(prefix + "-tint", String(value.tint));
}
function builtinProperties(style: WorkbookBuiltinTableDifferential): WorkbookStyleProperties {
  const result = new Map<string, string>();
  if (style.font?.bold) result.set("font-weight", "bold");
  if (style.font?.color) { color(result, "--office-cell-foreground", style.font.color); result.set("color", "var(--office-cell-foreground-color)"); }
  if (style.fill) { result.set("--office-cell-fill-kind", "solid"); color(result, "--office-cell-fill", style.fill); }
  for (const [side, border] of Object.entries(style.borders ?? {})) {
    result.set(`--office-cell-border-${side}-kind`, border.style); if (border.color) color(result, `--office-cell-border-${side}`, border.color);
  }
  return result;
}
function namedProperty(name: string, apply: DeepReadonly<WorkbookStyleApplication>): boolean {
  if (name.startsWith("--office-cell-fill-") || name.startsWith("--office-cell-gradient-") || name === "background-color") return apply.fill;
  if (name.startsWith("--office-cell-border-") || name.startsWith("border-")) return apply.border;
  if (name.startsWith("--office-font-") || name.startsWith("--office-cell-foreground-") || name.startsWith("font-") || name === "color"
    || name.startsWith("text-decoration-") || name === "vertical-align" || name === "letter-spacing" || name === "word-spacing") return apply.font;
  return apply.alignment;
}

/** Resolve region border slots before they enter the ordinary cell border
 * reader. A stripe's outside border is not an outline around every cell. */
function regionalProperties(properties: WorkbookStyleProperties, range: Range, coordinate: WorkbookCoordinate, rightToLeft: boolean, borderOnly = false): Map<string, string> {
  const result = new Map([...properties].filter(([name]) => !borderOnly && !name.startsWith("--office-cell-border-") && !name.startsWith("border-")));
  const slots: Partial<Record<"left" | "right" | "top" | "bottom" | "horizontal" | "vertical", Map<string, string>>> = {};
  for (const [name, value] of properties) {
    const match = /^--office-cell-border-(left|right|top|bottom|horizontal|vertical)(-.+)$/.exec(name);
    if (match) { const side = match[1] as keyof typeof slots; (slots[side] ??= new Map()).set(match[2]!, value); }
  }
  const borders = workbookTableRegionBorders(range, coordinate, slots);
  for (const [side, values] of Object.entries(borders)) for (const [suffix, value] of values) {
    const native = side === "before" ? "left" : side === "after" ? "right" : side;
    result.set(`--office-cell-border-${native}${suffix}`, value);
  }
  for (const side of ["left", "right", "top", "bottom"] as const) {
    const outside = side === "left" ? coordinate.column === (rightToLeft ? range.to.column : range.from.column) : side === "right" ? coordinate.column === (rightToLeft ? range.from.column : range.to.column)
      : side === "top" ? coordinate.row === range.from.row : coordinate.row === range.to.row;
    if (!outside) continue;
    for (const kind of ["style", "width", "color"]) {
      const property = `border-${side}-${kind}`, value = properties.get(property);
      if (value !== undefined) {
        result.set(property, value);
        const native = rightToLeft && side === "left" ? "right" : rightToLeft && side === "right" ? "left" : side;
        result.set(`--office-cell-border-${native}-kind`, "initial");
        if (native === "left" || native === "right") result.set(`--office-cell-border-${native === "left" ? "start" : "end"}-kind`, "initial");
      }
    }
  }
  return result;
}

/** Window-only composition over shared source styles. Neither the snapshot nor
 * its HTML/CSS is modified; only temporary source measurement cells receive
 * the final property overrides. The region index never enumerates table cells. */
export class WorkbookTableAppearance {
  private readonly tables: WorkbookMergeIndex;
  private readonly byRange: ReadonlyMap<Range, Table>;
  private readonly regions = new Map<Table, WorkbookTableRegions<WorkbookBuiltinTableDifferential | Differential>>();
  private readonly styles: WorkbookStyleLayers;
  private readonly visibleRows: AxisDimensions;
  private readonly visibleColumns: AxisDimensions;
  private readonly definitions: ReadonlyMap<string, DeepReadonly<NonNullable<WorkbookSnapshot["metadata"]["workbook"]["cellStyles"]>["definitions"][number]>>;
  private readonly namedStyles: ReadonlyMap<string, string>;
  private units = 0;
  private readonly format: (at: WorkbookCoordinate, source: WorkbookCellFormat) => WorkbookCellFormat;
  constructor(private readonly snapshot: WorkbookSnapshot, private readonly sheet: DeepReadonly<WorkbookSheet>, document: Document,
    account: (amount: number) => void, private readonly maxCodeUnits: number = 8 * 1024 * 1024) {
    const tables = sheet.tables ?? [];
    this.tables = new WorkbookMergeIndex(tables.map(table => table.range)); this.byRange = new Map(tables.map(table => [table.range, table]));
    this.styles = new WorkbookStyleLayers(document, account);
    this.definitions = new Map(snapshot.metadata.workbook.cellStyles?.definitions.map(definition => [definition.styleId, definition]));
    this.namedStyles = new Map(snapshot.metadata.workbook.cellStyles?.names.map(name => [name.name.toUpperCase(), name.styleId]));
    this.visibleRows = new AxisDimensions(1_048_576, sheet.defaultRowsHidden ? 0 : 1, sheet.rowDimensions.map(run => ({ from: run.from, to: run.to,
      sizePt: (run.hidden ?? sheet.defaultRowsHidden) || run.heightPt === 0 ? 0 : 1 })));
    this.visibleColumns = new AxisDimensions(16_384, 1, sheet.columnDimensions.map(run => ({ from: run.from, to: run.to, sizePt: run.hidden || run.width === 0 ? 0 : 1 })));
    this.format = createWorkbookTableFormatResolver(snapshot, sheet);
  }
  private layers(table: Table, coordinate: WorkbookCoordinate): readonly Layer[] {
    const result: Layer[] = [], { row, column } = coordinate;
    let regions = this.regions.get(table);
    if (!regions) {
      const reference = table.style, elements = reference?.kind === "builtin" ? workbookBuiltinTableStyle(reference.name).elements
        : reference?.kind === "custom" ? this.snapshot.metadata.workbook.tableStyles?.find(style => style.styleId === reference.styleId)?.elements ?? [] : [];
      regions = new WorkbookTableRegions<WorkbookBuiltinTableDifferential | Differential>(table, elements, this.visibleRows); this.regions.set(table, regions);
    }
    for (const region of regions.at(coordinate)) result.push({ range: region.range, style: region.style });
    const header = table.header && row === table.range.from.row, total = table.totals && row === table.range.to.row;
    const key = header ? "headerStyle" : total ? "totalStyle" : "dataStyle", namedKey = header ? "headerCellStyleName" : total ? "totalCellStyleName" : "dataCellStyleName";
    const range: Range = { from: { row: header || total ? row : table.range.from.row + Number(table.header), column: table.range.from.column },
      to: { row: header || total ? row : table.range.to.row - Number(table.totals), column: table.range.to.column } };
    const tableColumn = table.columns[column - table.range.from.column]!;
    for (const source of [table, tableColumn]) {
      const definition = this.definitions.get(this.namedStyles.get(source[namedKey]?.toUpperCase() ?? "") ?? "");
      if (definition) result.push({ range, style: definition, apply: definition.apply });
      if (source[key]) result.push({ range: source === table ? range : { from: { row: range.from.row, column }, to: { row: range.to.row, column } }, style: source[key]! });
    }
    if (table.borderStyle) result.push({ range: table.range, style: table.borderStyle, borderOnly: true });
    const border = header ? table.headerBorderStyle : total ? table.totalBorderStyle : undefined;
    if (border) result.push({ range, style: border, borderOnly: true });
    return result;
  }
  apply(coordinate: WorkbookCoordinate, source: WorkbookSourceCell): WorkbookSourceCell {
    const range = this.tables.at(coordinate.row, coordinate.column), table = range ? this.byRange.get(range) : undefined;
    if (!table || !this.visibleRows.size(coordinate.row) || !this.visibleColumns.size(coordinate.column)) return source;
    const layers = this.layers(table, coordinate); if (!layers.length) return source;
    const normal = this.styles.read(source.element, this.sheet.defaultCssClass), direct = this.styles.read(source.element, source.format.cssClass);
    const patch = new Map<string, string>();
    for (const layer of layers) {
      const properties = "cssClass" in layer.style ? this.styles.read(source.element, layer.style.cssClass, true) : builtinProperties(layer.style);
      const values = layer.apply ? new Map([...properties].filter(([name]) => namedProperty(name, layer.apply!))) : properties;
      const visibleRange: Range = { from: { row: this.visibleRows.hit(this.visibleRows.position(layer.range.from.row) + .5)!, column: this.visibleColumns.hit(this.visibleColumns.position(layer.range.from.column) + .5)! },
        to: { row: this.visibleRows.hit(this.visibleRows.position(layer.range.to.row + 1) - .5)!, column: this.visibleColumns.hit(this.visibleColumns.position(layer.range.to.column + 1) - .5)! } };
      for (const [name, value] of layer.apply ? values : regionalProperties(values, visibleRange, coordinate, this.sheet.rightToLeft, layer.borderOnly)) patch.set(name, value);
    }
    // Current native XFs are full records. Their default properties must not
    // erase table formatting, and their apply* propagation flags are unrelated.
    for (const [name, value] of workbookStyleDifferences(direct, normal)) patch.set(name, value);
    for (const [name, value] of patch) {
      this.units += name.length + value.length + 2;
      if (this.units > this.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Composed workbook appearance exceeds its window style budget.");
      source.element.style.setProperty(name, value, "important");
    }
    return Object.freeze({ element: source.element, format: this.format(coordinate, source.format) });
  }
}
