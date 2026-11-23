import { DocumentError, resolveWorkbookDisplayLocale, SourceStyleResolver, workbookBuiltinNumberFormat } from "@doffi/core";
import type { DeepReadonly, DocumentReadLimits, SourceStyleMap, SourceStyleNode, WorkbookSnapshot, WorkbookStyleApplication, WorkbookTableStyleReference } from "@doffi/core";
import type { OutputOpc } from "../export/opc.js";
import { declaration, empty } from "../export/xml.js";
import { alignment, border, fill, flag, font, pattern } from "./style-values.js";
import { bool, element, sheetContentType, sheetNamespaces, workbookPart, xstring } from "./xml.js";
import type { XlsxExportLoss } from "./options.js";

class StylesTable {
  readonly entries: string[] = [];
  private readonly indices = new Map<string, number>();
  add(xml: string): number {
    const previous = this.indices.get(xml);
    if (previous !== undefined) return previous;
    const index = this.entries.length;
    if (index >= 65_490) throw new DocumentError("ResourceLimitExceeded", "The native workbook style table is full.");
    this.entries.push(xml);
    this.indices.set(xml, index);
    return index;
  }
}
const groupNames = [["Font", "font"], ["Fill", "fill"], ["Border", "border"], ["Alignment", "alignment"],
  ["NumberFormat", "numberFormat"], ["Protection", "protection"]] as const;
const application = (style?: DeepReadonly<WorkbookStyleApplication>): Record<string, number> =>
  style ? Object.fromEntries(groupNames.map(([name, key]) => ["apply" + name, Number(style[key])])) : {};
const counted = (tag: string, items: readonly string[]): string => element(tag, items.join(""), { count: items.length });

export class WorkbookStylesWriter {
  readonly resolver: SourceStyleResolver;
  readonly theme: SourceStyleMap;
  readonly themeRaw: SourceStyleMap;
  readonly themeScope: SourceStyleNode;
  private readonly fonts = new StylesTable();
  private readonly fills = new StylesTable();
  private readonly borders = new StylesTable();
  private readonly cells = new StylesTable();
  private readonly differentials = new StylesTable();
  private readonly numberFormats = new Map<string, { id: number; code?: string }>();
  private readonly namedIds = new Map<string, number>();
  private readonly cellIndices = new Map<string, number>();
  private readonly dxfIndices = new Map<string, number>();
  private readonly bindings: ReadonlyMap<string, NonNullable<WorkbookSnapshot["metadata"]["workbook"]["cellStyles"]>["bindings"][number]>;
  private readonly named: readonly string[];
  private readonly names: readonly string[];
  private readonly normal: number;
  private readonly normalRaw: SourceStyleMap;
  private readonly normalCss: SourceStyleMap;

  private readonly reportedFormats = new Set<string>();
  readonly reportFormatLoss = (message: string): void => {
    if (!this.reportedFormats.has(message)) {
      this.reportedFormats.add(message);
      this.losses.push({ kind: "format", message });
    }
  };

  constructor(private readonly snapshot: WorkbookSnapshot, limits: DocumentReadLimits, private readonly losses: XlsxExportLoss[]) {
    this.resolver = new SourceStyleResolver(snapshot.styles.map(style => style.css), limits);
    const { workbook } = snapshot.metadata;
    this.themeScope = { tag: "section", attrs: { class: workbook.themeCssClass ?? "" }, parent: undefined };
    this.themeRaw = workbook.themeCssClass ? this.resolver.specifiedClass(workbook.themeCssClass) : {};
    this.theme = workbook.themeCssClass ? this.resolver.classStyle(workbook.themeCssClass) : {};
    for (const entry of workbook.numberFormats)
      this.numberFormats.set(entry.formatId, "builtinId" in entry ? { id: entry.builtinId } :
        /^general$/i.test(entry.code) ? { id: 0, code: entry.code } : { id: this.numberFormats.size + 164, code: entry.code });
    const definitions = workbook.cellStyles?.definitions ?? [], styles = workbook.cellStyles?.names ?? [];
    const nativeNormal = styles.find(style => style.builtinId === 0);
    const normal = nativeNormal && definitions.find(style => style.styleId === nativeNormal.styleId);
    const normalClass = normal?.cssClass ?? workbook.columnWidthFontCssClass ?? workbook.sheets[0]!.defaultCssClass;
    const normalFormat = normal?.numberFormatId ?? workbook.sheets[0]!.defaultNumberFormatId;
    this.normalRaw = this.raw(normalClass);
    this.normalCss = this.css(normalClass);
    this.fills.add(pattern("none"));
    this.fills.add(pattern("gray125"));
    this.borders.add("<border/>");
    this.fonts.add(font(this.normalRaw, this.normalCss, "cell", this.reportFormatLoss));
    this.bindings = new Map(workbook.cellStyles?.bindings.map(binding => [binding.cssClass, binding]) ?? []);
    const named = definitions.map((definition, index) => {
      this.namedIds.set(definition.styleId, index);
      return this.format(this.raw(definition.cssClass), this.css(definition.cssClass), definition.numberFormatId, undefined, definition.apply);
    });
    this.normal = nativeNormal ? this.namedIds.get(nativeNormal.styleId)! : named.length;
    if (!nativeNormal) named.push(this.format(this.normalRaw, this.normalCss, normalFormat));
    this.named = named;
    this.names = styles.map(style => empty("cellStyle", {
      name: xstring(style.name), xfId: this.namedIds.get(style.styleId), builtinId: style.builtinId, iLevel: style.level,
      hidden: bool(style.hidden), customBuiltin: bool(style.customBuiltin),
    }));
    if (!nativeNormal) {
      const existing = new Set(styles.map(style => style.name.toUpperCase()));
      let name = "Normal", suffix = 2;
      while (existing.has(name.toUpperCase())) name = "Normal " + suffix++;
      this.names = [...this.names, empty("cellStyle", { name, xfId: this.normal, builtinId: 0 })];
    }
    this.cells.add(this.format(this.normalRaw, this.normalCss, normalFormat, this.normal));
  }

  private projection(cssClass: string): SourceStyleNode {
    return { tag: "td", attrs: { class: cssClass }, parent: this.themeScope };
  }
  raw(cssClass: string): SourceStyleMap {
    return /\s/.test(cssClass) ? this.resolver.specified(this.projection(cssClass)) : this.resolver.specifiedClass(cssClass);
  }
  css(cssClass: string): SourceStyleMap {
    return /\s/.test(cssClass) ? this.resolver.get(this.projection(cssClass)) : this.resolver.classStyle(cssClass, this.theme);
  }
  private format(raw: SourceStyleMap, css: SourceStyleMap, formatId?: string, parent?: number, apply?: DeepReadonly<WorkbookStyleApplication>): string {
    return element("xf", alignment(css, this.reportFormatLoss), {
      numFmtId: this.numberFormats.get(formatId ?? "")?.id ?? 0,
      fontId: this.fonts.add(font(raw, css, "cell", this.reportFormatLoss)), fillId: this.fills.add(fill(raw, css) || pattern("none")),
      borderId: this.borders.add(border(raw, css, this.reportFormatLoss) || "<border/>"), xfId: parent,
      quotePrefix: flag(css["--office-cell-quote-prefix"]), ...application(apply),
    });
  }
  cell(cssClass: string, formatId?: string): number {
    const key = JSON.stringify([cssClass, formatId]), known = this.cellIndices.get(key);
    if (known !== undefined) return known;
    const binding = cssClass.split(/\s+/).map(name => this.bindings.get(name)).find(Boolean);
    const node = this.projection(cssClass);
    const raw = this.resolver.specified(node), css = this.resolver.get(node);
    const index = this.cells.add(this.format(raw, css, formatId, binding ? this.namedIds.get(binding.styleId) : this.normal, binding?.apply ??
      { font: true, fill: true, border: true, alignment: true, numberFormat: true, protection: false }));
    this.cellIndices.set(key, index);
    return index;
  }

  differential(cssClass?: string, formatId?: string): number | undefined {
    if (!cssClass && !formatId) return undefined;
    const key = JSON.stringify([cssClass, formatId]), previous = this.dxfIndices.get(key);
    if (previous !== undefined) return previous;
    const raw = cssClass ? this.raw(cssClass) : {}, computed = cssClass ? this.css(cssClass) : {};
    // Differential appearance writes only declarations of the rule. Theme
    // variables participate in resolving values without becoming overrides.
    const css = Object.fromEntries(Object.keys(raw).map(name => [name, computed[name]]).filter(entry => entry[1] !== undefined)) as SourceStyleMap;
    const format = formatId ? this.numberFormats.get(formatId) : undefined;
    let fmt = "";
    if (format) {
      let code = format.code;
      if (code === undefined) {
        try {
          const documentLocale = this.snapshot.metadata.locale;
          code = workbookBuiltinNumberFormat(format.id, resolveWorkbookDisplayLocale(documentLocale ? { documentLocale } : {}));
        } catch (error) {
          if (!(error instanceof DocumentError) || error.code !== "UnsupportedFeature") throw error;
          this.reportFormatLoss("XLSX requires a format code for differential built-in number format " + format.id + ". No supported code is available for this workbook locale.");
        }
      }
      if (code !== undefined) fmt = empty("numFmt", { numFmtId: format.id, formatCode: xstring(code) });
    }
    const index = this.differentials.add(element("dxf", font(raw, css, "differential", this.reportFormatLoss) + fmt + fill(raw, css) + alignment(css, this.reportFormatLoss) + border(raw, css, this.reportFormatLoss)));
    this.dxfIndices.set(key, index);
    return index;
  }

  tableName(style: DeepReadonly<WorkbookTableStyleReference> | undefined): string | undefined {
    return style?.kind === "builtin" ? style.name : this.snapshot.metadata.workbook.tableStyles?.find(item => item.styleId === style?.styleId)?.name;
  }
  write(opc: OutputOpc): void {
    const { workbook } = this.snapshot.metadata;
    const custom = (workbook.tableStyles ?? []).map(style => element("tableStyle", style.elements.map(entry => empty("tableStyleElement", {
      type: entry.region, size: entry.size, dxfId: this.differential(entry.style?.cssClass, entry.style?.numberFormatId),
    })).join(""), { name: xstring(style.name), pivot: bool(style.pivotCompatible), table: 1, count: style.elements.length }));
    const tableStyles = element("tableStyles", custom.join(""), { count: custom.length, defaultTableStyle: this.tableName(workbook.defaultTableStyle) });
    const formats = [...this.numberFormats.values()].flatMap(format => format.id >= 164 && format.code !== undefined
      ? [empty("numFmt", { numFmtId: format.id, formatCode: xstring(format.code) })] : []);
    const indexed = Array.from({ length: 64 }, (_, index) => {
      const rgb = this.theme[`--office-indexed-${index}`];
      const alpha = Number(this.theme[`--office-indexed-${index}-native-alpha`] ?? 0);
      return rgb ? empty("rgbColor", { rgb: (alpha.toString(16).padStart(2, "0") + rgb.slice(1)).toUpperCase() }) : "";
    }).filter(Boolean);
    const part = opc.part("xl/styles.xml", sheetContentType("styles"));
    part.add(declaration + "<styleSheet" + sheetNamespaces + ">");
    if (formats.length) part.add(counted("numFmts", formats));
    for (const [name, entries] of [["fonts", this.fonts.entries], ["fills", this.fills.entries], ["borders", this.borders.entries],
      ["cellStyleXfs", this.named], ["cellXfs", this.cells.entries], ["cellStyles", this.names], ["dxfs", this.differentials.entries]] as const) {
      part.add("<" + name + ' count="' + entries.length + '">');
      for (const entry of entries) part.add(entry);
      part.add("</" + name + ">");
    }
    part.add(tableStyles);
    if (indexed.length) part.add(element("colors", element("indexedColors", indexed.join(""))));
    part.add("</styleSheet>");
    part.finish();
    opc.relate(workbookPart, "styles", "styles.xml");
  }
}
