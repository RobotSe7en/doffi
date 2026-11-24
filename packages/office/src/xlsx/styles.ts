import { DocumentError } from "@doffi/core";
import type { WorkbookCellStyles, WorkbookNumberFormat, WorkbookStyleApplication } from "@doffi/core";
import type { OpcPackage } from "../ooxml/package.js";
import { declarations } from "../ooxml/styles.js";
import type { Css } from "../ooxml/styles.js";
import { check, elements, fail, S, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, integer, one, spreadsheetText, value } from "./values.js";
import { SpreadsheetColors, colorDeclarations } from "./colors.js";
import { readWorkbookTheme } from "./theme.js";
import type { WorkbookTheme } from "./theme.js";
import { readDrawingTheme } from "./drawing-theme.js";
import type { WorkbookDrawingTheme } from "./drawing-theme.js";
import { fontDeclarations } from "./font.js";
import { alignmentDeclarations, borderDeclarations, fillDeclarations } from "./appearance.js";
import type { CellStyleReference, SheetStyles } from "./sheet.js";
import type { RichTextStyles } from "./strings.js";
import { excelRootAttributes, readExcelStyleDefaults } from "./excel-metadata.js";

const groups = ["Font", "Fill", "Border", "Alignment", "NumberFormat", "Protection"] as const;
type FormatGroup = typeof groups[number];
export interface NativeFormat extends CellStyleReference {
  readonly parent?: number;
  /** In cellXfs these govern future named-style changes, not current display. */
  readonly apply: Readonly<Record<FormatGroup, boolean>>;
}
export interface NativeNamedStyle {
  readonly name: string; readonly format: number; readonly builtinId?: number; readonly level?: number; readonly hidden: boolean; readonly customBuiltin: boolean;
}
export interface DifferentialStyleReference { readonly cssClass: string; readonly numberFormatId?: string }
function table(root: XmlNode, name: string, child: string): XmlNode[] | undefined {
  const parent = one(root, name); if (!parent) return undefined;
  attributes(parent, ["count"]); check(parent, [child], S);
  const values = elements(parent, child, S);
  if (value(parent, "count") !== undefined && integer(parent, "count") !== values.length) fail(parent, "A native style table count is inconsistent.");
  return values;
}
const defaultFont: Css = {
  "--office-font-source-name": '"Liberation Sans"', "--office-font-ascii": '"Liberation Sans"', "--office-font-h-ansi": '"Liberation Sans"',
  "--office-font-east-asia": '"Noto Sans CJK SC"', "--office-font-complex-script": '"Liberation Sans"', "--office-font-scheme": "none",
  "font-family": "var(--office-font-ascii),var(--office-font-east-asia)", "font-size": "11pt", "font-weight": "normal", "font-style": "normal",
  "vertical-align": "baseline", "--office-font-underline": "none", "--office-font-strike": "false",
  ...colorDeclarations("--office-cell-foreground", { css: "var(--office-system-foreground)", tint: 0 }), color: "var(--office-cell-foreground-color)",
};
const defaultBorder = Object.fromEntries(["left", "right", "top", "bottom"].map(side => [`--office-cell-border-${side}-kind`, "none"]));

/** Native style records are transient input. Their only saved appearance values
 * are CSS declarations; number-format definitions remain typed source metadata.
 * Named definitions, aliases and update bindings persist without duplicating
 * effective cell appearance in the metadata. */
export class XlsxStyles implements SheetStyles, RichTextStyles {
  readonly theme: WorkbookTheme;
  readonly drawingTheme: WorkbookDrawingTheme;
  readonly colors: SpreadsheetColors;
  readonly cellFormats: readonly NativeFormat[];
  readonly styleFormats: readonly NativeFormat[];
  readonly namedStyles: readonly NativeNamedStyle[];
  readonly tableStyles: XmlNode | undefined;
  private readonly rules: string[] = [];
  private readonly formats = new Map<number, WorkbookNumberFormat>();
  private readonly sharedRules = new Map<string, string>();
  private readonly differentials: readonly XmlNode[];
  private readonly dxfCache = new Map<number, DifferentialStyleReference>();
  private ruleUnits = 0;
  private readonly fallbackNode: XmlNode;
  private readonly normalFormat: CellStyleReference | undefined;
  constructor(opc: OpcPackage, workbookPart: string, private readonly maxCssCodeUnits: number) {
    this.theme = readWorkbookTheme(opc, workbookPart);
    const part = opc.related(workbookPart, "styles"), root = part ? opc.read(part) : undefined;
    this.fallbackNode = { part: part ?? workbookPart, name: "xf", uri: S, attrs: new Map(), namespaces: {}, content: [] };
    if (root) {
      if (root.uri !== S || root.name !== "styleSheet" || opc.contentType(part!) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml") fail(root, "The native style part is invalid.");
      excelRootAttributes(root); check(root, ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"], S);
      readExcelStyleDefaults(one(root, "extLst"));
    }
    this.colors = new SpreadsheetColors(one(root, "colors"), this.theme.colors);
    this.addRule("xlsx-theme", { ...this.theme.css, ...this.colors.css });
    this.drawingTheme = readDrawingTheme(this.theme.source, (name, css) => this.addRule(name, css));
    // Native text defaults precede local object/list/run declarations. Keeping
    // them in shared source classes avoids browser paragraph margins and still
    // lets explicit zero insets, theme styles and local typography win.
    this.addRule("xlsx-drawing-text", { "font-size": "18pt", "line-height": "normal", color: "var(--office-theme-tx1,#000000)",
      "padding-left": "7.2pt", "padding-right": "7.2pt", "padding-top": "3.6pt", "padding-bottom": "3.6pt" });
    this.addRule("xlsx-drawing-paragraph", { margin: "0" });
    for (const node of root ? table(root, "numFmts", "numFmt") ?? [] : []) {
      attributes(node, ["numFmtId", "formatCode"]); check(node, [], S);
      const id = integer(node, "numFmtId"), code = spreadsheetText(value(node, "formatCode") ?? "", node, 2048);
      if (!code || this.formats.has(id)) fail(node, "A number format is empty or repeated.");
      this.formats.set(id, { formatId: `num-${id}`, code });
    }
    const fonts = root ? table(root, "fonts", "font") : undefined, fills = root ? table(root, "fills", "fill") : undefined, borders = root ? table(root, "borders", "border") : undefined;
    const fontCache = new Map<number, Css>(), fillCache = new Map<number, Css>(), borderCache = new Map<number, Css>();
    const component = (list: readonly XmlNode[] | undefined, index: number, node: XmlNode, cache: Map<number, Css>, fallback: Css, read: (node: XmlNode) => Css): Css => {
      if (list === undefined && index === 0) return fallback;
      const item = list?.[index]; if (!item) fail(node, "A native format references an unknown style component.");
      let result = cache.get(index); if (!result) { result = read(item); cache.set(index, result); } return result;
    };
    const format = (node: XmlNode, index: number, named: boolean, parentCount: number): NativeFormat => {
      attributes(node, ["numFmtId", "fontId", "fillId", "borderId", "xfId", "applyAlignment", "applyBorder", "applyFill", "applyFont", "applyNumberFormat", "applyProtection", "pivotButton", "quotePrefix"]); check(node, ["alignment", "protection"], S);
      if (boolean(node, "pivotButton")) unsupported(node, "Pivot button formatting is not supported.");
      if (named && value(node, "xfId") !== undefined) fail(node, "A named style format must not have a parent format.");
      const parent = value(node, "xfId") === undefined ? undefined : integer(node, "xfId", 0, parentCount - 1);
      const apply = Object.fromEntries(groups.map(group => [group, boolean(node, "apply" + group, named)])) as Record<FormatGroup, boolean>;
      const css: Css = {
        ...component(fonts, integer(node, "fontId", 0, 0xffff_ffff, 0), node, fontCache, defaultFont, font => ({ ...defaultFont, ...fontDeclarations(font, this.colors, true) })),
        ...component(fills, integer(node, "fillId", 0, 0xffff_ffff, 0), node, fillCache, { "--office-cell-fill-kind": "none" }, fill => fillDeclarations(fill, this.colors)),
        ...component(borders, integer(node, "borderId", 0, 0xffff_ffff, 0), node, borderCache, defaultBorder, border => borderDeclarations(border, this.colors, true)),
        ...alignmentDeclarations(one(node, "alignment"), true), "--office-cell-quote-prefix": String(boolean(node, "quotePrefix")),
      };
      this.protection(one(node, "protection"));
      const cssClass = `xlsx-${named ? "named" : "cell"}-${index}`; this.addRule(cssClass, css);
      return { cssClass, numberFormatId: this.numberFormat(integer(node, "numFmtId", 0, 0xffff_ffff, 0), node), ...(parent === undefined ? {} : { parent }), apply };
    };
    const named = root ? table(root, "cellStyleXfs", "xf") ?? [] : [], cells = root ? table(root, "cellXfs", "xf") : undefined;
    this.styleFormats = named.map((node, index) => format(node, index, true, 0));
    if (cells && !cells.length) fail(root!, "The cell format table is empty.");
    this.cellFormats = cells ? cells.map((node, index) => format(node, index, false, named.length)) : [format(this.fallbackNode, 0, false, named.length)];
    const names = new Set<string>();
    this.namedStyles = (root ? table(root, "cellStyles", "cellStyle") ?? [] : []).map(node => {
      attributes(node, ["name", "xfId", "builtinId", "iLevel", "hidden", "customBuiltin"]); check(node, [], S);
      const name = spreadsheetText(value(node, "name") ?? "", node, 255);
      if (!name || names.has(name.toUpperCase())) fail(node, "Named cell styles must have unique names."); names.add(name.toUpperCase());
      return { name, format: integer(node, "xfId", 0, named.length - 1), hidden: boolean(node, "hidden"), customBuiltin: boolean(node, "customBuiltin"),
        ...(value(node, "builtinId") === undefined ? {} : { builtinId: integer(node, "builtinId") }),
        ...(value(node, "iLevel") === undefined ? {} : { level: integer(node, "iLevel") }) };
    });
    const normal = this.namedStyles.filter(style => style.builtinId === 0);
    if (normal.length > 1) fail(root!, "A workbook must not declare competing built-in Normal styles.");
    this.normalFormat = normal[0] ? this.styleFormats[normal[0].format]! : root ? undefined : this.cellFormats[0]!;
    this.differentials = root ? table(root, "dxfs", "dxf") ?? [] : [];
    this.tableStyles = one(root, "tableStyles");
  }
  private addRule(name: string, css: Css): void {
    const source = `.${name}{${declarations(css)}}`;
    if (this.ruleUnits + source.length > this.maxCssCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Generated workbook styles exceed their source budget.");
    this.ruleUnits += source.length; this.rules.push(source);
  }
  private sharedRule(css: Css): string | undefined {
    if (!Object.keys(css).length) return undefined;
    const key = declarations(css), previous = this.sharedRules.get(key); if (previous) return previous;
    const name = `xlsx-style-${this.sharedRules.size}`; this.addRule(name, css); this.sharedRules.set(key, name); return name;
  }
  private numberFormat(id: number, node: XmlNode): string {
    if (!this.formats.has(id)) {
      if (id >= 164) fail(node, "A native number format definition is missing.");
      this.formats.set(id, { formatId: `num-${id}`, builtinId: id });
    }
    return `num-${id}`;
  }
  private protection(node: XmlNode | undefined): void {
    if (!node) return; attributes(node, ["locked", "hidden"]); check(node, [], S);
    if (!boolean(node, "locked", true) || boolean(node, "hidden")) unsupported(node, "Nondefault cell protection formatting is not implemented.");
  }
  cell(index: number, node: XmlNode): CellStyleReference {
    const format = this.cellFormats[index]; if (!Number.isSafeInteger(index) || index < 0 || !format) fail(node, "A cell references an unknown format.");
    return { cssClass: format.cssClass, numberFormatId: format.numberFormatId };
  }
  columnWidthFontCssClass(): string {
    return this.defaultCell().cssClass;
  }
  defaultCell(): CellStyleReference {
    if (!this.normalFormat) unsupported(this.fallbackNode, "Native worksheet defaults require an unambiguous declared Normal style.");
    return { cssClass: this.normalFormat.cssClass, numberFormatId: this.normalFormat.numberFormatId };
  }
  objectStyle(css: Css): string | undefined { return this.sharedRule(css); }
  drawingRule(name: string, css: Css): void { this.addRule(name, css); }
  run(node: XmlNode): string | undefined { return this.sharedRule(fontDeclarations(node, this.colors, false)); }
  tabStyle(node: XmlNode): string { return this.sharedRule(colorDeclarations("--office-tab", this.colors.read(node)))!; }
  scaleColor(node: XmlNode): string { return this.sharedRule(colorDeclarations("--office-scale", this.colors.read(node)))!; }
  differential(index: number, source: XmlNode): DifferentialStyleReference {
    const previous = this.dxfCache.get(index); if (previous) return previous;
    const node = this.differentials[index]; if (!Number.isSafeInteger(index) || index < 0 || !node) fail(source, "A differential style reference is invalid.");
    attributes(node, []); check(node, ["font", "numFmt", "fill", "alignment", "border", "protection"], S);
    const font = one(node, "font"), fill = one(node, "fill"), border = one(node, "border"), num = one(node, "numFmt");
    this.protection(one(node, "protection"));
    const css = { ...(font ? fontDeclarations(font, this.colors, false) : {}), ...(fill ? fillDeclarations(fill, this.colors) : {}), ...(border ? borderDeclarations(border, this.colors, false) : {}), ...alignmentDeclarations(one(node, "alignment"), false) };
    // A differential rule without a number format must not override its cell.
    let numberFormatId: string | undefined;
    if (num) {
      attributes(num, ["numFmtId", "formatCode"]); check(num, [], S); integer(num, "numFmtId");
      const code = spreadsheetText(value(num, "formatCode") ?? "", num, 2048); if (!code) fail(num, "A differential number format is empty.");
      numberFormatId = `dxf-num-${index}`;
      // Negative keys cannot collide with any native unsigned numFmtId.
      this.formats.set(-index - 1, { formatId: numberFormatId, code });
    }
    const cssClass = `xlsx-dxf-${index}`; this.addRule(cssClass, css);
    const result = { cssClass, ...(numberFormatId === undefined ? {} : { numberFormatId }) }; this.dxfCache.set(index, result); return result;
  }
  numberFormats(): WorkbookNumberFormat[] { return [...this.formats.values()]; }
  cellStyles(): WorkbookCellStyles {
    const id = (index: number): string => `xlsx-style-definition-${index}`;
    const apply = (format: NativeFormat): WorkbookStyleApplication => ({ font: format.apply.Font, fill: format.apply.Fill, border: format.apply.Border,
      alignment: format.apply.Alignment, numberFormat: format.apply.NumberFormat, protection: format.apply.Protection });
    return {
      definitions: this.styleFormats.map((format, index) => ({ styleId: id(index), cssClass: format.cssClass, numberFormatId: format.numberFormatId, apply: apply(format) })),
      names: this.namedStyles.map(({ format, ...style }) => ({ ...style, styleId: id(format) })),
      bindings: this.cellFormats.flatMap(format => format.parent === undefined ? [] : [{ cssClass: format.cssClass, styleId: id(format.parent), apply: apply(format) }]),
    };
  }
  stylesheet(): string { return this.rules.join("\n"); }
}
