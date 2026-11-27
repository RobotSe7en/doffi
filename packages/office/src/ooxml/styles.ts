import { wordRowExceptionCss } from "@doffi/core";
import { readWordTablePosition } from "./word-table-position.js";
import {
  passiveWordRunBooleans,
  passiveWordRunValues,
  passiveWordRunCss,
  wordUnderlineValues,
  wordEmphasisValues,
  wordOpenTypeLigatures,
  wordOpenTypeNumSpacings,
  wordOpenTypeNumForms,
} from "./word-passive-run-properties.js";
import {
  passiveWordParagraphBooleans,
  passiveWordParagraphCss,
} from "./word-passive-paragraph-properties.js";
import { measurePoints } from "./measure.js";
import { wordColorLuminance } from "@doffi/core";
import { DocumentError } from "@doffi/core";
import type { WordMetadata, WordNumberLevel } from "@doffi/core";
import type { OpcPackage } from "./package.js";
import {
  A,
  R,
  W,
  W14,
  WP,
  attr,
  check,
  child,
  elements,
  fail,
  number,
  on,
  unsupported,
} from "./xml.js";
import type { XmlNode } from "./xml.js";

export type Css = Record<string, string>;
type TabStops = WordMetadata["tabs"][number]["stops"];
/** Tab declarations amend the inherited set by position, including removals. */
export function mergeTabs(inherited: TabStops = [], properties?: XmlNode): TabStops {
  const stops = new Map(inherited.map((stop) => [stop.position, stop]));
  check(properties, ["tab"]);
  for (const stop of elements(properties, "tab")) {
    const type = attr(stop, "val") ?? "left";
    const alignment = type === "num" || type === "start" ? "left" : type === "end" ? "right" : type;
    const leader = attr(stop, "leader") ?? "none";
    if (
      !["left", "center", "right", "decimal", "clear"].includes(alignment) ||
      !["none", "dot", "hyphen", "underscore", "heavy", "middleDot"].includes(leader)
    )
      unsupported(stop, "A tab alignment or leader is not supported.");
    const position = number(attr(stop, "pos"), stop) / 20;
    if (alignment === "clear") stops.delete(position);
    else
      stops.set(position, {
        position,
        alignment: alignment as TabStops[number]["alignment"],
        leader: leader as TabStops[number]["leader"],
      });
  }
  return [...stops.values()].sort((left, right) => left.position - right.position);
}
export function nonInheritedRun(css: Css): Css {
  return Object.fromEntries(
    Object.entries(css).filter(
      ([name]) => name === "background-color" || name.startsWith("text-decoration"),
    ),
  );
}
export function inheritedRun(css: Css): Css {
  return Object.fromEntries(
    Object.entries(css).filter(
      ([name]) => name !== "background-color" && !name.startsWith("text-decoration"),
    ),
  );
}
export function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/\r/g, "&#13;");
}
export function declarations(css: Css): string {
  return Object.entries(css)
    .map(([name, value]) => name + ":" + value)
    .join(";");
}
export function pt(value: number): string {
  return String(Math.round(value * 10000) / 10000) + "pt";
}
const val = (node: XmlNode | undefined): string | undefined => attr(node, "val");
const quoteFont = (name: string): string =>
  '"' + name.replace(/["\\]/g, "\\$&").replace(/[\u0000-\u001f]/g, "") + '"';
export interface StyleValues {
  runClasses?: string[];
  p: Css;
  r: Css;
  table: Css;
  cell: Css;
  row?: Css;
  numbering?: { id?: string; indent: Css };
  tabs?: TabStops;
  outline?: number;
}
const empty = (): StyleValues => ({ p: {}, r: {}, table: {}, cell: {} });

export function tableLookMask(node: XmlNode | undefined): number | undefined {
  if (!node) return undefined;
  const value = attr(node, "val");
  if (value !== undefined && !/^[0-9a-f]{1,4}$/i.test(value))
    fail(node, "A table look mask is invalid.");
  let mask = parseInt(value ?? "0000", 16);
  for (const [name, bit] of [
    ["firstRow", 32],
    ["lastRow", 64],
    ["firstColumn", 128],
    ["lastColumn", 256],
    ["noHBand", 512],
    ["noVBand", 1024],
  ] as const) {
    const flag = attr(node, name);
    if (flag !== undefined) {
      if (!["0", "1", "true", "false", "on", "off"].includes(flag))
        fail(node, "A table look flag is invalid.");
      mask = ["1", "true", "on"].includes(flag) ? mask | bit : mask & ~bit;
    }
  }
  return mask;
}
export class DocxStyles {
  readonly css: string[] = [];
  private readonly independentRules = new Map<string, string[]>();
  readonly styles: WordMetadata["styles"] = [];
  readonly numbering: WordMetadata["numbering"] = [];
  readonly defaults: StyleValues;
  readonly defaultParagraph: string | undefined;
  readonly defaultCharacter: string | undefined;
  readonly defaultTable: string | undefined;
  private readonly numberProperties = new Map<string, { paragraph: Css; tabs?: XmlNode }>();
  private readonly numberStyleLevels = new Map<string, number>();
  private readonly definitions = new Map<string, XmlNode>();
  private readonly resolved = new Map<string, StyleValues>();
  private readonly themeColors = new Map<string, string>();
  private readonly themeFonts = new Map<string, string>();
  private themeName: string | undefined;
  private colorSchemeName: string | undefined;
  private fontSchemeName: string | undefined;
  private readonly locale: string;
  private readonly bidiLocale: string;
  readonly language: string | undefined;

  private readonly stylePart: string | undefined;
  private readonly opc: OpcPackage;
  constructor(
    opc: OpcPackage,
    main: string,
    private readonly compatibleImport = false,
    private readonly onImage?: (part: string, source: XmlNode) => string,
  ) {
    this.opc = opc;
    this.stylePart = opc.related(main, "styles");
    const stylePath = opc.related(main, "styles");
    const styleRoot = stylePath ? opc.read(stylePath) : undefined;
    const defaultLanguage = child(
      child(child(child(styleRoot, "docDefaults"), "rPrDefault"), "rPr"),
      "lang",
    );
    this.language = attr(defaultLanguage, "val");
    const settingsPath = opc.related(main, "settings"),
      themeLanguage = settingsPath ? child(opc.read(settingsPath), "themeFontLang") : undefined;
    this.locale = attr(defaultLanguage, "eastAsia") ?? attr(themeLanguage, "eastAsia") ?? attr(defaultLanguage, "val") ?? "en-US";
    this.bidiLocale = attr(defaultLanguage, "bidi") ?? attr(themeLanguage, "bidi") ?? "ar-SA";
    const themePath = opc.related(main, "theme");
    if (themePath) {
      const root = opc.read(themePath);
      this.themeName = attr(root, "name", "");
      const theme = child(root, "themeElements", A);
      const clrScheme = child(theme, "clrScheme", A);
      this.colorSchemeName = attr(clrScheme, "name", "");
      for (const color of elements(clrScheme, undefined, A)) {
        const values = elements(color, undefined, A);
        const item = values[0];
        const value = attr(item, item?.name === "sysClr" ? "lastClr" : "val", "");
        if (value && /^[0-9a-f]{6}$/i.test(value)) this.themeColors.set(color.name, value);
      }
      const fonts = child(theme, "fontScheme", A);
      this.fontSchemeName = attr(fonts, "name", "");
      for (const family of ["major", "minor"]) {
        const group = child(fonts, family + "Font", A);
        for (const [suffix, name] of [
          ["Ascii", "latin"],
          ["HAnsi", "latin"],
          ["EastAsia", "ea"],
          ["Bidi", "cs"],
        ]) {
          const face = attr(child(group, name!, A), "typeface", "");
          if (face) this.themeFonts.set(family + suffix, face);
        }
        for (const font of elements(group, "font", A)) {
          const script = attr(font, "script", "");
          const face = attr(font, "typeface", "");
          if (script && face) for (const role of ["EastAsia", "Bidi"])
            this.themeFonts.set(family + role + ":" + script, face);
        }
      }
    }
    const root = styleRoot;
    if (root && (root.uri !== W || root.name !== "styles"))
      fail(root, "The styles part has an invalid root.");
    check(root, ["docDefaults", "latentStyles", "style"]);
    const docDefaults = child(root, "docDefaults");
    this.defaults = empty();
    // Omitted native spacing means single automatic spacing independently of
    // the browser's default paragraph line-height.
    this.defaults.p["line-height"] = "1";
    Object.assign(this.defaults.r, this.run(child(child(docDefaults, "rPrDefault"), "rPr")));
    Object.assign(this.defaults.p, this.paragraph(child(child(docDefaults, "pPrDefault"), "pPr")));
    this.defaults.tabs = mergeTabs(
      [],
      child(child(child(docDefaults, "pPrDefault"), "pPr"), "tabs"),
    );
    for (const style of elements(root, "style")) {
      const id = attr(style, "styleId");
      const kind = attr(style, "type");
      if (
        !id ||
        !kind ||
        !["paragraph", "character", "table", "numbering"].includes(kind) ||
        this.definitions.has(id)
      )
        fail(style, "A native style identity or type is invalid.");
      check(style, [
        "name",
        "aliases",
        "basedOn",
        "next",
        "link",
        "autoRedefine",
        "hidden",
        "uiPriority",
        "semiHidden",
        "unhideWhenUsed",
        "qFormat",
        "locked",
        "personal",
        "personalCompose",
        "personalReply",
        "rsid",
        "pPr",
        "rPr",
        "tblPr",
        "trPr",
        "tcPr",
        "tblStylePr",
      ]);

      this.definitions.set(id, style);
      const basedOn = val(child(style, "basedOn"));
      const linkedStyleId = val(child(style, "link"));
      const nextStyleId = val(child(style, "next"));
      const conditions: NonNullable<WordMetadata["styles"][number]["conditions"]> = [];
      const conditional = new Set<string>();
      this.styles.push({
        styleId: id,
        kind: kind as "paragraph" | "character" | "table" | "numbering",
        cssClass: "ws" + this.styles.length,
        ...(basedOn ? { basedOn } : {}),
        ...(linkedStyleId ? { linkedStyleId } : {}),
        ...(nextStyleId ? { nextStyleId } : {}),
        conditions,
      });
      const metadata = this.styles[this.styles.length - 1]!;
      for (const key of ["name", "aliases"] as const) {
        const value = val(child(style, key));
        if (value !== undefined) metadata[key] = value;
      }
      for (const key of ["default", "customStyle"] as const) {
        const value = attr(style, key);
        if (value !== undefined) metadata[key] = ["1", "true", "on"].includes(value);
      }
      for (const key of ["hidden", "semiHidden", "unhideWhenUsed", "qFormat", "locked"] as const) {
        const value = child(style, key);
        if (value) metadata[key] = on(value);
      }
      const priority = child(style, "uiPriority");
      if (priority) metadata.uiPriority = number(val(priority), priority);
      const outline = child(child(style, "pPr"), "outlineLvl");
      if (outline) metadata.outlineLevel = number(val(outline), outline);

      const tblPr = child(style, "tblPr");
      const look = tableLookMask(child(tblPr, "tblLook"));
      if (look !== undefined) metadata.tableLookMask = look;
      for (const [name, key] of [
        ["tblStyleRowBandSize", "rowBandSize"],
        ["tblStyleColBandSize", "columnBandSize"],
      ] as const) {
        const value = child(tblPr, name);
        if (value) metadata[key] = number(val(value), value);
      }
      for (const region of elements(style, "tblStylePr")) {
        const type = attr(region, "type");
        if (
          !type ||
          ![
            "wholeTable",
            "firstRow",
            "lastRow",
            "firstCol",
            "lastCol",
            "band1Vert",
            "band2Vert",
            "band1Horz",
            "band2Horz",
            "neCell",
            "nwCell",
            "seCell",
            "swCell",
          ].includes(type) ||
          conditional.has(type)
        )
          fail(region, "A conditional table region is invalid.");
        check(region, ["pPr", "rPr", "tblPr", "tcPr", "trPr"]);
        const row = child(region, "trPr");
        const table = child(region, "tblPr");
        if (type !== "wholeTable" && child(table, "tblInd"))
          unsupported(table!, "Table indentation cannot vary by conditional cell region.");
        const values: StyleValues = {
          p: this.paragraph(child(region, "pPr")),
          r: this.run(child(region, "rPr")),
          table: this.table(table),
          row: this.row(row),
          cell: {
            ...this.margins(child(table, "tblCellMar")),
            ...this.cell(child(region, "tcPr")),
          },
        };
        conditional.add(type);
        const cssClass = metadata.cssClass + "-" + type;
        conditions.push({
          region: type as NonNullable<typeof metadata.conditions>[number]["region"],
          cssClass,
          components: Object.fromEntries(
            (["p", "r", "table", "cell", "row"] as const).map((key) => [
              key,
              this.component(cssClass + "-" + key, values[key] ?? {}),
            ]),
          ),
        });
        this.css.push(
          "." +
            cssClass +
            "{" +
            declarations({
              ...values.p,
              ...values.r,
              ...values.table,
              ...values.cell,
            }) +
            "}",
        );
      }
      if (["1", "true", "on"].includes(attr(style, "default") ?? "")) {
        if (kind === "paragraph") this.defaultParagraph = id;
        if (kind === "character") this.defaultCharacter = id;
        if (kind === "table") this.defaultTable = id;
      }
    }
    this.css.unshift(
      ".word-main{" + declarations(inheritedRun(this.defaults.r)) + "}",
      "p,h1,h2,h3,h4,h5,h6{" + declarations(this.defaults.p) + "}",
      ".word-default-run{" + declarations(nonInheritedRun(this.defaults.r)) + "}",
    );
    for (const style of this.styles) {
      const value = this.resolve(style.styleId);
      const definition = this.definitions.get(style.styleId)!;
      const parent = style.basedOn ? this.resolve(style.basedOn) : empty();
      const intrinsic = this.intrinsicTableStyle(definition);
      const fallback = (values: Css, inherited: Css): Css => Object.fromEntries(
        Object.entries(values).filter(([key]) => inherited[key] === undefined),
      );
      style.paragraphProperties = Object.keys(this.paragraph(child(definition, "pPr")));
      style.runProperties = Object.keys(
        this.run(child(definition, "rPr"), parent.r, style.kind !== "table"),
      );
      if (style.kind === "table" || style.kind === "paragraph" || style.kind === "character")
        style.components = {
          p: this.component(style.cssClass + "-p", this.paragraph(child(definition, "pPr"))),
          r: this.component(
            style.cssClass + "-r",
            this.run(child(definition, "rPr"), {}, style.kind !== "table"),
          ),
          ...(style.kind === "table"
            ? {
                table: this.component(
                  style.cssClass + "-table",
                  { ...fallback(intrinsic.table, parent.table), ...this.table(child(definition, "tblPr")) },
                ),
                row: this.component(style.cssClass + "-row", this.row(child(definition, "trPr"))),
                cell: this.component(style.cssClass + "-cell", {
                  ...fallback(intrinsic.cell, parent.cell),
                  ...this.margins(child(child(definition, "tblPr"), "tblCellMar")),
                  ...this.cell(child(definition, "tcPr")),
                }),
              }
            : {}),
        };
      const weight = style.kind === "character" ? 4 : style.kind === "paragraph" ? 3 : 1;
      const rules =
        style.kind === "character"
          ? value.r
          : style.kind === "table"
            ? { ...value.table, ...inheritedRun(value.r) }
            : { ...value.p, ...inheritedRun(value.r) };
      this.css.push(("." + style.cssClass).repeat(weight) + "{" + declarations(rules) + "}");
      if (style.kind !== "character") {
        const special = nonInheritedRun(value.r);
        if (Object.keys(special).length) {
          style.runCssClass = style.cssClass + "-run";
          this.css.push(
            ("." + style.runCssClass).repeat(style.kind === "paragraph" ? 3 : 1) +
              "{" +
              declarations(special) +
              "}",
          );
        }
      }
    }
    const numPath = opc.related(main, "numbering");
    if (numPath) this.readNumbering(opc.read(numPath), numPath);
  }
  private readonly componentValues = new Map<string, Css>();
  component(name: string, values: Css): string {
    this.componentValues.set(name, values);
    this.addIndependentRule("." + name, values);
    return name;
  }
  /** These selectors bind distinct source identities and have no shared cascade order. */
  addIndependentRule(selector: string, values: Css): void {
    const block = declarations(values);
    const group = this.independentRules.get(block);
    if (group) group.push(selector);
    else this.independentRules.set(block, [selector]);
  }
  stylesheet(): string {
    const grouped: string[] = [];
    for (const [block, selectors] of this.independentRules)
      for (let start = 0; start < selectors.length; start += 256)
        grouped.push(selectors.slice(start, start + 256).join(",") + "{" + block + "}");
    return [...this.css, ...grouped].join("\n");
  }
  readComponent = (name: string): Css => this.componentValues.get(name) ?? {};
  themeMetadata(): WordMetadata["theme"] {
    if (!this.themeColors.size && !this.themeFonts.size) return undefined;
    const allowed = new Set([
      "dk1",
      "lt1",
      "dk2",
      "lt2",
      "accent1",
      "accent2",
      "accent3",
      "accent4",
      "accent5",
      "accent6",
      "hlink",
      "folHlink",
    ]);
    const colors: Record<string, string> = {};
    for (const [role, value] of this.themeColors)
      if (allowed.has(role)) colors[role] = value.toUpperCase();
    const majorLatin = this.themeFonts.get("majorAscii");
    const minorLatin = this.themeFonts.get("minorAscii");
    const majorEastAsia =
      this.themeFonts.get("majorEastAsia") ?? this.themeFonts.get("majorEastAsia:Hans");
    const minorEastAsia =
      this.themeFonts.get("minorEastAsia") ?? this.themeFonts.get("minorEastAsia:Hans");
    return {
      ...(this.themeName ? { name: this.themeName } : {}),
      ...(this.colorSchemeName ? { colorSchemeName: this.colorSchemeName } : {}),
      ...(Object.keys(colors).length ? { colors } : {}),
      ...(this.fontSchemeName ? { fontSchemeName: this.fontSchemeName } : {}),
      ...(majorLatin ? { majorLatin } : {}),
      ...(minorLatin ? { minorLatin } : {}),
      ...(majorEastAsia ? { majorEastAsia } : {}),
      ...(minorEastAsia ? { minorEastAsia } : {}),
    };
  }
  drawingThemeProperties(): Css {
    const result: Css = {};
    for (const [role, value] of this.themeColors) result["--office-theme-" + role] = "#" + value;
    for (const [role, name] of Object.entries({ bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" }))
      if (result["--office-theme-" + name])
        result["--office-theme-" + role] = result["--office-theme-" + name]!;
    return result;
  }
  color(node: XmlNode, attribute = "val"): string | undefined {
    if (attribute === "fill") {
      const pattern = attr(node, "val");
      if (pattern === "nil") return "transparent";
      if (pattern && pattern !== "clear")
        unsupported(node, "Pattern shading is not yet supported.");
    }
    const theme = attr(node, attribute === "fill" ? "themeFill" : "themeColor");
    const roles: Record<string, string> = {
      dark1: "dk1",
      light1: "lt1",
      dark2: "dk2",
      light2: "lt2",
      text1: "dk1",
      background1: "lt1",
      text2: "dk2",
      background2: "lt2",
      hyperlink: "hlink",
      followedHyperlink: "folHlink",
    };
    let value = theme ? this.themeColors.get(roles[theme] ?? theme) : attr(node, attribute);
    if (value === "auto") return undefined;
    if (!value || !/^[0-9a-f]{6}$/i.test(value))
      return unsupported(node, "A document color or theme role cannot be resolved.");
    const tint = attr(node, attribute === "fill" ? "themeFillTint" : "themeTint");
    const shade = attr(node, attribute === "fill" ? "themeFillShade" : "themeShade");
    const modifier = tint ?? shade;
    if (modifier !== undefined) {
      if (!/^[0-9a-f]{2}$/i.test(modifier)) fail(node, "A theme color modifier is invalid.");
      value = wordColorLuminance(value, tint, shade);
    }
    return "#" + value.toLowerCase();
  }
  run(
    properties: XmlNode | undefined,
    inherited: Css = {},
    toggle = false,
    paragraphMark = false,
    allowRevision = false,
  ): Css {
    if (!properties) return {};
    if (properties.part === this.stylePart) {
      const change = elements(properties).find((node) => /PrChange$/.test(node.name));
      if (change)
        unsupported(change, "Revision properties in a shared style are preserved natively.");
    }
    check(properties, [
      "rStyle",
      "rFonts",
      "b",
      "bCs",
      "i",
      "iCs",
      "caps",
      "smallCaps",
      "strike",
      "dstrike",
      "outline",
      "shadow",
      "emboss",
      "imprint",
      "noProof",
      "snapToGrid",
      "vanish",
      "webHidden",
      "color",
      "spacing",
      "w",
      "kern",
      "position",
      "sz",
      "szCs",
      "highlight",
      "u",
      "effect",
      "fitText",
      "bdr",
      "shd",
      "vertAlign",
      "rtl",
      "cs",
      "em",
      "lang",
      "eastAsianLayout",
      "specVanish",
      "oMath",
      ...(allowRevision ? ["rPrChange"] : []),
      ...(paragraphMark ? ["ins", "del", "moveFrom", "moveTo"] : []),
    ], W, [
      {
        uri: W14,
        names: [
          "glow",
          "shadow",
          "reflection",
          "textOutline",
          "textFill",
          "scene3d",
          "props3d",
          "ligatures",
          "numForm",
          "numSpacing",
          "stylisticSets",
          "cntxtAlts",
        ],
      },
    ]);
    const css: Css = {};
    if (child(properties, "snapToGrid"))
      css["--word-run-snap-to-grid"] = on(child(properties, "snapToGrid")) ? "1" : "0";
    for (const role of ["val", "eastAsia", "bidi"]) {
      const language = attr(child(properties, "lang"), role);
      if (language !== undefined) css["--word-language-" + role] = JSON.stringify(language);
    }
    const fonts = child(properties, "rFonts");
    if (fonts) {
      const hint = attr(fonts, "hint");
      if (hint !== undefined) {
        if (!["default", "eastAsia", "cs"].includes(hint)) fail(fonts, "Invalid run font hint.");
        css["--word-font-hint"] = hint;
      }
      for (const [name, themeName] of [
        ["ascii", "asciiTheme"],
        ["hAnsi", "hAnsiTheme"],
        ["eastAsia", "eastAsiaTheme"],
        ["cs", "cstheme"],
      ]) {
        const role = attr(fonts, themeName!);
        const language = name === "cs" ? attr(child(properties, "lang"), "bidi") ?? this.bidiLocale :
          attr(child(properties, "lang"), "eastAsia") ??
          attr(child(properties, "lang"), "val") ??
          this.locale;
        let script = /^ja/i.test(language)
          ? "Jpan"
          : /^ko/i.test(language)
            ? "Hang"
            : /^zh-(?:TW|HK|MO|Hant)/i.test(language)
              ? "Hant"
              : "Hans";
        if (name === "cs") {
          try { script = new Intl.Locale(language).maximize().script ?? "Arab"; }
          catch { script = "Arab"; }
        }
        const face = role
          ? (this.themeFonts.get(role) ?? this.themeFonts.get(role + ":" + script)
            ?? (/^(major|minor)Bidi$/.test(role) ? this.themeFonts.get(role.replace("Bidi", "HAnsi")) : undefined))
          : attr(fonts, name!);
        if (role && !face) unsupported(fonts, "A theme font cannot be resolved.");
        if (face) {
          css["--word-font-" + name!] = quoteFont(face);
        }
      }
      if (css["--word-font-ascii"] || css["--word-font-hAnsi"])
        css["font-family"] = "var(--word-font-ascii,var(--word-font-hAnsi,inherit))";
    }
    for (const [name, property, active, inactive] of [
      ["b", "font-weight", "bold", "normal"],
      ["i", "font-style", "italic", "normal"],
      ["caps", "text-transform", "uppercase", "none"],
      ["smallCaps", "font-variant", "small-caps", "normal"],
    ]) {
      const italic = name === "i" && paragraphMark && this.compatibleImport ? elements(properties, "i") : [];
      // LibreOffice repeats the same disabled italic flag on some table
      // paragraph marks. Only the redundant, semantically identical case is safe.
      const redundant = italic.length === 2 && italic.every(item => !item.content.length
        && [...item.attrs.keys()].every(key => key === W + " val" || key.startsWith("http://www.w3.org/2000/xmlns/ ")) && !on(item));
      const item = redundant ? italic[0] : child(properties, name!);
      if (item) {
        const enabled = on(item);
        if (toggle && !enabled) continue;
        css[property!] = (toggle && enabled && inherited[property!] === active ? false : enabled)
          ? active!
          : inactive!;
      }
    }
    const fitText = child(properties, "fitText");
    if (fitText) {
      const twips = number(val(fitText), fitText);
      if (twips <= 0) fail(fitText, "Fit text width must be positive.");
      css["--word-fit-text"] = pt(twips / 20);
    }
    const eaLayout = child(properties, "eastAsianLayout");
    if (eaLayout) {
      const flagOn = (raw: string | undefined): boolean =>
        raw === "1" || raw === "true" || raw === "on";
      if (flagOn(attr(eaLayout, "combine"))) {
        const brackets = attr(eaLayout, "combineBrackets") ?? "none";
        if (!["none", "round", "square", "angle", "curly"].includes(brackets))
          fail(eaLayout, "Invalid East Asian combine brackets.");
        css["--word-ea-combine"] = brackets;
      }
      if (flagOn(attr(eaLayout, "vert"))) {
        css["--word-ea-vert"] = flagOn(attr(eaLayout, "vertCompress")) ? "compress" : "1";
      }
    }
    for (const feature of ["effect", "oMath"]) {
      const item = child(properties, feature);
      if (item && (feature === "effect" ? val(item) !== "none" : on(item))) unsupported(item);
    }
    for (const feature of ["webHidden", "rtl", "cs", "specVanish"])
      if (on(child(properties, feature))) unsupported(child(properties, feature)!);
    const kerning = child(properties, "kern");
    if (kerning) css["--word-kerning-threshold"] = pt(number(val(kerning), kerning) / 2);
    const complexSize = child(properties, "szCs");
    if (complexSize) css["--word-font-size-cs"] = pt(number(val(complexSize), complexSize) / 2);
    for (const [name, key] of [
      ["bCs", "--word-bold-cs"],
      ["iCs", "--word-italic-cs"],
    ] as const) {
      const value = child(properties, name);
      if (value) css[key] = on(value) ? "1" : "0";
    }
    const size = child(properties, "sz");
    if (size) {
      const value = number(val(size), size);
      if (value <= 0) fail(size, "Font size must be positive.");
      css["font-size"] = pt(value / 2);
    }
    const color = child(properties, "color");
    if (color) css.color = this.color(color) ?? "#000000";
    const shade = child(properties, "shd");
    if (shade && attr(shade, "fill") !== "auto") {
      const fill = this.color(shade, "fill");
      if (fill) {
        css["--word-run-shading"] = fill;
        css["background-color"] = fill;
      }
    }
    const highlight = val(child(properties, "highlight"));
    if (highlight === "none") css["background-color"] = "var(--word-run-shading,transparent)";
    if (highlight && highlight !== "none") {
      const colors: Record<string, string> = {
        yellow: "#ffff00",
        green: "#00ff00",
        cyan: "#00ffff",
        magenta: "#ff00ff",
        blue: "#0000ff",
        red: "#ff0000",
        darkBlue: "#000080",
        darkCyan: "#008080",
        darkGreen: "#008000",
        darkMagenta: "#800080",
        darkRed: "#800000",
        darkYellow: "#808000",
        darkGray: "#808080",
        lightGray: "#c0c0c0",
        black: "#000000",
        white: "#ffffff",
      };
      if (!colors[highlight]) unsupported(properties, "An unknown highlight color is used.");
      css["background-color"] = colors[highlight]!;
    }
    const underline = child(properties, "u");
    const strike = child(properties, "strike"),
      doubleStrike = child(properties, "dstrike");
    if (doubleStrike) {
      const enabled = on(doubleStrike);
      if (!toggle || enabled)
        css["--word-dstrike"] =
          toggle && enabled && inherited["--word-dstrike"] === "1" ? "0" : enabled ? "1" : "0";
    }
    if (underline || strike) {
      const lines = new Set(
        (inherited["text-decoration-line"] ?? "")
          .split(/\s+/)
          .filter((value) => value && value !== "none"),
      );
      const type = val(underline) ?? "single";
      if (underline) {
        if (!(wordUnderlineValues as readonly string[]).includes(type)) unsupported(underline);
        if (type === "none") lines.delete("underline");
        else lines.add("underline");
        css["--word-underline-style"] = type;
      }
      if (strike) {
        const enabled = on(strike);
        if (!toggle || enabled) {
          if (enabled && !(toggle && lines.has("line-through"))) lines.add("line-through");
          else lines.delete("line-through");
        }
      }
      css["text-decoration-line"] = [...lines].join(" ") || "none";
      if (underline)
        css["text-decoration-style"] =
          type === "double"
            ? "double"
            : type === "dotted" || type === "dottedHeavy"
              ? "dotted"
              : type === "wave" || type === "wavyHeavy" || type === "wavyDouble"
                ? "wavy"
                : [
                      "dash",
                      "dashedHeavy",
                      "dashLong",
                      "dashLongHeavy",
                      "dotDash",
                      "dashDotHeavy",
                      "dotDotDash",
                      "dashDotDotHeavy",
                    ].includes(type)
                  ? "dashed"
                  : "solid";
      if (underline && (attr(underline, "color") || attr(underline, "themeColor")))
        css["text-decoration-color"] = this.color(underline, "color") ?? "currentColor";
    }
    const align = val(child(properties, "vertAlign"));
    if (align) {
      if (!["baseline", "superscript", "subscript"].includes(align)) unsupported(properties);
      css["vertical-align"] =
        align === "superscript" ? "super" : align === "subscript" ? "sub" : "baseline";
    }
    const spacing = child(properties, "spacing");
    if (spacing) css["letter-spacing"] = pt(number(val(spacing), spacing) / 20);
    const width = child(properties, "w");
    if (width) {
      const raw = val(width)?.trim() ?? "",
        scale = Number(raw.replace(/%$/, ""));
      if (!/^\+?\d+%?$/.test(raw) || scale < 1 || scale > 600)
        fail(width, "A run character scale must be a whole percent between 1 and 600.");
      css["--word-run-native-w"] = String(scale);
    }
    const position = child(properties, "position");
    if (position) {
      const parsed = number(val(position), position);
      if (parsed < -3170 || parsed > 3170)
        fail(position, "A run character position must be between -3170 and 3170 half-points.");
    }
    const emphasis = child(properties, "em");
    if (emphasis) {
      const value = val(emphasis);
      if (!value || !(wordEmphasisValues as readonly string[]).includes(value))
        unsupported(emphasis, "An emphasis mark style is not supported.");
    }
    const runBorder = child(properties, "bdr");
    if (runBorder) {
      const style = attr(runBorder, "val") ?? "none";
      const styles: Record<string, string> = {
        none: "none",
        nil: "none",
        single: "solid",
        double: "double",
        dotted: "dotted",
        dashed: "dashed",
        dashSmallGap: "dashed",
        thick: "solid",
      };
      if (
        !styles[style] ||
        (attr(runBorder, "space") && number(attr(runBorder, "space"), runBorder) !== 0) ||
        (attr(runBorder, "shadow") &&
          !["0", "false", "off"].includes(attr(runBorder, "shadow")!)) ||
        (attr(runBorder, "frame") && !["0", "false", "off"].includes(attr(runBorder, "frame")!))
      )
        unsupported(runBorder, "A character border style is not supported.");
      if (styles[style] === "none") css["--word-run-border"] = "none";
      else {
        const width = number(attr(runBorder, "sz"), runBorder, 4) / 8,
          hasColor =
            attr(runBorder, "color") !== undefined || attr(runBorder, "themeColor") !== undefined,
          borderColor = !hasColor
            ? "currentColor"
            : (this.color(runBorder, "color") ?? "currentColor");
        css["--word-run-border"] = pt(width) + " " + styles[style] + " " + borderColor;
      }
    }
    const w14Color = (container: XmlNode | undefined): string => {
      if (!container) return "#4472c4";
      const srgb = child(container, "srgbClr", W14) ?? child(container, "srgbClr", A);
      const raw = attr(srgb, "val", W14) ?? attr(srgb, "val", "");
      if (raw && /^[0-9a-f]{6}$/i.test(raw)) return "#" + raw.toLowerCase();
      const scheme = child(container, "schemeClr", W14) ?? child(container, "schemeClr", A);
      const key = attr(scheme, "val", W14) ?? attr(scheme, "val", "");
      if (key) {
        const mapped = this.themeColors.get(key) ?? this.themeColors.get(key === "tx1" ? "dk1" : key === "bg1" ? "lt1" : key);
        if (mapped) return "#" + mapped.toLowerCase();
      }
      return "#4472c4";
    };
    const glow = child(properties, "glow", W14);
    if (glow) {
      const rad = Number(attr(glow, "rad", W14) ?? "63500");
      if (!Number.isFinite(rad) || rad < 0) fail(glow, "Invalid w14 glow radius.");
      css["--word-text-glow"] = pt(rad / 12700) + " " + w14Color(glow);
    }
    const textOutline = child(properties, "textOutline", W14);
    if (textOutline) {
      const w = Number(attr(textOutline, "w", W14) ?? "9525");
      if (!Number.isFinite(w) || w < 0) fail(textOutline, "Invalid w14 textOutline width.");
      const solidFill = child(textOutline, "solidFill", W14);
      css["--word-text-outline"] = pt(w / 12700) + " " + w14Color(solidFill);
    }
    const w14Shadow = child(properties, "shadow", W14);
    if (w14Shadow) {
      const blur = Number(attr(w14Shadow, "blurRad", W14) ?? "38100") / 12700;
      const dist = Number(attr(w14Shadow, "dist", W14) ?? "25400") / 12700;
      const dir = Math.round(Number(attr(w14Shadow, "dir", W14) ?? "2700000") / 60000);
      css["--word-text-shadow"] = `${pt(blur)} ${pt(dist)} ${dir} ${w14Color(w14Shadow)}`;
    }
    const reflection = child(properties, "reflection", W14);
    if (reflection) css["--word-text-reflection"] = "1";
    const textFill = child(properties, "textFill", W14);
    if (textFill) {
      const gradFill = child(textFill, "gradFill", W14);
      const gsLst = child(gradFill, "gsLst", W14);
      if (gradFill && gsLst) {
        const stops = elements(gsLst, "gs", W14)
          .map((gs) => {
            const pos = Math.max(0, Math.min(100, Math.round(Number(attr(gs, "pos", W14) ?? "0") / 1000)));
            return `${w14Color(gs)} ${pos}%`;
          })
          .filter(Boolean);
        if (stops.length >= 2) {
          const lin = child(gradFill, "lin", W14);
          const ang = lin ? Math.round(Number(attr(lin, "ang", W14) ?? "5400000") / 60000) : 90;
          css["--word-text-fill-gradient"] = `${ang}deg ${stops.join(" ")}`;
        }
      }
    }
    const props3d = child(properties, "props3d", W14);
    if (props3d) {
      const bevelT = child(props3d, "bevelT", W14);
      if (bevelT) {
        const prst = attr(bevelT, "prst", W14) ?? "circle";
        const w = Number(attr(bevelT, "w", W14) ?? "38100") / 12700;
        const h = Number(attr(bevelT, "h", W14) ?? "38100") / 12700;
        css["--word-text-bevel"] = `${prst} ${pt(w)} ${pt(h)}`;
      }
    }
    const scene3d = child(properties, "scene3d", W14);
    if (scene3d) {
      const camera = child(scene3d, "camera", W14);
      const prst = attr(camera, "prst", W14);
      if (prst) css["--word-text-3d"] = prst;
    }
    const ligatures = child(properties, "ligatures", W14);
    if (ligatures) {
      const value = attr(ligatures, "val", W14);
      if (!value || !(wordOpenTypeLigatures as readonly string[]).includes(value))
        unsupported(ligatures, "An OpenType ligature setting is not supported.");
      css["--word-opentype-ligatures"] = value;
    }
    const numSpacing = child(properties, "numSpacing", W14);
    if (numSpacing) {
      const value = attr(numSpacing, "val", W14);
      if (!value || !(wordOpenTypeNumSpacings as readonly string[]).includes(value))
        unsupported(numSpacing, "An OpenType number spacing setting is not supported.");
      css["--word-opentype-num-spacing"] = value;
    }
    const numForm = child(properties, "numForm", W14);
    if (numForm) {
      const value = attr(numForm, "val", W14);
      if (!value || !(wordOpenTypeNumForms as readonly string[]).includes(value))
        unsupported(numForm, "An OpenType number form setting is not supported.");
      css["--word-opentype-num-form"] = value;
    }
    const stylisticSets = child(properties, "stylisticSets", W14);
    if (stylisticSets) {
      check(stylisticSets, ["styleSet"], W14);
      const active: number[] = [];
      for (const item of elements(stylisticSets, "styleSet", W14)) {
        const id = number(attr(item, "id", W14), item);
        if (id < 1 || id > 20)
          fail(item, "An OpenType stylistic set id must be between 1 and 20.");
        const state = attr(item, "val", W14);
        if (state !== undefined && !["on", "off", "1", "0", "true", "false"].includes(state))
          fail(item, "An OpenType stylistic set value is invalid.");
        if (state === undefined || ["on", "1", "true"].includes(state)) active.push(id);
      }
      const unique = [...new Set(active)].sort((a, b) => a - b);
      css["--word-opentype-stylistic-set"] = unique.length ? unique.join(",") : "none";
    }
    const cntxtAlts = child(properties, "cntxtAlts", W14);
    if (cntxtAlts) {
      const state = attr(cntxtAlts, "val", W14);
      if (state !== undefined && !["on", "off", "1", "0", "true", "false"].includes(state))
        fail(cntxtAlts, "An OpenType contextual alternates value is invalid.");
      css["--word-opentype-contextual-alternates"] =
        state === undefined || ["on", "1", "true"].includes(state) ? "1" : "0";
    }
    for (const name of passiveWordRunBooleans) {
      const item = child(properties, name);
      if (item) css[passiveWordRunCss(name)] = on(item) ? "1" : "0";
    }
    for (const name of passiveWordRunValues) {
      const item = child(properties, name);
      if (item) css[passiveWordRunCss(name)] = val(item)!;
    }
    return css;
  }
  paragraph(properties: XmlNode | undefined): Css {
    if (properties?.part === this.stylePart && child(properties, "cnfStyle"))
      unsupported(
        child(properties, "cnfStyle")!,
        "Shared-style conditional cache flags remain native.",
      );
    if (!properties) return {};
    if (properties.part === this.stylePart) {
      const change = elements(properties).find((node) => /PrChange$/.test(node.name));
      if (change)
        unsupported(change, "Revision properties in a shared style are preserved natively.");
    }
    check(properties, [
      "pStyle",
      "keepNext",
      "keepLines",
      "pageBreakBefore",
      "framePr",
      "widowControl",
      "numPr",
      "suppressLineNumbers",
      "pBdr",
      "shd",
      "tabs",
      "suppressAutoHyphens",
      "kinsoku",
      "wordWrap",
      "overflowPunct",
      "topLinePunct",
      "autoSpaceDE",
      "autoSpaceDN",
      "bidi",
      "adjustRightInd",
      "snapToGrid",
      "spacing",
      "ind",
      "contextualSpacing",
      "mirrorIndents",
      "suppressOverlap",
      "jc",
      "textDirection",
      "textAlignment",
      "textboxTightWrap",
      "outlineLvl",
      "divId",
      "cnfStyle",
      "rPr",
      "sectPr",
      "pPrChange",
    ]);
    for (const name of [
      "adjustRightInd",
      "mirrorIndents",
      "suppressOverlap",
      "textAlignment",
      "textboxTightWrap",
    ]) {
      const item = child(properties, name);
      if (item) unsupported(item, "This paragraph layout property is not yet mapped.");
    }
    const css: Css = {};
    for (const [name, key] of [
      ["contextualSpacing", "--word-contextual-spacing"],
      ["kinsoku", "--word-kinsoku"],
      ["wordWrap", "--word-word-wrap"],
      ["overflowPunct", "--word-overflow-punctuation"],
      ["topLinePunct", "--word-top-line-punctuation"],
      ["autoSpaceDE", "--word-auto-space-de"],
      ["autoSpaceDN", "--word-auto-space-dn"],
    ] as const) {
      const item = child(properties, name);
      if (item) css[key] = on(item) ? "1" : "0";
    }
    if (child(properties, "snapToGrid"))
      css["--word-paragraph-snap-to-grid"] = on(child(properties, "snapToGrid")) ? "1" : "0";
    const suppressHyphens = child(properties, "suppressAutoHyphens");
    if (suppressHyphens) css["--word-suppress-auto-hyphens"] = on(suppressHyphens) ? "1" : "0";
    const align = val(child(properties, "jc"));
    if (align) {
      const values: Record<string, string> = {
        left: "left",
        start: "start",
        right: "right",
        end: "end",
        center: "center",
        both: "justify",
        distribute: "justify",
      };
      if (!values[align]) unsupported(properties, "A paragraph alignment is not supported.");
      css["text-align"] = values[align]!;
      css["--word-distributed"] = align === "distribute" ? "1" : "0";
    }
    const spacing = child(properties, "spacing");
    if (spacing) {
      for (const [name, property] of [
        ["before", "margin-top"],
        ["after", "margin-bottom"],
      ]) {
        const value = attr(spacing, name!);
        if (value !== undefined) css[property!] = pt(number(value, spacing) / 20);
      }
      // Keep the rest of w:spacing in the compatible view. Its line rule
      // can independently override an inherited fixed line height.
      for (const name of ["beforeLines", "afterLines"])
        if (attr(spacing, name) !== undefined)
          unsupported(spacing, "Line-based paragraph spacing remains native.", W + " " + name);
      for (const name of ["beforeAutospacing", "afterAutospacing"])
        if (["1", "true", "on"].includes(attr(spacing, name) ?? ""))
          unsupported(spacing, "Automatic paragraph spacing remains native.", W + " " + name);
      const line = attr(spacing, "line");
      if (line) {
        const rule = attr(spacing, "lineRule") ?? "auto";
        if (!["auto", "exact", "atLeast"].includes(rule))
          fail(spacing, "A line spacing rule is invalid.");
        css["--word-line-rule"] = rule;
        css["line-height"] =
          rule === "auto" ? String(number(line, spacing) / 240) : pt(number(line, spacing) / 20);
      }
    }
    const indent = child(properties, "ind");
    if (indent) {
      for (const [name, property] of [
        ["left", "margin-left"],
        ["right", "margin-right"],
        ["start", "margin-left"],
        ["end", "margin-right"],
        ["firstLine", "text-indent"],
      ]) {
        const value = attr(indent, name!);
        if (value !== undefined) css[property!] = pt(number(value, indent) / 20);
      }
      const hanging = attr(indent, "hanging");
      if (hanging !== undefined) css["text-indent"] = pt(-number(hanging, indent) / 20);
      const characterIndent = ["leftChars", "rightChars", "firstLineChars", "hangingChars"].find(
        (name) => attr(indent, name) !== undefined,
      );
      if (characterIndent)
        unsupported(indent, "Character-unit indentation remains native.", W + " " + characterIndent);
    }
    if (child(properties, "keepNext"))
      css["break-after"] = on(child(properties, "keepNext")) ? "avoid" : "auto";
    if (child(properties, "keepLines"))
      css["break-inside"] = on(child(properties, "keepLines")) ? "avoid" : "auto";
    if (child(properties, "pageBreakBefore"))
      css["break-before"] = on(child(properties, "pageBreakBefore")) ? "page" : "auto";
    if (child(properties, "widowControl")) {
      const count = on(child(properties, "widowControl")) ? "2" : "1";
      css.widows = count;
      css.orphans = count;
    }
    const frame = child(properties, "framePr");
    if (frame) {
      const dropCap = attr(frame, "dropCap");
      if (dropCap === "drop" || dropCap === "margin") {
        const lines = number(attr(frame, "lines"), frame, 3);
        const hSpace = number(attr(frame, "hSpace"), frame, 0);
        if (lines < 1 || lines > 10) fail(frame, "Drop cap lines must be between 1 and 10.");
        css["--word-drop-cap"] = dropCap;
        css["--word-drop-cap-lines"] = String(lines);
        css["--word-drop-cap-space"] = pt(hSpace / 20);
      } else {
        unsupported(frame, "A paragraph frame is not supported.");
      }
    }
    for (const name of ["bidi", "textDirection"]) {
      const item = child(properties, name);
      if (
        item &&
        ((name === "bidi" && on(item)) || (name === "textDirection" && val(item) !== "lrTb"))
      )
        unsupported(item, "A paragraph writing direction is not supported.");
    }
    for (const name of passiveWordParagraphBooleans) {
      const item = child(properties, name);
      if (item) css[passiveWordParagraphCss(name)] = on(item) ? "1" : "0";
    }
    if (child(properties, "textDirection")) css[passiveWordParagraphCss("textDirection")] = "lrTb";
    const div = child(properties, "divId");
    if (div) {
      const id = number(val(div), div);
      if (!Number.isSafeInteger(id) || id < 0)
        fail(div, "A paragraph div identity must be a nonnegative integer.");
      css[passiveWordParagraphCss("divId")] = String(id);
    }
    Object.assign(css, this.borders(child(properties, "pBdr")));
    const shade = child(properties, "shd");
    if (shade && attr(shade, "fill") !== "auto")
      css["background-color"] = this.color(shade, "fill") ?? "transparent";
    return css;
  }
  borders(properties: XmlNode | undefined): Css {
    if (!properties) return {};
    const result: Css = {};
    for (const border of elements(properties)) {
      const sides: Record<string, string> = {
        top: "top",
        bottom: "bottom",
        left: "left",
        right: "right",
        start: "left",
        end: "right",
      };
      if (border.name === "between")
        return unsupported(border, "Between-paragraph borders are not yet supported.");
      if (!sides[border.name] && !["insideH", "insideV", "tl2br", "tr2bl"].includes(border.name))
        unsupported(border);
      if (
        number(attr(border, "space"), border, 0) !== 0 ||
        ["shadow", "frame"].some((name) => ["true", "1", "on"].includes(attr(border, name) ?? ""))
      )
        unsupported(border, "Border spacing and border effects are not supported.");
      const style = val(border) ?? "single";
      const styles: Record<string, string> = {
        nil: "none",
        none: "none",
        single: "solid",
        double: "double",
        dotted: "dotted",
        dashed: "dashed",
        dashSmallGap: "dashed",
        thick: "solid",
      };
      if (!styles[style]) unsupported(border, "A border style is not supported.");
      const width = number(attr(border, "sz"), border, 4) / 8,
        hasColor = attr(border, "color") !== undefined || attr(border, "themeColor") !== undefined,
        color =
          styles[style] === "none" || !hasColor
            ? "#000000"
            : (this.color(border, "color") ?? "#000000");
      result[
        border.name === "insideH"
          ? "--word-border-inside-h"
          : border.name === "insideV"
            ? "--word-border-inside-v"
            : border.name === "tl2br"
              ? "--word-border-tl2br"
              : border.name === "tr2bl"
                ? "--word-border-tr2bl"
                : "border-" + sides[border.name]
      ] = pt(width) + " " + styles[style] + " " + color;
    }
    return result;
  }
  table(properties: XmlNode | undefined): Css {
    if (!properties) return {};
    if (properties.part === this.stylePart) {
      const change = elements(properties).find((node) => /PrChange$/.test(node.name));
      if (change)
        unsupported(change, "Revision properties in a shared style are preserved natively.");
    }
    check(properties, [
      "tblStyle",
      "tblpPr",
      "tblOverlap",
      "bidiVisual",
      "tblStyleRowBandSize",
      "tblStyleColBandSize",
      "tblW",
      "jc",
      "tblCellSpacing",
      "tblInd",
      "tblBorders",
      "shd",
      "tblLayout",
      "tblCellMar",
      "tblLook",
      "tblCaption",
      "tblDescription",
      "tblPrChange",
    ]);
    const result: Css = {
      "border-collapse": "collapse",
      ...this.cellSpacing(properties, true),
      ...readWordTablePosition(properties),
    };
    if (child(properties, "bidiVisual"))
      result["--word-table-bidi"] = on(child(properties, "bidiVisual")) ? "1" : "0";
    Object.assign(result, this.borders(child(properties, "tblBorders")));
    for (const [key, value] of Object.entries(this.margins(child(properties, "tblCellMar"))))
      result["--word-table-cell-" + key] = value;
    const width = child(properties, "tblW");
    if (width) {
      Object.assign(result, this.width(width, "width"));
      if (["auto", "nil"].includes(attr(width, "type") ?? "")) {
        result.width = "auto";
        result["--word-table-width-type"] = attr(width, "type")!;
      }
    }
    const layout =
      val(child(properties, "tblLayout")) ?? attr(child(properties, "tblLayout"), "type");
    if (layout) {
      if (!["fixed", "autofit"].includes(layout)) fail(properties, "Table layout is invalid.");
      result["table-layout"] = layout === "fixed" ? "fixed" : "auto";
    }
    const align = val(child(properties, "jc"));
    if (align === "center") {
      result["margin-left"] = "auto";
      result["margin-right"] = "auto";
    } else if (align === "right") result["margin-left"] = "auto";
    else if (align && align !== "left" && align !== "start") unsupported(properties);
    const indent = child(properties, "tblInd");
    if (indent) {
      const width = this.width(indent, "margin-left");
      if (!["auto", "pct"].includes(attr(indent, "type") ?? "dxa")) {
        result["margin-left"] = width["margin-left"] ?? "0pt";
        result["--word-table-indent-origin"] = "native";
      }
    }
    const shade = child(properties, "shd");
    if (shade && attr(shade, "fill") !== "auto")
      result["background-color"] = this.color(shade, "fill") ?? "transparent";
    return result;
  }
  rowException(properties: XmlNode | undefined): Css {
    if (!properties) return {};
    check(properties, [
      "tblW",
      "jc",
      "tblCellSpacing",
      "tblInd",
      "tblBorders",
      "shd",
      "tblLayout",
      "tblCellMar",
      "tblLook",
      "tblPrExChange",
    ]);
    const current = {
      ...properties,
      content: properties.content.filter(
        (item) => typeof item === "string" || item.name !== "tblPrExChange",
      ),
    };
    const css = { ...this.table(current), ...this.margins(child(current, "tblCellMar")) };
    delete css["--word-table-indent-origin"];
    delete css["border-collapse"];
    delete css["border-spacing"];
    for (const key of Object.keys(css))
      if (key.startsWith("--word-table-cell-padding-")) delete css[key];
    const alignment = val(child(current, "jc")),
      indent = child(current, "tblInd");
    if (indent) css.indent = this.width(indent, "margin-left")["margin-left"] ?? "0pt";
    if (alignment) {
      css.alignment = alignment;
      if (alignment === "center" || alignment === "right") css["margin-left"] = "auto";
      css["margin-right"] ??= "0pt";
      css["margin-left"] ??= "0pt";
    }
    const look = tableLookMask(child(current, "tblLook"));
    if (look !== undefined) css["look-mask"] = String(look);
    return wordRowExceptionCss(css);
  }
  row(properties: XmlNode | undefined): Css {
    if (!properties) return {};
    const shared = properties.part === this.stylePart;
    check(properties, [
      "cnfStyle",
      "divId",
      "gridBefore",
      "gridAfter",
      "wBefore",
      "wAfter",
      "cantSplit",
      "trHeight",
      "tblHeader",
      "tblCellSpacing",
      "jc",
      "hidden",
      "trPrChange",
      "ins",
      "del",
    ]);
    for (const name of [
      "hidden",
      ...(shared
        ? [
            "gridBefore",
            "gridAfter",
            "trHeight",
            "cnfStyle",
            "divId",
            "wBefore",
            "wAfter",
            "trPrChange",
            "ins",
            "del",
          ]
        : []),
    ] as const) {
      const value = child(properties, name);
      if (value) unsupported(value, "This row property remains native.");
    }
    const result: Css = { ...this.cellSpacing(properties) };
    const alignment = child(properties, "jc");
    if (alignment) {
      const value = val(alignment);
      if (!value || !["left", "center", "right", "start"].includes(value))
        unsupported(alignment, "This row alignment remains native.");
      result["--word-row-alignment"] = value;
    }
    const split = child(properties, "cantSplit");
    if (split) result["break-inside"] = on(split) ? "avoid" : "auto";
    const header = child(properties, "tblHeader");
    if (header) result["--word-row-repeat-header"] = on(header) ? "1" : "0";
    const height = child(properties, "trHeight");
    if (height) {
      const rule = attr(height, "hRule") ?? "atLeast";
      if (!["auto", "atLeast", "exact"].includes(rule))
        unsupported(height, "This row height rule remains native.");
      result.height = pt(number(attr(height, "val"), height, 0) / 20);
      result["--word-row-height-rule"] = rule;
    }
    return result;
  }
  cellSpacing(properties: XmlNode | undefined, table = false): Css {
    const spacing = child(properties, "tblCellSpacing");
    if (!spacing) return {};
    const type = attr(spacing, "type") ?? "dxa",
      raw = attr(spacing, "w");
    if (!["dxa", "nil", "auto", "pct"].includes(type))
      fail(spacing, "A table spacing unit is invalid.");
    if (raw !== undefined) {
      if (raw.endsWith("%")) {
        if (!/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)%$/.test(raw))
          fail(spacing, "A table spacing percentage is invalid.");
      } else measurePoints(raw, spacing);
    }
    const amount = type === "dxa" ? measurePoints(raw, spacing, 20, 0) : 0;
    if (amount < 0) fail(spacing, "A table cell spacing cannot be negative.");
    const prefix = table ? "--word-table-cell-spacing" : "--word-row-cell-spacing";
    return {
      [prefix]: pt(amount),
      [prefix + "-type"]: type,
      ...(raw !== undefined ? { [prefix + "-value"]: raw } : {}),
      ...(table
        ? { "border-collapse": amount > 0 ? "separate" : "collapse", "border-spacing": pt(amount) }
        : {}),
    };
  }
  cell(properties: XmlNode | undefined): Css {
    if (properties?.part === this.stylePart && child(properties, "cnfStyle"))
      unsupported(
        child(properties, "cnfStyle")!,
        "Shared-style conditional cache flags remain native.",
      );
    if (!properties) return {};
    if (properties.part === this.stylePart) {
      const change = elements(properties).find((node) => /PrChange$/.test(node.name));
      if (change)
        unsupported(change, "Revision properties in a shared style are preserved natively.");
    }
    check(properties, [
      "cnfStyle",
      "tcW",
      "gridSpan",
      "hMerge",
      "vMerge",
      "tcBorders",
      "shd",
      "noWrap",
      "tcMar",
      "textDirection",
      "tcFitText",
      "vAlign",
      "hideMark",
      "tcPrChange",
      "cellIns",
      "cellDel",
    ]);
    if (child(properties, "hMerge") || on(child(properties, "tcFitText"))) unsupported(properties);
    const result: Css = { ...this.borders(child(properties, "tcBorders")) };
    const width = child(properties, "tcW");
    if (width) Object.assign(result, this.width(width, "width"));
    const shade = child(properties, "shd");
    if (shade && attr(shade, "fill") !== "auto")
      result["background-color"] = this.color(shade, "fill") ?? "transparent";
    const align = val(child(properties, "vAlign"));
    if (align) {
      if (!["top", "center", "bottom"].includes(align)) unsupported(properties);
      result["vertical-align"] = align === "center" ? "middle" : align;
    }
    if (on(child(properties, "noWrap"))) result["white-space"] = "nowrap";
    const hideMark = child(properties, "hideMark");
    if (hideMark) result["--word-cell-hide-mark"] = on(hideMark) ? "1" : "0";
    const direction = val(child(properties, "textDirection"));
    if (direction) {
      if (!["lrTb", "tbRl", "btLr"].includes(direction)) unsupported(properties!);
      result["--word-cell-text-direction"] = direction;
    }
    Object.assign(result, this.margins(child(properties, "tcMar")));
    return result;
  }
  margins(properties: XmlNode | undefined): Css {
    const result: Css = {};
    for (const item of elements(properties)) {
      const name = item.name === "start" ? "left" : item.name === "end" ? "right" : item.name;
      if (!["top", "bottom", "left", "right"].includes(name)) unsupported(item);
      Object.assign(result, this.width(item, "padding-" + name));
    }
    return result;
  }
  width(node: XmlNode, property: string): Css {
    const type = attr(node, "type") ?? "dxa";
    const raw = attr(node, "w");
    if (type === "auto" || type === "nil") return {};
    if (type === "pct") {
      const literal = raw && /^([+-]?(?:\d+(?:\.\d+)?|\.\d+))%$/.exec(raw),
        value = literal ? Number(literal[1]) : number(raw, node) / 50;
      if (value < 0 && property !== "margin-left")
        unsupported(node, "This negative preferred width remains native.");
      return { [property]: String(value) + "%" };
    }
    if (type !== "dxa") return unsupported(node, "An OOXML width unit is not supported.");
    const value = measurePoints(raw, node);
    if (value < 0 && property !== "margin-left")
      unsupported(node, "This negative preferred width remains native.");
    return { [property]: pt(value) };
  }
  resolve(id: string, visiting: Set<string> = new Set<string>()): StyleValues {
    const cached = this.resolved.get(id);
    if (cached) return cached;
    const source = this.definitions.get(id);
    if (!source)
      throw new DocumentError("InvalidDocument", "A referenced native style is missing.");
    if (visiting.has(id)) return fail(source, "Style inheritance is cyclic.");
    visiting.add(id);
    const parentId = val(child(source, "basedOn"));
    if (parentId && !this.definitions.has(parentId))
      fail(source, "A native base style is missing.");
    const parent = parentId ? this.resolve(parentId, visiting) : empty();
    const pPr = child(source, "pPr");
    const numPr = child(pPr, "numPr");
    const num = val(child(numPr, "numId"));
    const paragraph = this.paragraph(pPr);
    const numbering =
      num !== undefined || child(numPr, "ilvl") || parent.numbering
        ? {
            ...(parent.numbering?.id !== undefined ? { id: parent.numbering.id } : {}),
            ...(num !== undefined ? { id: num } : {}),
            indent: {
              ...(num !== undefined && num !== parent.numbering?.id ? {} : parent.numbering?.indent),
              ...Object.fromEntries(
                Object.entries(paragraph).filter(([name]) =>
                  ["margin-left", "margin-right", "text-indent"].includes(name),
                ),
              ),
            },
          }
        : undefined;
    const outline = child(pPr, "outlineLvl");
    const tableProperties = child(source, "tblPr");
    const intrinsic = this.intrinsicTableStyle(source);
    this.cellSpacing(child(source, "trPr"));
    const result: StyleValues = {
      p: { ...parent.p, ...paragraph },
      r: {
        ...parent.r,
        ...this.run(child(source, "rPr"), parent.r, attr(source, "type") !== "table"),
      },
      table: { ...intrinsic.table, ...parent.table, ...this.table(tableProperties) },
      cell: {
        ...intrinsic.cell,
        ...parent.cell,
        ...this.margins(child(tableProperties, "tblCellMar")),
        ...this.cell(child(source, "tcPr")),
      },
      ...(numbering ? { numbering } : {}),
      tabs: mergeTabs(parent.tabs ?? this.defaults.tabs, child(pPr, "tabs")),
      ...(parent.outline !== undefined ? { outline: parent.outline } : {}),
      ...(outline ? { outline: number(val(outline), outline) } : {}),
    };
    visiting.delete(id);
    this.resolved.set(id, result);
    return result;
  }
  /** Word identifies a stock style by its canonical/localized name, not its
   * document-local styleId. Materialize its fallback before ordinary inheritance
   * so core editing and export need no stock-name branches. */
  private intrinsicTableStyle(source: XmlNode): { table: Css; cell: Css } {
    const name = val(child(source, "name"))?.toLowerCase();
    if (attr(source, "type") !== "table" || ["1", "true", "on"].includes(attr(source, "customStyle") ?? "")
      || (name !== "normal table" && name !== "普通表格")) return { table: {}, cell: {} };
    return {
      table: { "margin-left": "0pt", "--word-table-indent-origin": "native" },
      cell: { "padding-top": "0pt", "padding-bottom": "0pt", "padding-left": "5.4pt", "padding-right": "5.4pt" },
    };
  }
  hasStyle(id: string, kind: string): boolean {
    return attr(this.definitions.get(id), "type") === kind;
  }
  runStyle(id: string | undefined, inherited: Css): Css {
    if (!id) return { ...inherited };
    const chain: XmlNode[] = [];
    const seen = new Set<string>();
    let current: string | undefined = id;
    while (current) {
      const definition = this.definitions.get(current);
      if (!definition)
        throw new DocumentError("InvalidDocument", "A referenced native style is missing.");
      if (seen.has(current)) fail(definition, "Style inheritance is cyclic.");
      seen.add(current);
      chain.push(definition);
      current = val(child(definition, "basedOn"));
    }
    let result = { ...inherited };
    for (const definition of chain.reverse())
      result = {
        ...result,
        ...this.run(child(definition, "rPr"), result, true),
      };
    return result;
  }
  get(id: string | undefined): StyleValues {
    return id ? this.resolve(id) : empty();
  }
  tableSetting(
    id: string | undefined,
    key: "tableLookMask" | "rowBandSize" | "columnBandSize",
  ): number | undefined {
    let current = id;
    const seen = new Set<string>();
    while (current && !seen.has(current)) {
      seen.add(current);
      const definition = this.styles.find((item) => item.styleId === current);
      if (definition?.[key] !== undefined) return definition[key];
      current = definition?.basedOn;
    }
    return undefined;
  }
  runClass(id: string | undefined): string[] {
    const metadata = this.styles.find((item) => item.styleId === id);
    return metadata?.runCssClass ? [metadata.runCssClass] : [];
  }
  styleClass(id: string): string {
    const item = this.styles.find((style) => style.styleId === id);
    if (!item) throw new Error("Missing style");
    return item.cssClass;
  }
  paragraphNumbering(
    source: XmlNode,
    properties: XmlNode | undefined,
    styleId: string | undefined,
  ): { id: string; level: number; paragraph: Css; tabs?: XmlNode } | undefined {
    const num = child(properties, "numPr"),
      directId = val(child(num, "numId")),
      style = this.get(styleId).numbering,
      id = directId ?? style?.id;
    if (id === undefined) return;
    if (id === "0") return { id, level: 0, paragraph: {} };
    const directLevel = child(num, "ilvl");
    let level = directLevel ? number(val(directLevel), directLevel) : undefined;
    for (let current = styleId; level === undefined && current;) {
      const definition = this.definitions.get(current),
        associated = this.numberStyleLevels.get(id + ":" + current),
        styleLevel = child(child(child(definition, "pPr"), "numPr"), "ilvl");
      level = associated ?? (styleLevel ? number(val(styleLevel), styleLevel) : undefined);
      current = val(child(definition, "basedOn"));
    }
    level ??= 0;
    const definition = this.numberProperties.get(id + ":" + level);
    if (!definition) fail(source, "A paragraph refers to missing numbering.");
    return {
      ...definition,
      id,
      level,
      paragraph: {
        ...definition.paragraph,
        // A style can explicitly position its own list, including zero
        // indentation. A list applied directly still uses the selected level.
        ...(directId === undefined && directLevel === undefined ? style?.indent : {}),
      },
    };
  }
  private readNumbering(root: XmlNode, numberingPart: string): void {
    check(root, ["numPicBullet", "abstractNum", "num", "numIdMacAtCleanup"]);
    const picBullets = new Map<string, { assetPath: string; widthPt?: number; heightPt?: number }>();
    const findDescendant = (node: XmlNode, predicate: (item: XmlNode) => boolean): XmlNode | undefined => {
      for (const item of node.content) {
        if (typeof item === "string") continue;
        if (predicate(item)) return item;
        const found = findDescendant(item, predicate);
        if (found) return found;
      }
      return undefined;
    };
    const numRels = this.opc.relationships(numberingPart);
    for (const bullet of elements(root, "numPicBullet")) {
      const bulletId = attr(bullet, "numPicBulletId");
      if (!bulletId) continue;
      if (!this.onImage) unsupported(bullet, "Picture bullets require image asset registration.");
      let relId: string | undefined;
      let widthPt: number | undefined;
      let heightPt: number | undefined;
      const drawing = child(bullet, "drawing");
      if (drawing) {
        const blip = findDescendant(drawing, (item) => item.uri === A && item.name === "blip");
        relId = attr(blip, "embed", R) ?? attr(blip, "link", R);
        const extent = findDescendant(drawing, (item) => item.uri === WP && item.name === "extent");
        const cx = Number(attr(extent, "cx", ""));
        const cy = Number(attr(extent, "cy", ""));
        if (Number.isFinite(cx) && cx > 0) widthPt = Math.round((cx / 12700) * 100) / 100;
        if (Number.isFinite(cy) && cy > 0) heightPt = Math.round((cy / 12700) * 100) / 100;
      } else {
        const pict = child(bullet, "pict");
        const imagedata = pict ? findDescendant(pict, (item) => item.name === "imagedata") : undefined;
        relId = attr(imagedata, "id", R);
        const shape = pict ? findDescendant(pict, (item) => item.name === "shape") : undefined;
        const styleText = attr(shape, "style", "") ?? "";
        const wm = /width\s*:\s*([0-9.]+)pt/i.exec(styleText);
        const hm = /height\s*:\s*([0-9.]+)pt/i.exec(styleText);
        if (wm) widthPt = Number(wm[1]);
        if (hm) heightPt = Number(hm[1]);
      }
      if (!relId) unsupported(bullet, "Picture bullet image reference is missing.");
      const rel = numRels.get(relId);
      if (!rel || rel.external) unsupported(bullet, "External or missing picture bullet image.");
      const assetPath = this.onImage(rel.target, bullet);
      picBullets.set(bulletId, {
        assetPath,
        ...(widthPt && widthPt > 0 && widthPt <= 144 ? { widthPt } : {}),
        ...(heightPt && heightPt > 0 && heightPt <= 144 ? { heightPt } : {}),
      });
    }
    const abstract = new Map<string, XmlNode>();
    const abstractByStyleLink = new Map<string, XmlNode>();
    for (const definition of elements(root, "abstractNum")) {
      const id = attr(definition, "abstractNumId");
      if (!id || abstract.has(id)) fail(definition, "An abstract numbering identity is invalid.");
      check(definition, [
        "nsid",
        "multiLevelType",
        "tmpl",
        "name",
        "styleLink",
        "numStyleLink",
        "lvl",
      ]);
      abstract.set(id, definition);
      const styleLink = val(child(definition, "styleLink"));
      if (styleLink) abstractByStyleLink.set(styleLink, definition);
    }
    const seen = new Set<string>();
    for (const definition of elements(root, "num")) {
      const id = attr(definition, "numId");
      const ref = val(child(definition, "abstractNumId"));
      if (!id || seen.has(id) || !ref || !abstract.has(ref))
        fail(definition, "A numbering instance is invalid.");
      seen.add(id);
      const absNode = abstract.get(ref)!;
      const numStyleLink = val(child(absNode, "numStyleLink"));
      const styleLink = val(child(absNode, "styleLink"));
      const name = val(child(absNode, "name"));
      const rawMulti = val(child(absNode, "multiLevelType"));
      const multiLevelType =
        rawMulti === "singleLevel" || rawMulti === "multilevel" || rawMulti === "hybridMultilevel"
          ? rawMulti
          : undefined;
      const levelOwner =
        elements(absNode, "lvl").length > 0
          ? absNode
          : numStyleLink && abstractByStyleLink.has(numStyleLink)
            ? abstractByStyleLink.get(numStyleLink)!
            : absNode;
      const levels: WordNumberLevel[] = [];
      for (const source of elements(levelOwner, "lvl")) {
        check(source, [
          "start",
          "numFmt",
          "lvlRestart",
          "pStyle",
          "isLgl",
          "suff",
          "lvlText",
          "lvlPicBulletId",
          "legacy",
          "lvlJc",
          "pPr",
          "rPr",
        ]);
        if (child(source, "legacy")) unsupported(source);
        const level = number(attr(source, "ilvl"), source);
        const override = elements(definition, "lvlOverride").find(
          (item) => number(attr(item, "ilvl"), item) === level,
        );
        if (child(override, "lvl"))
          unsupported(override!, "Full numbering-level overrides are not supported.");
        const format = val(child(source, "numFmt")) ?? "decimal";
        if (
          ![
            "decimal",
            "bullet",
            "lowerLetter",
            "upperLetter",
            "lowerRoman",
            "upperRoman",
            "none",
          ].includes(format)
        )
          unsupported(source, "The numbering format is not supported.");
        const cssClass = "wn" + this.numbering.length + "l" + level;
        const pPr = child(source, "pPr"),
          paragraph = this.paragraph(pPr),
          tabs = child(pPr, "tabs"),
          styleId = val(child(source, "pStyle"));
        this.numberProperties.set(id + ":" + level, { paragraph, ...(tabs ? { tabs } : {}) });
        if (styleId && !this.numberStyleLevels.has(id + ":" + styleId))
          this.numberStyleLevels.set(id + ":" + styleId, level);
        this.css.push(
          "." +
            cssClass +
            "{" +
            declarations({ ...paragraph, ...this.run(child(source, "rPr")) }) +
            "}",
        );
        const restart = child(source, "lvlRestart");
        const picBulletRef = val(child(source, "lvlPicBulletId"));
        const picBullet = picBulletRef !== undefined ? picBullets.get(picBulletRef) : undefined;
        const isLgl = on(child(source, "isLgl")) ? true : undefined;
        levels.push({
          level,
          start: number(
            val(child(override, "startOverride")) ?? val(child(source, "start")),
            source,
            1,
          ),
          format: format as WordNumberLevel["format"],
          text: val(child(source, "lvlText")) ?? "%" + (level + 1) + ".",
          suffix: (val(child(source, "suff")) ?? "tab") as WordNumberLevel["suffix"],
          cssClass,
          ...(restart ? { restart: number(val(restart), restart) } : {}),
          ...(picBullet
            ? {
                pictureBulletAssetPath: picBullet.assetPath,
                ...(picBullet.widthPt !== undefined
                  ? { pictureBulletWidthPt: picBullet.widthPt }
                  : {}),
                ...(picBullet.heightPt !== undefined
                  ? { pictureBulletHeightPt: picBullet.heightPt }
                  : {}),
              }
            : {}),
          ...(styleId ? { pStyle: styleId } : {}),
          ...(isLgl ? { isLgl: true } : {}),
        });
      }
      if (!levels.length) continue;
      this.numbering.push({
        numberingId: id,
        ...(name ? { name } : {}),
        ...(multiLevelType ? { multiLevelType } : {}),
        ...(styleLink ? { styleLink } : {}),
        ...(numStyleLink ? { numStyleLink } : {}),
        levels,
      });
    }
  }
}
