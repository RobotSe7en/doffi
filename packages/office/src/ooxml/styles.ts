import {passiveWordRunBooleans,passiveWordRunValues,passiveWordRunCss} from './word-passive-run-properties.js';
import {passiveWordParagraphBooleans,passiveWordParagraphCss} from './word-passive-paragraph-properties.js';
import {measurePoints} from './measure.js';
import {wordColorLuminance} from '@doffi/core';
import { DocumentError } from "@doffi/core";
import type { WordMetadata, WordNumberLevel } from "@doffi/core";
import type { OpcPackage } from "./package.js";
import {
  A,
  W,
  attr,
  check,
  child,
  elements,
  fail,
  number,
  on,
  unsupported,
  unsupportedType,
} from "./xml.js";
import type { XmlNode } from "./xml.js";

export type Css = Record<string, string>;
type TabStops = WordMetadata["tabs"][number]["stops"];
/** Tab declarations amend the inherited set by position, including removals. */
export function mergeTabs(
  inherited: TabStops = [],
  properties?: XmlNode,
): TabStops {
  const stops = new Map(inherited.map((stop) => [stop.position, stop]));
  check(properties, ["tab"]);
  for (const stop of elements(properties, "tab")) {
    const alignment = attr(stop, "val") ?? "left";
    const leader = attr(stop, "leader") ?? "none";
    if (
      !["left", "center", "right", "decimal", "clear"].includes(alignment) ||
      !["none", "dot", "hyphen", "underscore", "heavy", "middleDot"].includes(
        leader,
      )
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
      ([name]) =>
        name === "background-color" || name.startsWith("text-decoration"),
    ),
  );
}
export function inheritedRun(css: Css): Css {
  return Object.fromEntries(
    Object.entries(css).filter(
      ([name]) =>
        name !== "background-color" && !name.startsWith("text-decoration"),
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
const val = (node: XmlNode | undefined): string | undefined =>
  attr(node, "val");
const quoteFont = (name: string): string =>
  '"' + name.replace(/["\\]/g, "\\$&").replace(/[\u0000-\u001f]/g, "") + '"';
export interface StyleValues {
  runClasses?: string[];
  p: Css;
  r: Css;
  table: Css;
  cell: Css;
  row?:Css;
  numbering?: { id: string; level: number };
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
  readonly styles: WordMetadata["styles"] = [];
  readonly numbering: WordMetadata["numbering"] = [];
  readonly defaults: StyleValues;
  readonly defaultParagraph: string | undefined;
  readonly defaultCharacter: string | undefined;
  readonly defaultTable: string | undefined;
  private readonly numberProperties = new Map<string, Css>();
  private readonly definitions = new Map<string, XmlNode>();
  private readonly resolved = new Map<string, StyleValues>();
  private readonly themeColors = new Map<string, string>();
  private readonly themeFonts = new Map<string, string>();
  private readonly locale: string;
  readonly language: string | undefined;

  private readonly stylePart:string|undefined;
  constructor(opc: OpcPackage, main: string) {
    this.stylePart=opc.related(main,"styles");
    const stylePath = opc.related(main, "styles");
    const styleRoot = stylePath ? opc.read(stylePath) : undefined;
    const defaultLanguage = child(
      child(child(child(styleRoot, "docDefaults"), "rPrDefault"), "rPr"),
      "lang",
    );
    this.language = attr(defaultLanguage, "val");
    this.locale =
      attr(defaultLanguage, "eastAsia") ??
      attr(defaultLanguage, "val") ??
      "en-US";
    const themePath = opc.related(main, "theme");
    if (themePath) {
      const root = opc.read(themePath);
      const theme = child(root, "themeElements", A);
      for (const color of elements(
        child(theme, "clrScheme", A),
        undefined,
        A,
      )) {
        const values = elements(color, undefined, A);
        const item = values[0];
        const value = attr(
          item,
          item?.name === "sysClr" ? "lastClr" : "val",
          "",
        );
        if (value && /^[0-9a-f]{6}$/i.test(value))
          this.themeColors.set(color.name, value);
      }
      const fonts = child(theme, "fontScheme", A);
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
          if (script && face)
            this.themeFonts.set(family + "EastAsia:" + script, face);
        }
      }
    }
    const root = styleRoot;
    if (root && (root.uri !== W || root.name !== "styles"))
      fail(root, "The styles part has an invalid root.");
    check(root, ["docDefaults", "latentStyles", "style"]);
    const docDefaults = child(root, "docDefaults");
    this.defaults = empty();
    Object.assign(
      this.defaults.r,
      this.run(child(child(docDefaults, "rPrDefault"), "rPr")),
    );
    Object.assign(
      this.defaults.p,
      this.paragraph(child(child(docDefaults, "pPrDefault"), "pPr")),
    );
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
      const conditions: NonNullable<
        WordMetadata["styles"][number]["conditions"]
      > = [];
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
      for(const key of ['name','aliases']as const){const value=val(child(style,key));if(value!==undefined)metadata[key]=value;}
      for(const key of ['default','customStyle']as const){const value=attr(style,key);if(value!==undefined)metadata[key]=['1','true','on'].includes(value);}
      for(const key of ['hidden','semiHidden','unhideWhenUsed','qFormat','locked']as const){const value=child(style,key);if(value)metadata[key]=on(value);}
      const priority=child(style,'uiPriority');if(priority)metadata.uiPriority=number(val(priority),priority);
      const outline=child(child(style,'pPr'),'outlineLvl');if(outline)metadata.outlineLevel=number(val(outline),outline);

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
          unsupported(
            table!,
            "Table indentation cannot vary by conditional cell region.",
          );
        const values: StyleValues = {
          p: this.paragraph(child(region, "pPr")),
          r: this.run(child(region, "rPr")),
          table: this.table(table),row:this.row(row),
          cell: {
            ...this.margins(child(table, "tblCellMar")),
            ...this.cell(child(region, "tcPr")),
          },
        };
        conditional.add(type);
        const cssClass = metadata.cssClass + "-" + type;
        conditions.push({
          region: type as NonNullable<
            typeof metadata.conditions
          >[number]["region"],
          cssClass,
          components:Object.fromEntries((['p','r','table','cell','row']as const).map(key=>[key,this.component(cssClass+'-'+key,values[key]??{})])),
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
      ".word-default-run{" +
        declarations(nonInheritedRun(this.defaults.r)) +
        "}",
    );
    for (const style of this.styles) {
      const value = this.resolve(style.styleId);
      const definition = this.definitions.get(style.styleId)!;
      const parent = style.basedOn ? this.resolve(style.basedOn) : empty();
      style.paragraphProperties = Object.keys(
        this.paragraph(child(definition, "pPr")),
      );
      style.runProperties = Object.keys(
        this.run(child(definition, "rPr"), parent.r, style.kind!=="table"),
      );
      if(style.kind==='table'||style.kind==='paragraph'||style.kind==='character')style.components={
        p:this.component(style.cssClass+'-p',this.paragraph(child(definition,'pPr'))),
        r:this.component(style.cssClass+'-r',this.run(child(definition,'rPr'),{},style.kind!=='table')),
        ...(style.kind==='table'?{table:this.component(style.cssClass+'-table',this.table(child(definition,'tblPr'))),row:this.component(style.cssClass+'-row',this.row(child(definition,'trPr'))),cell:this.component(style.cssClass+'-cell',{...this.margins(child(child(definition,'tblPr'),'tblCellMar')),...this.cell(child(definition,'tcPr'))})}:{})
      };
      const weight =
        style.kind === "character" ? 4 : style.kind === "paragraph" ? 3 : 1;
      const rules =
        style.kind === "character"
          ? value.r
          : style.kind === "table"
            ? { ...value.table, ...inheritedRun(value.r) }
            : { ...value.p, ...inheritedRun(value.r) };
      this.css.push(
        ("." + style.cssClass).repeat(weight) + "{" + declarations(rules) + "}",
      );
      if (style.kind !== "character") {
        const special = nonInheritedRun(value.r);
        if (Object.keys(special).length) {
          style.runCssClass = style.cssClass + "-run";
          this.css.push(
            ("." + style.runCssClass).repeat(
              style.kind === "paragraph" ? 3 : 1,
            ) +
              "{" +
              declarations(special) +
              "}",
          );
        }
      }
    }
    const numPath = opc.related(main, "numbering");
    if (numPath) this.readNumbering(opc.read(numPath));
  }
  private readonly componentValues=new Map<string,Css>();
  component(name:string,values:Css):string{this.componentValues.set(name,values);this.css.push('.'+name+'{'+declarations(values)+'}');return name;}
  readComponent=(name:string):Css=>this.componentValues.get(name)??{};
  drawingThemeProperties():Css{
    const result:Css={};for(const[role,value]of this.themeColors)result['--office-theme-'+role]='#'+value;
    for(const[role,name]of Object.entries({bg1:'lt1',tx1:'dk1',bg2:'lt2',tx2:'dk2'}))if(result['--office-theme-'+name])result['--office-theme-'+role]=result['--office-theme-'+name]!;
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
    let value = theme
      ? this.themeColors.get(roles[theme] ?? theme)
      : attr(node, attribute);
    if (value === "auto") return undefined;
    if (!value || !/^[0-9a-f]{6}$/i.test(value))
      return unsupported(
        node,
        "A document color or theme role cannot be resolved.",
      );
    const tint = attr(
      node,
      attribute === "fill" ? "themeFillTint" : "themeTint",
    );
    const shade = attr(
      node,
      attribute === "fill" ? "themeFillShade" : "themeShade",
    );
    const modifier = tint ?? shade;
    if (modifier !== undefined) {
      if (!/^[0-9a-f]{2}$/i.test(modifier))
        fail(node, "A theme color modifier is invalid.");
      value = wordColorLuminance(value, tint, shade);
    }
    return "#" + value.toLowerCase();
  }
  run(
    properties: XmlNode | undefined,
    inherited: Css = {},
    toggle = false,paragraphMark=false,allowRevision=false,
  ): Css {
    if (!properties) return {};
    if(properties.part===this.stylePart){const change=elements(properties).find(node=>/PrChange$/.test(node.name));if(change)unsupported(change,'Revision properties in a shared style are preserved natively.');}
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
      "oMath",...(allowRevision?["rPrChange"]:[]),...(paragraphMark?["ins","del","moveFrom","moveTo"]:[]),
    ]);
    const css: Css = {};
    if(child(properties,"snapToGrid")) css["--word-run-snap-to-grid"]=on(child(properties,"snapToGrid"))?"1":"0";
    for(const role of ['val','eastAsia','bidi']){const language=attr(child(properties,'lang'),role);if(language!==undefined)css['--word-language-'+role]=JSON.stringify(language);}
    const fonts = child(properties, "rFonts");
    if (fonts) {
      for (const [name, themeName] of [
        ["ascii", "asciiTheme"],
        ["hAnsi", "hAnsiTheme"],
        ["eastAsia", "eastAsiaTheme"],
        ["cs", "cstheme"],
      ]) {
        const role = attr(fonts, themeName!);
        const language =
          attr(child(properties, "lang"), "eastAsia") ??
          attr(child(properties, "lang"), "val") ??
          this.locale;
        const script = /^ja/i.test(language)
          ? "Jpan"
          : /^ko/i.test(language)
            ? "Hang"
            : /^zh-(?:TW|HK|MO|Hant)/i.test(language)
              ? "Hant"
              : "Hans";
        const face = role
          ? (this.themeFonts.get(role) ??
            this.themeFonts.get(role + ":" + script))
          : attr(fonts, name!);
        if (role && !face)
          unsupported(fonts, "A theme font cannot be resolved.");
        if (face) {
          css["--word-font-" + name!] = quoteFont(face);
        }
      }
      if (css["--word-font-ascii"] || css["--word-font-hAnsi"])
        css["font-family"] =
          "var(--word-font-ascii,var(--word-font-hAnsi,inherit))";
    }
    for (const [name, property, active, inactive] of [
      ["b", "font-weight", "bold", "normal"],
      ["i", "font-style", "italic", "normal"],
      ["caps", "text-transform", "uppercase", "none"],
      ["smallCaps", "font-variant", "small-caps", "normal"],
    ]) {
      const item = child(properties, name!);
      if (item) {
        const enabled = on(item);
        if (toggle && !enabled) continue;
        css[property!] = (
          toggle && enabled && inherited[property!] === active ? false : enabled
        )
          ? active!
          : inactive!;
      }
    }
    if(child(properties,"fitText"))unsupportedType(child(properties,"fitText")!);
    for (const feature of [
      "outline",
      "shadow",
      "emboss",
      "imprint",
      "effect",
      "eastAsianLayout",
      "oMath",
    ]) {
      const item = child(properties, feature);
      if (item && (feature === "effect" ? val(item) !== "none" : on(item)))
        unsupported(item);
    }
    for (const feature of ["vanish", "webHidden", "rtl", "cs", "specVanish"])
      if (on(child(properties, feature)))
        unsupported(child(properties, feature)!);
    const kerning = child(properties, "kern");
    if (kerning)
      css["--word-kerning-threshold"] = pt(number(val(kerning), kerning) / 2);
    const complexSize = child(properties, "szCs");
    if (complexSize)
      css["--word-font-size-cs"] = pt(
        number(val(complexSize), complexSize) / 2,
      );
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
    if (highlight === "none")
      css["background-color"] = "var(--word-run-shading,transparent)";
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
      if (!colors[highlight])
        unsupported(properties, "An unknown highlight color is used.");
      css["background-color"] = colors[highlight]!;
    }
    const underline = child(properties, "u");
    const strike = child(properties, "strike"),
      doubleStrike = child(properties, "dstrike");
    if (underline || strike || doubleStrike) {
      const lines = new Set(
        (inherited["text-decoration-line"] ?? "")
          .split(/\s+/)
          .filter((value) => value && value !== "none"),
      );
      const type = val(underline) ?? "single";
      if (underline) {
        if (
          !["none", "single", "double", "dotted", "dash", "wave"].includes(type)
        )
          unsupported(underline);
        if (type === "none") lines.delete("underline");
        else lines.add("underline");
      }
      const strikeNode = doubleStrike ?? strike;
      if (strikeNode) {
        const enabled = on(strikeNode);
        if (!toggle || enabled) {
          if (enabled && !(toggle && lines.has("line-through")))
            lines.add("line-through");
          else lines.delete("line-through");
        }
      }
      css["text-decoration-line"] = [...lines].join(" ") || "none";
      if (
        doubleStrike &&
        on(doubleStrike) &&
        underline &&
        !["none", "double"].includes(type)
      )
        unsupported(
          properties,
          "Mixed underline and double-strike decoration styles are not supported.",
        );
      if (doubleStrike && on(doubleStrike))
        css["text-decoration-style"] = "double";
      else if (underline)
        css["text-decoration-style"] =
          type === "double"
            ? "double"
            : type === "dotted"
              ? "dotted"
              : type === "dash"
                ? "dashed"
                : type === "wave"
                  ? "wavy"
                  : "solid";
      if (
        underline &&
        (attr(underline, "color") || attr(underline, "themeColor"))
      )
        css["text-decoration-color"] =
          this.color(underline, "color") ?? "currentColor";
    }
    const align = val(child(properties, "vertAlign"));
    if (align) {
      if (!["baseline", "superscript", "subscript"].includes(align))
        unsupported(properties);
      css["vertical-align"] =
        align === "superscript"
          ? "super"
          : align === "subscript"
            ? "sub"
            : "baseline";
    }
    const spacing = child(properties, "spacing");
    if (spacing) css["letter-spacing"] = pt(number(val(spacing), spacing) / 20);
    for (const feature of ["w", "position", "bdr", "em"]) {
      const item = child(properties, feature);
      if (
        item &&
        !(feature === "w" && val(item) === "100") &&
        !(feature === "position" && val(item) === "0") &&
        !(feature === "em" && val(item) === "none")
      )
        unsupported(item);
    }
    for(const name of passiveWordRunBooleans){const item=child(properties,name);if(item)css[passiveWordRunCss(name)]=on(item)?'1':'0';}
    for(const name of passiveWordRunValues){const item=child(properties,name);if(item)css[passiveWordRunCss(name)]=val(item)!;}
    return css;
  }
  paragraph(properties: XmlNode | undefined): Css {
    if(properties?.part===this.stylePart&&child(properties,'cnfStyle'))unsupported(child(properties,'cnfStyle')!,'Shared-style conditional cache flags remain native.');
    if (!properties) return {};
    if(properties.part===this.stylePart){const change=elements(properties).find(node=>/PrChange$/.test(node.name));if(change)unsupported(change,'Revision properties in a shared style are preserved natively.');}
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
      "sectPr", "pPrChange",
    ]);
    for (const name of [
      "topLinePunct",
      "adjustRightInd",
      "contextualSpacing",
      "mirrorIndents",
      "suppressOverlap",
      "textAlignment",
      "textboxTightWrap",
    ]) {
      const item = child(properties, name);
      if (item)
        unsupported(item, "This paragraph layout property is not yet mapped.");
    }
    const css: Css = {};
    for(const [name,key] of [['kinsoku','--word-kinsoku'],['wordWrap','--word-word-wrap'],['autoSpaceDE','--word-auto-space-de'],['autoSpaceDN','--word-auto-space-dn']] as const){const item=child(properties,name);if(item)css[key]=on(item)?'1':'0';}
    if(child(properties,"snapToGrid")) css["--word-paragraph-snap-to-grid"]=on(child(properties,"snapToGrid"))?"1":"0";
    const overflowPunctuation = child(properties, "overflowPunct");
    if (overflowPunctuation) {
      if (on(overflowPunctuation))
        unsupported(
          overflowPunctuation,
          "Hanging punctuation is not yet supported.",
        );
      css["--word-overflow-punctuation"] = "0";
    }
    const suppressHyphens = child(properties, "suppressAutoHyphens");
    if (suppressHyphens)
      css["--word-suppress-auto-hyphens"] = on(suppressHyphens) ? "1" : "0";
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
      if (!values[align])
        unsupported(properties, "A paragraph alignment is not supported.");
      css["text-align"] = values[align]!;css["--word-distributed"]=align==="distribute"?"1":"0";
    }
    const spacing = child(properties, "spacing");
    if (spacing) {
      for (const [name, property] of [
        ["before", "margin-top"],
        ["after", "margin-bottom"],
      ]) {
        const value = attr(spacing, name!);
        if (value !== undefined)
          css[property!] = pt(number(value, spacing) / 20);
      }
      if (
        attr(spacing, "beforeLines") ||
        attr(spacing, "afterLines") ||
        attr(spacing, "beforeAutospacing") === "1" ||
        attr(spacing, "afterAutospacing") === "1"
      )
        unsupported(spacing, "Automatic paragraph spacing is not supported.");
      const line = attr(spacing, "line");
      if (line) {
        const rule = attr(spacing, "lineRule") ?? "auto";
        if (!["auto", "exact", "atLeast"].includes(rule))
          fail(spacing, "A line spacing rule is invalid.");
        css["--word-line-rule"]=rule;
        css["line-height"] =
          rule === "auto"
            ? String(number(line, spacing) / 240)
            : pt(number(line, spacing) / 20);
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
        if (value !== undefined)
          css[property!] = pt(number(value, indent) / 20);
      }
      const hanging = attr(indent, "hanging");
      if (hanging !== undefined)
        css["text-indent"] = pt(-number(hanging, indent) / 20);
      if (
        ["leftChars", "rightChars", "firstLineChars", "hangingChars"].some(
          (name) => attr(indent, name) !== undefined,
        )
      )
        unsupported(indent, "Character-unit indentation is not supported.");
    }
    if (child(properties, "keepNext"))
      css["break-after"] = on(child(properties, "keepNext")) ? "avoid" : "auto";
    if (child(properties, "keepLines"))
      css["break-inside"] = on(child(properties, "keepLines"))
        ? "avoid"
        : "auto";
    if (child(properties, "pageBreakBefore"))
      css["break-before"] = on(child(properties, "pageBreakBefore"))
        ? "page"
        : "auto";
    if (child(properties, "widowControl")) {
      const count = on(child(properties, "widowControl")) ? "2" : "1";
      css.widows = count;
      css.orphans = count;
    }
    for(const name of ['framePr','bidi','textDirection']){const item=child(properties,name);if(item&&(name==='framePr'||name==='bidi'&&on(item)||name==='textDirection'&&val(item)!=='lrTb'))unsupported(item,'A paragraph frame or writing direction is not supported.');}
    for(const name of passiveWordParagraphBooleans){const item=child(properties,name);if(item)css[passiveWordParagraphCss(name)]=on(item)?'1':'0';}
    if(child(properties,'textDirection'))css[passiveWordParagraphCss('textDirection')]='lrTb';
    const div=child(properties,'divId');if(div){const id=number(val(div),div);if(!Number.isSafeInteger(id)||id<0)fail(div,'A paragraph div identity must be a nonnegative integer.');css[passiveWordParagraphCss('divId')]=String(id);}
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
        return unsupported(
          border,
          "Between-paragraph borders are not yet supported.",
        );
      if (!sides[border.name] && !["insideH", "insideV"].includes(border.name))
        unsupported(border);
      if (
        number(attr(border, "space"), border, 0) !== 0 ||
        ["shadow", "frame"].some((name) =>
          ["true", "1", "on"].includes(attr(border, name) ?? ""),
        )
      )
        unsupported(
          border,
          "Border spacing and border effects are not supported.",
        );
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
      if (!styles[style])
        unsupported(border, "A border style is not supported.");
      const width = number(attr(border, "sz"), border, 4) / 8;
      result[
        border.name === "insideH"
          ? "--word-border-inside-h"
          : border.name === "insideV"
            ? "--word-border-inside-v"
            : "border-" + sides[border.name]
      ] =
        pt(width) +
        " " +
        styles[style] +
        " " +
        (this.color(border, "color") ?? "#000000");
    }
    return result;
  }
  table(properties: XmlNode | undefined): Css {
    if (!properties) return {};
    if(properties.part===this.stylePart){const change=elements(properties).find(node=>/PrChange$/.test(node.name));if(change)unsupported(change,'Revision properties in a shared style are preserved natively.');}
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
      "tblDescription", "tblPrChange",
    ]);
    if (child(properties, "tblpPr"))
      unsupported(properties, "This floating table retains its native positioning.");
    const result: Css = { "border-collapse": "collapse",...this.cellSpacing(properties,true) };
    if(child(properties,'bidiVisual'))result['--word-table-bidi']=on(child(properties,'bidiVisual'))?'1':'0';
    Object.assign(result, this.borders(child(properties, "tblBorders")));
    const width = child(properties, "tblW");
    if (width) Object.assign(result, this.width(width, "width"));
    const layout =
      val(child(properties, "tblLayout")) ??
      attr(child(properties, "tblLayout"), "type");
    if (layout) {
      if (!["fixed", "autofit"].includes(layout))
        fail(properties, "Table layout is invalid.");
      result["table-layout"] = layout === "fixed" ? "fixed" : "auto";
    }
    const align = val(child(properties, "jc"));
    if (align === "center") {
      result["margin-left"] = "auto";
      result["margin-right"] = "auto";
    } else if (align === "right") result["margin-left"] = "auto";
    else if (align && align !== "left" && align !== "start")
      unsupported(properties);
    const indent = child(properties, "tblInd");
    if (indent)
      result["margin-left"] =
        this.width(indent, "margin-left")["margin-left"] ?? "0pt";
    const shade = child(properties, "shd");
    if (shade && attr(shade, "fill") !== "auto")
      result["background-color"] = this.color(shade, "fill") ?? "transparent";
    return result;
  }
  row(properties:XmlNode|undefined):Css{
    if(!properties)return{};const shared=properties.part===this.stylePart;
    check(properties,['cnfStyle','divId','gridBefore','gridAfter','wBefore','wAfter','cantSplit','trHeight','tblHeader','tblCellSpacing','jc','hidden','trPrChange','ins','del']);
    for(const name of ['hidden','jc',...(shared?['gridBefore','gridAfter','trHeight','cnfStyle','divId','wBefore','wAfter','trPrChange','ins','del']:[])] as const){const value=child(properties,name);if(value)unsupported(value,'This row property remains native.');}
    const result:Css={...this.cellSpacing(properties)};
    const split=child(properties,'cantSplit');if(split)result['break-inside']=on(split)?'avoid':'auto';
    const header=child(properties,'tblHeader');if(header)result['--word-row-repeat-header']=on(header)?'1':'0';
    const height=child(properties,'trHeight');if(height){const rule=attr(height,'hRule')??'atLeast';if(!['auto','atLeast','exact'].includes(rule))unsupported(height,'This row height rule remains native.');result.height=pt(number(attr(height,'val'),height,0)/20);result['--word-row-height-rule']=rule;}
    return result;
  }
  cellSpacing(properties: XmlNode | undefined,table=false): Css {
    const spacing = child(properties, "tblCellSpacing");
    if(!spacing)return{};
    const type=attr(spacing,'type')??'dxa',raw=attr(spacing,'w');if(!['dxa','nil','auto','pct'].includes(type))fail(spacing,'A table spacing unit is invalid.');
    if(raw!==undefined){if(raw.endsWith('%')){if(!/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)%$/.test(raw))fail(spacing,'A table spacing percentage is invalid.');}else measurePoints(raw,spacing);}
    const amount=type==='dxa'?measurePoints(raw,spacing,20,0):0;if(amount<0)fail(spacing,'A table cell spacing cannot be negative.');
    const prefix=table?'--word-table-cell-spacing':'--word-row-cell-spacing';
    return{[prefix]:pt(amount),[prefix+'-type']:type,...(raw!==undefined?{[prefix+'-value']:raw}:{}),...(table?{'border-collapse':amount>0?'separate':'collapse','border-spacing':pt(amount)}:{})};
  }
  cell(properties: XmlNode | undefined): Css {
    if(properties?.part===this.stylePart&&child(properties,'cnfStyle'))unsupported(child(properties,'cnfStyle')!,'Shared-style conditional cache flags remain native.');
    if (!properties) return {};
    if(properties.part===this.stylePart){const change=elements(properties).find(node=>/PrChange$/.test(node.name));if(change)unsupported(change,'Revision properties in a shared style are preserved natively.');}
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
      "hideMark", "tcPrChange", "cellIns", "cellDel",
    ]);
    if (child(properties, "hMerge") || on(child(properties, "tcFitText")))
      unsupported(properties);
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
    const direction=val(child(properties,"textDirection"));
    if(direction){if(!["lrTb","tbRl","btLr"].includes(direction))unsupported(properties!);result["--word-cell-text-direction"]=direction;}
    Object.assign(result, this.margins(child(properties, "tcMar")));
    return result;
  }
  margins(properties: XmlNode | undefined): Css {
    const result: Css = {};
    for (const item of elements(properties)) {
      const name =
        item.name === "start"
          ? "left"
          : item.name === "end"
            ? "right"
            : item.name;
      if (!["top", "bottom", "left", "right"].includes(name)) unsupported(item);
      Object.assign(result, this.width(item, "padding-" + name));
    }
    return result;
  }
  width(node: XmlNode, property: string): Css {
    const type = attr(node, "type") ?? "dxa";
    const raw = attr(node, "w");
    if (type === "auto" || type === "nil") return {};
    if (type === "pct"){
      const literal=raw&&/^([+-]?(?:\d+(?:\.\d+)?|\.\d+))%$/.exec(raw),value=literal?Number(literal[1]):number(raw,node)/50;
      if(value<0&&property!=='margin-left')unsupported(node,'This negative preferred width remains native.');
      return { [property]: String(value) + "%" };
    }
    if (type !== "dxa")
      return unsupported(node, "An OOXML width unit is not supported.");
    const value=measurePoints(raw,node);if(value<0&&property!=='margin-left')unsupported(node,'This negative preferred width remains native.');return { [property]: pt(value) };
  }
  resolve(id: string, visiting: Set<string> = new Set<string>()): StyleValues {
    const cached = this.resolved.get(id);
    if (cached) return cached;
    const source = this.definitions.get(id);
    if (!source)
      throw new DocumentError(
        "InvalidDocument",
        "A referenced native style is missing.",
      );
    if (visiting.has(id)) return fail(source, "Style inheritance is cyclic.");
    visiting.add(id);
    const parentId = val(child(source, "basedOn"));
    if (parentId && !this.definitions.has(parentId))
      fail(source, "A native base style is missing.");
    const parent = parentId ? this.resolve(parentId, visiting) : empty();
    const pPr = child(source, "pPr");
    const numPr = child(pPr, "numPr");
    const num = val(child(numPr, "numId"));
    const outline = child(pPr, "outlineLvl");
    const tableProperties = child(source, "tblPr");
    this.cellSpacing(child(source, "trPr"));
    const result: StyleValues = {
      p: { ...parent.p, ...this.paragraph(pPr) },
      r: { ...parent.r, ...this.run(child(source, "rPr"), parent.r, attr(source,"type")!=="table") },
      table: { ...parent.table, ...this.table(tableProperties) },
      cell: {
        ...parent.cell,
        ...this.margins(child(tableProperties, "tblCellMar")),
        ...this.cell(child(source, "tcPr")),
      },
      ...(parent.numbering ? { numbering: parent.numbering } : {}),
      tabs: mergeTabs(parent.tabs ?? this.defaults.tabs, child(pPr, "tabs")),
      ...(parent.outline !== undefined ? { outline: parent.outline } : {}),
      ...(num
        ? {
            numbering: {
              id: num,
              level: number(val(child(numPr, "ilvl")), numPr!, 0),
            },
          }
        : {}),
      ...(outline ? { outline: number(val(outline), outline) } : {}),
    };
    visiting.delete(id);
    this.resolved.set(id, result);
    return result;
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
        throw new DocumentError(
          "InvalidDocument",
          "A referenced native style is missing.",
        );
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
  numberingParagraph(id: string, level: number): Css {
    return this.numberProperties.get(id + ":" + level) ?? {};
  }
  private readNumbering(root: XmlNode): void {
    check(root, ["numPicBullet", "abstractNum", "num", "numIdMacAtCleanup"]);
    if (elements(root, "numPicBullet").length)
      unsupported(root, "Picture bullets are not supported.");
    const abstract = new Map<string, XmlNode>();
    for (const definition of elements(root, "abstractNum")) {
      const id = attr(definition, "abstractNumId");
      if (!id || abstract.has(id))
        fail(definition, "An abstract numbering identity is invalid.");
      check(definition, [
        "nsid",
        "multiLevelType",
        "tmpl",
        "name",
        "styleLink",
        "numStyleLink",
        "lvl",
      ]);
      if (child(definition, "numStyleLink"))
        unsupported(definition, "Numbering style links are not supported.");
      abstract.set(id, definition);
    }
    const seen = new Set<string>();
    for (const definition of elements(root, "num")) {
      const id = attr(definition, "numId");
      const ref = val(child(definition, "abstractNumId"));
      if (!id || seen.has(id) || !ref || !abstract.has(ref))
        fail(definition, "A numbering instance is invalid.");
      seen.add(id);
      const levels: WordNumberLevel[] = [];
      for (const source of elements(abstract.get(ref), "lvl")) {
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
        if (
          child(source, "lvlPicBulletId") ||
          child(source, "legacy") ||
          on(child(source, "isLgl"))
        )
          unsupported(source);
        const level = number(attr(source, "ilvl"), source);
        const override = elements(definition, "lvlOverride").find(
          (item) => number(attr(item, "ilvl"), item) === level,
        );
        if (child(override, "lvl"))
          unsupported(
            override!,
            "Full numbering-level overrides are not supported.",
          );
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
        const paragraph = this.paragraph(child(source, "pPr"));
        this.numberProperties.set(id + ":" + level, paragraph);
        this.css.push(
          "." +
            cssClass +
            "{" +
            declarations({ ...paragraph, ...this.run(child(source, "rPr")) }) +
            "}",
        );
        const restart = child(source, "lvlRestart");
        levels.push({
          level,
          start: number(
            val(child(override, "startOverride")) ??
              val(child(source, "start")),
            source,
            1,
          ),
          format: format as WordNumberLevel["format"],
          text: val(child(source, "lvlText")) ?? "%" + (level + 1) + ".",
          suffix: (val(child(source, "suff")) ??
            "tab") as WordNumberLevel["suffix"],
          cssClass,
          ...(restart ? { restart: number(val(restart), restart) } : {}),
        });
      }
      this.numbering.push({ numberingId: id, levels });
    }
  }
}
