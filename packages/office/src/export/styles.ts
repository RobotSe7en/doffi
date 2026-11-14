import { DocumentError } from "@doffi/core";
import type { DeepReadonly, DocumentReadLimits, WordMetadata } from "@doffi/core";
import { variables } from "./css.js";
import type { Css, ExportCss } from "./css.js";
import { DocxStyles } from "../ooxml/styles.js";
import { xml } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { cellProperties, paragraphKeys, paragraphProperties, pick, runKeys, runProperties, tableProperties } from "./properties.js";
import { orderedProperties, attributes, declaration, empty, escapeXml, namespaces, value } from "./xml.js";
import { contentType } from "./opc.js";
import type { OutputOpc } from "./opc.js";

type Style = DeepReadonly<WordMetadata["styles"][number]>;
export interface HtmlNumbering { readonly id: string; readonly start: number; readonly level: number; readonly format: string; readonly text: string }
export class ExportStyles {
  readonly defaults: Css;
  readonly paragraphDefaults: Css;
  readonly ids: ReadonlyMap<string, string>;
  readonly numberingIds: Map<string, string> = new Map<string, string>();
  private readonly definitions: Map<string, Style>;
  private readonly native = new Map<string, { p: Css; r: Css }>();
  private readonly contexts = new Map<string, Css>();
  private contextWork = 0;
  private readonly properties = new Map<string, { source: string; node: XmlNode }>();
  private readonly propertyBudget = { nodes: 0 };
  private readonly reader: DocxStyles;
  constructor(readonly css: ExportCss, readonly word: DeepReadonly<WordMetadata> | undefined, private readonly locale: string | undefined, private readonly limits: DocumentReadLimits, main: Css) {
    this.definitions = new Map(word?.styles.map((style) => [style.styleId, style]) ?? []);
    this.ids = new Map([...this.definitions.keys()].map((id) => [id, id]));
    const special = runKeys.filter((key) => key === "background-color" || key.startsWith("text-decoration"));
    this.defaults = { ...pick(main, runKeys.filter((key) => !special.includes(key) && (word || key !== "--word-run-shading"))), ...(word?.defaultRunCssClass ? pick(css.classStyle(word.defaultRunCssClass, main), special) : pick(css.emptyStyle(), special)) };
    // Paragraph source nodes carry resolved defaults; a native default of zero
    // spacing is explicit so producers do not substitute their UI defaults.
    this.paragraphDefaults = { "margin-top": "0pt", "margin-bottom": "0pt" };
    // Replay only our generated, theme-free run properties through the shared
    // Word semantics. This reader has no external parts or I/O capability.
    this.reader = new DocxStyles({ files: new Map(), limits, xmlBudget: { nodes: 0 }, related: () => undefined, relationships: () => new Map(), contentType: () => undefined,
      read: () => { throw new DocumentError("ExportFailed", "Native style baseline cannot read an external part."); } }, "word/document.xml");
    word?.numbering.forEach((definition, index) => this.numberingIds.set(definition.numberingId, String(index + 1)));
  }
  style(id: string | undefined): Style | undefined { return id ? this.definitions.get(id) : undefined; }
  private paragraphStyle(id: string | undefined): Style | undefined { return this.style(id) ?? (!id ? [...this.definitions.values()].find((style) => style.default??(style.kind === "paragraph" && style.styleId === "Normal"&&!this.word?.styles.some(item=>item.kind===style.kind&&item.default))) : undefined); }
  paragraphContext(id: string | undefined): Css { const style = this.paragraphStyle(id); return { ...this.css.emptyStyle(), ...(style ? this.nativeValues(style).p : {}) }; }
  runContext(paragraph: string | undefined, character?: string): Css {
    const key = JSON.stringify([paragraph ?? null, character ?? null]), known = this.contexts.get(key); if (known) return known;
    const paragraphStyle = this.paragraphStyle(paragraph);
    let result = paragraphStyle ? this.nativeValues(paragraphStyle).r : this.defaults;
    const chain: Style[] = []; let current = this.style(character);
    while (current) { if (chain.length >= this.limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "Native style inheritance exceeds its depth budget."); chain.push(current); current = this.style(current.basedOn); }
    this.contextWork += chain.length + 1;
    if (this.contextWork > this.limits.maxNodes * 128) throw new DocumentError("ResourceLimitExceeded", "Native style mapping exceeds its work budget.");
    for (const style of chain.reverse()) result = this.applyRun(style, result);
    this.contexts.set(key, result);
    return result;
  }
  private runSource(style: Style): { source: string; node: XmlNode } {
    const known = this.properties.get(style.styleId); if (known) return known;
    const css = this.values(style), parent = style.basedOn ? this.values(this.style(style.basedOn)!) : this.defaults;
    const source = runProperties(this.own(style, "r", css, parent), style.components?.r?{}:parent, true);
    const encoder = new (globalThis as unknown as { TextEncoder: new () => { encode(value: string): Uint8Array } }).TextEncoder();
    const node = xml(encoder.encode(declaration + "<w:rPr" + namespaces + ">" + source + "</w:rPr>"), "generated-style-properties", this.limits, this.propertyBudget);
    const result = { source, node }; this.properties.set(style.styleId, result); return result;
  }
  private applyRun(style: Style, inherited: Css): Css {
    const result = { ...inherited, ...this.reader.run(this.runSource(style).node, { ...inherited }, true) };
    for (const [key, raw] of Object.entries(result)) if (!key.startsWith("--") && raw.includes("var(")) result[key] = variables(raw, (name) => result[name], 0, this.limits.maxFileBytes) ?? inherited[key]!;
    return result;
  }
  private component(name:string,parent:Css=this.defaults):Css{return pick(this.css.classStyle(name,parent),this.css.classKeys(name));}
  private values(style: Style): Css {
    if(style.components){const p=style.components.p?this.component(style.components.p):{},r=style.components.r?this.component(style.components.r):{};return{...p,...r};}
    const result = this.css.classStyle(style.cssClass, this.defaults);
    return style.runCssClass ? { ...result, ...pick(this.css.classStyle(style.runCssClass, result), ["background-color", "text-decoration-line", "text-decoration-style", "text-decoration-color"]) } : result;
  }
  private own(style: Style, kind: "p" | "r", css: Css, parent: Css): Css {
    if(style.components?.[kind])return this.component(style.components[kind]!);
    const original = kind === "p" ? style.paragraphProperties : style.runProperties;
    const allowed = kind === "p" ? paragraphKeys : runKeys;
    const declared = new Set([...this.css.classKeys(style.cssClass), ...(style.runCssClass && kind === "r" ? this.css.classKeys(style.runCssClass) : [])]);
    const owns = (key: string): boolean => original === undefined || original.some((name) => name === key || key.startsWith(name + "-") && ["border", "margin", "padding"].some((prefix) => name.startsWith(prefix)));
    return pick(css, allowed.filter((key) => owns(key) || declared.has(key) && css[key] !== parent[key] && !(kind === "r" && style.kind !== "character" && key === "background-color" && !style.runCssClass)));
  }
  private nativeValues(style: Style, depth = 0): { p: Css; r: Css } {
    const known = this.native.get(style.styleId); if (known) return known;
    if (depth > this.limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "Native style inheritance exceeds its depth budget.");
    const css = this.values(style), parentStyle = style.basedOn ? this.style(style.basedOn) : undefined;
    const parent = parentStyle ? this.nativeValues(parentStyle, depth + 1) : { p: {}, r: this.defaults };
    const parentCss = parentStyle ? this.values(parentStyle) : this.defaults;
    const result = { p: { ...parent.p, ...this.own(style, "p", css, parentCss) }, r: this.applyRun(style, parent.r) };
    this.native.set(style.styleId, result); return result;
  }
  write(opc: OutputOpc): void {
    const out = opc.part("word/styles.xml", contentType("styles"));
    out.add(declaration + "<w:styles" + namespaces + ">");
    out.add("<w:docDefaults><w:rPrDefault><w:rPr>" + runProperties({...pick(this.defaults, runKeys.filter((key) => key !== "vertical-align" || this.defaults[key] !== "baseline")),...(this.defaults["--word-language-val"]===undefined&&this.locale?{"--word-language-val":JSON.stringify(this.locale)}:{})}) + "</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>" + paragraphProperties(this.paragraphDefaults) + "</w:pPr></w:pPrDefault></w:docDefaults>");
    for (const style of this.definitions.values()) {
      const css = this.values(style), parent = style.basedOn ? this.values(this.definitions.get(style.basedOn)!) : this.defaults;
      out.add("<w:style" + attributes({ "w:type": style.kind, "w:styleId": this.ids.get(style.styleId)!,"w:customStyle":style.customStyle===undefined?undefined:Number(style.customStyle), ...(style.default??(style.kind === "paragraph" && style.styleId === "Normal"&&!this.word?.styles.some(item=>item.kind===style.kind&&item.default)) ? { "w:default": 1 } : {}) }) + ">" + value("name", style.name??style.styleId)+(style.aliases!==undefined?value("aliases",style.aliases):""));
      for (const [key, name] of [["basedOn", "basedOn"], ["nextStyleId", "next"], ["linkedStyleId", "link"]] as const) if (style[key]) out.add(value(name, this.ids.get(style[key])));
      for(const key of ['hidden','uiPriority','semiHidden','unhideWhenUsed','qFormat','locked']as const)if(style[key]!==undefined)out.add(value(key,typeof style[key]==='boolean'?Number(style[key]):style[key]));
      out.add("<w:pPr>" + orderedProperties("pPr",paragraphProperties({ ...this.own(style, "p", css, parent), "font-size": css["font-size"]! }) +(style.outlineLevel!==undefined?value("outlineLvl",style.outlineLevel):""))+ "</w:pPr>");
      out.add("<w:rPr>" + this.runSource(style).source + "</w:rPr>");
      if (style.kind === "table") {
        const tableCss=style.components?.table?this.component(style.components.table):css,cellCss=style.components?.cell?this.component(style.components.cell):pick(css,Object.keys(css).filter(key=>key.startsWith('padding-')));
        out.add("<w:tblPr>" + (style.rowBandSize !== undefined ? value("tblStyleRowBandSize", style.rowBandSize) : "") + (style.columnBandSize !== undefined ? value("tblStyleColBandSize", style.columnBandSize) : "") + tableProperties(tableCss) + (style.tableLookMask !== undefined ? empty("w:tblLook", { "w:val": style.tableLookMask.toString(16).padStart(4, "0") }) : "") + "</w:tblPr>");
        out.add("<w:tcPr>" + cellProperties(cellCss) + "</w:tcPr>");
        for (const condition of style.conditions ?? []) {
          const conditional = pick(this.css.classStyle(condition.cssClass, this.defaults), this.css.classKeys(condition.cssClass));
          const component=(key:'p'|'r'|'table'|'cell')=>condition.components?.[key]?this.component(condition.components[key]!):conditional;
          out.add('<w:tblStylePr w:type="' + escapeXml(condition.region) + '"><w:pPr>' + paragraphProperties(pick(component('p'),paragraphKeys)) + "</w:pPr><w:rPr>" + runProperties(pick(component('r'),runKeys)) + "</w:rPr><w:tblPr>" + tableProperties(component('table')) + "</w:tblPr><w:tcPr>" + cellProperties(component('cell')) + "</w:tcPr></w:tblStylePr>");
        }
      }
      out.add("</w:style>");
    }
    out.add("</w:styles>"); out.finish(); opc.relate("word/document.xml", "styles", "styles.xml");
  }
  writeNumbering(opc: OutputOpc, html: readonly HtmlNumbering[] = []): void {
    if (!this.word?.numbering.length && !html.length) return;
    const out = opc.part("word/numbering.xml", contentType("numbering")); out.add(declaration + "<w:numbering" + namespaces + ">");
    for (const definition of this.word?.numbering ?? []) {
      const id = this.numberingIds.get(definition.numberingId)!;
      out.add('<w:abstractNum w:abstractNumId="' + id + '">' + value("multiLevelType", "multilevel"));
      for (const level of definition.levels) {
        const css = pick(this.css.classStyle(level.cssClass, this.defaults), this.css.classKeys(level.cssClass));
        out.add('<w:lvl w:ilvl="' + level.level + '">' + value("start", level.start) + value("numFmt", level.format) + (level.restart !== undefined ? value("lvlRestart", level.restart) : "") + value("suff", level.suffix) + value("lvlText", level.text) + '<w:pPr>' + paragraphProperties(pick(css, paragraphKeys.filter((key) => !["background-color"].includes(key)))) + '</w:pPr><w:rPr>' + runProperties(pick(css, runKeys)) + "</w:rPr></w:lvl>");
      }
      out.add("</w:abstractNum>");
    }
    for (const item of html) out.add('<w:abstractNum w:abstractNumId="' + item.id + '"><w:lvl w:ilvl="' + item.level + '">' + value("start", item.start) + value("numFmt", item.format) + value("lvlText", item.text) + value("suff", "tab") + '<w:pPr><w:tabs><w:tab w:val="left" w:pos="' + ((item.level + 1) * 720) + '"/></w:tabs><w:ind w:left="' + ((item.level + 1) * 720) + '" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>');
    for (const definition of this.word?.numbering ?? []) { const id = this.numberingIds.get(definition.numberingId)!; out.add('<w:num w:numId="' + id + '">' + value("abstractNumId", id) + "</w:num>"); }
    for (const item of html) out.add('<w:num w:numId="' + item.id + '">' + value("abstractNumId", item.id) + "</w:num>");
    out.add("</w:numbering>"); out.finish(); opc.relate("word/document.xml", "numbering", "numbering.xml");
  }
}
