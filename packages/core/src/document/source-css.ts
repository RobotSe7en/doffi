import { generate, ident, lexer, List, parse, walk } from "css-tree";
import type { CssNode, Declaration, Selector } from "css-tree";
import { DocumentError } from "../errors.js";
import type { DocumentReadLimits } from "../runtime.js";
import { normalizeCss } from "./css.js";
import type { CssBudget } from "./css.js";
import { color, length } from "./css-values.js";
import { checkLimit, readLimits } from "./validation.js";
import { encode } from "./package.js";

/** A source snapshot projection. Keep nodes and their ancestry immutable for
 * the lifetime of one resolver; editor DOM and layout substitutions are excluded. */
export interface SourceStyleNode { readonly tag: string; readonly attrs: Readonly<Record<string, string>>; readonly parent: SourceStyleNode | undefined }
export interface SourceStyleResolverOptions {
  readonly initial?: Readonly<Record<string, string>>;
  /** Native run export needs effective ancestor paints. Source formatting uses
   * false to inspect each element's own paint before normalizing selected text. */
  readonly combineDecorations?: boolean;
}

export type SourceStyleMap = Readonly<Record<string, string>>;
type Css = SourceStyleMap;
interface Compound { tag?: string; classes: string[]; attrs: { name: string; value?: string }[]; root: boolean }
type Specificity = readonly [number, number, number];
interface Rule { compounds: Compound[]; declarations: Declaration[]; specificity: Specificity; order: number; key: string; revisionDisplay: boolean }
interface Entry { value: string; important: boolean; specificity: Specificity; order: number }
interface DeclarationEntry extends Entry { fallback: boolean; property: string; sourceValue: string }
type CascadeRule = Pick<Rule,"declarations"|"specificity"|"order"> & {fallback?:boolean};
const inherited = new Set(["color", "font-family", "font-size", "font-weight", "font-style", "font-variant", "text-transform", "text-align", "text-indent", "line-height", "letter-spacing", "word-spacing", "white-space", "widows", "orphans", "tab-size", "overflow-wrap", "list-style-type", "list-style-position"]);
const inline = new Set(["span", "a", "strong", "b", "em", "i", "u", "s", "sub", "sup", "ins", "del", "code"]);
const initial: Css = { "margin-top": "0pt", "margin-right": "0pt", "margin-bottom": "0pt", "margin-left": "0pt", color: "#000000", "font-family": '"Noto Sans CJK SC",serif', "font-size": "12pt", "font-weight": "normal", "font-style": "normal", "font-variant": "normal", "text-transform": "none", "text-align": "left", "text-indent": "0pt", "line-height": "normal", "letter-spacing": "normal", "word-spacing": "normal", "white-space": "normal", "vertical-align": "baseline", "background-color": "transparent", "text-decoration-line": "none", "text-decoration-style": "solid", "text-decoration-color": "currentColor", widows: "2", orphans: "2", "break-before": "auto", "break-after": "auto", "break-inside": "auto" };
function valueString(nodes: CssNode[]): string { return generate({ type: "Value", children: new List<CssNode>().fromArray(nodes) }); }
/** Resolve CSS variables as values, never by executing CSS or JavaScript. */
export function resolveSourceCssVariables(value: string, get: (name: string) => string | undefined, depth = 0, maximum: number = 8 * 1024 * 1024): string | undefined {
  if (depth > 128) throw new DocumentError("ResourceLimitExceeded", "CSS variable nesting exceeds the Office budget.");
  let size = value.length;
  if (size > maximum) throw new DocumentError("ResourceLimitExceeded", "Expanded CSS exceeds the Office value budget.");
  const ast = parse(value, { context: "value", parseCustomProperty: true }); let valid = true;
  walk(ast, function (node, item, list) {
    if (node.type !== "Function" || ident.decode(node.name).toLowerCase() !== "var") return;
    const nodes = node.children.toArray(), name = nodes[0];
    if (name?.type !== "Identifier") { valid = false; return this.skip; }
    let replacement = get(ident.decode(name.name));
    if (replacement === undefined) {
      const comma = nodes.findIndex((child) => child.type === "Operator" && child.value === ",");
      if (comma >= 0) replacement = resolveSourceCssVariables(valueString(nodes.slice(comma + 1)), get, depth + 1, maximum);
    }
    if (replacement === undefined) valid = false;
    else {
      size += replacement.length;
      if (size > maximum) throw new DocumentError("ResourceLimitExceeded", "Expanded CSS exceeds the Office value budget.");
      const parsed = parse(replacement, { context: "value", parseCustomProperty: true });
      if (parsed.type !== "Value" || !item || !list) valid = false; else list.replace(item, parsed.children);
    }
    return this.skip;
  });
  return valid ? generate(ast) : undefined;
}
function selector(source: Selector): { compounds: Compound[]; specificity: Specificity; key: string } {
  const compounds: Compound[] = [{ classes: [], attrs: [], root: false }]; const specificity: [number, number, number] = [0, 0, 0];
  for (const part of source.children) {
    let compound = compounds.at(-1)!;
    if (part.type === "Combinator") { compounds.push({ classes: [], attrs: [], root: false }); continue; }
    if (part.type === "TypeSelector") { compound.tag = ident.decode(part.name).toLowerCase(); specificity[2]++; }
    else if (part.type === "ClassSelector") { compound.classes.push(ident.decode(part.name)); specificity[1]++; }
    else if (part.type === "AttributeSelector") {
      compound.attrs.push({ name: ident.decode(part.name.name), ...(part.value ? { value: part.value.type === "String" ? part.value.value : ident.decode(part.value.name) } : {}) }); specificity[1]++;
    } else if (part.type === "PseudoClassSelector") { compound.root = true; specificity[1]++; }
    else throw new DocumentError("UnsupportedFeature", "Office export encountered an unsupported selector.");
  }
  const last = compounds.at(-1)!, key = last.classes[0] ? "." + last.classes[0] : last.attrs[0] ? "[" + last.attrs[0].name : last.tag ?? (last.root ? ":root" : "*");
  return { compounds, specificity, key };
}
function tokens(value: string): string[] {
  const ast = parse(value, { context: "value", parseCustomProperty: true }); return ast.type === "Value" ? ast.children.toArray().map((node) => generate(node)) : [];
}
export function expandSourceCssProperty(property: string, value: string): [string, string][] {
  if (["inherit", "initial", "unset", "revert"].includes(value)
    && (/^(margin|padding|border(?:-(?:width|style|color|top|right|bottom|left))?)$/.test(property) || property === "text-decoration"))
    return expandSourceCssProperty(property, property === "margin" || property === "padding" || /^border-(width|style|color)$/.test(property) ? "0" : "none").map(([name]) => [name, value]);
  const parts = tokens(value), box = /^(margin|padding|border-(?:width|style|color))$/.exec(property);
  if (box && parts.length >= 1 && parts.length <= 4) {
    const [a, b = a, c = a, d = b] = parts;
    return ["top", "right", "bottom", "left"].map((side, index) => [property.startsWith("border-") ? "border-" + side + property.slice(6) : property + "-" + side, [a, b, c, d][index]!] as [string, string]);
  }
  if (property === "border" || /^border-(top|right|bottom|left)$/.test(property)) {
    let width = "medium", style = "none", color = "currentColor";
    for (const part of parts) {
      if (/^(none|hidden|solid|double|dotted|dashed|groove|ridge|inset|outset)$/.test(part)) style = part;
      else if (/^(thin|medium|thick|[-+.\d])/.test(part)) width = part;
      else color = part;
    }
    const sides = property === "border" ? ["top", "right", "bottom", "left"] : [property.slice(7)];
    return sides.flatMap((side) => [["border-" + side + "-width", width], ["border-" + side + "-style", style], ["border-" + side + "-color", color]] as [string, string][]);
  }
  if (property === "text-decoration") {
    const lines: string[] = []; let style = "solid", color = "currentColor";
    for (const part of parts) {
      if (["none", "underline", "overline", "line-through"].includes(part)) lines.push(part);
      else if (["solid", "double", "dotted", "dashed", "wavy"].includes(part)) style = part;
      else color = part;
    }
    return [["text-decoration-line", lines.join(" ") || "none"], ["text-decoration-style", style], ["text-decoration-color", color]];
  }
  return [[property, value]];
}
function higher(next: Entry, previous: Entry | undefined): boolean {
  if (!previous) return true;
  if (next.important !== previous.important) return next.important;
  for (const index of [0, 1, 2] as const) if (next.specificity[index] !== previous.specificity[index]) return next.specificity[index] > previous.specificity[index];
  return next.order >= previous.order;
}
/** Portable cascade for the validated document CSS subset. This resolves
 * authored values; it does not measure layout or substitute fonts. */
export class SourceStyleResolver {
  private readonly indexed = new Map<string, Rule[]>();
  private readonly cache = new WeakMap<SourceStyleNode, Css>();
  private readonly specifiedCache = new WeakMap<SourceStyleNode, Css>();
  private readonly declarationCache = new WeakMap<SourceStyleNode, ReadonlyMap<string, DeclarationEntry>>();
  private readonly entriesCache = new WeakMap<SourceStyleNode, readonly CascadeRule[]>();
  private work = 0;
  private rootFont = 12;
  private readonly initial: Css;
  private readonly resolvingNodes = new Set<SourceStyleNode>();
  private readonly limits: DocumentReadLimits;
  private readonly revisionIds: ReadonlySet<string>;
  private readonly combineDecorations: boolean;
  private readonly cssBudget: CssBudget = { nodes: 0 };
  private sourceBytes = 0;
  constructor(styles: readonly string[], limits: DocumentReadLimits, revisionIds: ReadonlySet<string> = new Set(), options: SourceStyleResolverOptions = {}) {
    this.limits = readLimits(limits);
    this.revisionIds = new Set(revisionIds);
    this.combineDecorations = options.combineDecorations !== false;
    this.initial = Object.freeze({ ...initial, ...options.initial });
    checkLimit(styles.length, this.limits.maxFiles);
    let order = 0;
    for (const [index, source] of ["p,h1,h2,h3,h4,h5,h6{margin:0;font-weight:inherit;line-height:normal}td,th{padding:0;font-weight:inherit;text-align:inherit;vertical-align:top}", ...styles].entries()) {
      const ast = parse(index === 0 ? source : this.validated(source, "stylesheet"), { context: "stylesheet", parseCustomProperty: true });
      if (ast.type !== "StyleSheet") continue;
      for (const item of ast.children) {
        if (item.type !== "Rule" || item.prelude?.type !== "SelectorList") throw new DocumentError("UnsupportedFeature", "Only validated source styles can be compiled to Office.");
        const declarations = item.block.children.toArray().filter((node): node is Declaration => node.type === "Declaration");
        for (const choice of item.prelude.children) {
          if (choice.type !== "Selector") continue;
          const compiled = selector(choice), description = generate(choice);
          const revisionDisplay = [".word-main ins", ".word-main del"].includes(description) && declarations.every((decl) => {
            const value = generate(decl.value).toLowerCase();
            return decl.property === "color" && ["#006400", "#a00000"].includes(value) || decl.property === "text-decoration" && ["underline", "line-through"].includes(value);
          });
          const rule = { ...compiled, declarations, order: order++, revisionDisplay };
          const list = this.indexed.get(rule.key) ?? []; list.push(rule); this.indexed.set(rule.key, list);
        }
      }
    }
  }
  private validated(source: string, context: "stylesheet" | "declarationList"): string {
    checkLimit(source.length, this.limits.maxFileBytes);
    const bytes = encode(source).byteLength;
    checkLimit(bytes, this.limits.maxFileBytes);
    this.sourceBytes += bytes;
    checkLimit(this.sourceBytes, this.limits.maxPackageBytes);
    return normalizeCss(source, context, this.limits, this.cssBudget);
  }
  private step(): void {
    if (++this.work > this.limits.maxCssNodes + this.limits.maxNodes * 128) throw new DocumentError("ResourceLimitExceeded", "Office style resolution exceeds its work budget.");
  }
  private matches(node: SourceStyleNode, rule: Rule): boolean {
    const compound = (target: SourceStyleNode, part: Compound): boolean => {
      this.step(); const classes = new Set((target.attrs.class ?? "").split(/\s+/));
      return (!part.tag || part.tag === target.tag) && (!part.root || target.tag === "html") && part.classes.every((name) => classes.has(name))
        && part.attrs.every((attr) => target.attrs[attr.name] !== undefined && (attr.value === undefined || target.attrs[attr.name] === attr.value));
    };
    let current: SourceStyleNode | undefined = node;
    for (let index = rule.compounds.length - 1; index >= 0; index--) {
      if (index === rule.compounds.length - 1) { if (!compound(current!, rule.compounds[index]!)) return false; }
      else { let depth = 0; while (current && !compound(current, rule.compounds[index]!)) { if (++depth > this.limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "Source selector ancestry exceeds its depth budget."); current = current.parent; } if (!current) return false; }
      current = current?.parent;
    }
    return true;
  }
  private entries(node:SourceStyleNode):CascadeRule[] {
    const known = this.entriesCache.get(node); if (known) return [...known];
    const selected = new Set<Rule>();
    for (const key of ["*", node.tag, ...(node.tag === "html" ? [":root"] : []), ...(node.attrs.class ?? "").split(/\s+/).filter(Boolean).map((name) => "." + name), ...Object.keys(node.attrs).map((name) => "[" + name)])
      for (const rule of this.indexed.get(key) ?? []) if (this.matches(node, rule) && !(rule.revisionDisplay && this.revisionIds.has(node.attrs["data-node-id"] ?? ""))) selected.add(rule);
    const entries:CascadeRule[]=[...selected];
    if (node.attrs.style) {
      const ast = parse(this.validated(node.attrs.style, "declarationList"), { context: "declarationList", parseCustomProperty: true });
      if (ast.type === "DeclarationList") entries.push({ declarations: ast.children.toArray().filter((item): item is Declaration => item.type === "Declaration"), specificity: [1, 0, 0], order: Number.MAX_SAFE_INTEGER });
    }
    this.entriesCache.set(node, entries); return [...entries];
  }
  get(node: SourceStyleNode): Css {
    const known=this.cache.get(node);if(known)return known;
    this.step();
    if (this.resolvingNodes.has(node)) throw new DocumentError("InvalidOperation", "Source style ancestry contains a cycle.");
    if (this.resolvingNodes.size >= this.limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "Source style ancestry exceeds its depth budget.");
    this.resolvingNodes.add(node);
    try {
      const parent=node.parent?this.get(node.parent):this.initial,entries=this.entries(node);
      const result=this.compute(parent,entries,node);this.cache.set(node,result);return result;
    } finally { this.resolvingNodes.delete(node); }
  }
  /** Match the preview's native inheritance fallback: a late type selector;
   * authored classes, attribute selectors, inline styles and priority still win. */
  withFallback(node:SourceStyleNode,defaults:Css,initialDefaults:Css={},nativeDefaults:Css={},options:Pick<SourceStyleResolverOptions,"combineDecorations">={}):Css {
    const parent = node.parent ? this.get(node.parent) : this.initial;
    const ast=parse(Object.entries(defaults).map(([name,value])=>name+":"+value).join(";"),{context:"declarationList",parseCustomProperty:true});
    if(ast.type!=="DeclarationList")throw new DocumentError("ExportFailed","Native style defaults are invalid.");
    const entries=this.entries(node),baseAst=parse(Object.entries(initialDefaults).map(([name,value])=>name+":"+value).join(";"),{context:"declarationList",parseCustomProperty:true});
    if(baseAst.type==="DeclarationList")entries.unshift({declarations:baseAst.children.toArray().filter((n):n is Declaration=>n.type==="Declaration"),specificity:[0,0,0],order:-1,fallback:true});
    const nativeAst=parse(Object.entries(nativeDefaults).map(([name,value])=>name+":"+value).join(";"),{context:"declarationList",parseCustomProperty:true});
    if(nativeAst.type==="DeclarationList")entries.unshift({declarations:nativeAst.children.toArray().filter((n):n is Declaration=>n.type==="Declaration"),specificity:[0,0,1],order:-2,fallback:true});
    entries.push({declarations:ast.children.toArray().filter((n):n is Declaration=>n.type==="Declaration"),specificity:[0,0,1],order:Number.MAX_SAFE_INTEGER-1,fallback:true});
    const declared = this.declarations(entries); this.declarationCache.set(node, declared); this.specifiedCache.set(node, Object.freeze(Object.fromEntries([...declared].filter(([, entry]) => !entry.fallback).map(([name, entry]) => [name, entry.value]))));
    const result=this.compute(parent,entries,node,options.combineDecorations);this.cache.set(node,result);return result;
  }
  private classRules(name:string):Rule[] {
    return (this.indexed.get("." + name) ?? []).filter((rule) => rule.compounds.length === 1 && rule.compounds[0]!.classes.length > 0 && rule.compounds[0]!.classes.every((value) => value === name)
      && !rule.compounds[0]!.tag && !rule.compounds[0]!.root && rule.compounds[0]!.attrs.length === 0);
  }
  private declarations(rules: readonly CascadeRule[]): ReadonlyMap<string, DeclarationEntry> {
    const winners = new Map<string, DeclarationEntry>(); let serial = 0;
    for (const rule of [...rules].sort((a, b) => a.order - b.order)) for (const declaration of rule.declarations) {
      this.step(); const decoded = ident.decode(declaration.property), property = decoded.startsWith("--") ? decoded : decoded.toLowerCase(), raw = generate(declaration.value);
      // A variable shorthand has not been substituted yet. Keep its origin for
      // editing/no-op checks instead of pretending that var() is only a color.
      const keys = property.startsWith("--") ? [[property, raw]] : /\bvar\(/i.test(raw) && expandSourceCssProperty(property, "initial").some(([key]) => key !== property)
        ? expandSourceCssProperty(property, "initial").map(([key]) => [key, raw]) : expandSourceCssProperty(property, raw);
      for (const [name, value] of keys) {
        const entry = { value: value!, property, sourceValue: raw, important: Boolean(declaration.important), specificity: rule.specificity, order: serial++, fallback: Boolean(rule.fallback) };
        if (higher(entry, winners.get(name!))) winners.set(name!, entry);
      }
    }
    return winners;
  }
  private rawDeclarations(rules: readonly CascadeRule[]): Css {
    return Object.freeze(Object.fromEntries([...this.declarations(rules)].filter(([, entry]) => !entry.fallback).map(([name, entry]) => [name, entry.value])));
  }
  /** Winning authored declarations, retaining native role references before var(). */
  specified(node: SourceStyleNode): Css {
    let result = this.specifiedCache.get(node);
    if (!result) { const entries = this.declarations(this.entries(node)); this.declarationCache.set(node, entries); result = Object.freeze(Object.fromEntries([...entries].filter(([, entry]) => !entry.fallback).map(([name, entry]) => [name, entry.value]))); this.specifiedCache.set(node, result); }
    return result;
  }
  /** Compare source declarations without resolving away variable provenance.
   * A later longhand prevents an opaque shorthand from being treated as equal. */
  hasSpecified(node: SourceStyleNode, properties: Css): boolean {
    this.specified(node); const entries = this.declarationCache.get(node)!;
    let bytes = 0; const budget: CssBudget = { nodes: 0 }, expectedProperties: [string, string, boolean][] = [];
    for (const [name, value] of Object.entries(properties)) {
      const source = ident.encode(name) + ":" + value;
      checkLimit(source.length, this.limits.maxFileBytes); bytes += encode(source).byteLength; checkLimit(bytes, this.limits.maxFileBytes);
      const ast = parse(normalizeCss(source, "declarationList", this.limits, budget), { context: "declarationList", parseCustomProperty: true });
      if (ast.type !== "DeclarationList" || ast.children.size !== 1 || ast.children.first?.type !== "Declaration") throw new DocumentError("InvalidOperation", "A source declaration query requires one value per property.");
      const declaration = ast.children.first, decoded = ident.decode(declaration.property), property = decoded.startsWith("--") ? decoded : decoded.toLowerCase();
      expectedProperties.push([property, generate(declaration.value), Boolean(declaration.important)]);
    }
    return expectedProperties.every(([property, value, important]) => {
      const opaque = /\bvar\(/i.test(value) && expandSourceCssProperty(property, "initial").some(([key]) => key !== property);
      return (opaque ? expandSourceCssProperty(property, "initial") : expandSourceCssProperty(property, value)).every(([name, expected]) => {
        const entry = entries.get(name); return !!entry && !entry.fallback && (!important || entry.important) && (opaque ? entry.property === property && entry.sourceValue === value : entry.value === expected);
      });
    });
  }
  specifiedClass(name:string):Css {return this.rawDeclarations(this.classRules(name));}
  classStyle(name: string, parent: Css = this.initial): Css {return this.compute(parent,this.classRules(name));}
  emptyStyle(): Css { return this.compute(this.initial, []); }
  classKeys(name: string): readonly string[] {
    const keys = new Set<string>();
    for (const rule of this.indexed.get("." + name) ?? []) if (rule.compounds.length === 1 && rule.compounds[0]!.classes.every((value) => value === name))
      for (const declaration of rule.declarations) for (const [key] of expandSourceCssProperty(declaration.property, generate(declaration.value))) keys.add(key);
    return [...keys];
  }
  private compute(parent: Css, rules: readonly { declarations: Declaration[]; specificity: Specificity; order: number }[], node?: SourceStyleNode, combineDecorations = this.combineDecorations): Css {
    const result: Record<string, string> = { ...this.initial }, custom = new Map<string, Entry>(), winners = new Map<string, Entry>();
    for (const [name, value] of Object.entries(parent)) if (name.startsWith("--") || inherited.has(name)) result[name] = value;
    if (node?.parent && inline.has(node.parent.tag)) result["background-color"] = parent["background-color"]?.toLowerCase() === "currentcolor" ? parent.color! : parent["background-color"] ?? this.initial["background-color"]!;
    if (node) {
      if (node.tag === "ul") result["list-style-type"] = "disc";
      if (node.tag === "ol") result["list-style-type"] = ({ "1": "decimal", a: "lower-alpha", A: "upper-alpha", i: "lower-roman", I: "upper-roman" } as Record<string, string>)[node.attrs.type ?? "1"] ?? "decimal";
      if (["b", "strong"].includes(node.tag)) result["font-weight"] = "bold";
      if (["em", "i"].includes(node.tag)) result["font-style"] = "italic";
      if (node.tag === "u" || node.tag === "ins" && !this.revisionIds.has(node.attrs["data-node-id"] ?? "")) result["text-decoration-line"] = "underline";
      if (node.tag === "s" || node.tag === "del" && !this.revisionIds.has(node.attrs["data-node-id"] ?? "")) result["text-decoration-line"] = "line-through";
      if (/^h[1-6]$/.test(node.tag)) result["font-size"] = [2, 1.5, 1.17, 1, .83, .67][Number(node.tag[1]) - 1] + "em";
      if (["sub", "sup"].includes(node.tag)) result["vertical-align"] = node.tag === "sup" ? "super" : "sub";
      if (["td", "th"].includes(node.tag)) result["vertical-align"] = "top";
    }
    let serial = 0;
    for (const rule of [...rules].sort((a, b) => a.order - b.order)) for (const declaration of rule.declarations) {
      this.step(); const decoded = ident.decode(declaration.property), name = decoded.startsWith("--") ? decoded : decoded.toLowerCase(), entry = { value: generate(declaration.value), important: Boolean(declaration.important), specificity: rule.specificity, order: serial++ };
      if (name.startsWith("--") && higher(entry, custom.get(name))) custom.set(name, entry);
    }
    const cycles = new Set<string>(), visited = new Set<string>(), path: string[] = [];
    const dependencies = (name: string): void => {
      const start = path.indexOf(name); if (start >= 0) { for (const member of path.slice(start)) cycles.add(member); return; }
      if (visited.has(name) || !custom.has(name)) return;
      if (path.length >= this.limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "CSS variable graph exceeds its depth budget.");
      path.push(name);
      walk(parse(custom.get(name)!.value, { context: "value", parseCustomProperty: true }), (item) => {
        this.step(); if (item.type === "Function" && ident.decode(item.name).toLowerCase() === "var" && item.children.first?.type === "Identifier") dependencies(ident.decode(item.children.first.name));
      });
      path.pop(); visited.add(name);
    };
    for (const name of custom.keys()) dependencies(name);
    const memo = new Map<string, string | undefined>(), resolving = new Set<string>();
    const variable = (name: string): string | undefined => {
      if (memo.has(name)) return memo.get(name);
      if (cycles.has(name)) return undefined;
      if (resolving.has(name)) return undefined;
      if (resolving.size >= this.limits.maxDepth) throw new DocumentError("ResourceLimitExceeded", "CSS variable resolution exceeds its depth budget.");
      const local = custom.get(name); if (!local) return parent[name];
      resolving.add(name); const value = ["inherit", "unset"].includes(local.value) ? parent[name] : local.value === "initial" ? undefined : resolveSourceCssVariables(local.value, variable, 0, this.limits.maxFileBytes);
      resolving.delete(name); memo.set(name, value); return value;
    };
    for (const name of custom.keys()) { const value = variable(name); if (value === undefined) delete result[name]; else result[name] = value; }
    for (const rule of [...rules].sort((a, b) => a.order - b.order)) for (const declaration of rule.declarations) {
      const decoded = ident.decode(declaration.property), property = decoded.startsWith("--") ? decoded : decoded.toLowerCase(); if (property.startsWith("--")) continue;
      const resolved = resolveSourceCssVariables(generate(declaration.value), (name) => result[name], 0, this.limits.maxFileBytes);
      const value = resolved !== undefined && !lexer.matchProperty(property, resolved).error ? resolved : undefined;
      // Invalid-at-computed-value declarations still win cascade and compute to unset.
      for (const [name, expanded] of expandSourceCssProperty(property, value ?? "unset")) {
        const entry = { value: expanded, important: Boolean(declaration.important), specificity: rule.specificity, order: serial++ };
        if (higher(entry, winners.get(name))) winners.set(name, entry);
      }
    }
    for (const [name, entry] of winners) {
      const value = entry.value;
      if (value === "inherit" || value === "unset" && inherited.has(name)) result[name] = parent[name] ?? this.initial[name] ?? "initial";
      else if (value === "initial" || value === "unset" || value === "revert") result[name] = this.initial[name] ?? "initial";
      else result[name] = lexer.matchProperty(name, value).error ? (inherited.has(name) ? parent[name] ?? this.initial[name] ?? "initial" : this.initial[name] ?? "initial") : value;
    }
    const base = parseFloat(parent["font-size"] ?? "12pt"), size = result["font-size"] ?? "12pt";
    const sizes: Record<string, number> = { "xx-small": 7.2, "x-small": 9, small: 10, medium: 12, large: 13.5, "x-large": 18, "xx-large": 24, "xxx-large": 36, smaller: base / 1.2, larger: base * 1.2 };
    result["font-size"] = String(sizes[size] ?? length(size, { font: base, rootFont: node?.tag === "html" ? 12 : this.rootFont, percent: base })) + "pt";
    const font = parseFloat(result["font-size"]);
    if (node?.tag === "html") this.rootFont = font;
    // Relative lengths compute on the element that declares them, before they
    // are inherited. Percentages stay relative to layout except for line-height.
    for (const [name, raw] of Object.entries(result)) {
      if (name.startsWith("--") || name === "font-family" || name === "font-size" || !/[\d.]r?em\b/i.test(raw)) continue;
      const ast = parse(raw, { context: "value" });
      walk(ast, (item) => {
        if (item.type === "Dimension" && ["em", "rem"].includes(item.unit.toLowerCase())) { item.value = String(Number(item.value) * (item.unit.toLowerCase() === "em" ? font : this.rootFont)); item.unit = "pt"; }
      });
      result[name] = generate(ast);
    }
    if (result["line-height"] !== "normal" && !/^[\d.]+$/.test(result["line-height"]!)) result["line-height"] = length(result["line-height"]!, { font, rootFont: this.rootFont, percent: font }) + "pt";
    for (const key of ["letter-spacing", "word-spacing"]) if (result[key] !== "normal") result[key] = length(result[key]!, { font, rootFont: this.rootFont, percent: font }) + "pt";
    if (result.color?.toLowerCase() === "currentcolor") result.color = parent.color ?? this.initial.color!;
    // Decorations propagate through descendant inline boxes, including from a
    // block ancestor. `none` on a descendant does not erase an ancestor's paint.
    const ownLines = result["text-decoration-line"]!, parentLines = node?.parent ? parent["text-decoration-line"] ?? "none" : "none";
    const lines = [...new Set([...parentLines.split(/\s+/), ...ownLines.split(/\s+/)].filter((line) => line !== "none"))];
    if (lines.length && combineDecorations) {
      if (parentLines !== "none" && ownLines !== "none") {
        const parentColor = parent["text-decoration-color"]?.toLowerCase() === "currentcolor" ? parent.color : parent["text-decoration-color"];
        const ownColor = result["text-decoration-color"]?.toLowerCase() === "currentcolor" ? result.color : result["text-decoration-color"];
        if (color(parentColor ?? "#000000", parent.color) !== color(ownColor ?? "#000000", result.color) || parent["text-decoration-style"] !== result["text-decoration-style"]) throw new DocumentError("UnsupportedFeature", "Multiple decoration paints cannot share one native Word run property.");
      }
      if (parentLines !== "none" && ownLines === "none") {
        result["text-decoration-style"] = parent["text-decoration-style"]!;
        result["text-decoration-color"] = parent["text-decoration-color"]?.toLowerCase() === "currentcolor" ? parent.color! : parent["text-decoration-color"]!;
      }
      result["text-decoration-line"] = lines.join(" ");
    }
    if (["bolder", "lighter"].includes(result["font-weight"] ?? "")) {
      const weight = parent["font-weight"] === "bold" ? 700 : Number(parent["font-weight"]) || 400;
      result["font-weight"] = String(result["font-weight"] === "bolder" ? weight < 350 ? 400 : weight < 550 ? 700 : 900 : weight < 550 ? 100 : weight < 750 ? 400 : 700);
    }
    return Object.freeze(result);
  }
}
