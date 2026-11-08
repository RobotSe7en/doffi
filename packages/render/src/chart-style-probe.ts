import { DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, SourceStyleResolver, resolveSourceCssVariables } from "@doffi/core";
import type { SourceStyleNode } from "@doffi/core";
import { markViewNode } from "./view-nodes.js";
import type { DrawingPaintStyle } from "./presentation/appearance.js";

type Classes = string | readonly string[] | undefined;
type Decoration = { line: string; style: string; color: string };
const fontDecorations = new WeakMap<CSSStyleDeclaration, Decoration>();
const rolePaint = new WeakMap<CSSStyleDeclaration, ReadonlySet<string>>();
/** Capture the metadata role's paint before its probes are removed. The outer
 * chart's frame is not a default background/border for each label. */
export function captureChartRolePaint(style: CSSStyleDeclaration, accountText: (units: number) => void): DrawingPaintStyle | undefined {
  const keys = rolePaint.get(style); if (!keys?.size) return undefined;
  const values = new Map<string, string>();
  const take = (name: string) => { if (values.has(name)) return values.get(name)!; const value = style.getPropertyValue(name).trim(); accountText(name.length + value.length); values.set(name, value); return value; };
  for (const key of keys) take(key);
  // Reference colors can be supplied by a theme rather than the local role.
  const references = ["--office-fill-reference", "--office-line-reference", "--office-background-reference", "--office-effect-reference"];
  for (let depth = 0; depth < 5; depth++) {
    const prefixes = [...values].filter(([name, value]) => name.endsWith("-source") && value === "reference").map(([name]) => take(name.slice(0, -7) + "-reference-prefix"));
    let added = false;
    for (const prefix of prefixes) if (references.includes(prefix)) for (const suffix of ["color", "source", "reference-prefix", "alpha", "lum-mod", "lum-off", "tint", "shade"]) { const key = prefix + "-" + suffix; if (!values.has(key)) { take(key); added = true; } }
    if (!added) break;
  }
  const fill = values.get("--office-fill-kind"), line = values.get("--office-line-kind");
  if (!(fill !== "none" && (["solid", "linear"].includes(fill ?? "") || values.get("--office-fill-color"))) && !(line !== "none" && (line === "solid" || values.get("--office-line-color")))) return undefined;
  return Object.freeze({ color: style.color, getPropertyValue: (name: string) => values.get(name) ?? "" });
}
/** Native text color transforms also color currentColor decorations. */
export function chartFontDecoration(style: CSSStyleDeclaration, foreground: string): Readonly<Record<string, string>> {
  const decoration = fontDecorations.get(style);
  return { "text-decoration-line": decoration?.line ?? style.textDecorationLine, "text-decoration-style": decoration?.style ?? style.textDecorationStyle,
    "text-decoration-color": decoration?.color.toLowerCase() === "currentcolor" ? foreground : style.textDecorationColor };
}
const decorationKeys = ["text-decoration", "text-decoration-line", "text-decoration-style", "text-decoration-color"] as const;
function sourceStyles(document: Document): SourceStyleResolver {
  const sheets = [...document.head.querySelectorAll<HTMLStyleElement>("style[data-office-source-styles],style[data-sdk-source-styles]")], css = sheets.map(sheet => sheet.textContent ?? "");
  if (!sheets.length) throw new DocumentError("InvalidOperation", "Chart font roles require their owning source styles.");
  // These styles have already passed their package/surface admission budgets.
  // Bound this second, mask-only parse by those captured input strings, rather
  // than imposing a smaller package limit on a host's admitted source.
  const units = css.reduce((total, text) => total + text.length, 0), resolver = new SourceStyleResolver(css, {
    ...DEFAULT_DOCUMENT_READ_LIMITS, maxFiles: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxFiles, sheets.length),
    maxCssNodes: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxCssNodes, units * 2), maxDepth: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxDepth, units),
    maxFileBytes: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxFileBytes, units * 4), maxPackageBytes: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxPackageBytes, units * 4),
  });
  return resolver;
}
function projection(element: HTMLElement): SourceStyleNode {
  let parent: SourceStyleNode | undefined; const ancestors: HTMLElement[] = [];
  for (let current: HTMLElement | null = element; current; current = current.parentElement) ancestors.push(current);
  for (const current of ancestors.reverse()) {
    const attrs = Object.fromEntries([...current.attributes].filter(attribute => attribute.name !== "style").map(attribute => [attribute.name, attribute.value]));
    // Layout adds private positioning to these nodes. Only authored decoration
    // declarations participate in this source-presence query; the browser still
    // computes every actual font, variable and physical value.
    attrs.style = decorationKeys.flatMap(name => { const value = current.style.getPropertyValue(name); return value ? [name + ":" + value + (current.style.getPropertyPriority(name) ? "!important" : "")] : []; }).join(";");
    parent = { tag: current.localName, attrs, parent };
  }
  return parent!;
}
/** Materialize Office font defaults across metadata roles. This is distinct
 * from real HTML ancestor paint: explicit none on a point clears its default. */
export function createChartStyleProbe(scope: HTMLElement, source: HTMLElement, account: (amount: number) => void, probes: HTMLElement[]): (classes?: Classes) => CSSStyleDeclaration {
  const document = scope.ownerDocument, view = document.defaultView!, resolver = sourceStyles(document), decorations = new WeakMap<HTMLElement, Decoration>(), paintKeys = new WeakMap<HTMLElement, ReadonlySet<string>>();
  account(1); const parser = markViewNode(document.createElement("span")); probes.push(parser);
  const own = (node: HTMLElement) => resolver.specified(projection(node));
  const color = (raw: string | undefined, computed: CSSStyleDeclaration, inherited: string): string => {
    if (raw === undefined) return inherited;
    const value = resolveSourceCssVariables(raw, name => computed.getPropertyValue(name).trim() || undefined)?.trim();
    if (value === "inherit") return inherited;
    if (value === undefined || ["initial", "unset", "revert"].includes(value)) return "currentColor";
    // An unresolved variable shorthand is opaque in the presence query. The
    // browser's own value parser identifies its omitted/currentColor component.
    parser.style.textDecoration = ""; parser.style.textDecoration = value;
    const resolved = parser.style.textDecorationColor;
    return !resolved || resolved.toLowerCase() === "currentcolor" ? "currentColor" : computed.textDecorationColor;
  };
  const base = view.getComputedStyle(source), raw = own(source);
  decorations.set(source, { line: base.textDecorationLine, style: base.textDecorationStyle, color: color(raw["text-decoration-color"], base, "currentColor") });
  paintKeys.set(source, new Set());
  function role(probe: HTMLElement, parent: HTMLElement): CSSStyleDeclaration {
    const raw = own(probe), computed = view.getComputedStyle(probe), inherited = decorations.get(parent)!;
    const decoration = { line: raw["text-decoration-line"] === undefined ? inherited.line : computed.textDecorationLine,
      style: raw["text-decoration-style"] === undefined ? inherited.style : computed.textDecorationStyle,
      color: color(raw["text-decoration-color"], computed, inherited.color) };
    // Only private probes are written. Canonical source elements and their
    // classes stay untouched, and no probe is serialized into output.
    probe.style.setProperty("text-decoration-line", decoration.line, "important"); probe.style.setProperty("text-decoration-style", decoration.style, "important"); probe.style.setProperty("text-decoration-color", decoration.color, "important");
    const paint = new Set([...paintKeys.get(parent) ?? [], ...Object.keys(raw).filter(name => /^--office-(?:fill|line)-/.test(name))]);
    paintKeys.set(probe, paint); decorations.set(probe, decoration); const result = view.getComputedStyle(probe); fontDecorations.set(result, decoration); rolePaint.set(result, paint); return result;
  }
  if (scope !== source) role(scope, source);
  return classes => {
    let parent = scope, result: CSSStyleDeclaration | undefined;
    const layers = typeof classes === "string" ? [classes] : classes?.length ? classes : [""];
    for (const name of layers) {
      account(1); const probe = markViewNode(document.createElement("span")); probe.className = name;
      probe.style.cssText = "position:absolute;visibility:hidden"; parent.append(probe); probes.push(probe); result = role(probe, parent); parent = probe;
    }
    return result!;
  };
}
