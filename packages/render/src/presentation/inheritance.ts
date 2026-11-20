import { DocumentError } from "@doffi/core";
import { markViewNode, ownViewAttribute } from "../view-nodes.js";

const ordinary = ["font-family", "font-size", "font-weight", "font-style", "color", "line-height", "letter-spacing", "text-align", "text-indent", "margin-left", "margin-right", "margin-top", "margin-bottom", "padding-top", "padding-right", "padding-bottom", "padding-left", "tab-size", "text-transform", "font-variant"];
export function inheritedLeading(node: HTMLElement, computed: CSSStyleDeclaration): string {
  const raw = computed.lineHeight, size = parseFloat(computed.fontSize), height = parseFloat(raw);
  if (!Number.isFinite(height) || !Number.isFinite(size) || size <= 0) return raw;
  const typed = node as HTMLElement & { computedStyleMap?: () => { get(name: string): { unit?: string; value?: number } | undefined } };
  const value = typed.computedStyleMap?.().get("line-height") as { unit?: string; value?: number } | undefined;
  if (value?.unit === "number") return String(value.value);
  const before = node.style.getPropertyValue("font-size"), priority = node.style.getPropertyPriority("font-size");
  node.style.setProperty("font-size", size * 2 + "px", "important");
  const doubled = parseFloat(node.ownerDocument.defaultView!.getComputedStyle(node).lineHeight);
  if (before) node.style.setProperty("font-size", before, priority); else node.style.removeProperty("font-size");
  return Math.abs(doubled - height * 2) < 0.05 ? String(height / size) : raw;
}
/** The browser resolves source CSS and variables in the instance's theme. Only
 * low-specificity fallback declarations are added; local source rules win. */
export class PresentationInheritance {
  private readonly sheet: HTMLStyleElement;
  private serial = 0;
  private work = 0;
  private bytes = 0;
  private readonly names: readonly string[];
  constructor(private readonly document: Document, sourceCss: string, private readonly budget: number) {
    this.names = [...new Set([...ordinary, ...sourceCss.matchAll(/(--office-[a-zA-Z0-9-]+)\s*:/g)].map(item => typeof item === "string" ? item : item[1]!))];
    this.sheet = markViewNode(document.createElement("style")); document.head.append(this.sheet);
  }
  private rule(node: HTMLElement, values: ReadonlyMap<string, string>): void {
    if (++this.serial > this.budget * 8) throw new DocumentError("ResourceLimitExceeded", "Presentation inheritance exceeds its node budget.");
    ownViewAttribute(node, "data-sdk-style-key"); node.dataset.sdkStyleKey = [node.dataset.sdkStyleKey, String(this.serial)].filter(Boolean).join(" ");
    const declaration = this.document.createElement("div").style;
    for (const [key, value] of values) if (value) declaration.setProperty(key, value);
    const rule = `${node.tagName.toLowerCase()}:where([data-sdk-style-key~="${this.serial}"]){${declaration.cssText}}`;
    this.bytes += rule.length * 2;
    if (this.bytes > this.budget * 1024) throw new DocumentError("ResourceLimitExceeded", "Presentation inheritance exceeds its style budget.");
    this.sheet.sheet!.insertRule(rule);
  }
  apply(node: HTMLElement, prototypes: readonly HTMLElement[], initial: ReadonlyMap<string, string> = new Map()): void {
    const values = new Map(initial);
    for (const prototype of prototypes) {
      const probe = markViewNode(this.document.createElement(node.tagName));
      probe.className = prototype.className; probe.setAttribute("style", prototype.getAttribute("style") ?? "");
      if (prototype.hasAttribute("data-office-kind")) probe.setAttribute("data-office-kind", prototype.getAttribute("data-office-kind")!);
      probe.setAttribute("data-sdk-probe", ""); node.parentElement!.append(probe);
      this.rule(probe, values);
      const names = new Set([...this.names, ...values.keys(), ...[...probe.style].filter(name => name.startsWith("--office-"))]);
      this.work += names.size;
      if (this.work > this.budget * 256) throw new DocumentError("ResourceLimitExceeded", "Presentation inheritance exceeds its work budget.");
      const computed = this.document.defaultView!.getComputedStyle(probe);
      const leading = inheritedLeading(probe, computed);
      for (const key of names) {
        const value = key === "line-height" ? leading : computed.getPropertyValue(key).trim(); if (value) values.set(key, value);
      }
      probe.remove();
    }
    if (values.size) this.rule(node, values);
  }
}
