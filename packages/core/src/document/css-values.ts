import { parse } from "css-tree";
import type { CssNode } from "css-tree";
import { DocumentError } from "../errors.js";

export interface LengthContext { readonly font: number; readonly rootFont?: number; readonly percent: number }
interface Quantity { value: number; dimension: boolean }
function unsupported(value: string): never { throw new DocumentError("UnsupportedFeature", "A CSS value cannot be represented in native Office: " + value.slice(0, 80)); }
/** Evaluate validated dimensional arithmetic with an explicit document context. */
export function length(value: string, context: LengthContext): number {
  const ast = parse(value, { context: "value" });
  const evaluate = (nodes: CssNode[], depth: number): Quantity => {
    if (depth > 128) throw new DocumentError("ResourceLimitExceeded", "Office length arithmetic exceeds its depth budget.");
    let cursor = 0;
    const primary = (): Quantity => {
      const node = nodes[cursor++]; if (!node) return unsupported(value);
      if (node.type === "Number") return { value: Number(node.value), dimension: false };
      if (node.type === "Percentage") return { value: Number(node.value) * context.percent / 100, dimension: true };
      if (node.type === "Dimension") {
        const factors: Record<string, number> = { pt: 1, px: .75, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, q: 72 / 101.6, pc: 12, em: context.font, rem: context.rootFont ?? 12 };
        const factor = factors[node.unit.toLowerCase()]; if (factor === undefined) return unsupported(value);
        return { value: Number(node.value) * factor, dimension: true };
      }
      if (node.type === "Parentheses") return evaluate(node.children.toArray(), depth + 1);
      if (node.type === "Function") {
        const name = node.name.toLowerCase(), children = node.children.toArray();
        if (name === "calc") return evaluate(children, depth + 1);
        if (!["min", "max", "clamp"].includes(name)) return unsupported(value);
        const groups: CssNode[][] = [[]]; for (const child of children) { if (child.type === "Operator" && child.value === ",") groups.push([]); else groups.at(-1)!.push(child); }
        const values = groups.map((group) => evaluate(group, depth + 1));
        if (values.some((item) => item.dimension !== values[0]!.dimension) || name === "clamp" && values.length !== 3) return unsupported(value);
        const numbers = values.map((item) => item.value);
        return { value: name === "min" ? numbers.reduce((left, right) => Math.min(left, right)) : name === "max" ? numbers.reduce((left, right) => Math.max(left, right)) : Math.max(numbers[0]!, Math.min(numbers[1]!, numbers[2]!)), dimension: values[0]!.dimension };
      }
      if (node.type === "Operator" && ["+", "-"].includes(node.value)) { const result = primary(); return { ...result, value: result.value * (node.value === "-" ? -1 : 1) }; }
      return unsupported(value);
    };
    const product = (): Quantity => {
      let result = primary();
      while (nodes[cursor]?.type === "Operator" && ["*", "/"].includes((nodes[cursor] as { value: string }).value.trim())) {
        const operation = (nodes[cursor++] as { value: string }).value.trim(), right = primary();
        if (operation === "*" && result.dimension && right.dimension || operation === "/" && (!result.dimension && right.dimension || right.value === 0)) return unsupported(value);
        result = { value: operation === "*" ? result.value * right.value : result.value / right.value,
          dimension: operation === "*" ? result.dimension || right.dimension : result.dimension && !right.dimension };
      }
      return result;
    };
    let result = product();
    while (cursor < nodes.length) {
      const operation = nodes[cursor++]; if (operation?.type !== "Operator" || !["+", "-"].includes(operation.value.trim())) return unsupported(value);
      const right = product(); if (result.dimension !== right.dimension) return unsupported(value);
      result = { value: result.value + right.value * (operation.value.trim() === "-" ? -1 : 1), dimension: result.dimension };
    }
    return result;
  };
  if (ast.type !== "Value") return unsupported(value);
  const result = evaluate(ast.children.toArray(), 0);
  if (!Number.isFinite(result.value) || Math.abs(result.value) > 100_000_000 || !result.dimension && result.value !== 0) return unsupported(value);
  return result.value;
}
const named: Readonly<Record<string, string>> = { black: "000000", silver: "c0c0c0", gray: "808080", grey: "808080", white: "ffffff", maroon: "800000", red: "ff0000", purple: "800080", fuchsia: "ff00ff", green: "008000", lime: "00ff00", olive: "808000", yellow: "ffff00", navy: "000080", blue: "0000ff", teal: "008080", aqua: "00ffff", orange: "ffa500", rebeccapurple: "663399", darkblue: "00008b", darkgreen: "006400", darkred: "8b0000", lightgray: "d3d3d3", lightgrey: "d3d3d3" };
export function color(value: string, current = "#000000"): string | undefined {
  const input = value.trim().toLowerCase();
  if (input === "transparent") return undefined;
  if (input === "currentcolor") return color(current === "currentColor" ? "#000000" : current);
  if (named[input]) return named[input];
  const hex = /^#([a-f0-9]{3,8})$/.exec(input)?.[1];
  if (hex) {
    if (hex.length === 3) return [...hex].map((part) => part + part).join("");
    if (hex.length === 6) return hex;
    const alpha = hex.length === 4 ? hex[3]!.repeat(2) : hex.length === 8 ? hex.slice(6) : undefined;
    if (alpha === "00") return undefined; if (alpha !== "ff") return unsupported(value);
    return hex.length === 4 ? [...hex.slice(0, 3)].map((part) => part + part).join("") : hex.slice(0, 6);
  }
  const ast = parse(input, { context: "value" });
  if (ast.type !== "Value" || ast.children.size !== 1 || ast.children.first?.type !== "Function") return unsupported(value);
  const fn = ast.children.first, values = fn.children.toArray().filter((node) => node.type !== "Operator");
  if (values.length < 3 || values.length > 4) return unsupported(value);
  const component = (node: CssNode | undefined, percent: number): number => node?.type === "Number" ? Number(node.value) : node?.type === "Percentage" ? Number(node.value) * percent / 100 : unsupported(value);
  const alpha = values[3] ? component(values[3], 1) : 1;
  if (alpha <= 0) return undefined; if (alpha < 1) return unsupported(value);
  let rgb: number[];
  if (["rgb", "rgba"].includes(fn.name)) rgb = values.slice(0, 3).map((node) => component(node, 255));
  else if (["hsl", "hsla"].includes(fn.name)) {
    const first = values[0]!, factors: Record<string, number> = { deg: 1, grad: .9, rad: 180 / Math.PI, turn: 360 };
    const hue = first.type === "Dimension" ? Number(first.value) * (factors[first.unit] ?? unsupported(value)) : component(first, 360);
    const h = ((hue % 360) + 360) % 360 / 60, s = Math.min(1, Math.max(0, component(values[1], 1))), l = Math.min(1, Math.max(0, component(values[2], 1)));
    const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(h % 2 - 1)), m = l - c / 2;
    const parts = h < 1 ? [c, x, 0] : h < 2 ? [x, c, 0] : h < 3 ? [0, c, x] : h < 4 ? [0, x, c] : h < 5 ? [x, 0, c] : [c, 0, x]; rgb = parts.map((part) => (part + m) * 255);
  } else return unsupported(value);
  return rgb.map((part) => Math.round(Math.max(0, Math.min(255, part))).toString(16).padStart(2, "0")).join("");
}
