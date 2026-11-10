import type { DocumentReadLimits } from "../runtime.js";
import { setInlineStyle } from "./css.js";
import { ident, parse, walk } from "css-tree";
import type { CssNode } from "css-tree";
import { DocumentError } from "../errors.js";

export function drawingStyleVariable(value: CssNode): { name: string; fallback: boolean } | undefined {
  if (value.type !== "Value" || value.children.size !== 1) return;
  const call = value.children.first;
  if (call?.type !== "Function" || ident.decode(call.name).toLowerCase() !== "var" || call.children.first?.type !== "Identifier") return;
  return { name: ident.decode(call.children.first.name), fallback: call.children.toArray().some(node => node.type === "Operator" && node.value === ",") };
}

/** Direct drawing colors replace native theme transforms. The existing
 * declarations remain the sole appearance source for both PPT and workbooks. */
export function setDrawingInlineStyle(source: string, property: string, value: string, limits: DocumentReadLimits): string {
  let style = setInlineStyle(source, property, value, limits);
  const name = property.startsWith("--") ? property : property.toLowerCase(), prefix = name === "color" ? "--office-text" : name === "background-color" ? "--office-fill"
    : /^--office-[a-z0-9-]+-color$/.test(property) ? property.slice(0, -6) : undefined;
  const references = new Set<string>(); let variable: ReturnType<typeof drawingStyleVariable>;
  if (value.includes("(") && (prefix || name === "font-family" || name.startsWith("--"))) {
    const ast = parse(style, { context: "declarationList", parseCustomProperty: true });
    if (ast.type !== "DeclarationList") throw new DocumentError("InvalidOperation", "A drawing style requires valid declarations.");
    const declaration = ast.children.toArray().reverse().find(node => node.type === "Declaration" && ident.decode(node.property) === name);
    if (declaration?.type !== "Declaration") throw new DocumentError("InvalidOperation", "A drawing style requires its requested declaration.");
    variable = drawingStyleVariable(declaration.value);
    walk(declaration.value, node => { if (node.type === "Function" && ident.decode(node.name).toLowerCase() === "var" && node.children.first?.type === "Identifier") references.add(ident.decode(node.children.first.name)); });
  }
  if (name.startsWith("--") && references.has(name)) throw new DocumentError("InvalidOperation", "A native drawing style variable cannot refer to itself.");
  const retain = (target: string): boolean => {
    if (!references.has(target)) return false;
    if (variable?.name !== target) throw new DocumentError("UnsupportedFeature", "A native style self-reference requires a direct variable expression.");
    return true;
  };
  if (prefix) {
    if (!retain(prefix + "-color")) style = setInlineStyle(style, prefix + "-color", value, limits);
    for (const [name, value] of Object.entries({ source: "direct", alpha: "1", "lum-mod": "1", "lum-off": "0", tint: "0", shade: "1" })) style = setInlineStyle(style, prefix + "-" + name, value, limits);
    if (name === "background-color") style = setInlineStyle(style, "--office-fill-kind", "solid", limits);
  }
  if (name === "font-family") for (const role of ["ascii", "h-ansi", "east-asia", "complex-script"]) if (!retain("--office-font-" + role)) style = setInlineStyle(style, "--office-font-" + role, value, limits);
  return style;
}
