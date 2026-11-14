import { DocumentError } from "@doffi/core";
import { generate, parse } from "css-tree";
import { color, length } from "./values.js";
import type { Css } from "./css.js";
export function unsupported(message: string): never { throw new DocumentError("UnsupportedFeature", message); }
export function integer(value: number, factor = 1, minimum = -27273042329600, maximum = 27273042316900): number {
    const result = Math.round(value * factor);
    if (!Number.isSafeInteger(result) || result < minimum || result > maximum)
        unsupported("A source value is outside its native DrawingML range.");
    return result;
}
export const emu = (value: number, minimum = -27273042329600, maximum = 27273042316900): number => integer(value, 12700, minimum, maximum);
export function points(value: string | undefined, css: Css, fallback = 0, percent?: number): number { if (value?.includes("%") && percent === undefined)
    unsupported("Percentage box dimensions require a resolved native reference and are outside this DrawingML profile."); return value === undefined || value === "initial" || value === "auto" ? fallback : length(value, { font: parseFloat(css["font-size"] ?? "18pt"), percent: percent ?? 0 }); }
export function cssString(value: string): string {
    const ast = parse(value, { context: "value" });
    if (ast.type === "Value" && ast.children.size === 1 && ast.children.first?.type === "String")
        return ast.children.first.value;
    return value;
}
/** DrawingML carries RGB and opacity independently. Word's opaque-color profile
 * continues to reject partial alpha through its existing color() boundary. */
export function rgba(value: string, current = "#000000"): {
    rgb: string;
    alpha: number;
} {
    let source = value.trim();
    if (source.toLowerCase() === "currentcolor")
        source = current.toLowerCase() === "currentcolor" ? "#000000" : current;
    if (source.toLowerCase() === "transparent")
        return { rgb: "000000", alpha: 0 };
    const hex = /^#([a-f\d]{4}|[a-f\d]{8})$/i.exec(source)?.[1];
    if (hex) {
        const alpha = parseInt(hex.length === 4 ? hex[3]!.repeat(2) : hex.slice(6), 16) / 255;
        source = "#" + (hex.length === 4 ? [...hex.slice(0, 3)].map(c => c + c).join("") : hex.slice(0, 6));
        return { rgb: color(source)!, alpha };
    }
    const ast = parse(source, { context: "value" });
    if (ast.type === "Value" && ast.children.size === 1 && ast.children.first?.type === "Function") {
        const fn = ast.children.first, values = fn.children.toArray().filter(node => node.type !== "Operator");
        if (values.length === 4) {
            const a = values[3]!, alpha = a.type === "Number" ? Number(a.value) : a.type === "Percentage" ? Number(a.value) / 100 : NaN;
            if (!Number.isFinite(alpha) || alpha < 0 || alpha > 1)
                unsupported("A source opacity is outside its native range.");
            const name = fn.name.toLowerCase().startsWith("hsl") ? "hsl" : fn.name.toLowerCase().startsWith("rgb") ? "rgb" : unsupported("This color function is outside the PPTX profile.");
            return { rgb: color(name + "(" + values.slice(0, 3).map(node => generate(node)).join(",") + ")", current)!, alpha };
        }
    }
    const rgb = color(source, current);
    return { rgb: rgb ?? "000000", alpha: rgb === undefined ? 0 : 1 };
}
