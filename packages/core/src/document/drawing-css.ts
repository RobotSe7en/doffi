import type { DocumentReadLimits } from "../runtime.js";
import { setInlineStyle } from "./css.js";

/** Direct drawing colors replace native theme transforms. The existing
 * declarations remain the sole appearance source for both PPT and workbooks. */
export function setDrawingInlineStyle(source: string, property: string, value: string, limits: DocumentReadLimits): string {
  let style = setInlineStyle(source, property, value, limits);
  const name = property.toLowerCase(), prefix = name === "color" ? "--office-text" : name === "background-color" ? "--office-fill"
    : /^--office-[a-z0-9-]+-color$/.test(property) ? property.slice(0, -6) : undefined;
  if (prefix) {
    style = setInlineStyle(style, prefix + "-color", value, limits);
    for (const [name, value] of Object.entries({ source: "direct", alpha: "1", "lum-mod": "1", "lum-off": "0", tint: "0", shade: "1" })) style = setInlineStyle(style, prefix + "-" + name, value, limits);
    if (name === "background-color") style = setInlineStyle(style, "--office-fill-kind", "solid", limits);
  }
  if (name === "font-family") for (const role of ["ascii", "h-ansi", "east-asia", "complex-script"]) style = setInlineStyle(style, "--office-font-" + role, value, limits);
  return style;
}
