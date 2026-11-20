import { A, check, child, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { fontName, value } from "../drawingml/values.js";
import { effects, fill, onlyColor, outline, themeColorRoles as roles } from "../drawingml/styles.js";
import type { Context } from "./context.js";
export * from "../drawingml/styles.js";

export function importTheme(context: Context, path: string): string {
  const known = context.surfaces.get(path);
  if (known) return known.id;
  const source = context.part(path, "theme", "theme", A);
  check(source, ["themeElements", "objectDefaults", "extraClrSchemeLst", "custClrLst", "extLst"], A);
  const theme = child(source, "themeElements", A);
  if (!theme) fail(source, "Theme elements are missing.");
  check(theme, ["clrScheme", "fontScheme", "fmtScheme"], A);
  const colors = child(theme, "clrScheme", A), fonts = child(theme, "fontScheme", A), format = child(theme, "fmtScheme", A);
  if (!colors || !fonts || !format) fail(source, "A presentation theme is incomplete.");
  const css: Css = {}, id = context.id("theme");
  for (const role of roles) {
    const item = child(colors, role, A);
    if (!item) fail(colors, "A required theme color is missing.");
    const v = onlyColor(item);
    if (!v.css.startsWith("#") || Object.keys(v).length !== 1) unsupported(item, "Theme base colors must be untransformed RGB values.");
    css["--office-palette-" + role] = v.css;
    css["--office-theme-" + role] = "var(--office-palette-" + role + ")";
  }
  for (const kind of ["major", "minor"] as const) {
    const parent = child(fonts, kind + "Font", A);
    if (!parent) fail(fonts, "A theme font role is missing.");
    for (const [tag, role] of [["latin", "ascii"], ["ea", "east-asia"], ["cs", "complex-script"]] as const) {
      const name = value(child(parent, tag, A)!, "typeface");
      if (name === undefined) fail(parent, "A theme font name is missing.");
      // An empty East Asian font uses the declared Hans supplemental role.
      const supplement = elements(parent, "font", A).find(n => value(n, "script") === (role === "east-asia" ? "Hans" : "Arab"));
      const resolved = name || (supplement && value(supplement, "typeface"));
      css["--office-" + kind + "-" + role] = resolved ? fontName(resolved) : "sans-serif";
      if (role === "ascii") css["--office-" + kind + "-h-ansi"] = css["--office-" + kind + "-" + role]!;
    }
    for (const item of elements(parent, "font", A)) {
      const script = value(item, "script"), name = value(item, "typeface");
      if (!script || !/^[A-Za-z]{4}$/.test(script) || !name) fail(item, "A supplemental theme font is invalid.");
      css["--office-" + kind + "-script-" + script] = fontName(name);
    }
  }
  css.color = "var(--office-theme-dk1)"; css["font-family"] = "var(--office-minor-ascii),var(--office-minor-east-asia)"; css["font-size"] = "18pt";
  context.rule(id, css);
  for (const [name, prefix, map] of [["fillStyleLst", "fill", (n: XmlNode) => fill({ ...n, content: [n] })],
    ["bgFillStyleLst", "background", (n: XmlNode) => fill({ ...n, content: [n] })],
    ["lnStyleLst", "line", outline], ["effectStyleLst", "effect", (n: XmlNode) => { check(n, ["effectLst", "scene3d", "sp3d"], A); if (child(n, "scene3d", A) || child(n, "sp3d", A)) unsupported(n); return effects(child(n, "effectLst", A)); }]] as const) {
    const list = child(format, name, A);
    if (!list) fail(format, "A theme format list is missing.");
    elements(list, undefined, A).forEach((item, index) => {
      const css = map(item);
      for (const key of Object.keys(css)) css[key] = css[key]!.replaceAll("--office-placeholder-color", "--office-" + prefix + "-reference-color");
      for (const key of Object.keys(css)) if (key.endsWith("-source") && css[key] === "placeholder") {
        css[key] = "reference"; css[key.slice(0, -7) + "-reference-prefix"] = "--office-" + prefix + "-reference";
      }
      context.rule(id + "-" + prefix + "-" + (index + 1), css);
    });
  }
  const defaults = child(source, "objectDefaults", A);
  check(defaults, ["spDef", "lnDef", "txDef", "extLst"], A);
  if (defaults && elements(defaults, undefined, A).length) {
    for (const kind of ["spDef", "lnDef", "txDef"]) {
      const definition = child(defaults, kind, A);
      if (!definition) continue;
      check(definition, ["spPr", "bodyPr", "lstStyle", "style", "extLst"], A);
      for (const required of ["spPr", "bodyPr", "lstStyle"])
        if (!child(definition, required, A)) fail(definition, "A theme object default is incomplete.");
    }
    context.retainPreview("theme-defaults", id, defaults, "Theme object defaults require compatible preview.");
  }
  context.metadata.themes.push({ themeId: id, cssClass: id });
  context.surfaces.set(path, { id, root: { tag: "div", attrs: {}, children: [] }, source, themeId: id });
  return id;
}
