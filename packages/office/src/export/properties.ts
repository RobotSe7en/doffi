import { wordRowExceptionProperties } from "@doffi/core";
import { writeWordTablePosition } from "./table-position.js";
import {
  passiveWordRunBooleans,
  passiveWordRunValues,
  passiveWordRunCss,
  wordUnderlineValues,
  wordEmphasisValues,
  wordOpenTypeLigatures,
  wordOpenTypeNumSpacings,
  wordOpenTypeNumForms,
} from "../ooxml/word-passive-run-properties.js";
import {
  passiveWordParagraphBooleans,
  passiveWordParagraphValues,
  passiveWordParagraphCss,
} from "../ooxml/word-passive-paragraph-properties.js";
import { DocumentError } from "@doffi/core";
import type { Css } from "./css.js";
import { color, fontName, length, nativeInteger } from "./values.js";
import { empty, orderedProperties, value } from "./xml.js";

export const runKeys: readonly string[] = [
  ...passiveWordRunBooleans.map(passiveWordRunCss),
  ...passiveWordRunValues.map(passiveWordRunCss),
  "font-family",
  "font-size",
  "font-weight",
  "font-style",
  "font-variant",
  "text-transform",
  "color",
  "letter-spacing",
  "vertical-align",
  "background-color",
  "text-decoration-line",
  "text-decoration-style",
  "text-decoration-color",
  "--word-font-ascii",
  "--word-font-hAnsi",
  "--word-font-eastAsia",
  "--word-font-cs",
  "--word-font-hint",
  "--word-font-size-cs",
  "--word-bold-cs",
  "--word-italic-cs",
  "--word-kerning-threshold",
  "--word-run-shading",
  "--word-run-snap-to-grid",
  "--word-run-native-w",
  "--word-dstrike",
  "--word-underline-style",
  "--word-run-border",
  "--word-text-glow",
  "--word-text-outline",
  "--word-text-shadow",
  "--word-text-reflection",
  "--word-text-fill-gradient",
  "--word-text-bevel",
  "--word-text-3d",
  "--word-ea-combine",
  "--word-ea-vert",
  "--word-fit-text",
  "--word-opentype-ligatures",
  "--word-opentype-num-spacing",
  "--word-opentype-num-form",
  "--word-opentype-stylistic-set",
  "--word-opentype-contextual-alternates",
  ...["val", "eastAsia", "bidi"].map((role) => "--word-language-" + role),
];
export const paragraphKeys: readonly string[] = [
  ...passiveWordParagraphBooleans.map(passiveWordParagraphCss),
  ...passiveWordParagraphValues.map(passiveWordParagraphCss),
  "margin-top",
  "margin-bottom",
  "margin-left",
  "margin-right",
  "text-indent",
  "text-align",
  "line-height",
  "break-before",
  "break-after",
  "break-inside",
  "widows",
  "orphans",
  "background-color",
  "--word-contextual-spacing",
  "--word-suppress-auto-hyphens",
  "--word-overflow-punctuation",
  "--word-top-line-punctuation",
  "--word-kinsoku",
  "--word-word-wrap",
  "--word-auto-space-de",
  "--word-auto-space-dn",
  "--word-paragraph-snap-to-grid",
  "--word-distributed",
  "--word-line-rule",
  "--word-drop-cap",
  "--word-drop-cap-lines",
  "--word-drop-cap-space",
  ...["top", "right", "bottom", "left"].flatMap((side) =>
    ["width", "style", "color"].map((part) => "border-" + side + "-" + part),
  ),
];
export const pick = (css: Css, keys: readonly string[]): Css =>
  Object.fromEntries(Object.entries(css).filter(([key]) => keys.includes(key)));
export const difference = (css: Css, expected: Css, keys: readonly string[]): Css =>
  Object.fromEntries(
    Object.entries(css).filter(([key, val]) => keys.includes(key) && val !== expected[key]),
  );
export function unsupported(message: string): never {
  throw new DocumentError("UnsupportedFeature", message);
}
export function points(raw: string, css: Css = {}, width = 468): number {
  return length(raw, { font: parseFloat(css["font-size"] ?? "12pt"), percent: width });
}
export function rowHeightProperties(css: Css): string {
  const rule = css["--word-row-height-rule"] ?? "atLeast";
  if (!["auto", "atLeast", "exact"].includes(rule))
    unsupported("The native table row height rule is invalid.");
  if ((!css.height || css.height === "auto") && rule !== "auto") return "";
  return empty("w:trHeight", {
    "w:val": nativeInteger(
      css.height && css.height !== "auto" ? points(css.height, css) : 0,
      20,
      0,
    ),
    "w:hRule": rule,
  });
}
const toggle = (name: string, enabled: boolean): string => value(name, enabled ? "1" : "0");
const shade = (css: Css): string =>
  css["background-color"] === undefined
    ? ""
    : empty("w:shd", {
        "w:val": "clear",
        "w:color": "auto",
        "w:fill": color(css["background-color"], css.color) ?? "auto",
      });
const bold = (value: string | undefined): boolean => value === "bold" || Number(value) >= 600;
export function rowProperties(css: Css): string {
  return (
    (css["break-inside"] !== undefined
      ? toggle("cantSplit", css["break-inside"] === "avoid" || css["break-inside"] === "avoid-page")
      : "") +
    rowHeightProperties(css) +
    (css["--word-row-alignment"] !== undefined ? value("jc", css["--word-row-alignment"]) : "") +
    wordCellSpacing(css, "row") +
    (css["--word-row-repeat-header"] !== undefined
      ? toggle("tblHeader", css["--word-row-repeat-header"] === "1")
      : "")
  );
}
/** Style toggle values are relative to basedOn; direct formatting is absolute. */
export function runProperties(css: Css, inherited: Css = {}, style = false, role?: string): string {
  let out = "";
  const fonts: Record<string, string> = {};
  const hint = css["--word-font-hint"];
  for (const name of ["ascii", "hAnsi", "eastAsia", "cs"]) {
    const raw = css["--word-font-" + name];
    if (raw) fonts["w:" + name] = fontName(raw);
  }
  if (css["font-family"]) {
    const name = fontName(css["font-family"]);
    if (role) fonts["w:" + role] = name;
    else if (!Object.keys(fonts).length)
      for (const key of ["ascii", "hAnsi", "eastAsia", "cs"]) fonts["w:" + key] = name;
  }
  if (hint !== undefined) {
    if (!["default", "eastAsia", "cs"].includes(hint)) unsupported("Invalid native run font hint.");
    fonts["w:hint"] = hint;
  }
  for (const name of passiveWordRunBooleans) {
    const raw = css[passiveWordRunCss(name)];
    if (raw !== undefined) {
      if (raw !== "0" && raw !== "1") unsupported("A native run Boolean must be zero or one.");
      out += toggle(name, raw === "1");
    }
  }
  const width = css["--word-run-native-w"];
  if (width !== undefined) {
    const scale = Number(width);
    if (!/^\d+$/.test(width) || scale < 1 || scale > 600)
      unsupported("A native run character scale must be a whole percent between 1 and 600.");
    out += value("w", scale);
  }
  for (const name of passiveWordRunValues) {
    const raw = css[passiveWordRunCss(name)];
    if (raw !== undefined) {
      if (name === "position") {
        const parsed = Number(raw);
        if (!/^-?\d+$/.test(raw) || parsed < -3170 || parsed > 3170)
          unsupported("A native run character position must be an integer between -3170 and 3170.");
        out += value(name, raw);
      } else if (name === "em") {
        if (!(wordEmphasisValues as readonly string[]).includes(raw))
          unsupported("An emphasis mark style is outside the supported DOCX profile.");
        out += value(name, raw);
      } else {
        if (raw !== "none")
          unsupported("An active native run effect is outside the editable profile.");
        out += value(name, raw);
      }
    }
  }
  if (css["--word-run-snap-to-grid"] !== undefined)
    out += toggle("snapToGrid", css["--word-run-snap-to-grid"] === "1");
  if (Object.keys(fonts).length) out += empty("w:rFonts", fonts);
  for (const [prop, name, active] of [
    ["font-weight", "b", "bold"],
    ["font-style", "i", "italic"],
    ["text-transform", "caps", "uppercase"],
    ["font-variant", "smallCaps", "small-caps"],
  ]) {
    if (css[prop!] === undefined) continue;
    const enabled = prop === "font-weight" ? bold(css[prop!]) : css[prop!] === active;
    if (prop === "font-style" && !["normal", "italic"].includes(css[prop!]!))
      unsupported("Oblique fonts cannot be expressed by the supported DOCX profile.");
    if (prop === "text-transform" && !["none", "uppercase"].includes(css[prop!]!))
      unsupported("This text transform cannot be expressed by the supported DOCX profile.");
    const previous = prop === "font-weight" ? bold(inherited[prop!]) : inherited[prop!] === active;
    if (!style || enabled !== previous) out += toggle(name!, style || enabled);
  }
  for (const [prop, name] of [
    ["--word-bold-cs", "bCs"],
    ["--word-italic-cs", "iCs"],
  ])
    if (css[prop!] !== undefined) out += toggle(name!, css[prop!] === "1");
  let emittedDstrike = false;
  if (css["--word-dstrike"] !== undefined) {
    if (css["--word-dstrike"] !== "0" && css["--word-dstrike"] !== "1")
      unsupported("A native double-strikethrough value must be zero or one.");
    const enabled = css["--word-dstrike"] === "1",
      previous = inherited["--word-dstrike"] === "1";
    if (!style || enabled !== previous) {
      out += toggle("dstrike", style || enabled);
      emittedDstrike = true;
    }
  }
  if (css["text-decoration-line"] !== undefined) {
    const lines = css["text-decoration-line"].split(/\s+/);
    if (lines.includes("overline")) unsupported("DOCX has no supported overline property.");
    const enabled = lines.includes("line-through"),
      previous = (inherited["text-decoration-line"] ?? "").includes("line-through");
    if (!style || enabled !== previous) {
      if (!style && !enabled)
        out += toggle("strike", false) + (emittedDstrike ? "" : toggle("dstrike", false));
      else if (
        css["--word-dstrike"] === undefined &&
        !lines.includes("underline") &&
        css["--word-underline-style"] === undefined &&
        css["text-decoration-style"] === "double"
      ) {
        if (!emittedDstrike) {
          out += toggle("dstrike", style || enabled);
          emittedDstrike = true;
        }
      } else out += toggle("strike", style || enabled);
    }
  }
  if (css.color !== undefined)
    out += value(
      "color",
      color(css.color) ?? unsupported("Transparent native text is not supported."),
    );
  if (css["letter-spacing"] !== undefined)
    out += value(
      "spacing",
      css["letter-spacing"] === "normal"
        ? 0
        : nativeInteger(points(css["letter-spacing"], { ...inherited, ...css })),
    );
  if (css["--word-kerning-threshold"] !== undefined)
    out += value(
      "kern",
      nativeInteger(points(css["--word-kerning-threshold"], { ...inherited, ...css }), 2, 0),
    );
  if (css["font-size"] !== undefined)
    out += value("sz", nativeInteger(points(css["font-size"], css), 2, 1));
  if (css["--word-font-size-cs"] !== undefined)
    out += value("szCs", nativeInteger(points(css["--word-font-size-cs"], css), 2, 1));
  const background = css["background-color"],
    base = css["--word-run-shading"];
  const highlights: Record<string, string> = {
    ffff00: "yellow",
    "00ff00": "green",
    "00ffff": "cyan",
    ff00ff: "magenta",
    "0000ff": "blue",
    ff0000: "red",
    "000080": "darkBlue",
    "008080": "darkCyan",
    "008000": "darkGreen",
    "800080": "darkMagenta",
    "800000": "darkRed",
    "808000": "darkYellow",
    "808080": "darkGray",
    c0c0c0: "lightGray",
    "000000": "black",
    ffffff: "white",
  };
  const fill = background !== undefined ? color(background, css.color) : undefined,
    baseFill = base ? color(base, css.color) : undefined;
  if (
    background !== undefined &&
    base !== undefined &&
    fill !== baseFill &&
    fill &&
    highlights[fill]
  )
    out += value("highlight", highlights[fill]);
  else if (background !== undefined) out += value("highlight", "none");
  if (
    css["text-decoration-line"] !== undefined ||
    css["text-decoration-style"] !== undefined ||
    css["text-decoration-color"] !== undefined ||
    css["--word-underline-style"] !== undefined
  ) {
    const lines = css["text-decoration-line"] ?? inherited["text-decoration-line"] ?? "none";
    const styles: Record<string, string> = {
      solid: "single",
      double: "double",
      dotted: "dotted",
      dashed: "dash",
      wavy: "wave",
    };
    const directExplicit = css["--word-underline-style"],
      inheritedExplicit = inherited["--word-underline-style"];
    if (
      directExplicit !== undefined &&
      !(wordUnderlineValues as readonly string[]).includes(directExplicit)
    )
      unsupported("Unsupported native underline style.");
    if (
      inheritedExplicit !== undefined &&
      !(wordUnderlineValues as readonly string[]).includes(inheritedExplicit)
    )
      unsupported("Unsupported native underline style.");
    const actualCssStyle = css["text-decoration-style"] ?? inherited["text-decoration-style"] ?? "solid";
    const fallbackStyle =
      styles[actualCssStyle] ?? unsupported("Unsupported native underline style.");
    const nativeCssStyle = (value: string): string =>
      value === "double"
        ? "double"
        : value === "dotted" || value === "dottedHeavy"
          ? "dotted"
          : value === "wave" || value === "wavyHeavy" || value === "wavyDouble"
            ? "wavy"
            : [
                  "dash",
                  "dashedHeavy",
                  "dashLong",
                  "dashLongHeavy",
                  "dotDash",
                  "dashDotHeavy",
                  "dotDotDash",
                  "dashDotDotHeavy",
                ].includes(value)
              ? "dashed"
              : "solid";
    const explicit =
      directExplicit !== undefined
        ? directExplicit
        : css["text-decoration-style"] !== undefined
          ? undefined
          : inheritedExplicit;
    const underlineVal = !lines.includes("underline")
      ? "none"
      : explicit && explicit !== "none" && nativeCssStyle(explicit) === actualCssStyle
        ? explicit
        : fallbackStyle;
    out += empty("w:u", {
      "w:val": underlineVal,
      ...(css["text-decoration-color"] &&
      css["text-decoration-color"].toLowerCase() !== "currentcolor"
        ? { "w:color": color(css["text-decoration-color"], css.color ?? inherited.color) }
        : {}),
    });
  }
  const runBorder = css["--word-run-border"];
  if (runBorder !== undefined) {
    const trimmed = runBorder.trim();
    if (trimmed === "none") out += empty("w:bdr", { "w:val": "none", "w:sz": 0, "w:space": 0, "w:color": "auto" });
    else {
      const match = /^(\S+)\s+(\S+)\s+(.+)$/.exec(trimmed);
      if (!match) unsupported("Invalid native character border.");
      const [, rawWidth, borderStyle, rawColor] = match!;
      const borderStyles: Record<string, string> = {
        none: "none",
        hidden: "none",
        solid: "single",
        double: "double",
        dotted: "dotted",
        dashed: "dashed",
      };
      const widths: Record<string, number> = { thin: 0.75, medium: 2.25, thick: 3.75 };
      const width = widths[rawWidth!] ?? points(rawWidth!, css);
      const nativeStyle =
        width === 0
          ? "none"
          : (borderStyles[borderStyle!] ??
            unsupported("This character border style has no supported DOCX representation."));
      out += empty("w:bdr", {
        "w:val": nativeStyle,
        "w:sz": nativeStyle === "none" ? 0 : Math.max(2, nativeInteger(width, 8, 0)),
        "w:space": 0,
        "w:color":
          rawColor!.toLowerCase() === "currentcolor"
            ? "auto"
            : (color(rawColor!, css.color ?? inherited.color) ?? "auto"),
      });
    }
  }
  const fitText = css["--word-fit-text"];
  if (fitText !== undefined && fitText.trim() !== "none") {
    const valTwips = nativeInteger(points(fitText, css), 20, 1);
    out += empty("w:fitText", { "w:val": valTwips });
  }
  const eaCombine = css["--word-ea-combine"];
  const eaVert = css["--word-ea-vert"];
  if (
    (eaCombine && eaCombine !== "none") ||
    (eaVert && eaVert !== "0")
  ) {
    out += empty("w:eastAsianLayout", {
      ...(eaCombine && eaCombine !== "none"
        ? { "w:combine": 1, "w:combineBrackets": eaCombine }
        : {}),
      ...(eaVert && eaVert !== "0"
        ? { "w:vert": 1, ...(eaVert === "compress" ? { "w:vertCompress": 1 } : {}) }
        : {}),
    });
  }
  const textGlow = css["--word-text-glow"];
  if (textGlow && textGlow.trim() !== "none") {
    const m = /^(\S+)\s+(.+)$/.exec(textGlow.trim());
    if (m) {
      const rad = Math.max(0, Math.round(points(m[1]!, css) * 12700));
      const hex = (color(m[2]!, css.color) ?? "4472C4").toUpperCase();
      out += `<w14:glow w14:rad="${rad}"><w14:srgbClr w14:val="${hex}"/></w14:glow>`;
    }
  }
  const textOutline = css["--word-text-outline"];
  if (textOutline && textOutline.trim() !== "none") {
    const m = /^(\S+)\s+(.+)$/.exec(textOutline.trim());
    if (m) {
      const w = Math.max(0, Math.round(points(m[1]!, css) * 12700));
      const hex = (color(m[2]!, css.color) ?? "1F4E79").toUpperCase();
      out += `<w14:textOutline w14:w="${w}" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:srgbClr w14:val="${hex}"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>`;
    }
  }
  const textShadow = css["--word-text-shadow"];
  if (textShadow && textShadow.trim() !== "none") {
    const m = /^(\S+)\s+(\S+)\s+(\S+)\s+(.+)$/.exec(textShadow.trim());
    if (m) {
      const blur = Math.max(0, Math.round(points(m[1]!, css) * 12700));
      const dist = Math.max(0, Math.round(points(m[2]!, css) * 12700));
      const dir = Math.max(0, Math.round(Number(m[3]) * 60000));
      const hex = (color(m[4]!, css.color) ?? "000000").toUpperCase();
      out += `<w14:shadow w14:blurRad="${blur}" w14:dist="${dist}" w14:dir="${dir}" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:srgbClr w14:val="${hex}"><w14:alpha w14:val="40000"/></w14:srgbClr></w14:shadow>`;
    }
  }
  if (css["--word-text-reflection"] === "1") {
    out +=
      '<w14:reflection w14:blurRad="6350" w14:stA="50000" w14:stPos="0" w14:endA="300" w14:endPos="35000" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>';
  }
  const textGradient = css["--word-text-fill-gradient"];
  if (textGradient && textGradient.trim() !== "none") {
    const raw = textGradient.trim();
    const angleMatch = /^(-?\d+(?:\.\d+)?)deg\b/i.exec(raw);
    const angleDeg = angleMatch ? Number(angleMatch[1]) : 90;
    const body = angleMatch ? raw.slice(angleMatch[0].length) : raw;
    const stops: { hex: string; pos: number }[] = [];
    const stopRegex = /(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\)|[a-zA-Z]+)\s*(\d+(?:\.\d+)?%)?/g;
    let m: RegExpExecArray | null;
    while ((m = stopRegex.exec(body)) !== null) {
      const c = color(m[1]!, css.color);
      if (!c) continue;
      const pos = m[2]
        ? Math.max(0, Math.min(100, Math.round(Number(m[2].slice(0, -1)))))
        : stops.length === 0
          ? 0
          : 100;
      stops.push({ hex: c.toUpperCase(), pos });
    }
    if (stops.length >= 2) {
      const ang = (((Math.round(angleDeg) % 360) + 360) % 360) * 60000;
      out +=
        "<w14:textFill><w14:gradFill><w14:gsLst>" +
        stops
          .map(
            (stop) =>
              `<w14:gs w14:pos="${stop.pos * 1000}"><w14:srgbClr w14:val="${stop.hex}"/></w14:gs>`,
          )
          .join("") +
        `</w14:gsLst><w14:lin w14:ang="${ang}" w14:scaled="1"/></w14:gradFill></w14:textFill>`;
    }
  }
  const text3d = css["--word-text-3d"];
  if (text3d && text3d.trim() !== "none") {
    const prst = /^[A-Za-z0-9]+$/.test(text3d.trim()) ? text3d.trim() : "orthographicFront";
    out += `<w14:scene3d><w14:camera w14:prst="${prst}"/><w14:lightRig w14:rig="threePt" w14:dir="t"/></w14:scene3d>`;
  }
  const textBevel = css["--word-text-bevel"];
  if (textBevel && textBevel.trim() !== "none") {
    const parts = textBevel.trim().split(/\s+/);
    const prst = /^[A-Za-z0-9]+$/.test(parts[0] ?? "") ? parts[0]! : "circle";
    const w = parts[1] ? Math.max(0, Math.round(points(parts[1], css) * 12700)) : 38100;
    const h = parts[2] ? Math.max(0, Math.round(points(parts[2], css) * 12700)) : 38100;
    out += `<w14:props3d><w14:bevelT w14:w="${w}" w14:h="${h}" w14:prst="${prst}"/></w14:props3d>`;
  }
  const ligatures = css["--word-opentype-ligatures"];
  if (ligatures !== undefined) {
    if (!(wordOpenTypeLigatures as readonly string[]).includes(ligatures))
      unsupported("Unsupported OpenType ligature setting.");
    out += empty("w14:ligatures", { "w14:val": ligatures });
  }
  const numForm = css["--word-opentype-num-form"];
  if (numForm !== undefined) {
    if (!(wordOpenTypeNumForms as readonly string[]).includes(numForm))
      unsupported("Unsupported OpenType number form setting.");
    out += empty("w14:numForm", { "w14:val": numForm });
  }
  const numSpacing = css["--word-opentype-num-spacing"];
  if (numSpacing !== undefined) {
    if (!(wordOpenTypeNumSpacings as readonly string[]).includes(numSpacing))
      unsupported("Unsupported OpenType number spacing setting.");
    out += empty("w14:numSpacing", { "w14:val": numSpacing });
  }
  const stylisticSet = css["--word-opentype-stylistic-set"];
  if (stylisticSet !== undefined) {
    const trimmed = stylisticSet.trim();
    if (trimmed === "none" || trimmed === "0" || trimmed === "") out += empty("w14:stylisticSets");
    else {
      const ids = trimmed.split(",").map((part) => part.trim());
      let items = "";
      for (const rawId of ids) {
        const id = Number(rawId);
        if (!/^\d+$/.test(rawId) || id < 1 || id > 20)
          unsupported("An OpenType stylistic set id must be between 1 and 20.");
        items += empty("w14:styleSet", { "w14:id": id, "w14:val": "on" });
      }
      out += "<w14:stylisticSets>" + items + "</w14:stylisticSets>";
    }
  }
  const contextualAlternates = css["--word-opentype-contextual-alternates"];
  if (contextualAlternates !== undefined) {
    if (contextualAlternates !== "0" && contextualAlternates !== "1")
      unsupported("An OpenType contextual alternates value must be zero or one.");
    out += empty("w14:cntxtAlts", { "w14:val": contextualAlternates });
  }
  if (base !== undefined || background !== undefined)
    out += shade({
      ...css,
      "background-color":
        base !== undefined && fill !== baseFill && fill && highlights[fill]
          ? base
          : (background ?? base!),
    });
  if (css["vertical-align"] !== undefined) {
    const values: Record<string, string> = {
      baseline: "baseline",
      super: "superscript",
      sub: "subscript",
    };
    out += value(
      "vertAlign",
      values[css["vertical-align"]] ??
        unsupported("This inline vertical alignment has no supported DOCX representation."),
    );
  }
  const languages: Record<string, string> = {};
  for (const role of ["val", "eastAsia", "bidi"])
    if (css["--word-language-" + role] !== undefined)
      languages["w:" + role] = fontName(css["--word-language-" + role]!);
  if (Object.keys(languages).length) out += empty("w:lang", languages);
  return orderedProperties("rPr", out);
}
export function borders(css: Css, container: string, inside = false): string {
  let out = "";
  for (const side of [
    "top",
    "left",
    "bottom",
    "right",
    ...(inside ? ["insideH", "insideV"] : []),
    ...(container === "tcBorders" ? ["tl2br", "tr2bl"] : []),
  ]) {
    let style = css["border-" + side + "-style"],
      rawWidth = css["border-" + side + "-width"],
      rawColor = css["border-" + side + "-color"];
    const custom =
      side === "insideH"
        ? css["--word-border-inside-h"]
        : side === "insideV"
          ? css["--word-border-inside-v"]
          : side === "tl2br"
            ? css["--word-border-tl2br"]
            : side === "tr2bl"
              ? css["--word-border-tr2bl"]
              : undefined;
    if (custom) {
      const match = /^(\S+)\s+(\S+)\s+(.+)$/.exec(custom);
      if (!match) unsupported("Invalid native table border.");
      [, rawWidth, style, rawColor] = match!;
    }
    if (style === undefined && rawWidth === undefined && rawColor === undefined) continue;
    const styles: Record<string, string> = {
      none: "nil",
      hidden: "nil",
      solid: "single",
      double: "double",
      dotted: "dotted",
      dashed: "dashed",
    };
    const widths: Record<string, number> = { thin: 0.75, medium: 2.25, thick: 3.75 };
    const width = rawWidth ? (widths[rawWidth] ?? points(rawWidth, css)) : 0.5;
    const nativeStyle =
      width === 0
        ? "nil"
        : (styles[style ?? "none"] ??
          unsupported("This border style has no supported DOCX representation."));
    out += empty("w:" + side, {
      "w:val": nativeStyle,
      "w:sz": Math.max(2, nativeInteger(width, 8, 0)),
      "w:space": 0,
      "w:color": color(rawColor ?? "currentColor", css.color) ?? "auto",
    });
  }
  return out ? "<w:" + container + ">" + out + "</w:" + container + ">" : "";
}
export function paragraphProperties(css: Css, width = 468): string {
  let out = "";
  for (const name of passiveWordParagraphBooleans) {
    const raw = css[passiveWordParagraphCss(name)];
    if (raw !== undefined) {
      if ((raw !== "0" && raw !== "1") || (name === "bidi" && raw !== "0"))
        unsupported(
          "An active or invalid native paragraph direction requires compatibility preservation.",
        );
      out += toggle(name, raw === "1");
    }
  }
  const direction = css[passiveWordParagraphCss("textDirection")];
  if (direction !== undefined) {
    if (direction !== "lrTb")
      unsupported(
        "An alternate native paragraph text direction requires compatibility preservation.",
      );
    out += value("textDirection", direction);
  }
  const div = css[passiveWordParagraphCss("divId")];
  if (div !== undefined) {
    if (!/^\d+$/.test(div) || !Number.isSafeInteger(Number(div)))
      unsupported("A native paragraph div identity must be a nonnegative integer.");
    out += value("divId", div);
  }
  for (const [prop, name, active] of [
    ["break-after", "keepNext", "avoid"],
    ["break-inside", "keepLines", "avoid"],
    ["break-before", "pageBreakBefore", "page"],
  ]) {
    if (css[prop!] !== undefined) {
      if (!["auto", active, "avoid-page"].includes(css[prop!]!))
        unsupported("This paragraph break setting has no supported DOCX representation.");
      out += toggle(name!, css[prop!] === active || css[prop!] === "avoid-page");
    }
  }
  if (css.widows !== undefined || css.orphans !== undefined) {
    if (
      [css.widows, css.orphans].some((value) => value !== undefined && !["1", "2"].includes(value))
    )
      unsupported("DOCX widow control supports one or two lines in this profile.");
    out += toggle("widowControl", css.widows !== "1" && css.orphans !== "1");
  }
  const dropCap = css["--word-drop-cap"];
  if (dropCap === "drop" || dropCap === "margin") {
    const lines = Math.max(1, Math.min(10, Math.round(Number(css["--word-drop-cap-lines"] ?? "3")) || 3));
    const space = css["--word-drop-cap-space"] ? nativeInteger(points(css["--word-drop-cap-space"], css), 20, 0) : 0;
    out += empty("w:framePr", {
      "w:dropCap": dropCap,
      "w:lines": lines,
      "w:hSpace": space,
      "w:wrap": "around",
      "w:vAnchor": "text",
      "w:hAnchor": "text",
    });
    if (css["--word-paragraph-native-suppressLineNumbers"] === undefined)
      out += "<w:suppressLineNumbers/>";
  }
  if (css["--word-paragraph-snap-to-grid"] !== undefined)
    out += toggle("snapToGrid", css["--word-paragraph-snap-to-grid"] === "1");
  out += borders(css, "pBdr") + shade(css);
  for (const [key, name] of [
    ["--word-contextual-spacing", "contextualSpacing"],
    ["--word-suppress-auto-hyphens", "suppressAutoHyphens"],
    ["--word-overflow-punctuation", "overflowPunct"],
    ["--word-top-line-punctuation", "topLinePunct"],
    ["--word-kinsoku", "kinsoku"],
    ["--word-word-wrap", "wordWrap"],
    ["--word-auto-space-de", "autoSpaceDE"],
    ["--word-auto-space-dn", "autoSpaceDN"],
  ])
    if (css[key!] !== undefined) {
      if (!["0", "1"].includes(css[key!]!))
        unsupported("A native paragraph Boolean must be zero or one.");
      out += toggle(name!, css[key!] === "1");
    }
  const spacing: Record<string, string | number> = {};
  for (const [key, name] of [
    ["margin-top", "before"],
    ["margin-bottom", "after"],
  ])
    if (css[key!] !== undefined)
      spacing["w:" + name!] = nativeInteger(points(css[key!]!, css, width), 20, 0);
  if (css["line-height"] === "normal") {
    spacing["w:line"] = 240;
    spacing["w:lineRule"] = "auto";
  } else if (css["line-height"] !== undefined) {
    const line = css["line-height"];
    if (/^[\d.]+$/.test(line)) {
      spacing["w:line"] = nativeInteger(Number(line), 240, 1);
      spacing["w:lineRule"] = "auto";
    } else {
      spacing["w:line"] = nativeInteger(
        points(line, css, parseFloat(css["font-size"] ?? "12pt")),
        20,
        1,
      );
      spacing["w:lineRule"] = css["--word-line-rule"] === "atLeast" ? "atLeast" : "exact";
    }
  }
  if (Object.keys(spacing).length) out += empty("w:spacing", spacing);
  const ind: Record<string, number> = {};
  for (const [key, name] of [
    ["margin-left", "left"],
    ["margin-right", "right"],
  ])
    if (css[key!] !== undefined) ind["w:" + name!] = nativeInteger(points(css[key!]!, css, width));
  if (css["text-indent"] !== undefined) {
    const amount = points(css["text-indent"], css, width);
    ind[amount < 0 ? "w:hanging" : "w:firstLine"] = nativeInteger(Math.abs(amount));
  }
  if (Object.keys(ind).length) out += empty("w:ind", ind);
  if (css["text-align"] !== undefined)
    out += value(
      "jc",
      css["--word-distributed"] === "1"
        ? "distribute"
        : css["text-align"] === "justify"
          ? "both"
          : css["text-align"],
    );
  return orderedProperties("pPr", out);
}
export function nativeWidth(name: string, raw: string | undefined, css: Css, width = 468): string {
  if (raw === undefined || raw === "auto" || raw === "initial") return "";
  return empty(
    "w:" + name,
    raw.endsWith("%")
      ? { "w:type": "pct", "w:w": nativeInteger(parseFloat(raw), 50, 0) }
      : { "w:type": "dxa", "w:w": nativeInteger(points(raw, css, width)) },
  );
}
export function cellMargins(css: Css, name: string): string {
  const out = ["top", "left", "bottom", "right"]
    .map((side) => nativeWidth(side, css["padding-" + side], css))
    .join("");
  return out ? "<w:" + name + ">" + out + "</w:" + name + ">" : "";
}
export function tableProperties(css: Css): string {
  let out =
    (css.width === "auto" && ["auto", "nil"].includes(css["--word-table-width-type"] ?? "")
      ? empty("w:tblW", { "w:type": css["--word-table-width-type"], "w:w": "0" })
      : nativeWidth("tblW", css.width, css)) + writeWordTablePosition(css);
  if (css["--word-table-bidi"] !== undefined) {
    if (!["0", "1"].includes(css["--word-table-bidi"]))
      unsupported("A native table direction must be zero or one.");
    out += toggle("bidiVisual", css["--word-table-bidi"] === "1");
  }
  out += wordCellSpacing(css, "table");
  out += cellMargins(
    Object.fromEntries(
      Object.entries(css)
        .filter(([key]) => key.startsWith("--word-table-cell-padding-"))
        .map(([key, value]) => [key.slice(18), value]),
    ),
    "tblCellMar",
  );
  if (css["margin-left"] === "auto")
    out += value("jc", css["margin-right"] === "auto" ? "center" : "right");
  else out += nativeWidth("tblInd", css["margin-left"], css);
  out += borders(css, "tblBorders", true) + shade(css);
  if (css["table-layout"] !== undefined)
    out += empty("w:tblLayout", {
      "w:type": css["table-layout"] === "fixed" ? "fixed" : "autofit",
    });
  return orderedProperties("tblPr", out);
}
export function rowExceptionProperties(css: Css): string {
  const properties = wordRowExceptionProperties(css),
    look = properties["look-mask"];
  return orderedProperties(
    "tblPrEx",
    tableProperties(properties).replace(
      /<w:tblInd[^>]*\/>/g,
      properties.indent !== undefined ? "" : "$&",
    ) +
      (properties.indent !== undefined
        ? nativeWidth("tblInd", properties.indent, properties)
        : "") +
      cellMargins(properties, "tblCellMar") +
      (look !== undefined
        ? empty("w:tblLook", { "w:val": Number(look).toString(16).padStart(4, "0") })
        : ""),
  );
}
function wordCellSpacing(css: Css, kind: "table" | "row" | "row-exception"): string {
  let out = "";
  const prefix = "--word-" + kind + "-cell-spacing";
  const spacing =
    css[prefix] ??
    (kind === "table" && css["border-collapse"] === "separate" ? css["border-spacing"] : undefined);
  if (spacing !== undefined) {
    const values = spacing.trim().split(/\s+/);
    if (values.length > 2 || (values.length === 2 && values[0] !== values[1]))
      unsupported("Native Word table cell spacing must be uniform.");
    const amount = points(values[0]!, css);
    if (!Number.isFinite(amount) || amount < 0)
      unsupported("Native table cell spacing must be a nonnegative length.");
    const type = css[prefix + "-type"] ?? "dxa",
      raw = css[prefix + "-value"];
    if (!["dxa", "nil", "auto", "pct"].includes(type))
      unsupported("A native table spacing unit is invalid.");
    if (amount === 0 && type !== "dxa")
      out += empty("w:tblCellSpacing", { "w:type": type, "w:w": raw ?? "0" });
    else {
      let retained = false;
      if (raw !== undefined && type === "dxa") {
        const parsed = /^-?\d+$/.test(raw)
          ? Number(raw) / 20
          : points(raw.replace(/pi$/, "pc"), css);
        retained = Math.abs(parsed - amount) < 1e-7;
      }
      out += empty("w:tblCellSpacing", {
        "w:type": "dxa",
        "w:w": retained ? raw : nativeInteger(amount),
      });
    }
  }
  return out;
}
export function cellProperties(css: Css): string {
  let out = nativeWidth("tcW", css.width, css) + borders(css, "tcBorders") + shade(css);
  if (css["white-space"] === "nowrap") out += value("noWrap", "1");
  out += cellMargins(css, "tcMar");
  if (css["vertical-align"] !== undefined)
    out += value("vAlign", css["vertical-align"] === "middle" ? "center" : css["vertical-align"]);
  const hideMark = css["--word-cell-hide-mark"];
  if (hideMark !== undefined) out += toggle("hideMark", hideMark === "1");
  return orderedProperties("tcPr", out);
}
