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
  ]) {
    let style = css["border-" + side + "-style"],
      rawWidth = css["border-" + side + "-width"],
      rawColor = css["border-" + side + "-color"];
    const custom =
      side === "insideH"
        ? css["--word-border-inside-h"]
        : side === "insideV"
          ? css["--word-border-inside-v"]
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
