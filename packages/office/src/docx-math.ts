import { DocumentError, validateWordMath } from "@doffi/core";
import type { DeepReadonly, WordMathNode, WordMathTree, WordMetadata } from "@doffi/core";
import {
  W,
  attr,
  child,
  elements,
  check,
  checkAttributes,
  number,
  text,
  fail,
  unsupported,
} from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import { empty, escapeXml } from "./export/xml.js";
export const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
type Style = NonNullable<WordMathNode["style"]>;
type Input<T> = T extends WordMathNode ? Omit<T, "mathId"> : never;
const flag = (node: XmlNode | undefined, fallback = false): boolean => {
  if (!node) return fallback;
  const value = attr(node, "val", node.uri);
  if (value === undefined) return true;
  if (["1", "true", "on"].includes(value)) return true;
  if (["0", "false", "off"].includes(value)) return false;
  return fail(node, "A math Boolean is invalid.");
};
function readStyle(node: XmlNode | undefined): Style | undefined {
  if (!node) return;
  check(node, ["rFonts", "sz", "color", "b", "i", "lang"]);
  const fonts = child(node, "rFonts"),
    size = child(node, "sz"),
    color = child(node, "color"),
    language = attr(child(node, "lang"), "val");
  if (child(node, "lang")) checkAttributes(child(node, "lang")!, ["val"], W);
  if (fonts) checkAttributes(fonts, ["ascii", "hAnsi", "eastAsia", "cs"], W);
  if (color) checkAttributes(color, ["val"], W);
  const result: Style = {
    ...(fonts
      ? {
          fonts: Object.fromEntries(
            ["ascii", "hAnsi", "eastAsia", "cs"].flatMap((role) => {
              const value = attr(fonts, role);
              return value ? [[role, value]] : [];
            }),
          ),
        }
      : {}),
    ...(size ? { fontSize: number(attr(size, "val"), size) / 2 } : {}),
    ...(color ? { color: attr(color, "val") === "auto" ? "auto" : "#" + attr(color, "val") } : {}),
    ...(child(node, "b") ? { bold: flag(child(node, "b")) } : {}),
    ...(child(node, "i") ? { italic: flag(child(node, "i")) } : {}),
    ...(language ? { language } : {}),
  };
  return Object.keys(result).length ? result : undefined;
}
function writeStyle(style: DeepReadonly<Style> | undefined): string {
  if (!style) return "";
  return (
    "<w:rPr>" +
    (style.fonts
      ? empty(
          "w:rFonts",
          Object.fromEntries(
            Object.entries(style.fonts).map(([key, value]) => ["w:" + key, value]),
          ),
        )
      : "") +
    (style.bold === undefined ? "" : empty("w:b", { "w:val": style.bold ? 1 : 0 })) +
    (style.italic === undefined ? "" : empty("w:i", { "w:val": style.italic ? 1 : 0 })) +
    (style.fontSize ? empty("w:sz", { "w:val": Math.round(style.fontSize * 2) }) : "") +
    (style.color
      ? empty("w:color", { "w:val": style.color === "auto" ? "auto" : style.color.slice(1) })
      : "") +
    (style.language ? empty("w:lang", { "w:val": style.language }) : "") +
    "</w:rPr>"
  );
}
export function readWordMath(
  source: XmlNode,
  options: {
    display?: "inline" | "block";
    integralLimits?: string;
    naryLimits?: string;
  } = {},
): WordMathTree {
  const nodes: WordMathNode[] = [];
  let serial = 0;
  const add = (node: Input<WordMathNode>): string => {
    const id = "m" + ++serial;
    nodes.push({ ...node, mathId: id } as WordMathNode);
    return id;
  };
  const row = (node: XmlNode | undefined): string => {
    if (!node) fail(source, "A math argument is missing.");
    checkAttributes(node, []);
    const properties = child(node, "argPr", M),
      size = child(properties, "argSz", M),
      style = readStyle(child(child(node, "ctrlPr", M), "rPr"));
    check(properties, ["argSz"], M);
    return add({
      kind: "row",
      children: node.content.flatMap((item) =>
        typeof item === "string"
          ? item.trim()
            ? fail(node, "Unexpected text in a math argument.")
            : []
          : item.uri === M && ["argPr", "ctrlPr"].includes(item.name)
            ? []
            : [parse(item)],
      ),
      ...(size ? { argumentSize: number(attr(size, "val", M), size) } : {}),
      ...(style ? { style } : {}),
    });
  };
  const parse = (node: XmlNode): string => {
    if (node.uri !== M) unsupported(node, "This math argument remains native.");
    checkAttributes(node, []);
    const property = child(node, node.name + "Pr", M),
      control = child(property, "ctrlPr", M);
    check(control, ["rPr"], W);
    const style = readStyle(child(control, "rPr")),
      base = style ? { style } : {};
    if (node.name === "r") {
      for (const item of node.content)
        if (
          typeof item !== "string" &&
          !(
            (item.uri === M && ["rPr", "t"].includes(item.name)) ||
            (item.uri === W && item.name === "rPr")
          )
        )
          unsupported(item);
      const run = child(node, "rPr", M);
      check(run, ["sty", "scr", "nor", "lit", "brk", "aln"], M);
      if (child(run, "brk", M) || child(run, "aln", M) || flag(child(run, "lit", M)))
        unsupported(node, "Math alignment controls remain native.");
      const native = attr(child(run, "sty", M), "val", M) ?? "i",
        script = attr(child(run, "scr", M), "val", M),
        variant = flag(child(run, "nor", M))
          ? "normal"
          : script
            ? (
                {
                  "roman": "italic",
                  "script": "script",
                  "fraktur": "fraktur",
                  "double-struck": "doubleStruck",
                  "sans-serif": "sansSerif",
                  "monospace": "monospace",
                } as const
              )[script as "roman"]
            : ({ p: "normal", i: "italic", b: "bold", bi: "boldItalic" } as const)[native as "i"];
      if (!variant) unsupported(node, "This math script style remains native.");
      const values = elements(node, "t", M);
      for (const item of values)
        if (item.content.some((child) => typeof child !== "string"))
          fail(item, "Math text contains XML.");
      return add({
        kind: "text",
        text: values.map(text).join(""),
        variant,
        ...(readStyle(child(node, "rPr", W)) ? { style: readStyle(child(node, "rPr", W))! } : {}),
      });
    }
    const args = (names: readonly string[]): string[] =>
      names.map((name) => row(child(node, name, M)));
    const prop = (names: readonly string[]): void => {
      check(property, [...names, "ctrlPr"], M);
      if (property)
        for (const item of elements(property, undefined, M))
          if (item.name !== "ctrlPr") checkAttributes(item, ["val"], M);
    };
    switch (node.name) {
      case "f": {
        check(node, ["fPr", "num", "den"], M);
        prop(["type"]);
        const mode = attr(child(property, "type", M), "val", M) ?? "bar";
        if (!["bar", "noBar", "skw", "lin"].includes(mode)) unsupported(node);
        return add({
          ...base,
          kind: "fraction",
          children: args(["num", "den"]),
          mode: mode === "skw" ? "skewed" : mode === "lin" ? "linear" : (mode as "bar" | "noBar"),
        });
      }
      case "rad": {
        check(node, ["radPr", "deg", "e"], M);
        prop(["degHide"]);
        const e = row(child(node, "e", M)),
          degree = child(node, "deg", M);
        return add({
          ...base,
          kind: "radical",
          children: degree ? [e, row(degree)] : [e],
          hideDegree: flag(child(property, "degHide", M)),
        });
      }
      case "sSub":
      case "sSup":
      case "sSubSup":
      case "sPre": {
        const mode =
            node.name === "sSub"
              ? "sub"
              : node.name === "sSup"
                ? "sup"
                : node.name === "sPre"
                  ? "pre"
                  : "subsup",
          names =
            mode === "sub" ? ["e", "sub"] : mode === "sup" ? ["e", "sup"] : ["e", "sub", "sup"];
        check(node, [node.name + "Pr", ...names], M);
        prop([]);
        return add({ ...base, kind: "script", mode, children: args(names) });
      }
      case "d": {
        check(node, ["dPr", "e"], M);
        prop(["begChr", "endChr", "sepChr", "grow", "shp"]);
        const shape = attr(child(property, "shp", M), "val", M);
        if (shape && shape !== "centered")
          unsupported(node, "This delimiter layout remains native.");
        return add({
          ...base,
          kind: "fence",
          children: elements(node, "e", M).map(row),
          open: attr(child(property, "begChr", M), "val", M) ?? "(",
          close: attr(child(property, "endChr", M), "val", M) ?? ")",
          separator: attr(child(property, "sepChr", M), "val", M) ?? "|",
          grow: flag(child(property, "grow", M), true),
        });
      }
      case "m": {
        check(node, ["mPr", "mr"], M);
        check(property, ["baseJc", "plcHide", "mcs", "ctrlPr"], M);
        const alignment = attr(child(property, "baseJc", M), "val", M) ?? "center";
        if (!["top", "center", "bot"].includes(alignment)) unsupported(node);
        const columns: ("left" | "center" | "right")[] = [];
        for (const column of elements(child(property, "mcs", M), "mc", M)) {
          check(column, ["mcPr"], M);
          const pr = child(column, "mcPr", M);
          check(pr, ["count", "mcJc"], M);
          const count = number(attr(child(pr, "count", M), "val", M), column, 1),
            align = attr(child(pr, "mcJc", M), "val", M) ?? "center";
          if (!Number.isSafeInteger(count) || count < 1 || count > 255)
            fail(column, "The matrix column count is invalid.");
          if (!["left", "center", "right"].includes(align)) unsupported(node);
          for (let i = 0; i < count; i++) columns.push(align as "center");
        }
        return add({
          ...base,
          kind: "matrix",
          children: elements(node, "mr", M).map(parse),
          verticalAlignment: alignment === "bot" ? "bottom" : (alignment as "top" | "center"),
          hidePlaceholders: flag(child(property, "plcHide", M)),
          ...(columns.length ? { columnAlignments: columns } : {}),
        });
      }
      case "mr": {
        check(node, ["e"], M);
        return add({ kind: "matrixRow", children: elements(node, "e", M).map(row) });
      }
      case "nary": {
        check(node, ["naryPr", "sub", "sup", "e"], M);
        prop(["chr", "limLoc", "grow", "subHide", "supHide"]);
        return add({
          ...base,
          kind: "nary",
          children: args(["e", "sub", "sup"]),
          operator: attr(child(property, "chr", M), "val", M) ?? "∫",
          limits:
            (attr(child(property, "limLoc", M), "val", M) ??
              (/[∫∬∭∮∯∰]/u.test(attr(child(property, "chr", M), "val", M) ?? "∫")
                ? options.integralLimits
                : options.naryLimits) ??
              (options.display === "block" ? "undOvr" : "subSup")) === "subSup"
              ? "subSup"
              : "aboveBelow",
          hideLower: flag(child(property, "subHide", M)),
          hideUpper: flag(child(property, "supHide", M)),
        });
      }
      case "func": {
        check(node, ["funcPr", "fName", "e"], M);
        prop([]);
        return add({ ...base, kind: "function", children: args(["fName", "e"]) });
      }
      case "acc": {
        check(node, ["accPr", "e"], M);
        prop(["chr"]);
        return add({
          ...base,
          kind: "accent",
          children: args(["e"]),
          symbol: attr(child(property, "chr", M), "val", M) ?? "̂",
        });
      }
      case "bar": {
        check(node, ["barPr", "e"], M);
        prop(["pos"]);
        return add({
          ...base,
          kind: "bar",
          children: args(["e"]),
          position: attr(child(property, "pos", M), "val", M) === "bot" ? "bottom" : "top",
        });
      }
      case "limLow":
      case "limUpp": {
        check(node, [node.name + "Pr", "e", "lim"], M);
        prop([]);
        return add({
          ...base,
          kind: "limit",
          children: args(["e", "lim"]),
          position: node.name === "limLow" ? "lower" : "upper",
        });
      }
      case "eqArr": {
        check(node, ["eqArrPr", "e"], M);
        prop([]);
        return add({ ...base, kind: "equationArray", children: elements(node, "e", M).map(row) });
      }
      case "box": {
        check(node, ["boxPr", "e"], M);
        prop([]);
        return add({ ...base, kind: "box", children: args(["e"]) });
      }
      default:
        return unsupported(node, "This Office math object remains native.");
    }
  };
  try {
    if (!["oMath", "oMathPara"].includes(source.name) || source.uri !== M) unsupported(source);
    const groups = source.name === "oMathPara" ? elements(source, "oMath", M) : [source];
    if (!groups.length) unsupported(source, "An empty native math paragraph remains preserved.");
    const roots = groups.map(row),
      rootMathId = roots.length === 1 ? roots[0]! : add({ kind: "equationGroup", children: roots }),
      math = { rootMathId, nodes };
    try {
      validateWordMath(math);
    } catch (error) {
      if (error instanceof DocumentError)
        unsupported(source, "This equation is outside the editable math tree.");
      throw error;
    }
    return math;
  } catch (error) {
    if (error instanceof DocumentError && error.code === "UnsupportedFeature")
      unsupported(source, "This equation has additional native formatting or structure.");
    throw error;
  }
}
export function writeWordMath(
  equation: DeepReadonly<NonNullable<WordMetadata["equations"]>[number]>,
): string {
  const nodes = new Map(equation.nodes.map((node) => [node.mathId, node]));
  const write = (id: string, inherited: DeepReadonly<Style> = {}): string => {
    const node = nodes.get(id)!,
      children = "children" in node ? node.children : [],
      style = {
        ...inherited,
        ...node.style,
        ...(inherited.fonts || node.style?.fonts
          ? { fonts: { ...inherited.fonts, ...node.style?.fonts } }
          : {}),
      },
      arg = (name: string, id: string): string => {
        const argument = nodes.get(id)!;
        return (
          "<m:" +
          name +
          ">" +
          (argument.kind === "row" && argument.argumentSize !== undefined
            ? "<m:argPr>" + empty("m:argSz", { "m:val": argument.argumentSize }) + "</m:argPr>"
            : "") +
          write(id, style) +
          "</m:" +
          name +
          ">"
        );
      },
      property = (name: string, content = ""): string =>
        content || Object.keys(style).length
          ? "<m:" +
            name +
            "Pr>" +
            content +
            (Object.keys(style).length ? "<m:ctrlPr>" + writeStyle(style) + "</m:ctrlPr>" : "") +
            "</m:" +
            name +
            "Pr>"
          : "",
      val = (name: string, value: string | number): string =>
        empty("m:" + name, { "m:val": value });
    switch (node.kind) {
      case "equationGroup":
        return children.map((id) => write(id, style)).join("");
      case "row":
        return children.map((id) => write(id, style)).join("");
      case "text": {
        const scripts: Record<string, string> = {
            doubleStruck: "double-struck",
            script: "script",
            fraktur: "fraktur",
            sansSerif: "sans-serif",
            monospace: "monospace",
          },
          variantXml =
            node.variant && scripts[node.variant]
              ? val("scr", scripts[node.variant]!)
              : val(
                  "sty",
                  node.variant === "normal"
                    ? "p"
                    : node.variant === "bold"
                      ? "b"
                      : node.variant === "boldItalic"
                        ? "bi"
                        : "i",
                );
        return (
          "<m:r><m:rPr>" +
          variantXml +
          "</m:rPr>" +
          writeStyle(
            equation.fontFamily
              ? {
                  ...style,
                  fonts: { ascii: equation.fontFamily, hAnsi: equation.fontFamily, ...style.fonts },
                }
              : style,
          ) +
          '<m:t xml:space="preserve">' +
          escapeXml(node.text) +
          "</m:t></m:r>"
        );
      }
      case "fraction":
        return (
          "<m:f>" +
          property(
            "f",
            val(
              "type",
              node.mode === "skewed"
                ? "skw"
                : node.mode === "linear"
                  ? "lin"
                  : (node.mode ?? "bar"),
            ),
          ) +
          arg("num", children[0]!) +
          arg("den", children[1]!) +
          "</m:f>"
        );
      case "radical":
        return (
          "<m:rad>" +
          property("rad", val("degHide", (node.hideDegree ?? children.length === 1) ? 1 : 0)) +
          (children[1] ? arg("deg", children[1]) : "<m:deg/>") +
          arg("e", children[0]!) +
          "</m:rad>"
        );
      case "script": {
        const name =
            node.mode === "sub"
              ? "sSub"
              : node.mode === "sup"
                ? "sSup"
                : node.mode === "pre"
                  ? "sPre"
                  : "sSubSup",
          argumentsXml =
            node.mode === "pre"
              ? arg("sub", children[1]!) + arg("sup", children[2]!) + arg("e", children[0]!)
              : arg("e", children[0]!) +
                (node.mode === "sup"
                  ? arg("sup", children[1]!)
                  : arg("sub", children[1]!) + (children[2] ? arg("sup", children[2]) : ""));
        return "<m:" + name + ">" + property(name) + argumentsXml + "</m:" + name + ">";
      }
      case "fence":
        return (
          "<m:d>" +
          property(
            "d",
            val("begChr", node.open ?? "(") +
              val("sepChr", node.separator ?? "|") +
              val("endChr", node.close ?? ")") +
              val("grow", node.grow === false ? 0 : 1),
          ) +
          children.map((id) => arg("e", id)).join("") +
          "</m:d>"
        );
      case "matrix":
        return (
          "<m:m>" +
          property(
            "m",
            (node.verticalAlignment
              ? val("baseJc", node.verticalAlignment === "bottom" ? "bot" : node.verticalAlignment)
              : "") +
              (node.hidePlaceholders === undefined
                ? ""
                : val("plcHide", node.hidePlaceholders ? 1 : 0)) +
              (node.columnAlignments
                ? "<m:mcs>" +
                  node.columnAlignments
                    .map(
                      (align) =>
                        "<m:mc><m:mcPr>" +
                        val("count", 1) +
                        val("mcJc", align) +
                        "</m:mcPr></m:mc>",
                    )
                    .join("") +
                  "</m:mcs>"
                : ""),
          ) +
          children.map((id) => write(id, style)).join("") +
          "</m:m>"
        );
      case "matrixRow":
        return "<m:mr>" + children.map((id) => arg("e", id)).join("") + "</m:mr>";
      case "nary":
        return (
          "<m:nary>" +
          property(
            "nary",
            val("chr", node.operator) +
              val("limLoc", node.limits === "subSup" ? "subSup" : "undOvr") +
              val("subHide", node.hideLower ? 1 : 0) +
              val("supHide", node.hideUpper ? 1 : 0),
          ) +
          arg("sub", children[1]!) +
          arg("sup", children[2]!) +
          arg("e", children[0]!) +
          "</m:nary>"
        );
      case "function":
        return (
          "<m:func>" +
          property("func") +
          arg("fName", children[0]!) +
          arg("e", children[1]!) +
          "</m:func>"
        );
      case "accent":
        return (
          "<m:acc>" + property("acc", val("chr", node.symbol)) + arg("e", children[0]!) + "</m:acc>"
        );
      case "bar":
        return (
          "<m:bar>" +
          property("bar", val("pos", node.position === "bottom" ? "bot" : "top")) +
          arg("e", children[0]!) +
          "</m:bar>"
        );
      case "limit": {
        const name = node.position === "lower" ? "limLow" : "limUpp";
        return (
          "<m:" +
          name +
          ">" +
          property(name) +
          arg("e", children[0]!) +
          arg("lim", children[1]!) +
          "</m:" +
          name +
          ">"
        );
      }
      case "equationArray":
        return (
          "<m:eqArr>" +
          property("eqArr") +
          children.map((id) => arg("e", id)).join("") +
          "</m:eqArr>"
        );
      case "box":
        return "<m:box>" + property("box") + arg("e", children[0]!) + "</m:box>";
    }
  };
  const root = nodes.get(equation.rootMathId)!,
    base = {
      ...(equation.fontFamily
        ? { fonts: { ascii: equation.fontFamily, hAnsi: equation.fontFamily } }
        : {}),
      ...root.style,
    },
    roots = root.kind === "equationGroup" ? root.children : [equation.rootMathId];
  const math = roots
    .map(
      (id) => '<m:oMath xmlns:m="' + M + '" xmlns:w="' + W + '">' + write(id, base) + "</m:oMath>",
    )
    .join("");
  return equation.display === "block"
    ? '<m:oMathPara xmlns:m="' +
        M +
        '"><m:oMathParaPr><m:jc m:val="' +
        (equation.alignment ?? "centerGroup") +
        '"/></m:oMathParaPr>' +
        math +
        "</m:oMathPara>"
    : math;
}
