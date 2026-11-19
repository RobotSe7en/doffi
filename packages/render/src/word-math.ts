import { officeFontSegments } from "./font-scripts.js";
import type { ContentDocumentModel, DeepReadonly, WordTextStyle } from "@doffi/core";
import { DocumentError } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { markViewNode, ownViewAttribute, ownViewStyle } from "./view-nodes.js";
const NS = "http://www.w3.org/1998/Math/MathML";
/** MathML is generated exclusively from typed math nodes. It is a derived view,
 * never accepted as authored HTML or a second editable document model. */
export function composeWordMath(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
  defaultFont: (stack: string) => string,
): void {
  let count = document.querySelectorAll("*").length;
  const make = (
    name: string,
    children: readonly Element[] = [],
    attributes: Readonly<Record<string, string>> = {},
  ): Element => {
    if (++count > limits.maxNodes)
      throw new DocumentError("ResourceLimitExceeded", "Math rendering exceeds its node budget.");
    const element = document.createElementNS(NS, name);
    for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, value);
    element.append(...children);
    return element;
  };
  for (const equation of model.snapshot().metadata.word?.equations ?? []) {
    const source = document.querySelector<HTMLElement>(
      '[data-node-id="' + CSS.escape(equation.nodeId) + '"]',
    );
    if (!source) continue;
    const nodes = new Map(equation.nodes.map((node) => [node.mathId, node])),
      original = document.defaultView!.getComputedStyle(source),
      size = original.fontSize;
    const render = (id: string, inherited: DeepReadonly<WordTextStyle> = {}): Element => {
      const node = nodes.get(id)!,
        style = {
          ...inherited,
          ...node.style,
          ...(inherited.fonts || node.style?.fonts
            ? { fonts: { ...inherited.fonts, ...node.style?.fonts } }
            : {}),
        },
        children = "children" in node ? node.children : [],
        childRender = (id: string): Element => render(id, style),
        child = (index: number): Element => childRender(children[index]!),
        operator = (value: string): Element => {
          const element = make("mo");
          element.textContent = value;
          return element;
        };
      let result: Element;
      switch (node.kind) {
        case "equationGroup":
          result = make(
            "mtable",
            children.map((id) => make("mtr", [make("mtd", [childRender(id)])])),
            {
              columnalign:
                equation.alignment === "left"
                  ? "left"
                  : equation.alignment === "right"
                    ? "right"
                    : "center",
              rowspacing: ".5em",
            },
          );
          break;
        case "row":
          result = make("mrow", children.map(childRender));
          if (node.argumentSize !== undefined)
            result = make("mstyle", [result], { scriptlevel: String(-node.argumentSize) });
          break;
        case "text": {
          const tag =
              /^\s|\s$/u.test(node.text) || (node.variant === "normal" && node.text.length > 1)
                ? "mtext"
                : /^\p{N}+$/u.test(node.text)
                  ? "mn"
                  : /^[\p{L}\p{M}\p{N}\u200d]+$/u.test(node.text)
                    ? "mi"
                    : "mo",
            variants: Readonly<Record<string, string>> = {
              normal: "normal",
              italic: "italic",
              bold: "bold",
              boldItalic: "bold-italic",
              doubleStruck: "double-struck",
              script: "script",
              fraktur: "fraktur",
              sansSerif: "sans-serif",
              monospace: "monospace",
            };
          const parts = officeFontSegments(node.text);
          const tokens = parts.map((part) => {
            const token = make(tag, [], { mathvariant: variants[node.variant ?? "italic"]! }),
              role =
                part.role === "east-asia"
                  ? "eastAsia"
                  : part.role === "h-ansi"
                    ? "hAnsi"
                    : part.role === "complex-script"
                      ? "cs"
                      : "ascii";
            token.textContent = part.text;
            token.setAttribute("data-word-script", role);
            const family = style.fonts?.[role];
            if (family) token.setAttribute("style", "font-family:" + JSON.stringify(family));
            return token;
          });
          result = tokens.length === 1 ? tokens[0]! : make("mrow", tokens);
          break;
        }
        case "fraction":
          result =
            node.mode === "linear"
              ? make("mrow", [child(0), operator("/"), child(1)])
              : make("mfrac", [child(0), child(1)], {
                  ...(node.mode === "noBar" ? { linethickness: "0" } : {}),
                  ...(node.mode === "skewed" ? { bevelled: "true" } : {}),
                });
          break;
        case "radical":
          result = make(
            children.length === 1 || node.hideDegree ? "msqrt" : "mroot",
            children.length === 1 || node.hideDegree ? [child(0)] : children.map(childRender),
          );
          break;
        case "script":
          result =
            node.mode === "pre"
              ? make("mmultiscripts", [child(0), make("mprescripts"), child(1), child(2)])
              : make(
                  node.mode === "sub" ? "msub" : node.mode === "sup" ? "msup" : "msubsup",
                  children.map(childRender),
                );
          break;
        case "fence": {
          const body: Element[] = [operator(node.open ?? "(")];
          for (const [index, id] of children.entries()) {
            if (index) body.push(operator(node.separator ?? "|"));
            body.push(childRender(id));
          }
          body.push(operator(node.close ?? ")"));
          for (const item of body)
            if (item.localName === "mo") item.setAttribute("stretchy", String(node.grow !== false));
          result = make("mrow", body);
          break;
        }
        case "matrix":
          result = make("mtable", children.map(childRender), {
            ...(node.columnAlignments ? { columnalign: node.columnAlignments.join(" ") } : {}),
            ...(node.verticalAlignment ? { align: node.verticalAlignment } : {}),
          });
          break;
        case "matrixRow":
          result = make(
            "mtr",
            children.map((id) => make("mtd", [childRender(id)])),
          );
          break;
        case "nary": {
          const sign = operator(node.operator);
          sign.setAttribute("largeop", "true");
          const lower = !node.hideLower,
            upper = !node.hideUpper,
            tag =
              node.limits === "subSup"
                ? lower && upper
                  ? "msubsup"
                  : lower
                    ? "msub"
                    : "msup"
                : lower && upper
                  ? "munderover"
                  : lower
                    ? "munder"
                    : "mover",
            limits =
              lower && upper
                ? [sign, child(1), child(2)]
                : lower
                  ? [sign, child(1)]
                  : upper
                    ? [sign, child(2)]
                    : [sign];
          result = make("mrow", [lower || upper ? make(tag, limits) : sign, child(0)]);
          break;
        }
        case "function":
          result = make("mrow", [child(0), operator("⁡"), child(1)]);
          break;
        case "accent":
          result = make("mover", [child(0), operator(node.symbol)], { accent: "true" });
          break;
        case "bar":
          result = make(
            node.position === "bottom" ? "munder" : "mover",
            [child(0), operator("¯")],
            node.position === "bottom" ? { accentunder: "true" } : { accent: "true" },
          );
          break;
        case "limit":
          result = make(node.position === "lower" ? "munder" : "mover", children.map(childRender));
          break;
        case "equationArray":
          result = make(
            "mtable",
            children.map((id) => make("mtr", [make("mtd", [childRender(id)])])),
          );
          break;
        case "box":
          result = make("mrow", children.map(childRender));
          break;
      }
      applyStyle(result, node.style, inherited.fontSize ?? parseFloat(size) * 0.75);
      return result;
    };
    const rendered = render(equation.rootMathId);
    if (
      equation.nodes.every((node) => node.kind === "row" || (node.kind === "text" && !node.text))
    ) {
      const placeholder = make("mtext");
      placeholder.textContent = "□";
      rendered.append(placeholder);
    }
    const math = markViewNode(
      make("math", [rendered], { display: equation.display === "block" ? "block" : "inline" }),
    );
    math.classList.add("sdk-word-equation-view");
    math.setAttribute(
      "style",
      "font-family:" +
        (equation.fontFamily ? JSON.stringify(equation.fontFamily) : defaultFont('"Cambria Math"')) +
        ";font-size:" +
        size +
        ";line-height:normal;color:" +
        original.color,
    );
    source.append(math);
    ownViewStyle(source, "font-size", "display", "line-height", "vertical-align");
    source.style.fontSize = "0";
    source.style.lineHeight = "normal";
    source.style.display = equation.display === "block" ? "block" : "inline-block";
    source.style.verticalAlign = "middle";
    if (equation.display === "block") {
      ownViewStyle(source, "text-align");
      source.style.textAlign =
        equation.alignment === "left"
          ? "left"
          : equation.alignment === "right"
            ? "right"
            : "center";
      math.setAttribute("style", math.getAttribute("style") + ";display:inline-block");
    }
    ownViewAttribute(source, "contenteditable");
    source.contentEditable = "false";
  }
}
function applyStyle(
  element: Element,
  style: DeepReadonly<WordTextStyle> | undefined,
  inheritedSize: number,
): void {
  if (!style) return;
  const values: string[] = [];
  if (style.fonts && !element.hasAttribute("data-word-script"))
    values.push(
      "font-family:" +
        Object.values(style.fonts)
          .map((value) => JSON.stringify(value))
          .join(","),
    );
  if (style.fontSize) values.push("font-size:" + style.fontSize / inheritedSize + "em");
  if (style.color && style.color !== "auto") values.push("color:" + style.color);
  if (style.bold !== undefined) values.push("font-weight:" + (style.bold ? "bold" : "normal"));
  if (style.italic !== undefined) values.push("font-style:" + (style.italic ? "italic" : "normal"));
  element.setAttribute("style", (element.getAttribute("style") ?? "") + ";" + values.join(";"));
  if (style.language) element.setAttribute("lang", style.language);
}
