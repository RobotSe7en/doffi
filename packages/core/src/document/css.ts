import { fork, generate, ident, lexer, walk } from "css-tree";
import type { CssNode, Declaration } from "css-tree";
import type { DocumentReadLimits } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { checkLimit, invalid, unsupported } from "./validation.js";

const properties = new Set([
  "color",
  "background-color",
  "font-family",
  "font-size",
  "font-weight",
  "font-style",
  "text-decoration",
  "text-decoration-line",
  "text-decoration-color",
  "text-decoration-style",
  "text-align",
  "text-indent",
  "line-height",
  "letter-spacing",
  "word-spacing",
  "white-space",
  "vertical-align",
  "width",
  "height",
  "min-width",
  "max-width",
  "min-height",
  "max-height",
  "widows",
  "orphans",
  "table-layout",
  "list-style-type",
  "list-style-position",
  "break-before",
  "break-after",
  "break-inside",
  "text-transform",
  "font-variant",
  "tab-size",
  "overflow-wrap",
  "margin",
  "margin-top",
  "margin-right",
  "margin-bottom",
  "margin-left",
  "padding",
  "padding-top",
  "padding-right",
  "padding-bottom",
  "padding-left",
  "border",
  "border-width",
  "border-style",
  "border-color",
  "border-collapse",
  "border-spacing",
  ...["top", "right", "bottom", "left"].flatMap((side) => [
    "border-" + side,
    "border-" + side + "-width",
    "border-" + side + "-style",
    "border-" + side + "-color",
  ]),
]);
const functions = new Set(["var", "calc", "min", "max", "clamp", "rgb", "rgba", "hsl", "hsla"]);

export interface CssBudget {
  nodes: number;
}

// CSS Tree's public syntax fork lets us guard each node parser before it
// allocates children. Its parser is synchronous; the per-call state is released
// even on failure and never retained between documents.
const parseBoundedCss = (() => {
  let active:
    | { nodes: number; depth: number; limits: DocumentReadLimits; parsed: WeakSet<CssNode> }
    | undefined;
  type NodeParser = (this: unknown, ...args: unknown[]) => CssNode;
  const syntax = fork((config) => {
    for (const definition of Object.values(config.node ?? {})) {
      const node = definition as { parse?: NodeParser };
      const parse = node.parse;
      if (!parse) continue;
      node.parse = function (...args) {
        const state = active!;
        checkLimit(++state.nodes, state.limits.maxCssNodes, "maxCssNodes", "CSS");
        checkLimit(++state.depth, state.limits.maxDepth, "maxDepth", "CSS");
        try {
          const result = parse.apply(this, args);
          // Some parsers also create nodes directly, such as an attribute's
          // name or whitespace. Count these without revisiting parsed subtrees.
          walk(result, function (child) {
            if (child === result) return undefined;
            if (state.parsed.has(child)) return this.skip;
            checkLimit(++state.nodes, state.limits.maxCssNodes, "maxCssNodes", "CSS");
            return undefined;
          });
          state.parsed.add(result);
          return result;
        } finally {
          state.depth--;
        }
      };
    }
    return config;
  });
  return (
    source: string,
    context: "stylesheet" | "declarationList",
    limits: DocumentReadLimits,
    budget: CssBudget,
  ): CssNode => {
    active = { nodes: budget.nodes, depth: 0, limits, parsed: new WeakSet() };
    try {
      return syntax.parse(source, {
        context,
        parseCustomProperty: true,
        onParseError: (error) => {
          if (error instanceof DocumentError) throw error;
          invalid("The CSS source is malformed.");
        },
      });
    } finally {
      active = undefined;
    }
  };
})();

function validateDeclaration(node: Declaration): void {
  const property = ident.decode(node.property);
  if (property === "--") invalid("A CSS custom property must have a name.");
  if (!property.startsWith("--") && !properties.has(property.toLowerCase())) {
    unsupported("The CSS property is outside the supported document style set.");
  }
  let hasVariable = false;
  walk(node.value, (child) => {
    if (child.type === "Url" || child.type === "Raw") {
      invalid("CSS resources and unparsed values are not supported.");
    }
    if (child.type === "Function") {
      const name = ident.decode(child.name).toLowerCase();
      if (!functions.has(name)) unsupported("The CSS function is not supported.");
      if (name === "var") {
        const first = child.children.first;
        if (
          first?.type !== "Identifier" ||
          !ident.decode(first.name).startsWith("--") ||
          ident.decode(first.name).length <= 2
        ) {
          invalid("CSS var() must reference a custom property.");
        }
        hasVariable = true;
      }
    }
  });
  if (
    !property.startsWith("--") &&
    !hasVariable &&
    lexer.matchProperty(property, node.value).error
  ) {
    invalid("A CSS declaration value does not match its property.");
  }
}

/** Parse and validate without resolving the cascade or materializing computed styles. */
export function normalizeCss(
  source: string,
  context: "stylesheet" | "declarationList",
  limits: DocumentReadLimits,
  budget: CssBudget,
): string {
  try {
    const ast = parseBoundedCss(source, context, limits, budget);
    let depth = 0;
    walk(ast, {
      enter(node: CssNode) {
        checkLimit(++budget.nodes, limits.maxCssNodes, "maxCssNodes", "CSS");
        checkLimit(++depth, limits.maxDepth, "maxDepth", "CSS");
        switch (node.type) {
          case "Atrule":
          case "IdSelector":
          case "PseudoElementSelector":
            unsupported("The CSS rule or selector is outside the supported style set.");
            break;
          case "Raw":
          case "Url":
            invalid("CSS contains unparsed content or a resource reference.");
            break;
          case "TypeSelector":
            if (!/^[a-z][a-z0-9-]*$/i.test(ident.decode(node.name)))
              unsupported("The CSS type selector is not supported.");
            break;
          case "Combinator":
            if (node.name !== " ") unsupported("Only descendant CSS combinators are supported.");
            break;
          case "PseudoClassSelector":
            if (ident.decode(node.name).toLowerCase() !== "root" || node.children !== null) {
              unsupported("Only the :root CSS pseudo-class is supported.");
            }
            break;
          case "AttributeSelector":
            if (
              !/^data-[a-z0-9-]+$/.test(ident.decode(node.name.name)) ||
              (node.matcher !== null && node.matcher !== "=") ||
              node.flags !== null
            ) {
              unsupported(
                "Only document data-attribute presence and equality selectors are supported.",
              );
            }
            break;
          case "Declaration":
            validateDeclaration(node);
            break;
          case "Rule":
            if (node.prelude?.type !== "SelectorList")
              unsupported("Only ordinary CSS style rules are supported.");
            if (node.block.children.some((child) => child.type !== "Declaration")) {
              unsupported("Nested CSS rules are not supported.");
            }
            break;
          case "StyleSheet":
            if (node.children.some((child) => child.type !== "Rule"))
              unsupported("Only CSS style rules are supported.");
            break;
          case "DeclarationList":
            if (node.children.some((child) => child.type !== "Declaration"))
              invalid("An inline style must contain only declarations.");
            break;
        }
      },
      leave() {
        depth--;
      },
    });
    return generate(ast);
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    return invalid("The CSS source could not be parsed or validated.");
  }
}

/** Set exactly one declaration, without allowing property/value text to inject siblings. */
export function removeInlineStyle(
  source: string,
  property: string,
  limits: DocumentReadLimits,
): string {
  const ast = parseBoundedCss(source, "declarationList", limits, { nodes: 0 });
  if (ast.type !== "DeclarationList")
    throw new DocumentError("InvalidOperation", "The inline style is invalid.");
  const name = property.startsWith("--") ? property : property.toLowerCase();
  ast.children.forEach((node, item, list) => {
    if (node.type === "Declaration") {
      const key = ident.decode(node.property);
      if ((key.startsWith("--") ? key : key.toLowerCase()) === name) list.remove(item);
    }
  });
  return normalizeCss(generate(ast), "declarationList", limits, { nodes: 0 });
}
export function setInlineStyle(
  source: string,
  property: string,
  value: string,
  limits: DocumentReadLimits,
): string {
  if (!/^(?:--[a-zA-Z0-9_-]+|[a-zA-Z][a-zA-Z-]*)$/.test(property)) {
    throw new DocumentError(
      "InvalidOperation",
      "A style operation requires one CSS property name.",
    );
  }
  const name = property.startsWith("--") ? property : property.toLowerCase();
  const normalized = normalizeCss(name + ":" + value, "declarationList", limits, { nodes: 0 });
  const incoming = parseBoundedCss(normalized, "declarationList", limits, { nodes: 0 });
  if (
    incoming.type !== "DeclarationList" ||
    incoming.children.size !== 1 ||
    incoming.children.first?.type !== "Declaration" ||
    incoming.children.first.property !== name
  ) {
    throw new DocumentError(
      "InvalidOperation",
      "A style operation must contain exactly one declaration.",
    );
  }
  const ast = parseBoundedCss(source, "declarationList", limits, { nodes: 0 });
  if (ast.type !== "DeclarationList")
    throw new DocumentError("InvalidOperation", "The inline style is invalid.");
  const declaration = incoming.children.first;
  // Append after existing shorthands so the explicitly set longhand takes effect.
  ast.children.forEach((node, item, list) => {
    if (node.type !== "Declaration") return;
    const existing = ident.decode(node.property);
    if ((existing.startsWith("--") ? existing : existing.toLowerCase()) !== name) return;
    list.remove(item);
  });
  ast.children.appendData(declaration);
  return normalizeCss(generate(ast), "declarationList", limits, { nodes: 0 });
}
