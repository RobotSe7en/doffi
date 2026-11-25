import type { DefaultTreeAdapterTypes } from "parse5";
import type { ContentDocumentManifest } from "../contracts/document.js";
import type {
  DeepReadonly,
  DocumentNode,
  DocumentPackage,
  DocumentReadLimits,
} from "../runtime.js";
import { DocumentError } from "../errors.js";
import { normalizeCss } from "./css.js";
import type { CssBudget } from "./css.js";
import { parseHtmlTree, serializeHtmlOuter, serializeHtmlTree } from "./html-parser.js";
import { validateImageAsset } from "./assets.js";
import { resolveReference } from "./package.js";
import { checkLimit, freeze, invalid, unsupported } from "./validation.js";
import { textPositions } from "./text.js";

type Element = DefaultTreeAdapterTypes.Element;
type Node = DefaultTreeAdapterTypes.Node;
export const bodyTags: ReadonlySet<string> = new Set([
  "article",
  "section",
  "div",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "p",
  "span",
  "strong",
  "em",
  "u",
  "s",
  "sub",
  "sup",
  "br",
  "ul",
  "ol",
  "li",
  "table",
  "thead",
  "tbody",
  "tfoot",
  "tr",
  "th",
  "td",
  "img",
  "ins",
  "del",
]);
const shellTags = new Set(["html", "head", "body", "meta", "title", "link"]);
export const supportedTags: ReadonlySet<string> = new Set([...bodyTags, ...shellTags]);
const structuralTags = new Set(["thead", "tbody", "tfoot", "br"]);
const containers = new Set(["article", "section", "div"]);
const globalAttributes = new Set(["id", "class", "title", "lang", "dir", "style"]);
const tagAttributes: Readonly<Record<string, readonly string[]>> = {
  img: ["src", "alt", "width", "height"],
  td: ["colspan", "rowspan"],
  th: ["colspan", "rowspan", "scope"],
  ol: ["start", "reversed", "type"],
  li: ["value"],
  meta: ["charset"],
  link: ["rel", "href"],
};

function isElement(node: Node): node is Element {
  return "tagName" in node;
}

function attribute(node: Element, name: string): string | undefined {
  return node.attrs.find((attr) => attr.name === name)?.value;
}

function textContent(node: Node): string {
  const parts: string[] = [];
  const stack = [node];
  while (stack.length) {
    const current = stack.pop()!;
    if (current.nodeName === "#text" && "value" in current) parts.push(current.value);
    else if (isElement(current) && current.tagName === "br") parts.push("\n");
    else if (isElement(current) && current.tagName === "img") parts.push("\ufffc");
    else if ("childNodes" in current) {
      for (let index = current.childNodes.length - 1; index >= 0; index--)
        stack.push(current.childNodes[index]!);
    }
  }
  return parts.join("");
}

export interface ParsedHtml {
  readonly html: string;
  directText(nodeId: string): string | undefined;
  nodeIds(): readonly string[];
  peekNode(nodeId: string): Omit<DocumentNode, "text" | "html"> | undefined;
  getNode(nodeId: string): DocumentNode | undefined;
  hasNode(nodeId: string): boolean;
  textOffset(nodeId: string): number | undefined;
  textLength(nodeId: string): number | undefined;
  textScope(nodeId: string, from: number, to: number): string | undefined;
  isContainer(nodeId: string): boolean;
  signature(nodeId: string): string | undefined;
}

/** The same scriptless HTML boundary serves document bodies and bounded
 * workbook object/comment content. Cell chunks retain their compact parser. */
export interface HtmlSourceDescriptor {
  readonly kind: "document" | "presentation" | "workbook-content";
  readonly entrypoint: string;
  readonly styles: readonly string[];
  readonly assets: DeepReadonly<ContentDocumentManifest["assets"]>;
}
export function parseHtml(
  source: string,
  manifest: HtmlSourceDescriptor,
  files: DocumentPackage,
  limits: DocumentReadLimits,
  cssBudget: CssBudget,
): ParsedHtml {
  try {
    const tree = parseHtmlTree(source, limits, supportedTags);
    const html = tree.childNodes.find((node) => isElement(node) && node.tagName === "html") as
      | Element
      | undefined;
    const head = html?.childNodes.find((node) => isElement(node) && node.tagName === "head") as
      | Element
      | undefined;
    const body = html?.childNodes.find((node) => isElement(node) && node.tagName === "body") as
      | Element
      | undefined;
    if (!html || !head || !body) invalid("The HTML document shell is missing.");
    const roots = body.childNodes.filter(isElement);
    const root = roots[0];
    if (
      roots.length !== 1 ||
      !root ||
      !containers.has(root.tagName) ||
      attribute(root, "data-office-kind") !== manifest.kind ||
      body.childNodes.some(
        (node) => node.nodeName === "#text" && "value" in node && node.value.trim() !== "",
      )
    ) {
      invalid("The body must contain one document root container.");
    }
    const index = new Map<
      string,
      { element: Element; parentId: string | null; children: string[]; signature?: string }
    >();
    const domIds = new Set<string>();
    const stylePaths: string[] = [];
    const assets = new Map(manifest.assets.map((asset) => [asset.path, asset]));
    const inlineStyles = new Map<string, string>();
    const stack: { node: Node; depth: number; inBody: boolean; parentId: string | null }[] = [
      { node: tree, depth: 0, inBody: false, parentId: null },
    ];
    let count = 0;
    let charsets = 0;
    let titles = 0;
    while (stack.length) {
      const { node, depth, inBody, parentId } = stack.pop()!;
      checkLimit(++count, limits.maxNodes, "maxNodes", "HTML");
      checkLimit(depth, limits.maxDepth, "maxDepth", "HTML");
      let childParentId = parentId;
      if (isElement(node)) {
        const tag = node.tagName;
        if (node.namespaceURI !== "http://www.w3.org/1999/xhtml")
          unsupported("Foreign HTML namespaces are not supported.");
        if (inBody ? !bodyTags.has(tag) : !shellTags.has(tag)) {
          unsupported("The HTML element is outside the supported document structure.");
        }
        if (["meta", "title", "link"].includes(tag) && node.parentNode !== head) {
          invalid("Document head elements must be inside head.");
        }
        for (const attr of node.attrs) {
          if (attr.namespace || attr.prefix || attr.name.startsWith("on"))
            invalid("Executable or namespaced HTML attributes are not allowed.");
          const extension = /^data-[a-z0-9-]+$/.test(attr.name) || /^aria-[a-z-]+$/.test(attr.name);
          if (
            !globalAttributes.has(attr.name) &&
            !extension &&
            !tagAttributes[tag]?.includes(attr.name)
          ) {
            unsupported("The HTML attribute is outside the supported document structure.");
          }
          if (attr.name === "style") {
            const source = attr.value;
            const cached = inlineStyles.get(source);
            if (cached !== undefined) {
              // An identical declaration list shares its parse, but each use costs one node.
              checkLimit(++cssBudget.nodes, limits.maxCssNodes, "maxCssNodes", "inline CSS");
              attr.value = cached;
            } else {
              attr.value = normalizeCss(source, "declarationList", limits, cssBudget);
              inlineStyles.set(source, attr.value);
            }
          }
          if (attr.name === "dir" && !["ltr", "rtl", "auto"].includes(attr.value))
            invalid("The HTML text direction is invalid.");
          if (
            ["width", "height", "colspan", "rowspan"].includes(attr.name) &&
            (!/^\d+$/.test(attr.value) ||
              !Number.isSafeInteger(Number(attr.value)) ||
              Number(attr.value) < (attr.name === "rowspan" ? 0 : 1))
          ) {
            invalid("HTML dimensions and table spans must be supported integer values.");
          }
          if (
            ["start", "value"].includes(attr.name) &&
            (!/^-?\d+$/.test(attr.value) || !Number.isSafeInteger(Number(attr.value)))
          ) {
            invalid("HTML list numbering must be a safe integer.");
          }
          if (attr.name === "scope" && !["row", "col", "rowgroup", "colgroup"].includes(attr.value))
            invalid("The table header scope is invalid.");
          if (attr.name === "type" && !["1", "a", "A", "i", "I"].includes(attr.value))
            invalid("The list numbering type is not supported.");
        }
        const domId = attribute(node, "id");
        if (domId !== undefined) {
          if (!domId || domIds.has(domId))
            invalid("HTML id attributes must be nonempty and unique.");
          domIds.add(domId);
        }
        const nodeId = attribute(node, "data-node-id");
        if (!inBody && nodeId !== undefined)
          invalid("Document node identities belong inside the body root.");
        if (
          inBody &&
          (manifest.kind === "workbook-content" || !structuralTags.has(tag)) &&
          nodeId === undefined
        )
          invalid("A document element is missing its node identity.");
        if (nodeId !== undefined) {
          if (!nodeId || index.has(nodeId))
            invalid("Document node identities must be nonempty and unique.");
          index.set(nodeId, { element: node, parentId, children: [] });
          if (parentId !== null) index.get(parentId)!.children.push(nodeId);
          childParentId = nodeId;
        }
        if (tag === "meta") {
          if (attribute(node, "charset")?.toLowerCase() !== "utf-8")
            unsupported("Only UTF-8 charset metadata is supported.");
          charsets++;
        }
        if (tag === "title") titles++;
        if (tag === "link") {
          const href = attribute(node, "href");
          if (attribute(node, "rel")?.toLowerCase() !== "stylesheet" || !href)
            unsupported("Only package stylesheet links are supported.");
          stylePaths.push(resolveReference(href, manifest.entrypoint));
        }
        if (tag === "img") {
          const src = attribute(node, "src");
          if (!src) invalid("Document images require a package resource reference.");
          const path = resolveReference(src, manifest.entrypoint);
          const asset = assets.get(path);
          if (!asset) invalid("An image refers to an undeclared package resource.");
          validateImageAsset(asset);
          if (!files.has(path))
            throw new DocumentError(
              "AssetUnavailable",
              "A referenced image is not supplied in the package.",
            );
        }
      } else if (node.nodeName === "#documentType" && "name" in node) {
        if (node.name !== "html" || node.publicId || node.systemId)
          unsupported("Only the HTML5 doctype is supported.");
      }
      if ("childNodes" in node) {
        for (let child = node.childNodes.length - 1; child >= 0; child--) {
          stack.push({
            node: node.childNodes[child]!,
            depth: depth + 1,
            inBody: inBody || node === body,
            parentId: childParentId,
          });
        }
      }
    }
    if (charsets !== 1 || titles > 1)
      invalid("The HTML head requires one UTF-8 charset and at most one title.");
    if (
      stylePaths.length !== manifest.styles.length ||
      stylePaths.some((path, order) => path !== manifest.styles[order])
    ) {
      invalid("HTML stylesheet links must match the manifest order exactly.");
    }
    let positions: ReturnType<typeof textPositions> | undefined;
    let scopes: { from: number; to: number; scope: string }[] | undefined;
    const inheritedScope = (element: Element): string => {
      for (
        let node: Node | null = element;
        node;
        node = "parentNode" in node ? node.parentNode : null
      ) {
        if (isElement(node) && attribute(node, "data-office-story") !== undefined) {
          const id = attribute(node, "data-node-id");
          if (id !== undefined) return id;
        }
      }
      return attribute(root, "data-node-id")!;
    };
    const readScopes = () => {
      if (scopes) return scopes;
      positions ??= textPositions(tree);
      scopes = [];
      const visit = (node: Node, scope: string): void => {
        if (isElement(node) && attribute(node, "data-office-story") !== undefined)
          scope = attribute(node, "data-node-id") ?? scope;
        if (
          node.nodeName === "#text" ||
          (isElement(node) && ["br", "img"].includes(node.tagName))
        ) {
          const range = positions!.get(node)!;
          if (range.from < range.to) scopes!.push({ ...range, scope });
        } else if ("childNodes" in node) for (const child of node.childNodes) visit(child, scope);
      };
      visit(root, inheritedScope(root));
      return scopes;
    };
    return {
      html: serializeHtmlTree(tree),
      directText: (nodeId) =>
        index
          .get(nodeId)
          ?.element.childNodes.filter(
            (node): node is DefaultTreeAdapterTypes.TextNode => node.nodeName === "#text",
          )
          .map((node) => node.value)
          .join(""),
      nodeIds: () => [...index.keys()],
      signature: (nodeId) => {
        const entry = index.get(nodeId);
        if (!entry) return undefined;
        if (entry.signature !== undefined) return entry.signature;
        const tokens = (node: Node): unknown => {
          if (isElement(node)) {
            const id = attribute(node, "data-node-id");
            if (node !== entry.element && id !== undefined) return { id };
            return [node.tagName, node.attrs, node.childNodes.map(tokens)];
          }
          return node.nodeName === "#text" && "value" in node
            ? node.value
            : node.nodeName === "#comment" && "data" in node
              ? { comment: node.data }
              : null;
        };
        return (entry.signature = JSON.stringify([entry.parentId, tokens(entry.element)]));
      },
      peekNode: (nodeId) => {
        const entry = index.get(nodeId);
        return entry
          ? {
              nodeId,
              tagName: entry.element.tagName,
              attributes: Object.fromEntries(
                entry.element.attrs.map((attr) => [attr.name, attr.value]),
              ),
              parentNodeId: entry.parentId,
              childNodeIds: entry.children,
            }
          : undefined;
      },
      textOffset: (nodeId) => {
        positions ??= textPositions(tree);
        const element = index.get(nodeId)?.element;
        return element ? positions.get(element)?.from : undefined;
      },
      textLength: (nodeId) => {
        positions ??= textPositions(tree);
        const element = index.get(nodeId)?.element,
          bounds = element ? positions.get(element) : undefined;
        return bounds ? bounds.to - bounds.from : undefined;
      },
      textScope: (nodeId, from, to) => {
        const element = index.get(nodeId)?.element;
        if (!element) return undefined;
        const runs = readScopes();
        const bounds = positions!.get(element)!;
        const offset = bounds.from;
        if (bounds.from === bounds.to) return inheritedScope(element);
        let low = 0;
        let high = runs.length;
        while (low < high) {
          const middle = Math.floor((low + high) / 2);
          if (runs[middle]!.to <= offset + from) low = middle + 1;
          else high = middle;
        }
        // Points follow the next source leaf, or the preceding leaf at this node's end.
        if (from === to) return runs[low - (offset + from === bounds.to ? 1 : 0)]?.scope;
        const scope = runs[low]?.scope;
        for (let index = low; index < runs.length && runs[index]!.from < offset + to; index++) {
          if (runs[index]!.scope !== scope) return undefined;
        }
        return scope;
      },
      hasNode: (nodeId) => index.has(nodeId),
      isContainer: (nodeId) => {
        const element = index.get(nodeId)?.element;
        return element !== undefined && containers.has(element.tagName);
      },
      getNode: (nodeId) => {
        const entry = index.get(nodeId);
        if (!entry) return undefined;
        return freeze({
          nodeId,
          tagName: entry.element.tagName,
          attributes: Object.fromEntries(
            entry.element.attrs.map((attr) => [attr.name, attr.value]),
          ),
          parentNodeId: entry.parentId,
          childNodeIds: [...entry.children],
          text: textContent(entry.element),
          html: serializeHtmlOuter(entry.element),
        });
      },
    };
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    return invalid("The HTML source could not be parsed or validated.");
  }
}
