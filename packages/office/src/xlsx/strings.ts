import { DocumentError } from "@doffi/core";
import type { DocumentReadLimits } from "@doffi/core";
import { attr, check, elements, fail, S, text, unsupported } from "../ooxml/xml.js";
import type { XmlBudget, XmlNode } from "../ooxml/xml.js";
import { streamXml } from "../ooxml/xml-stream.js";
import type { ImportControl } from "../import-control.js";
import { attributes, integer, one, spreadsheetText, textHtml, value } from "./values.js";

export interface RichString { readonly html: string; readonly codeUnits: number }
export interface RichTextStyles { run(node: XmlNode): string | undefined }
export interface SharedStrings {
  readonly size: number;
  get(index: number, node: XmlNode): RichString;
  validateReferences(): void;
}
const XML = "http://www.w3.org/XML/1998/namespace";
export function spreadsheetStringText(node: XmlNode): string {
  for (const [key] of node.attrs) if (!key.startsWith("http://www.w3.org/2000/xmlns/ ") && key !== XML + " space") unsupported(node);
  const space = attr(node, "space", XML);
  if (space !== undefined && space !== "preserve" && space !== "default") fail(node, "Invalid native text whitespace mode.");
  if (node.content.some(item => typeof item !== "string")) fail(node, "Native text must not contain child elements.");
  return spreadsheetText(text(node), node);
}
/** Phonetic display settings can exist without any pronunciation runs. */
export function validatePhoneticProperties(node: XmlNode | undefined): void {
  if (!node) return;
  attributes(node, ["fontId", "type", "alignment"]); check(node, [], S);
  integer(node, "fontId");
  const type = value(node, "type"), alignment = value(node, "alignment");
  if (type !== undefined && !["halfwidthKatakana", "fullwidthKatakana", "Hiragana", "noConversion"].includes(type))
    fail(node, "A phonetic conversion type is invalid.");
  if (alignment !== undefined && !["noControl", "left", "center", "distributed"].includes(alignment))
    fail(node, "A phonetic alignment is invalid.");
}
export function richString(node: XmlNode, styles: RichTextStyles, nodeId?: () => string): RichString {
  attributes(node, []); check(node, ["t", "r", "phoneticPr"], S);
  validatePhoneticProperties(one(node, "phoneticPr"));
  const simple = one(node, "t"), runs = elements(node, "r", S);
  if (simple && runs.length) fail(node, "A native string mixes plain and rich text containers.");
  if (simple) { const source = spreadsheetStringText(simple); return { html: textHtml(source), codeUnits: source.length }; }
  let html = "", codeUnits = 0;
  for (const run of runs) {
    attributes(run, []); check(run, ["rPr", "t"], S);
    const source = one(run, "t"); if (!source) fail(run, "A rich text run has no text element.");
    const decoded = spreadsheetStringText(source), properties = one(run, "rPr"); codeUnits += decoded.length;
    if (codeUnits > 32767) throw new DocumentError("ResourceLimitExceeded", "A rich string exceeds its UTF-16 budget.");
    const cssClass = properties ? styles.run(properties) : undefined;
    html += cssClass ? `<span${nodeId ? ` data-node-id="${textHtml(nodeId())}"` : ""} class="${textHtml(cssClass)}">${textHtml(decoded)}</span>` : textHtml(decoded);
  }
  return { html, codeUnits };
}
export interface SharedStringOptions {
  readonly limits: DocumentReadLimits;
  readonly budget: XmlBudget;
  readonly maxRetainedNodes: number;
  readonly maxStrings: number;
  readonly maxCodeUnits: number;
  readonly control: ImportControl;
}
export async function readSharedStrings(bytes: Uint8Array | undefined, part: string, styles: RichTextStyles, options: SharedStringOptions): Promise<SharedStrings> {
  const strings: RichString[] = []; let units = 0, references = 0, declaredCount: number | undefined;
  if (bytes) {
    let root: XmlNode | undefined;
    const parsed = await streamXml(bytes, part, { limits: options.limits, budget: options.budget, maxRetainedNodes: options.maxRetainedNodes, maxTextCodeUnits: 32767 * 7,
      control: options.control, progress: (completed, total) => ({ phase: "metadata", completed, total }),
      ignoreWhitespace: node => node.uri === S && node.name !== "t",
      open(node, parent) { if (!parent) { if (node.uri !== S || node.name !== "sst") fail(node, "The shared strings root is invalid."); root = node; } },
      complete(node, parent) {
        if (parent === root && node.uri === S && node.name === "si") {
          const source = richString(node, styles); units += source.html.length;
          if (units > options.maxCodeUnits || strings.length >= options.maxStrings) throw new DocumentError("ResourceLimitExceeded", "Shared strings exceed their cumulative budget.");
          strings.push(source); return false;
        }
        return true;
      },
    });
    attributes(parsed, ["count", "uniqueCount"]); check(parsed, [], S);
    if (value(parsed, "uniqueCount") !== undefined && integer(parsed, "uniqueCount") !== strings.length) fail(parsed, "Shared string unique count is inconsistent.");
    if (value(parsed, "count") !== undefined) declaredCount = integer(parsed, "count");
  }
  return Object.freeze({ size: strings.length,
    get(index: number, node: XmlNode): RichString {
      if (!Number.isSafeInteger(index) || index < 0 || index >= strings.length) fail(node, "A cell references an unknown shared string.");
      references++; return strings[index]!;
    },
    validateReferences(): void {
      if (declaredCount !== undefined && references !== declaredCount) fail({ part, name: "sst" }, "Shared string reference count is inconsistent.");
    },
  });
}
