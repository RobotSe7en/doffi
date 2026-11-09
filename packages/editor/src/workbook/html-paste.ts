import { copyWorkbookFormula, DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, parseWorkbookFormula, parseWorkbookR1C1Formula, printWorkbookFormula, SourceStyleResolver, WorkbookCoordinateSchema, WorkbookValueSchema } from "@doffi/core";
import type { DocumentTransaction, SourceStyleNode, WorkbookCellInput, WorkbookCoordinate, WorkbookHyperlink, WorkbookRange, WorkbookSnapshot } from "@doffi/core";
import { defaultTreeAdapter, parse as parseHtml } from "parse5";
import type { DefaultTreeAdapterTypes } from "parse5";
import { generate, parse } from "css-tree";
import type { CssNode } from "css-tree";
import { parseWorkbookTextInput } from "./text-input.js";
import type { WorkbookPasteOptions } from "./text-input.js";

type Node = DefaultTreeAdapterTypes.Node;
type Element = DefaultTreeAdapterTypes.Element;
type Properties = Record<string, string>;
interface Run { from: number; to: number; properties: Properties }
interface Cell { at: WorkbookCoordinate; element: Element; text: string; runs: Run[]; properties: Properties; numberFormat?: string; hyperlink?: Omit<WorkbookHyperlink, "range"> }
export interface WorkbookHtmlPastePlan {
  readonly transaction: DocumentTransaction;
  readonly range: WorkbookRange;
  /** Bounded reports for presentation declarations excluded from cell input. */
  readonly diagnostics: readonly string[];
}
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
function limit(message: string): never { throw new DocumentError("ResourceLimitExceeded", message); }
const isElement = (node: Node): node is Element => "tagName" in node;
const children = (node: Node): readonly Node[] => "childNodes" in node ? node.childNodes : [];
const attributes = (node: Element): Record<string, string> => Object.fromEntries(node.attrs.map(attr => [attr.prefix ? attr.prefix + ":" + attr.name : attr.name, attr.value]));
const escape = (text: string): string => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;");
const textProperties = new Set(["font-family", "font-size", "font-weight", "font-style", "font-variant", "color", "text-decoration-line", "text-decoration-style", "text-decoration-color", "vertical-align", "letter-spacing", "word-spacing",
  ...["ascii", "h-ansi", "east-asia", "complex-script", "underline", "strike", "decoration-reset"].map(name => "--office-font-" + name)]);
const supported = (name: string): boolean => textProperties.has(name) || /^(?:--office-cell-(?:reading-order|vertical)|font|text-decoration|background-color|text-align|text-indent|white-space|overflow-wrap|line-height|border(?:-(?:top|right|bottom|left))?(?:-(?:width|style|color))?)$/.test(name);
const tags = new Set(["html", "head", "body", "title", "meta", "style", "table", "colgroup", "col", "thead", "tbody", "tfoot", "tr", "td", "th", "p", "div", "span", "br", "b", "strong", "i", "em", "u", "s", "strike", "sub", "sup", "font", "a", "pre"]);
const inherited = (name: string): boolean => textProperties.has(name) && name !== "vertical-align" || ["text-align", "white-space", "line-height", "--office-cell-reading-order"].includes(name);

/** Read clipboard HTML as bounded data in Node or the owning Worker. Original
 * nodes, attributes and styles are never attached to an editor DOM. */
export function createWorkbookHtmlPastePlan(snapshot: WorkbookSnapshot, sheetId: string, origin: WorkbookCoordinate, html: string, options: WorkbookPasteOptions = {}): WorkbookHtmlPastePlan {
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["mode", "maxCells", "maxCodeUnits"].includes(key))) invalid("Invalid workbook HTML paste options.");
  const maxCells = options.maxCells ?? 25_000, maxUnits = options.maxCodeUnits ?? 2 * 1024 * 1024;
  if (![maxCells, maxUnits].every(value => Number.isSafeInteger(value) && value > 0) || typeof html !== "string" || !WorkbookCoordinateSchema.safeParse(origin).success
    || options.mode !== undefined && !["auto", "text", "number"].includes(options.mode)
    || !snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === sheetId)) invalid("HTML paste requires valid budgets and a coordinate on an existing worksheet.");
  if (html.length > maxUnits) limit("Clipboard HTML exceeds its input budget.");
  for (const character of html) { const value = character.codePointAt(0)!; if (!value || value >= 0xd800 && value <= 0xdfff) invalid("Clipboard HTML requires well-formed Unicode without null characters."); }
  let created = 0, insertWork = 0;
  const depths = new WeakMap<Node, number>();
  const insertion = (parent: Node, node: Node): void => {
    const pending = [{ node, depth: (depths.get(parent) ?? 0) + 1 }];
    while (pending.length) {
      const item = pending.pop()!;
      if (++insertWork > 1_000_000 || item.depth > 128) limit("Clipboard HTML construction exceeds its depth or work budget.");
      depths.set(item.node, item.depth);
      for (const child of children(item.node)) pending.push({ node: child, depth: item.depth + 1 });
    }
  };
  const count = (): void => { if (++created > 100_000) limit("Clipboard HTML construction exceeds its node budget."); };
  const fragment = parseHtml(html, { treeAdapter: {
    ...defaultTreeAdapter,
    createElement(...args) { count(); return defaultTreeAdapter.createElement(...args); },
    createCommentNode(...args) { count(); return defaultTreeAdapter.createCommentNode(...args); },
    appendChild(parent, node) { insertion(parent, node); defaultTreeAdapter.appendChild(parent, node); },
    insertBefore(parent, node, reference) { insertion(parent, node); defaultTreeAdapter.insertBefore(parent, node, reference); },
  }, onParseError(error) {
    if (/surrogate|outside-unicode-range|null-character/.test(error.code)) invalid("Clipboard HTML contains an invalid Unicode character reference.");
  } });
  const all: Element[] = [], tables: Element[] = [], styles: string[] = [], projections = new Map<Element, SourceStyleNode>(), formats = new Map<Element, string>();
  const diagnostics = new Set<string>(); let nodes = 0, cssNodes = 0;
  const report = (message: string): void => { if (diagnostics.size < 32) diagnostics.add(message.slice(0, 1024)); };
  const declarations = (source: string): { css: string; format?: string } => {
    let ast: CssNode;
    try { ast = parse(source, { context: "declarationList" }); } catch { return invalid("Clipboard CSS contains invalid declarations."); }
    if (ast.type !== "DeclarationList") return invalid("Clipboard CSS is not a declaration list.");
    const output: string[] = []; let format: string | undefined;
    for (const item of ast.children) {
      if (++cssNodes > 100_000) limit("Clipboard CSS exceeds its node budget.");
      if (item.type !== "Declaration") invalid("Clipboard CSS contains an unsupported declaration.");
      const name = item.property.toLowerCase(), value = generate(item.value);
      if (name === "mso-number-format") {
        const valueNode = item.value.type === "Value" ? item.value.children.first : undefined;
        format = valueNode?.type === "String" ? valueNode.value : value;
      } else if (supported(name)) {
        if (/\b(?:url|var|expression)\s*\(/i.test(value)) invalid("Clipboard styles cannot reference external resources or source variables.");
        output.push(generate(item));
      } else report("Ignored unsupported clipboard style: " + name);
    }
    return { css: output.join(";"), ...(format !== undefined ? { format } : {}) };
  };
  const visit = (node: Node, depth: number, parent?: SourceStyleNode): void => {
    if (++nodes > 100_000 || depth > 128) limit("Clipboard HTML exceeds its node or depth budget.");
    let next = parent;
    if (isElement(node)) {
      if (!tags.has(node.tagName) || node.namespaceURI !== "http://www.w3.org/1999/xhtml") invalid("Clipboard HTML contains unsupported content: " + node.tagName);
      const attrs = attributes(node);
      if (Object.keys(attrs).some(name => /^on/i.test(name) || ["src", "srcset", "action", "formaction", "background"].includes(name))) invalid("Clipboard HTML contains executable attributes or unauthorized resources.");
      if (attrs.href && (node.tagName !== "a" || !/^(?:https?:|mailto:|#)/i.test(attrs.href))) invalid("Clipboard hyperlinks require an explicit supported URL or document location.");
      if (node.tagName === "table") tables.push(node);
      if (node.tagName === "style") styles.push(children(node).map(child => "value" in child ? child.value : "").join(""));
      const authored = declarations(attrs.style ?? "");
      if (authored.format !== undefined) formats.set(node, authored.format);
      const hints: string[] = [];
      if (attrs.align) hints.push("text-align:" + attrs.align);
      if (attrs.bgcolor) hints.push("background-color:" + attrs.bgcolor);
      if (attrs.dir && ["ltr", "rtl"].includes(attrs.dir)) hints.push("--office-cell-reading-order:" + attrs.dir);
      if (attrs.valign && ["top", "middle", "bottom"].includes(attrs.valign)) hints.push("--office-cell-vertical:" + (attrs.valign === "middle" ? "center" : attrs.valign));
      if (node.tagName === "font") {
        if (attrs.face) hints.push(declarations("font-family:" + attrs.face).css);
        if (attrs.color) hints.push("color:" + attrs.color);
        if (attrs.size) { const size = Number(attrs.size); if (Number.isInteger(size) && size >= 1 && size <= 7) hints.push("font-size:" + [8, 10, 12, 14, 18, 24, 36][size - 1] + "pt"); else report("Ignored relative or unsupported HTML font size."); }
      }
      const style = [...hints, authored.css].filter(Boolean).join(";");
      next = { tag: node.tagName, attrs: { ...(attrs.class ? { class: attrs.class } : {}), ...(style ? { style } : {}) }, parent };
      projections.set(node, next); all.push(node);
    }
    for (const child of children(node)) visit(child, depth + 1, next);
  };
  visit(fragment, 0);
  if (tables.length > 1) invalid("Clipboard HTML requires one table without nested tables.");
  if (tables[0]) {
    const checkOutside = (node: Node): void => {
      if (node === tables[0] || isElement(node) && ["head", "style", "meta", "title"].includes(node.tagName)) return;
      if (node.nodeName === "#text" && "value" in node && node.value.trim()) invalid("Clipboard HTML contains text outside its single table.");
      for (const child of children(node)) checkOutside(child);
    };
    checkOutside(fragment);
  }
  const ruleFormats: { selector: string; code: string }[] = [];
  const cleanStyles = styles.map(source => {
    const ast = parse(source, { context: "stylesheet" }); if (ast.type !== "StyleSheet") return invalid("Invalid clipboard stylesheet.");
    const output: string[] = [];
    for (const rule of ast.children) {
      if (++cssNodes > 100_000) limit("Clipboard CSS exceeds its node budget.");
      if (rule.type !== "Rule" || rule.prelude?.type !== "SelectorList") { report("Ignored unsupported clipboard CSS rule."); continue; }
      const authored = declarations(rule.block.children.toArray().map(item => generate(item)).join(";"));
      for (const selector of rule.prelude.children) {
        const text = generate(selector);
        if (!/^(?:[a-zA-Z][\w-]*|\.[\w-]+)(?:(?:\s+|\.)[\w-]+)*$/.test(text) || text.length > 256) { report("Ignored unsupported clipboard selector."); continue; }
        if (authored.css) output.push(text + "{" + authored.css + "}");
        if (authored.format !== undefined) ruleFormats.push({ selector: text, code: authored.format });
      }
    }
    return output.join("");
  });
  const resolver = new SourceStyleResolver(["b,strong{font-weight:bold}i,em{font-style:italic}u{text-decoration-line:underline}s,strike{text-decoration-line:line-through}sup{vertical-align:super}sub{vertical-align:sub}pre{white-space:pre-wrap}", ...cleanStyles], DEFAULT_DOCUMENT_READ_LIMITS);
  const computed = new Map<SourceStyleNode, Properties>();
  const properties = (node: SourceStyleNode): Properties => {
    const cached = computed.get(node); if (cached) return cached;
    const parent = node.parent ? properties(node.parent) : {}, result: Properties = Object.fromEntries(Object.entries(parent).filter(([name]) => inherited(name)));
    const own = resolver.specified(node);
    for (const [name, value] of Object.entries(own)) {
      if (!supported(name)) continue;
      if (["inherit", "unset"].includes(value)) { if (parent[name]) result[name] = parent[name]!; else delete result[name]; }
      else if (["initial", "revert"].includes(value)) delete result[name];
      else result[name] = value;
    }
    const decorations = [...new Set([parent["text-decoration-line"], result["text-decoration-line"]].flatMap(value => value?.split(/\s+/).filter(value => value !== "none") ?? []))];
    if (own["--office-font-decoration-reset"] !== undefined) result["text-decoration-line"] = resolver.get(node)["text-decoration-line"]!;
    else if (decorations.length) result["text-decoration-line"] = decorations.join(" ");
    computed.set(node, result); return result;
  };
  const cells: Cell[] = [], merges: WorkbookRange[] = [], occupied = new Set<string>();
  let height = 1, width = 1, work = 0;
  const add = (element: Element, row: number, column: number): void => {
    const cell: Cell = { at: { row: origin.row + row, column: origin.column + column }, element, text: "", runs: [], properties: { ...properties(projections.get(element)!) } };
    let blockBreak = false;
    const hyperlinks = new Map<string, Omit<WorkbookHyperlink, "range">>();
    const collect = (node: Node, parent: SourceStyleNode): void => {
      const current = isElement(node) ? projections.get(node)! : parent;
      if (isElement(node) && ["style", "meta", "title"].includes(node.tagName)) return;
      if (isElement(node) && node.tagName === "a") {
        const attrs = attributes(node), href = attrs.href;
        if (href) hyperlinks.set(href, { ...(href.startsWith("#") ? { location: href.slice(1) } : { url: href }), ...(attrs.title ? { tooltip: attrs.title } : {}) });
      }
      const block = isElement(node) && ["p", "div"].includes(node.tagName);
      if (block && cell.text) blockBreak = true;
      const text = node.nodeName === "#text" && "value" in node ? node.value : isElement(node) && node.tagName === "br" ? "\n" : "";
      if (text) {
        if (blockBreak && cell.text && !cell.text.endsWith("\n")) cell.text += "\n";
        blockBreak = false;
        const from = cell.text.length; cell.text += text;
        if (cell.text.length > 32767) limit("A clipboard HTML cell exceeds its text budget.");
        const style = Object.fromEntries(Object.entries(properties(current)).filter(([name, value]) => textProperties.has(name) && value !== cell.properties[name]));
        if (Object.keys(style).length) {
          const previous = cell.runs.at(-1);
          if (previous?.to === from && JSON.stringify(previous.properties) === JSON.stringify(style)) previous.to = cell.text.length;
          else cell.runs.push({ from, to: cell.text.length, properties: style });
        }
      }
      for (const child of children(node)) collect(child, current);
      if (block) blockBreak = true;
    };
    collect(element, projections.get(element)!);
    if (cell.text.length > 32767) limit("A clipboard HTML cell exceeds its text budget.");
    // A full-cell font wrapper styles numeric/Boolean cells too. Promote it
    // before choosing a source type, and avoid one text operation per row.
    if (cell.runs.length === 1 && cell.runs[0]!.from === 0 && cell.runs[0]!.to === cell.text.length) {
      Object.assign(cell.properties, cell.runs[0]!.properties); cell.runs.length = 0;
    }
    const inlineFormat = formats.get(element);
    if (inlineFormat !== undefined) cell.numberFormat = inlineFormat;
    if (hyperlinks.size > 1) invalid("A clipboard cell cannot represent multiple different hyperlink targets.");
    const hyperlink = hyperlinks.values().next().value; if (hyperlink) cell.hyperlink = hyperlink;
    const nativeNumber = attributes(element).sdnum;
    if (cell.numberFormat === undefined && nativeNumber) {
      const code = /^[^;]*;[^;]*;([\s\S]+)$/.exec(nativeNumber)?.[1];
      if (code && code !== "BOOLEAN") cell.numberFormat = code;
    }
    const classNames = attributes(element).class?.split(/\s+/) ?? [];
    for (const rule of ruleFormats) if (rule.selector === element.tagName || rule.selector.startsWith(".") && classNames.includes(rule.selector.slice(1))) cell.numberFormat ??= rule.code;
    cells.push(cell);
  };
  if (tables[0]) {
    const rows = all.filter(node => node.tagName === "tr"); height = rows.length;
    if (!height) invalid("Clipboard table contains no rows.");
    rows.forEach((row, index) => {
      let column = 0;
      for (const cell of children(row).filter(isElement).filter(node => node.tagName === "td" || node.tagName === "th")) {
        while (occupied.has(index + ":" + column)) column++;
        const attrs = attributes(cell), rowSpan = attrs.rowspan === undefined ? 1 : Number(attrs.rowspan), columnSpan = attrs.colspan === undefined ? 1 : Number(attrs.colspan);
        if (![rowSpan, columnSpan].every(value => Number.isSafeInteger(value) && value > 0) || index + rowSpan > height) invalid("Clipboard table has invalid row or column spans.");
        if (rowSpan * columnSpan > maxCells || column + columnSpan > 16_384) limit("Clipboard table spans exceed their cell budget.");
        for (let y = index; y < index + rowSpan; y++) for (let x = column; x < column + columnSpan; x++) {
          if (++work > maxCells) limit("Clipboard table exceeds its cell budget.");
          const key = y + ":" + x; if (occupied.has(key)) invalid("Clipboard table spans overlap."); occupied.add(key);
        }
        add(cell, index, column);
        if (rowSpan > 1 || columnSpan > 1) merges.push({ from: { row: origin.row + index, column: origin.column + column }, to: { row: origin.row + index + rowSpan - 1, column: origin.column + column + columnSpan - 1 } });
        column += columnSpan; width = Math.max(width, column);
      }
    });
  } else {
    const container = all.find(node => node.tagName === "body") ?? { nodeName: "div", tagName: "div", namespaceURI: "http://www.w3.org/1999/xhtml", attrs: [], childNodes: fragment.childNodes, parentNode: null } as Element;
    if (!projections.has(container)) projections.set(container, { tag: "div", attrs: {}, parent: undefined });
    add(container, 0, 0);
  }
  if (height * width > maxCells) limit("Clipboard HTML rectangle exceeds its cell budget.");
  const range = { from: { ...origin }, to: { row: origin.row + height - 1, column: origin.column + width - 1 } };
  if (!WorkbookCoordinateSchema.safeParse(range.to).success) invalid("Clipboard HTML exceeds its destination worksheet.");
  const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!;
  for (const merge of sheet.merges) if (merge.from.row <= range.to.row && merge.to.row >= range.from.row && merge.from.column <= range.to.column && merge.to.column >= range.from.column) {
    if (!merges.some(next => JSON.stringify(next) === JSON.stringify(merge))) invalid("HTML paste cannot split or replace a differently merged destination.");
  }
  const operations: DocumentTransaction["operations"] = [], inputs: WorkbookCellInput[] = [], richCells = new Set<string>(), appearance = new Map<string, { properties: Properties; ranges: WorkbookRange[] }>();
  const addOperation = (operation: DocumentTransaction["operations"][number]): void => { if (operations.length >= 1024) limit("Clipboard HTML exceeds the transaction operation budget."); operations.push(operation); };
  const byCoordinate = new Map(cells.map(cell => [cell.at.row + ":" + cell.at.column, cell]));
  const formatsByCode = new Map(snapshot.metadata.workbook.numberFormats.filter(item => "code" in item).map(item => [item.code, item.formatId]));
  const ids = new Set(snapshot.metadata.workbook.numberFormats.map(item => item.formatId)), builtinFormats = new Map(snapshot.metadata.workbook.numberFormats.filter(item => "builtinId" in item).map(item => [item.builtinId, item.formatId])); let serial = 0;
  for (let row = range.from.row; row <= range.to.row; row++) for (let column = range.from.column; column <= range.to.column; column++) {
    const cell = byCoordinate.get(row + ":" + column), at = { row, column };
    if (!cell) { inputs.push({ ...at, source: { kind: "value", value: { type: "blank" } } }); continue; }
    const attrs = attributes(cell.element);
    const typed = ["x:str", "x:fmla", "data-sheets-value", "data-sheets-formula", "data-doffi-value", "data-doffi-formula-a1"].some(name => attrs[name] !== undefined);
    let source = parseWorkbookTextInput(cell.text, at, typed ? "text" : options.mode), formatId: string | undefined;
    if (options.mode !== "text" && (attrs["x:num"] !== undefined || attrs.sdval !== undefined && !attrs["data-sheets-formula"])) {
      const number = attrs["x:num"] || attrs.sdval || cell.text;
      source = parseWorkbookTextInput(number, at, "number");
    }
    if (options.mode !== "text" && attrs["data-sheets-value"] !== undefined) {
      let data: Record<string, unknown>;
      try { data = JSON.parse(attrs["data-sheets-value"]!); } catch { return invalid("Clipboard cell value metadata is invalid JSON."); }
      if (!data || typeof data !== "object" || Array.isArray(data)) invalid("Clipboard cell value metadata is invalid.");
      const type = data["1"], value = type === 2 ? { type: "text", value: data["2"] } : type === 3 ? { type: "number", value: data["3"] }
        : type === 4 ? { type: "boolean", value: data["4"] === 1 ? true : data["4"] === 0 ? false : data["4"] } : undefined;
      const parsed = WorkbookValueSchema.safeParse(value); if (!parsed.success) invalid("Clipboard cell value metadata has an unsupported type.");
      else if (type !== 3 || attrs["x:num"] === undefined && attrs.sdval === undefined) source = { kind: "value", value: parsed.data };
    } else if (options.mode !== "text" && attrs["data-sheets-formula"]?.startsWith("=#")) {
      const parsed = WorkbookValueSchema.safeParse({ type: "error", value: attrs["data-sheets-formula"].slice(1) });
      if (parsed.success) source = { kind: "value", value: parsed.data };
    }
    if (attrs["x:str"] !== undefined) source = parseWorkbookTextInput(cell.text, at, "text");
    const formulae: string[] = [];
    if (options.mode !== "text" && attrs["x:fmla"] !== undefined) {
      const local = { row: row - origin.row, column: column - origin.column };
      formulae.push(printWorkbookFormula(parseWorkbookFormula(attrs["x:fmla"], local), at));
    }
    if (attrs["data-doffi-formula"] !== undefined && (attrs["data-doffi-formula"] !== "R1C1" || !attrs["data-sheets-formula"])) invalid("Clipboard formula metadata requires its declared dialect.");
    if (options.mode !== "text" && attrs["data-sheets-formula"] && (attrs["data-doffi-formula"] === "R1C1" || !attrs["data-sheets-formula"].startsWith("=#"))) {
      const formula = attrs["data-sheets-formula"];
      // Calc also uses these exact constant carriers for native Boolean cells.
      if (!(attrs["data-doffi-formula"] === undefined && source.kind === "value" && source.value.type === "boolean" && /^=(?:TRUE|FALSE)\(\)$/i.test(formula))) formulae.push(printWorkbookFormula(parseWorkbookR1C1Formula(formula, at), at));
    }
    if (options.mode !== "text" && attrs["data-doffi-value"] !== undefined) {
      let value: unknown;
      try { value = JSON.parse(attrs["data-doffi-value"]!); } catch { return invalid("Clipboard value metadata is invalid JSON."); }
      const parsed = WorkbookValueSchema.safeParse(value); if (!parsed.success) invalid("Clipboard value metadata is invalid.");
      source = { kind: "value", value: parsed.data };
    }
    if (options.mode !== "text" && attrs["data-doffi-formula-a1"] !== undefined) {
      const origin = /^(\d+),(\d+)$/.exec(attrs["data-doffi-origin"] ?? ""), coordinate = origin ? WorkbookCoordinateSchema.safeParse({ row: Number(origin[1]), column: Number(origin[2]) }) : undefined;
      if (!coordinate?.success || !attrs["data-doffi-formula-a1"].startsWith("=")) invalid("Clipboard A1 formulas require valid source coordinates.");
      formulae.push(copyWorkbookFormula(attrs["data-doffi-formula-a1"], coordinate.data, at));
    }
    if (formulae.length) {
      if (formulae.some(formula => formula !== formulae[0])) invalid("Clipboard formula dialect metadata disagrees.");
      source = parseWorkbookTextInput(formulae[0]!, at);
    }
    if (cell.runs.length && source.kind === "value" && source.value.type === "text" && source.value.value === cell.text) {
      source = { kind: "richText", html: escape(cell.text).replace(/\n/g, "<br>") }; richCells.add(row + ":" + column);
    }
    if (cell.numberFormat !== undefined && attrs["data-doffi-number-format"] === undefined) {
      formatId = formatsByCode.get(cell.numberFormat);
      if (!formatId) {
        do { formatId = "clipboard-format-" + snapshot.manifest.revision + "-" + serial++; } while (ids.has(formatId)); ids.add(formatId); formatsByCode.set(cell.numberFormat, formatId);
        addOperation({ op: "setWorkbookNumberFormat", format: { formatId, code: cell.numberFormat } });
      }
    }
    if (attrs["data-doffi-number-format"] !== undefined) {
      const builtinId = Number(attrs["data-doffi-number-format"]);
      if (!/^\d+$/.test(attrs["data-doffi-number-format"]) || !Number.isInteger(builtinId) || builtinId < 0 || builtinId > 163) invalid("Clipboard builtin number format is invalid.");
      formatId = builtinFormats.get(builtinId);
      if (!formatId) {
        do { formatId = "clipboard-format-" + snapshot.manifest.revision + "-" + serial++; } while (ids.has(formatId)); ids.add(formatId);
        builtinFormats.set(builtinId, formatId); addOperation({ op: "setWorkbookNumberFormat", format: { formatId, builtinId } });
      }
    }
    inputs.push({ ...at, source, ...(formatId ? { formatId } : {}) });
    if (Object.keys(cell.properties).length) {
      const key = JSON.stringify(cell.properties), group = appearance.get(key) ?? { properties: cell.properties, ranges: [] };
      group.ranges.push({ from: at, to: at }); appearance.set(key, group);
    }
  }
  addOperation({ op: "setWorkbookCells", sheetId, cells: inputs });
  for (const group of appearance.values()) addOperation({ op: "formatWorkbookRanges", sheetId, ...group });
  for (const cell of cells) if (richCells.has(cell.at.row + ":" + cell.at.column))
    for (const run of cell.runs) addOperation({ op: "formatWorkbookText", sheetId, at: cell.at, from: run.from, to: run.to, expectedText: cell.text.slice(run.from, run.to), properties: run.properties });
  for (const merge of merges) addOperation({ op: "mergeWorkbookCells", sheetId, range: merge, coveredContent: "reject" });
  addOperation({ op: "replaceWorkbookHyperlinks", sheetId, range, links: cells.filter(cell => cell.hyperlink).map(cell => ({ ...cell.hyperlink!, range: { from: cell.at, to: cell.at } })) });
  return { transaction: { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations }, range, diagnostics: [...diagnostics] };
}
