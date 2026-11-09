import { captureWorkbookModel, createWorkbookCalculatedEvaluator, createWorkbookNumberFormatter, DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, parseWorkbookFormula, printWorkbookR1C1Formula, resolveWorkbookCellFormat, SourceStyleResolver, WorkbookRangeSchema } from "@doffi/core";
import type { DeepReadonly, SourceStyleNode, WorkbookCalculationResult, WorkbookModel, WorkbookRange, WorkbookValue } from "@doffi/core";
import { parseFragment } from "parse5";
import type { DefaultTreeAdapterTypes } from "parse5";
import { string } from "css-tree";

/** Portable clipboard content carries data, never session handles or resource
 * capabilities. Source coordinates only describe formula-copy offsets. */
export interface WorkbookClipboardData { readonly text: string; readonly html: string }
export const WORKBOOK_CLIPBOARD_MAX_CELLS = 25_000;
export interface WorkbookClipboardDataOptions {
  readonly calculation?: WorkbookCalculationResult;
  readonly locale?: string;
  readonly maxCells?: number;
  readonly maxCodeUnits?: number;
  readonly maxSteps?: number;
  readonly signal?: { readonly aborted: boolean };
  readonly yield?: () => Promise<void>;
}
type Node = DefaultTreeAdapterTypes.Node;
type Element = DefaultTreeAdapterTypes.Element;
const element = (node: Node): node is Element => "tagName" in node;
const children = (node: Node): readonly Node[] => "childNodes" in node ? node.childNodes : [];
const escape = (value: string): string => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\r/g, "&#13;");
const overlap = (a: DeepReadonly<WorkbookRange>, b: DeepReadonly<WorkbookRange>): boolean => a.from.row <= b.to.row && a.to.row >= b.from.row && a.from.column <= b.to.column && a.to.column >= b.from.column;
const contains = (a: DeepReadonly<WorkbookRange>, b: DeepReadonly<WorkbookRange>): boolean => a.from.row <= b.from.row && a.to.row >= b.to.row && a.from.column <= b.from.column && a.to.column >= b.to.column;
const supported = (name: string): boolean => /^(?:font-(?:family|size|weight|style|variant)|color|background-color|text-(?:align|indent|decoration(?:-line|-style|-color)?)|vertical-align|white-space|line-height|letter-spacing|word-spacing|border(?:-(?:top|right|bottom|left))?(?:-color|-style|-width)?|padding(?:-(?:top|right|bottom|left))?|--office-(?:font-(?:ascii|h-ansi|east-asia|complex-script|underline|strike|decoration-reset)|cell-[\w-]+))$/.test(name);

/** Bounded source serialization shared by direct and Worker editors. Formula
 * caches are never substituted for current calculated values. */
export async function createWorkbookClipboardData(input: WorkbookModel, sheetId: string, selection: DeepReadonly<WorkbookRange>, options: WorkbookClipboardDataOptions = {}): Promise<WorkbookClipboardData> {
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["calculation", "locale", "maxCells", "maxCodeUnits", "maxSteps", "signal", "yield"].includes(key))) throw new DocumentError("InvalidOperation", "Invalid workbook clipboard options.");
  const config = { maxCells: WORKBOOK_CLIPBOARD_MAX_CELLS, maxCodeUnits: 2 * 1024 * 1024, maxSteps: 10_000_000, ...options }, parsed = WorkbookRangeSchema.safeParse(selection);
  if (!parsed.success || parsed.data.from.row > parsed.data.to.row || parsed.data.from.column > parsed.data.to.column
    || ![config.maxCells, config.maxCodeUnits, config.maxSteps].every(value => Number.isSafeInteger(value) && value > 0)
    || config.locale !== undefined && (typeof config.locale !== "string" || !config.locale || config.locale.length > 255)
    || config.signal !== undefined && (!config.signal || typeof config.signal.aborted !== "boolean") || config.yield !== undefined && typeof config.yield !== "function") throw new DocumentError("InvalidOperation", "Clipboard output requires an ordered range and valid budgets.");
  const model = captureWorkbookModel(input), snapshot = model.snapshot(), range = parsed.data, sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId);
  if (!sheet) throw new DocumentError("InvalidOperation", "Clipboard output requires an existing worksheet.");
  const total = (range.to.row - range.from.row + 1) * (range.to.column - range.from.column + 1);
  if (total > config.maxCells) throw new DocumentError("ResourceLimitExceeded", "Clipboard output exceeds its cell budget.");
  let steps = 0, units = 0;
  const guard = (): void => {
    if (config.signal?.aborted) throw new DocumentError("OperationCancelled", "Clipboard output was cancelled.");
    if (captureWorkbookModel(input) !== model) throw new DocumentError("RevisionConflict", "Clipboard source changed before completion.");
  };
  const step = (): void => { if (++steps > config.maxSteps) throw new DocumentError("ResourceLimitExceeded", "Clipboard output exceeds its work budget."); };
  const account = (value: string): string => { units += value.length; if (units > config.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Clipboard output exceeds its text budget."); return value; };
  guard();
  const merges = sheet.merges.filter(merge => { step(); return overlap(range, merge); });
  if (merges.some(merge => !contains(range, merge))) throw new DocumentError("InvalidOperation", "Clipboard output cannot split a merged region.");
  const resolver = new SourceStyleResolver(snapshot.styles.map(style => style.css), DEFAULT_DOCUMENT_READ_LIMITS);
  // Chunk ancestry has a strict, public source contract. Cell/story node IDs
  // are unrelated: workbook getNode only reads the separate story document.
  const root: SourceStyleNode = { tag: "html", attrs: {}, parent: undefined }, body: SourceStyleNode = { tag: "body", attrs: {}, parent: root };
  const tables = new Map<string, SourceStyleNode>(), rows = new Map<string, SourceStyleNode>();
  const parent = (row: number, column: number): SourceStyleNode => {
    const chunkRow = Math.floor(row / 256) * 256, chunkColumn = Math.floor(column / 32) * 32, key = chunkRow + ":" + chunkColumn;
    let table = tables.get(key);
    if (!table) { table = { tag: "tbody", attrs: {}, parent: { tag: "table", attrs: { "data-sheet-id": sheetId, "data-row": String(chunkRow), "data-column": String(chunkColumn) }, parent: body } }; tables.set(key, table); }
    const rowKey = key + ":" + row; let result = rows.get(rowKey);
    if (!result) { result = { tag: "tr", attrs: { "data-row": String(row) }, parent: table }; rows.set(rowKey, result); } return result;
  };
  const css = (projection: SourceStyleNode, clearDecoration = false): string => Object.entries(resolver.get(projection)).filter(([name]) => supported(name)).map(([name, value]) => name + ":" + (clearDecoration && name === "text-decoration-line" ? "none" : value)).join(";");
  const rich = (source: string, cssClass: string, row: number, column: number): { style: string; content: string } => {
    account(source); const tree = parseFragment("<table><tbody><tr>" + source + "</tr></tbody></table>");
    let root: Element | undefined;
    const find = (node: Node): void => { step(); if (element(node) && node.tagName === "td") root ??= node; for (const child of children(node)) find(child); };
    find(tree);
    if (!root) throw new DocumentError("InvalidDocument", "Clipboard source has no cell element.");
    const attrs = Object.fromEntries(root.attrs.map(attr => [attr.name, attr.value]));
    const projection: SourceStyleNode = { tag: "td", attrs: { ...attrs, class: cssClass }, parent: parent(row, column) };
    const render = (node: Node, ancestor: SourceStyleNode, depth: number): string => {
      step(); if (depth > 128) throw new DocumentError("ResourceLimitExceeded", "Clipboard rich text exceeds its depth budget.");
      if (node.nodeName === "#text" && "value" in node) {
        const paint = resolver.get(ancestor), style = ["text-decoration-line", "text-decoration-style", "text-decoration-color"].map(name => name + ":" + paint[name]).join(";");
        return '<span style="' + escape(style) + '">' + escape(node.value) + "</span>";
      }
      if (!element(node)) return "";
      if (node.tagName === "br") return "<br>";
      const next: SourceStyleNode = { tag: node.tagName, attrs: Object.fromEntries(node.attrs.map(attr => [attr.name, attr.value])), parent: ancestor };
      return '<span style="' + escape(css(next, true)) + '">' + children(node).map(child => render(child, next, depth + 1)).join("") + "</span>";
    };
    return { style: css(projection), content: children(root).map(child => render(child, projection, 0)).join("") };
  };
  const evaluator = config.calculation === undefined ? undefined : createWorkbookCalculatedEvaluator(model, config.calculation, { maxSteps: config.maxSteps, maxCells: config.maxCells, maxCodeUnits: config.maxCodeUnits });
  const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem, ...(config.locale === undefined ? {} : { hostLocale: config.locale }), ...(snapshot.metadata.locale ? { documentLocale: snapshot.metadata.locale } : {}) });
  const formats = new Map(snapshot.metadata.workbook.numberFormats.map(format => [format.formatId, format]));
  const html: string[] = [account('<table style="border-collapse:collapse;white-space:pre-wrap"><tbody>')], text: string[] = [];
  try {
    let visited = 0;
    for (let row = range.from.row; row <= range.to.row; row++) {
      html.push(account("<tr>")); const line: string[] = [];
      for (let column = range.from.column; column <= range.to.column; column++) {
        guard(); step();
        const at = { row, column }, cell = model.getCell(sheetId, row, column), format = resolveWorkbookCellFormat(sheet, at, cell);
        let merge: typeof merges[number] | undefined;
        for (const item of merges) { step(); if (contains(item, { from: at, to: at })) { merge = item; break; } }
        const covered = merge && (row !== merge.from.row || column !== merge.from.column);
        let value: DeepReadonly<WorkbookValue> = covered ? { type: "blank" } : cell?.value ?? { type: "blank" };
        if (!covered && cell?.formula) {
          if (!evaluator) throw new DocumentError("UnsupportedFeature", "Copying formulas requires current calculation results.");
          const result = evaluator.cell({ sheetId, ...at }, { maxSteps: config.maxSteps - steps, maxCodeUnits: Math.max(0, config.maxCodeUnits - units), ...(config.signal ? { signal: config.signal } : {}) });
          steps += result.steps; units += result.codeUnits;
          if (result.kind !== "value") throw new DocumentError("UnsupportedFeature", "Clipboard formula results are unavailable."); value = result.value;
        }
        const definition = format.numberFormatId === undefined ? undefined : formats.get(format.numberFormatId);
        let display = value.type === "blank" ? "" : value.type === "boolean" ? value.value ? "TRUE" : "FALSE" : String(value.value);
        if (value.type === "number" || value.type === "text") {
          const formatted = formatter.format(value.value, definition);
          if (formatted.kind !== "parts") throw new DocumentError("UnsupportedFeature", "Clipboard number formatting is unavailable.");
          display = formatted.parts.map(part => part.kind === "text" ? part.text : "").join("");
        }
        line.push(account(/[\t\r\n"]/.test(display) ? '"' + display.replace(/"/g, '""') + '"' : display));
        if (!covered) {
          const appearance = rich(cell?.html ?? "<td></td>", format.cssClass, row, column), numberFormat = definition ? ";mso-number-format:" + string.encode(formatter.code(definition)) : "";
          const textSource = !cell?.formula && value.type === "text";
          let attrs = ' style="' + escape(appearance.style + numberFormat + (textSource ? ";text-decoration-line:none" : "")) + '" data-doffi-value="' + escape(JSON.stringify(value)) + '"';
          if (definition && "builtinId" in definition) attrs += ' data-doffi-number-format="' + definition.builtinId + '"';
          if (cell?.formula) attrs += ' data-doffi-formula="R1C1" data-sheets-formula="' + escape(printWorkbookR1C1Formula(parseWorkbookFormula(cell.formula.expression, at), at)) + '"';
          else if (value.type === "error") attrs += ' data-sheets-formula="=' + escape(value.value) + '"';
          if (value.type === "text") attrs += ' x:str="" data-sheets-value="' + escape(JSON.stringify({ "1": 2, "2": value.value })) + '"';
          if (value.type === "boolean" && !cell?.formula) attrs += ' data-sheets-value="' + escape(JSON.stringify({ "1": 4, "4": value.value ? 1 : 0 })) + '"';
          if (value.type === "number") {
            attrs += ' x:num="' + value.value + '" sdnum="' + escape("1033;0;" + formatter.code(definition)) + '"';
            // Calc's sdval is a constant carrier and takes precedence over its
            // R1C1 formula attribute. Keep it off formulas so they stay editable.
            if (!cell?.formula) attrs += ' sdval="' + value.value + '" data-sheets-value="' + escape(JSON.stringify({ "1": 3, "3": value.value })) + '"';
          }
          if (merge) attrs += ' rowspan="' + (merge.to.row - merge.from.row + 1) + '" colspan="' + (merge.to.column - merge.from.column + 1) + '"';
          let content = !cell?.formula && value.type === "text" ? appearance.content : escape(display).replace(/\n/g, "<br>");
          const link = sheet.hyperlinks?.find(link => { step(); return contains(link.range, { from: at, to: at }); });
          if (link) { const href = link.url ?? "#" + link.location; content = '<a href="' + escape(href) + '">' + content + "</a>"; }
          html.push(account("<td" + attrs + ">" + content + "</td>"));
        }
        if (++visited % 128 === 0 && visited < total && config.yield) { await config.yield(); guard(); }
      }
      text.push(line.join("\t")); html.push(account("</tr>"));
    }
    html.push(account("</tbody></table>")); const plain = text.join("\r\n"); account("\t".repeat(Math.max(0, total - text.length)) + "\r\n".repeat(Math.max(0, text.length - 1)));
    guard(); return Object.freeze({ text: plain, html: html.join("") });
  } finally { evaluator?.dispose(); formatter.dispose(); }
}
