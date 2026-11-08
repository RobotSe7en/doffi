import { ident, parse, walk } from "css-tree";
import type { WorkbookCellStyles, WorkbookMetadata, WorkbookStyleApplication } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import { DocumentError } from "../errors.js";
import { resolveWorkbookCellFormat } from "../workbook/cell-format.js";
import { WorkbookCellEdit } from "./workbook-cells.js";
import { SourceStyleEdit } from "./source-styles.js";
import { checkRange } from "../workbook/ranges.js";
import { formatWorkbookCellText } from "./workbook-text-format.js";

type StyleOperation = Extract<WorkbookOperation, { op: "formatWorkbookText" | "formatWorkbookRanges" | "formatWorkbookConditionalRule" | "setWorkbookStyleClass" | "setWorkbookCellStyleDefinition" | "setWorkbookCellStyleName" | "removeWorkbookCellStyleName" }>;
type Group = keyof WorkbookStyleApplication;
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
function group(property: string): Group {
  if (/^(font-|color$|text-decoration|vertical-align$|--office-font-|--office-cell-foreground)/.test(property)) return "font";
  if (/^(background-color$|--office-cell-(fill|gradient))/.test(property)) return "fill";
  if (/^(border(?:-|$)|--office-cell-border-)/.test(property)) return "border";
  if (/^(text-align$|text-indent$|white-space$|overflow-wrap$|line-height$|letter-spacing$|word-spacing$|--office-cell-(horizontal|vertical|indent|relative-indent|justify-last-line|reading-order|shrink-to-fit|text-rotation|wrap)$)/.test(property)) return "alignment";
  return invalid("The property is outside the workbook cell appearance groups.");
}
function validateNativeAppearance(properties: Readonly<Record<string, string>>, styles: SourceStyleEdit): void {
  const enums: readonly [RegExp, readonly string[]][] = [
    [/^--office-font-underline$/, ["none", "single", "double", "singleAccounting", "doubleAccounting"]],
    [/^--office-(?:font-strike|cell-(?:wrap|shrink-to-fit|justify-last-line|locked|formula-hidden|border-(?:diagonal-up|diagonal-down|outline)))$/, ["true", "false"]],
    [/^--office-cell-horizontal$/, ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"]],
    [/^--office-cell-vertical$/, ["top", "center", "bottom", "justify", "distributed"]],
    [/^--office-cell-reading-order$/, ["context", "ltr", "rtl"]],
    [/^--office-cell-fill-kind$/, ["none", "solid", "gradient", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"]],
    [/^--office-cell-gradient-kind$/, ["linear", "path"]],
    [/^--office-cell-border-[a-z]+-kind$/, ["none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"]],
  ];
  for (const [property, raw] of Object.entries(properties)) {
    const value = raw.trim().replace(/\s*!\s*important\s*$/i, ""); if (!property.startsWith("--") || /\bvar\(/i.test(value) || ["initial", "inherit", "unset", "revert"].includes(value)) continue;
    if (property.endsWith("-color")) styles.validateProperty("color", value);
    const border = /^--office-cell-border-([a-z]+)-kind$/.exec(property);
    if (border && !["left", "right", "top", "bottom", "start", "end", "diagonal", "horizontal", "vertical"].includes(border[1]!)) invalid("A native cell border side is invalid.");
    if (border && ["horizontal", "vertical"].includes(border[1]!) && value !== "none") throw new DocumentError("UnsupportedFeature", "Internal native cell borders are outside the declared workbook appearance scope.");
    for (const [pattern, values] of enums) if (pattern.test(property) && !values.includes(value)) invalid("A native workbook appearance token is invalid.");
    const numeric = (minimum: number, maximum: number, integer = false): void => {
      const number = Number(value);
      if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(value) || !Number.isFinite(number) || number < minimum || number > maximum || integer && !Number.isInteger(number)) invalid("A native workbook appearance parameter is outside its bounds.");
    };
    if (property === "--office-cell-text-rotation") { numeric(0, 255, true); if (Number(value) > 180 && Number(value) !== 255) invalid("Native text rotation must be 0–180 or 255."); }
    if (property.endsWith("-tint")) numeric(-1, 1);
    if (property.endsWith("-native-alpha")) numeric(0, 255, true);
    if (/^--office-cell-gradient-(?:left|right|top|bottom|stop-\d+-position)$/.test(property)) numeric(0, 1);
    if (property === "--office-cell-gradient-degree") numeric(-Number.MAX_VALUE, Number.MAX_VALUE);
    if (property === "--office-cell-gradient-stops") numeric(0, 4096, true);
    if (property === "--office-cell-indent") numeric(0, 255, true);
    if (property === "--office-cell-relative-indent") numeric(-15, 15, true);
  }
}
export function isWorkbookStyleOperation(operation: WorkbookOperation): operation is StyleOperation {
  return ["formatWorkbookText", "formatWorkbookRanges", "formatWorkbookConditionalRule", "setWorkbookStyleClass", "setWorkbookCellStyleDefinition", "setWorkbookCellStyleName", "removeWorkbookCellStyleName"].includes(operation.op);
}
/** Style identities and update flags remain metadata; effective appearance
 * stays solely in existing CSS and cell/row/column number-format sources. */
export function* editWorkbookStyle(workbook: WorkbookMetadata, cells: WorkbookCellEdit, styles: SourceStyleEdit, operation: StyleOperation,
  allocate: () => string): Generator<{ completed: number; total: number }, void> {
  if (operation.op === "setWorkbookCellStyleName") {
    if (!workbook.cellStyles?.definitions.some(style => style.styleId === operation.name.styleId)) invalid("A cell style name requires an existing definition.");
    const names = workbook.cellStyles!.names, index = names.findIndex(name => name.name.toUpperCase() === operation.name.name.toUpperCase());
    if (index < 0) names.push(operation.name); else names[index] = operation.name; return;
  }
  if (operation.op === "removeWorkbookCellStyleName") {
    if (workbook.cellStyles) workbook.cellStyles.names = workbook.cellStyles.names.filter(name => name.name.toUpperCase() !== operation.name.toUpperCase());
    return;
  }
  const properties: Record<string, string> = {};
  for (const [token, value] of Object.entries(operation.properties)) {
    const decoded = ident.decode(token), property = decoded.startsWith("--") ? decoded : decoded.toLowerCase();
    if (Object.hasOwn(properties, property)) invalid("A workbook style patch cannot repeat a property.");
    properties[property] = value;
  }
  // Native scalar color sources are consumed before ordinary CSS fallback.
  // A direct CSS color edit therefore updates the same native paint source;
  // old tint/ARGB provenance cannot recolor the user's new explicit value.
  const paint = (prefix: string, value: string): void => {
    if (value.includes(prefix + "-color")) return;
    properties[prefix + "-color"] ??= value; properties[prefix + "-tint"] ??= "0"; properties[prefix + "-native-alpha"] ??= "initial";
  };
  if (properties.color !== undefined) paint("--office-cell-foreground", properties.color);
  if (properties["background-color"] !== undefined) { properties["--office-cell-fill-kind"] ??= "solid"; paint("--office-cell-fill", properties["background-color"]); }
  if (properties["font-family"] !== undefined) for (const role of ["ascii", "h-ansi", "east-asia", "complex-script"]) properties["--office-font-" + role] ??= properties["font-family"];
  validateNativeAppearance(properties, styles);
  const groups = new Map(Object.keys(properties).map(property => [property, group(property)]));
  const classes = new Set<string>();
  for (const sheet of styles.snapshot()) walk(parse(sheet.css, { context: "stylesheet" }), node => { if (node.type === "ClassSelector") classes.add(ident.decode(node.name)); });
  if (operation.op === "formatWorkbookConditionalRule") {
    const rule = workbook.sheets.find(sheet => sheet.sheetId === operation.sheetId)?.conditionalRules?.find(rule => rule.ruleId === operation.ruleId)
      ?? invalid("Conditional formatting requires an existing rule.");
    let target: { cssClass?: string | undefined; numberFormatId?: string | undefined };
    if (rule.kind === "colorScale") {
      if (operation.stopIndex === undefined || !rule.stops[operation.stopIndex] || operation.numberFormatId !== undefined)
        invalid("A color scale requires an existing stop and cannot carry a number format.");
      if (Object.keys(properties).some(property => property !== "background-color" && !/^--office-cell-fill-(?:kind|color|tint|native-alpha)$/.test(property))
        || properties["--office-cell-fill-kind"] !== undefined && properties["--office-cell-fill-kind"].trim() !== "solid") invalid("Color-scale stops accept a solid background color.");
      target = rule.stops[operation.stopIndex]!;
    } else {
      if (operation.stopIndex !== undefined) invalid("Only color-scale rules have indexed stops.");
      target = rule;
      if (operation.numberFormatId !== undefined && operation.numberFormatId !== null && !workbook.numberFormats.some(format => format.formatId === operation.numberFormatId))
        invalid("Conditional formatting requires a defined number format.");
    }
    if (Object.keys(properties).length) {
      const current = target.cssClass;
      if (!current || !styles.resolver({ combineDecorations: false }).hasSpecifiedClass(current, properties)) {
        let local: string; do { local = "workbook-rule-" + allocate(); } while (classes.has(local));
        if (current) styles.aliasClass(current, local, new Set(Object.keys(properties)));
        styles.overrideLocalClass(local, properties, {}, ""); target.cssClass = local;
      }
    }
    if (operation.numberFormatId === null) delete target.numberFormatId;
    else if (operation.numberFormatId !== undefined) target.numberFormatId = operation.numberFormatId;
    return;
  }
  if (operation.op === "formatWorkbookText") {
    if ([...groups.values()].some(group => group !== "font")) invalid("Partial text formatting accepts only font properties.");
    const cell = cells.get(operation.sheetId, operation.at.row, operation.at.column);
    formatWorkbookCellText(cell, operation);
    if (!Object.keys(properties).length) return;
    let cssClass: string; do { cssClass = "workbook-run-" + allocate(); } while (classes.has(cssClass));
    const html = formatWorkbookCellText(cell, operation, cssClass);
    styles.overrideLocalClass(cssClass, properties, properties, "span");
    cells.set(operation.sheetId, { ...operation.at, source: { kind: "richText", html } }); return;
  }
  if (operation.op === "formatWorkbookRanges") {
    const sheet = workbook.sheets.find(sheet => sheet.sheetId === operation.sheetId) ?? invalid("The formatted worksheet does not exist.");
    let total = 0; for (const range of operation.ranges) { checkRange(range); total += (range.to.row - range.from.row + 1) * (range.to.column - range.from.column + 1); }
    cells.checkExpandedCells(total);
    if (operation.numberFormatId !== undefined && !workbook.numberFormats.some(format => format.formatId === operation.numberFormatId)) invalid("Cell formatting requires a defined number format.");
    if (!Object.keys(properties).length && operation.numberFormatId === undefined) return;
    const derived = new Map<string, string>(), definitions = new Map(workbook.cellStyles?.definitions.map(style => [style.cssClass, style])),
      bindings = new Map(workbook.cellStyles?.bindings.map(binding => [binding.cssClass, binding]));
    let completed = 0; const seen = new Set<string>();
    for (const range of operation.ranges) for (let row = range.from.row; row <= range.to.row; row++) for (let column = range.from.column; column <= range.to.column; column++) {
      const key = row + ":" + column; if (seen.has(key)) continue; seen.add(key);
      const cell = cells.get(sheet.sheetId, row, column), format = resolveWorkbookCellFormat(sheet, { row, column }, cell), tokens = format.cssClass.split(/\s+/).filter(Boolean),
        bindingToken = tokens.find(name => bindings.has(name) || definitions.has(name));
      let cssClass = derived.get(format.cssClass);
      if (cssClass === undefined) {
        const binding = bindingToken ? bindings.get(bindingToken) : undefined, definition = bindingToken ? definitions.get(bindingToken) : undefined;
        if (Object.keys(properties).length || binding || definition) {
          let local: string; do { local = "workbook-local-" + allocate(); } while (classes.has(local)); classes.add(local);
          if (bindingToken) styles.aliasClass(bindingToken, local);
          if (binding || definition) {
            const apply = binding ? { ...binding.apply } : { font: false, fill: false, border: false, alignment: false, numberFormat: false, protection: false };
            for (const value of groups.values()) apply[value] = true;
            if (operation.numberFormatId !== undefined) apply.numberFormat = true;
            workbook.cellStyles!.bindings.push({ cssClass: local, styleId: (binding ?? definition)!.styleId, apply });
          }
          const font = Object.fromEntries(Object.entries(properties).filter(([property]) => groups.get(property) === "font"));
          styles.overrideLocalClass(local, properties, font, "td");
          cssClass = [...tokens.filter(name => name !== bindingToken), local].join(" ");
        } else cssClass = format.cssClass;
        derived.set(format.cssClass, cssClass);
      }
      cells.format(sheet.sheetId, { row, column, ...(cssClass !== format.cssClass ? { cssClass } : {}), ...(operation.numberFormatId !== undefined ? { formatId: operation.numberFormatId } : {}) });
      if (++completed % 256 === 0) yield { completed, total };
    }
    return;
  }
  if (operation.op === "setWorkbookStyleClass") {
    if (!classes.has(operation.cssClass)) invalid("Style editing requires an existing workbook class.");
    if (workbook.cellStyles?.definitions.some(style => style.cssClass === operation.cssClass)) invalid("Named definitions must be updated through the named-style operation.");
    styles.setClass(operation.cssClass, properties);
    const binding = workbook.cellStyles?.bindings.find(binding => binding.cssClass === operation.cssClass);
    if (binding) for (const value of groups.values()) binding.apply[value] = true;
    return;
  }
  const definition = operation.definition;
  if (!workbook.numberFormats.some(format => format.formatId === definition.numberFormatId)) invalid("A named style requires a defined number format.");
  const model: WorkbookCellStyles = workbook.cellStyles ??= { definitions: [], names: [], bindings: [] };
  let previous = model.definitions.find(style => style.styleId === definition.styleId);
  if (!previous) {
    let cssClass: string; do { cssClass = "workbook-style-" + allocate(); } while (classes.has(cssClass));
    previous = { ...definition, cssClass }; model.definitions.push(previous);
    styles.addClass(cssClass, {});
  }
  const oldFormat = previous.numberFormatId, directClass = previous.cssClass;
  const followers = new Map(model.bindings.filter(binding => binding.styleId === definition.styleId).map(binding => [binding.cssClass, binding]));
  styles.setClass(directClass, properties);
  let followed = 0;
  for (const binding of followers.values()) {
    const inherited = Object.fromEntries(Object.entries(properties).filter(([property]) => !binding.apply[groups.get(property)!]));
    if (Object.keys(inherited).length) styles.setClass(binding.cssClass, inherited);
    if (++followed % 32 === 0) yield { completed: followed, total: followers.size };
  }
  if (oldFormat !== definition.numberFormatId) {
    const follows = (css: string): boolean => css.split(/\s+/).some(name => name === directClass || followers.has(name) && !followers.get(name)!.apply.numberFormat);
    // Capture before dimension defaults change; a cell without an explicit
    // format continues to inherit its updated source instead of being expanded.
    let completed = 0;
    for (const cell of cells.entries()) {
      const sheet = workbook.sheets.find(sheet => sheet.sheetId === cell.sheetId)!;
      if ((cell.formatId || cell.cssClass) && follows(resolveWorkbookCellFormat(sheet, cell, cell).cssClass)) cells.format(cell.sheetId, { row: cell.row, column: cell.column, formatId: definition.numberFormatId });
      if (++completed % 256 === 0) yield { completed, total: 0 };
    }
    for (const sheet of workbook.sheets) {
      if (follows(sheet.defaultCssClass)) sheet.defaultNumberFormatId = definition.numberFormatId;
      for (const dimension of [...sheet.rowDimensions, ...sheet.columnDimensions]) if (dimension.cssClass && follows(dimension.cssClass)) dimension.numberFormatId = definition.numberFormatId;
    }
  }
  Object.assign(previous, definition);
}
