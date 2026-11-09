import { checkLimit, freeze, invalid, unsupported } from "../document/validation.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "./chunk-index.js";
import { parseStructuredReference, readStructuredToken, type FormulaStructuredReference } from "./structured-reference.js";

/** First-party A1 grammar. Relative positions remain offsets for template sharing. */
export type FormulaExpression =
  | { readonly kind: "number"; readonly value: number }
  | { readonly kind: "string"; readonly value: string }
  | { readonly kind: "boolean"; readonly value: boolean }
  | { readonly kind: "error"; readonly value: string }
  | { readonly kind: "name"; readonly name: string; readonly sheet?: string; readonly workbook?: 0 }
  | { readonly kind: "cell"; readonly row: number; readonly column: number; readonly absoluteRow: boolean; readonly absoluteColumn: boolean; readonly sheet?: string }
  | { readonly kind: "range"; readonly from: FormulaCell; readonly to: FormulaCell }
  | { readonly kind: "axisRange"; readonly axis: "row" | "column"; readonly from: FormulaAxisPosition; readonly to: FormulaAxisPosition; readonly sheet?: string }
  | FormulaStructuredReference
  | { readonly kind: "unary"; readonly operator: "+" | "-" | "%"; readonly operand: FormulaExpression }
  | { readonly kind: "binary"; readonly operator: string; readonly left: FormulaExpression; readonly right: FormulaExpression }
  | { readonly kind: "call"; readonly name: string; readonly arguments: readonly (FormulaExpression | null)[] };
export type FormulaCell = Extract<FormulaExpression, { kind: "cell" }>;
export interface FormulaAxisPosition { readonly index: number; readonly absolute: boolean }
type AxisOperand = { readonly kind: "axis"; readonly axis: "row" | "column"; readonly position: FormulaAxisPosition; readonly sheet?: string };
type Operand = FormulaExpression | AxisOperand;
export interface FormulaOrigin { readonly row: number; readonly column: number }
interface Token { readonly kind: "number" | "string" | "quoted" | "word" | "error" | "operator" | "structured" | "space" | "end"; readonly text: string }

const operators: Readonly<Record<string, number>> = { "=": 10, "<>": 10, "<": 10, ">": 10, "<=": 10, ">=": 10, "&": 20, "+": 30, "-": 30, "*": 40, "/": 40, "^": 50, ",": 80, " ": 90, ":": 100 };
const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/;
const word = /^\$?[\p{L}\p{Nl}_\\][\p{L}\p{N}\p{M}_.?\\$]*/u;
const error = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA)/i;

const r1c1 = /^(?:R(?:\[[+-]?\d+\]|[1-9]\d*)?C(?:\[[+-]?\d+\]|[1-9]\d*)?|R(?:\[[+-]?\d+\]|[1-9]\d*)?|C(?:\[[+-]?\d+\]|[1-9]\d*)?)(?![\p{L}\p{N}\p{M}_.?$\\])/iu;
function tokenize(source: string, referenceStyle: "A1" | "R1C1"): Token[] {
  const tokens: Token[] = [];
  for (let cursor = 1; cursor < source.length;) {
    const character = source[cursor]!;
    if (/[ \r\n\t]/.test(character)) {
      while (cursor < source.length && /[ \r\n\t]/.test(source[cursor]!)) cursor++;
      tokens.push({ kind: "space", text: " " });
    } else if (character === "[") {
      const token = readStructuredToken(source, cursor); tokens.push({ kind: "structured", text: token.text }); cursor = token.end;
    } else if (character === '"' || character === "'") {
      const quote = character; let value = "", closed = false; cursor++;
      while (cursor < source.length) {
        const next = source[cursor++]!;
        if (next === quote) {
          if (source[cursor] === quote) { value += quote; cursor++; } else { closed = true; break; }
        } else value += next;
      }
      if (!closed) invalid("Unclosed formula string or sheet name.");
      tokens.push({ kind: quote === '"' ? "string" : "quoted", text: value });
    } else {
      const rest = source.slice(cursor), numeric = number.exec(rest), identifier = (referenceStyle === "R1C1" ? r1c1.exec(rest) : undefined) ?? word.exec(rest) ?? /^\$[1-9]\d*/.exec(rest), failure = error.exec(rest);
      if (numeric) { tokens.push({ kind: "number", text: numeric[0] }); cursor += numeric[0].length; }
      else if (identifier) { tokens.push({ kind: "word", text: identifier[0] }); cursor += identifier[0].length; }
      else if (failure) { tokens.push({ kind: "error", text: failure[0].toUpperCase() }); cursor += failure[0].length; }
      else {
        const pair = source.slice(cursor, cursor + 2);
        if (["<=", ">=", "<>"].includes(pair)) { tokens.push({ kind: "operator", text: pair }); cursor += 2; }
        else if ("+-*/^&=<>%:(),!".includes(character)) { tokens.push({ kind: "operator", text: character }); cursor++; }
        else unsupported("Formula syntax is outside the implemented parser grammar.");
      }
    }
    checkLimit(tokens.length, 4096);
  }
  tokens.push({ kind: "end", text: "" }); return tokens;
}

function axisOperand(text: string, origin: FormulaOrigin, sheet?: string): AxisOperand | undefined {
  const match = /^(\$?)([A-Za-z]{1,3}|[1-9]\d*)$/.exec(text);
  if (!match) return undefined;
  const axis = /^[0-9]/.test(match[2]!) ? "row" : "column";
  let position = axis === "row" ? Number(match[2]) : 0;
  if (axis === "column") for (const letter of match[2]!.toUpperCase()) position = position * 26 + letter.charCodeAt(0) - 64;
  if (position > (axis === "row" ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS)) return undefined;
  return { kind: "axis", axis, position: { index: position - 1 - (match[1] ? 0 : origin[axis]), absolute: !!match[1] }, ...(sheet === undefined ? {} : { sheet }) };
}

function reference(text: string, origin: FormulaOrigin, sheet?: string): Operand {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(text);
  const scope = sheet === undefined ? {} : { sheet };
  if (!match) {
    if (text.startsWith("$")) { const axis = axisOperand(text, origin, sheet); if (axis) return axis; }
    if (text.includes("$")) invalid("Invalid formula reference.");
    return { kind: "name", name: text, ...scope };
  }
  let column = 0;
  for (const letter of match[2]!.toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
  const row = Number(match[4]);
  if (row > MAX_SHEET_ROWS || column > MAX_SHEET_COLUMNS) {
    // An out-of-grid identifier may be a name, but a $ reference cannot be one.
    if (match[1] || match[3]) invalid("Formula reference exceeds worksheet bounds.");
    return { kind: "name", name: text, ...scope };
  }
  return { kind: "cell", row: row - 1 - (match[3] ? 0 : origin.row), column: column - 1 - (match[1] ? 0 : origin.column),
    absoluteRow: Boolean(match[3]), absoluteColumn: Boolean(match[1]), ...scope };
}

/** Relative references are offsets from origin, making copied formulas share an AST. */
export function parseFormula(source: string, origin: FormulaOrigin): FormulaExpression { return parseWithStyle(source, origin, "A1"); }
/** R1C1 addresses produce the same relative-offset AST and share all grammar,
 * structured references and budgets with persisted A1 formulas. */
export function parseR1C1Formula(source: string, origin: FormulaOrigin): FormulaExpression { return parseWithStyle(source, origin, "R1C1"); }
function parseWithStyle(source: string, origin: FormulaOrigin, referenceStyle: "A1" | "R1C1"): FormulaExpression {
  checkLimit(source.length, 8192);
  if (!source.startsWith("=") || source.length < 2) invalid("Formula source must begin with equals.");
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(source)) invalid("Invalid formula control character.");
  if (!Number.isInteger(origin.row) || !Number.isInteger(origin.column) || origin.row < 0 || origin.column < 0
    || origin.row >= MAX_SHEET_ROWS || origin.column >= MAX_SHEET_COLUMNS) invalid("Invalid formula origin.");
  const tokens = tokenize(source, referenceStyle); let index = 0, nodes = 0;
  const ref = (text: string, sheet?: string): Operand => {
    if (referenceStyle === "A1") return reference(text, origin, sheet);
    const cell = /^R(\[[+-]?\d+\]|[1-9]\d*)?C(\[[+-]?\d+\]|[1-9]\d*)?$/i.exec(text);
    const axis = /^([RC])(\[[+-]?\d+\]|[1-9]\d*)?$/i.exec(text);
    const position = (part: string | undefined, maximum: number): FormulaAxisPosition => {
      const absolute = !!part && !part.startsWith("["), index = !part ? 0 : absolute ? Number(part) - 1 : Number(part.slice(1, -1));
      if (!Number.isSafeInteger(index) || absolute && (index < 0 || index >= maximum)) invalid("R1C1 reference exceeds worksheet bounds.");
      return { index, absolute };
    };
    if (cell) {
      const row = position(cell[1], MAX_SHEET_ROWS), column = position(cell[2], MAX_SHEET_COLUMNS);
      return { kind: "cell", row: row.index, column: column.index, absoluteRow: row.absolute, absoluteColumn: column.absolute, ...(sheet === undefined ? {} : { sheet }) };
    }
    if (axis) { const dimension = axis[1]!.toUpperCase() === "R" ? "row" : "column"; return { kind: "axis", axis: dimension, position: position(axis[2], dimension === "row" ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS), ...(sheet === undefined ? {} : { sheet }) }; }
    if (text.includes("$") || /^(?:R(?:\d+|\[.*\])(?:C(?:\d+|\[.*\])?)?|C(?:\d+|\[.*\]))$/i.test(text)) invalid("Invalid R1C1 reference.");
    return { kind: "name", name: text, ...(sheet === undefined ? {} : { sheet }) };
  };
  const numericTokens = new Map<Operand, string>();
  const peek = (): Token => tokens[index]!;
  const take = (): Token => tokens[index++]!;
  const spaces = (): void => { if (peek().kind === "space") take(); };
  const at = (text: string): boolean => peek().kind === "operator" && peek().text === text;
  const consume = (text: string): void => { if (!at(text)) invalid("Unexpected formula token."); take(); };
  const operand = (value: Operand): FormulaExpression => { if (value.kind === "axis") invalid("A whole-axis reference requires a range."); return value; };
  const isReference = (value: Operand): boolean => ["cell", "range", "axisRange", "name", "structured", "axis"].includes(value.kind)
    || value.kind === "error" && value.value === "#REF!" || value.kind === "binary" && [",", " ", ":"].includes(value.operator)
    || value.kind === "call" && ["INDEX", "OFFSET", "INDIRECT"].includes(value.name);
  const asAxis = (value: Operand): AxisOperand | undefined => value.kind === "axis" ? value : referenceStyle === "R1C1" ? undefined
    : value.kind === "name" && value.workbook === undefined ? axisOperand(value.name, origin, value.sheet) : value.kind === "number" ? axisOperand(numericTokens.get(value) ?? "", origin) : undefined;
  function sheetReference(input: string): Operand {
    const sheet = input.replace(/^\[0+\]/, "");
    if (/[\[\]\\/:?*\u0000-\u001f]/.test(sheet)) unsupported("Unsupported formula sheet reference.");
    spaces(); const target = take();
    if (target.kind === "error" && target.text === "#REF!") return { kind: "error", value: "#REF!" };
    if (target.kind === "number" && referenceStyle === "A1") { const axis = axisOperand(target.text, origin, sheet); if (!axis) invalid("Invalid whole-row reference."); return axis; }
    if (target.kind !== "word") invalid("Expected a formula reference after a sheet name."); return ref(target.text, sheet);
  }
  function expression(minimum: number, depth: number, union = true): Operand {
    checkLimit(depth, 128); checkLimit(++nodes, 4096);
    spaces(); const token = take(); let left: Operand;
    // The empty sheet denotes the current name-evaluation worksheet (!A1,
    // !Name). Keep it distinct from an ordinary unqualified identifier.
    if (token.kind === "operator" && token.text === "!") left = sheetReference("");
    else if (token.kind === "operator" && (token.text === "+" || token.text === "-")) left = { kind: "unary", operator: token.text, operand: operand(expression(70, depth + 1, union)) };
    else if (token.kind === "operator" && token.text === "(") { left = expression(0, depth + 1); spaces(); consume(")"); }
    else if (token.kind === "number") {
      const value = Number(token.text); if (!Number.isFinite(value)) invalid("Invalid formula numeric constant.");
      left = { kind: "number", value }; numericTokens.set(left, token.text);
    } else if (token.kind === "string") left = { kind: "string", value: token.text };
    else if (token.kind === "error") left = { kind: "error", value: token.text };
    else if (token.kind === "structured" && /^\[\d+\]$/.test(token.text) && (at("!")
      || ["word", "quoted"].includes(peek().kind) && tokens[index + 1]?.kind === "operator" && tokens[index + 1]?.text === "!")) {
      if (!/^\[0+\]$/.test(token.text)) unsupported("External workbook references are not supported.");
      if (at("!")) {
        take(); spaces(); const target = take();
        if (target.kind !== "word") invalid("A workbook-qualified name requires a name token.");
        const name = ref(target.text); if (name.kind !== "name") invalid("A workbook prefix without a sheet requires a defined name.");
        left = peek().kind === "structured" ? parseStructuredReference(take().text, name.name) : { ...name, workbook: 0 };
      } else { const sheet = take(); if (!sheet.text) invalid("An explicit worksheet qualifier cannot be empty."); consume("!"); left = sheetReference(sheet.text); }
    }
    else if (token.kind === "structured") left = parseStructuredReference(token.text);
    else if (token.kind === "word" || token.kind === "quoted") {
      if (at("!")) {
        if (!token.text) invalid("An explicit worksheet qualifier cannot be empty.");
        take(); left = sheetReference(token.text);
      } else if (token.kind === "quoted") invalid("Quoted formula sheet names require a reference.");
      else if (peek().kind === "structured") {
        if (token.text.includes("$")) invalid("Invalid table name.");
        left = parseStructuredReference(take().text, token.text);
      }
      else if (at("(")) {
        if (token.text.includes("$") || token.text.includes("\\")) invalid("Invalid formula function name.");
        take(); spaces(); const args: (FormulaExpression | null)[] = [];
        if (!at(")")) {
          while (true) {
            spaces(); args.push(at(",") || at(")") ? null : operand(expression(0, depth + 1, false))); spaces();
            checkLimit(args.length, 255);
            if (!at(",")) break;
            take();
          }
        }
        consume(")"); left = { kind: "call", name: token.text.toUpperCase(), arguments: args };
      } else if (/^(TRUE|FALSE)$/i.test(token.text)) left = { kind: "boolean", value: token.text.toUpperCase() === "TRUE" };
      else left = ref(token.text);
    } else invalid("Expected a formula operand.");

    while (true) {
      if (peek().kind === "space") {
        const next = tokens[index + 1]!;
        if (!(isReference(left) && (["word", "quoted", "structured"].includes(next.kind) || next.kind === "number" && tokens[index + 2]?.text === ":" || next.kind === "operator" && (next.text === "(" || next.text === "!")))) { take(); continue; }
      }
      if (at("%") && 60 >= minimum) { take(); left = { kind: "unary", operator: "%", operand: operand(left) }; continue; }
      const operator = peek().kind === "operator" || peek().kind === "space" ? peek().text : "", precedence = operators[operator];
      if (precedence === undefined || precedence < minimum || operator === "," && !union) break;
      take(); const right = expression(precedence + 1, depth + 1, union);
      if (operator === ":") {
        if (left.kind === "cell" && right.kind === "cell") {
          if (right.sheet !== undefined && right.sheet.toLowerCase() !== left.sheet?.toLowerCase()) unsupported("Three-dimensional formula ranges are not supported.");
          left = { kind: "range", from: left, to: right.sheet === undefined && left.sheet !== undefined ? { ...right, sheet: left.sheet } : right };
        } else {
          const from = asAxis(left), to = asAxis(right);
          if (from && to && from.axis === to.axis) {
            if (to.sheet !== undefined && to.sheet.toLowerCase() !== from.sheet?.toLowerCase()) unsupported("Three-dimensional formula ranges are not supported.");
            left = { kind: "axisRange", axis: from.axis, from: from.position, to: to.position, ...(from.sheet === undefined ? {} : { sheet: from.sheet }) };
          } else {
            if (!isReference(left) || !isReference(right) || left.kind === "axis" || right.kind === "axis") unsupported("Unsupported formula range endpoints.");
            if (left.kind !== "error" && "sheet" in right && right.sheet !== undefined && (!("sheet" in left) || left.sheet?.toLowerCase() !== right.sheet.toLowerCase())) unsupported("Three-dimensional formula ranges are not supported.");
            left = { kind: "binary", operator, left, right };
          }
        }
      } else {
        if ([",", " "].includes(operator) && (!isReference(left) || !isReference(right))) invalid("Union/intersection operands must be references.");
        left = { kind: "binary", operator, left: operand(left), right: operand(right) };
      }
    }
    return left;
  }
  const ast = operand(expression(0, 0));
  if (peek().kind !== "end") unsupported("Unconsumed formula syntax is not supported.");
  // A long left-associative chain is shallow for the parser but deep as an AST.
  // Check it before recursive freezing or serialization can use the JS stack.
  const pending = [{ node: ast, depth: 0 }];
  while (pending.length) {
    const { node, depth } = pending.pop()!; checkLimit(depth, 128);
    const children = node.kind === "binary" ? [node.left, node.right] : node.kind === "unary" ? [node.operand]
      : node.kind === "call" ? node.arguments.filter((arg): arg is FormulaExpression => arg !== null) : node.kind === "range" ? [node.from, node.to] : [];
    for (const child of children) pending.push({ node: child, depth: depth + 1 });
  }
  return freeze(ast);
}

export interface FormulaTemplatePool {
  readonly size: number;
  readonly keyCodeUnits: number;
  intern(source: string, origin: FormulaOrigin): number;
  include(expression: FormulaExpression): number;
  get(id: number): FormulaExpression | undefined;
}
/** Per-workbook derived cache, bounded independently from source bytes. */
export function createFormulaTemplatePool(maxTemplates = 250_000, maxKeyCodeUnits = 32 * 1024 * 1024): FormulaTemplatePool {
  if (![maxTemplates, maxKeyCodeUnits].every(value => Number.isSafeInteger(value) && value > 0)) invalid("Invalid formula template budget.");
  const ids = new Map<string, number>(), expressions: FormulaExpression[] = []; let keyCodeUnits = 0;
  function include(ast: FormulaExpression): number {
    const key = JSON.stringify(ast), existing = ids.get(key);
    if (existing !== undefined) return existing;
    checkLimit(expressions.length + 1, maxTemplates); checkLimit(keyCodeUnits + key.length, maxKeyCodeUnits);
    const id = expressions.length; ids.set(key, id); expressions.push(ast); keyCodeUnits += key.length; return id;
  }
  return Object.freeze({
    get size() { return expressions.length; }, get keyCodeUnits() { return keyCodeUnits; },
    intern(source: string, origin: FormulaOrigin): number {
      return include(parseFormula(source, origin));
    },
    include,
    get(id: number) { return expressions[id]; },
  });
}
