import { afterEach, expect, test } from "vitest";
import { generate, parse } from "css-tree";
import { createDocumentSession, createWorkbook, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookSession, WorkbookStyleApplication } from "../src/index.js";

const flags: WorkbookStyleApplication = { font: false, fill: true, border: false, alignment: true, numberFormat: false, protection: false };
const sessions: WorkbookSession[] = [], decoder = new TextDecoder(), encoder = new TextEncoder();
function declarations(css: string, selector: string) {
  const ast = parse(css, { context: "stylesheet" }); if (ast.type !== "StyleSheet") throw Error("Invalid test CSS");
  return Object.fromEntries(ast.children.toArray().flatMap(rule => rule.type === "Rule" && generate(rule.prelude) === selector
    ? rule.block.children.toArray().flatMap(node => node.type === "Declaration" ? [[node.property, generate(node.value) + (node.important ? "!important" : "")]] : []) : []));
}
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
function apply(session: WorkbookSession, operations: DocumentOperation[]) {
  const manifest = session.snapshot().manifest;
  return session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations }));
}
function example(extraCss = "") {
  const files = new Map(createWorkbook({ documentId: "style-editing" })), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.cellStyles = { definitions: [{ styleId: "accent", cssClass: "accent", numberFormatId: "general", apply: { ...flags } }],
    names: [{ name: "汇总😀", styleId: "accent", hidden: true, customBuiltin: false }],
    bindings: [{ cssClass: "follows", styleId: "accent", apply: { ...flags } }, { cssClass: "fixed", styleId: "accent", apply: { ...flags, font: true, numberFormat: true } }] };
  metadata.workbook.sheets[0]!.rowDimensions = [{ from: 1, to: 3, cssClass: "follows", numberFormatId: "general" }];
  metadata.workbook.sheets[0]!.columnDimensions = [{ from: 1, to: 3, unit: "characters", cssClass: "follows", numberFormatId: "general" }];
  files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")!) + ".accent{color:red;background-color:blue}.follows{color:red!important;background-color:yellow}.fixed{color:green}.shared,.other{border-top-style:solid}" + extraCss));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" })); sessions.push(session);
  apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 123 } }, cssClass: "follows" },
    { row: 0, column: 1, source: { kind: "formula", expression: "=A1+1" }, cssClass: "fixed", formatId: "general" },
    { row: 1, column: 0, source: { kind: "value", value: { type: "text", value: "中文禰󠄀😀" } } },
    { row: 2, column: 0, source: { kind: "formula", expression: "=A1+2" }, formatId: "general" },
  ] }]); session.clearHistory(); return session;
}
const update = (numberFormatId = "general"): DocumentOperation => ({ op: "setWorkbookCellStyleDefinition", definition: { styleId: "accent", numberFormatId, apply: { ...flags } }, properties: { color: "purple", "background-color": "black", "border-top-style": "double" } });

test("named-style edits propagate each enabled group and sparse number-format sources, retain overrides and reopen/undo atomically", () => {
  const session = example(), before = session.snapshot(), formulaBody = session.getCell("sheet-1", 2, 0)!.html;
  apply(session, [{ op: "setWorkbookNumberFormat", format: { formatId: "money", code: "0.00" } }, update("money")]);
  const snapshot = session.snapshot(), css = snapshot.styles[0]!.css;
  expect(declarations(css, ".accent")).toMatchObject({ color: "purple", "background-color": "black", "border-top-style": "double", "--office-cell-foreground-color": "purple", "--office-cell-fill-color": "black" });
  expect(declarations(css, ".follows")).toMatchObject({ color: "purple!important", "background-color": "yellow", "border-top-style": "double", "--office-cell-foreground-color": "purple" });
  expect(css).toContain(".fixed{color:green;border-top-style:double}");
  expect(session.getCell("sheet-1", 0, 0)?.formatId).toBe("money"); expect(session.getCell("sheet-1", 0, 1)?.formatId).toBe("general");
  expect(session.getCell("sheet-1", 1, 0)?.formatId).toBeUndefined(); expect(session.getCell("sheet-1", 2, 0)?.formatId).toBe("money");
  expect(snapshot.metadata.workbook.sheets[0]!.rowDimensions).toMatchObject([{ from: 1, to: 3, numberFormatId: "money" }]);
  expect(snapshot.metadata.workbook.sheets[0]!.columnDimensions).toMatchObject([{ from: 1, to: 3, numberFormatId: "money" }]);
  expect(session.getCell("sheet-1", 2, 0)!.html.replace(' data-format-id="money"', ' data-format-id="general"')).toBe(formulaBody);
  expect(snapshot.manifest.chunks).toHaveLength(1); expect(session.history().undoSteps).toBe(1);
  expect(openDocument(session.serialize(), { kind: "workbook" }).snapshot()).toEqual(snapshot);
  session.undo(); expect(session.snapshot().metadata).toEqual(before.metadata); expect(session.snapshot().styles).toEqual(before.styles);
  expect(session.snapshot().manifest.chunks).toEqual(before.manifest.chunks); session.redo(); expect(session.snapshot().metadata).toEqual(snapshot.metadata);
});

test("local class edits protect their whole named-style group, preserve shared selectors and reject invalid later source atomically", () => {
  const session = example();
  apply(session, [{ op: "setWorkbookStyleClass", cssClass: "follows", properties: { color: "blue" } }, update(),
    { op: "setWorkbookStyleClass", cssClass: "shared", properties: { "border-top-style": "double" } }]);
  expect(declarations(session.snapshot().styles[0]!.css, ".follows")).toMatchObject({ color: "blue!important", "background-color": "yellow", "border-top-style": "double", "--office-cell-foreground-color": "blue" });
  expect(session.snapshot().styles[0]!.css).toContain(".other{border-top-style:solid}.shared{border-top-style:double}");
  expect(session.snapshot().metadata.workbook.cellStyles!.bindings[0]!.apply.font).toBe(true);
  const before = session.snapshot();
  for (const properties of [{ color: "url(https://private.invalid/)" }, { "--unknown": "value" }, { width: "12pt" }]) {
    expect(() => apply(session, [update(), { op: "setWorkbookStyleClass", cssClass: "shared", properties }])).toThrow();
    expect(session.snapshot()).toEqual(before);
  }
});

test("definition creation allocates source identity; aliases are independently editable and removing a name retains associated appearance", () => {
  const session = example(), before = session.snapshot();
  apply(session, [{ op: "setWorkbookCellStyleDefinition", definition: { styleId: "new", numberFormatId: "general", apply: { ...flags } }, properties: { color: "#123456", "--office-font-east-asia": '"Noto Sans CJK SC"' } },
    { op: "setWorkbookCellStyleName", name: { name: "新样式", styleId: "new", hidden: false, customBuiltin: false } },
    { op: "setWorkbookCellStyleName", name: { name: "别名", styleId: "new", hidden: true, customBuiltin: false } },
    { op: "removeWorkbookCellStyleName", name: "新样式" }]);
  const styles = session.snapshot().metadata.workbook.cellStyles!, definition = styles.definitions.find(style => style.styleId === "new")!;
  expect(styles.names).toEqual([before.metadata.workbook.cellStyles!.names[0], { name: "别名", styleId: "new", hidden: true, customBuiltin: false }]);
  expect(definition.cssClass).toMatch(/^workbook-style-/); expect(session.snapshot().styles[0]!.css).toContain("." + definition.cssClass + "{color:#123456;--office-font-east-asia:");
  expect(() => apply(session, [{ op: "setWorkbookStyleClass", cssClass: "accent", properties: { color: "black" } }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(openDocument(session.serialize(), { kind: "workbook" }).snapshot()).toEqual(session.snapshot());
  session.undo(); expect(session.snapshot().metadata).toEqual(before.metadata); expect(session.snapshot().styles).toEqual(before.styles);
});

test("same-value named updates retain redo and invalid definitions cannot escape an isolated candidate", async () => {
  const session = example(); apply(session, [update()]); session.undo(); const before = session.snapshot();
  const receipt = apply(session, [{ op: "setWorkbookCellStyleDefinition", definition: { styleId: "accent", numberFormatId: "general", apply: { ...flags } }, properties: {} }]);
  expect(receipt.revision).toBe(before.manifest.revision); expect(session.history()).toMatchObject({ undoSteps: 0, redoSteps: 1 });
  const abort = { aborted: false }, manifest = session.snapshot().manifest;
  await expect(session.prepareAsync({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision,
    operations: [update(), { op: "setWorkbookCellStyleName", name: { name: "invalid", styleId: "missing", hidden: false, customBuiltin: false } }] }, {},
    { signal: abort, yield: async () => { abort.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(session.snapshot()).toEqual(before); expect(session.history().redoSteps).toBe(1);
});

test("range formatting isolates existing shared and named classes, protects only changed groups and preserves rich bodies and merged covered formats", () => {
  const session = example(); apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 1, column: 0,
    source: { kind: "richText", html: '<span class="fixed">中文禰󠄀</span><br>😀' } }] },
    { op: "mergeWorkbookCells", sheetId: "sheet-1", range: { from: { row: 1, column: 0 }, to: { row: 1, column: 1 } } }]);
  session.clearHistory(); const before = session.snapshot(), body = session.getCell("sheet-1", 1, 0)!.html;
  apply(session, [{ op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 1, column: 1 } }],
    properties: { color: "orange", "--office-cell-text-rotation": "45" } }, update()]);
  const current = session.snapshot(), changed = current.metadata.workbook.cellStyles!.bindings.slice(2);
  expect(changed).toHaveLength(2); expect(changed.every(binding => binding.apply.font && binding.apply.alignment)).toBe(true);
  expect(changed[0]!.apply.fill).toBe(true); expect(changed[0]!.apply.border).toBe(false);
  for (const binding of changed) expect(declarations(current.styles[0]!.css, "td." + binding.cssClass)).toMatchObject({ color: "orange!important", "--office-cell-text-rotation": "45!important", "--office-cell-foreground-color": "orange!important" });
  expect(session.getCell("sheet-1", 1, 1)).toMatchObject({ value: { type: "blank" }, cssClass: changed[0]!.cssClass });
  expect(session.getCell("sheet-1", 1, 0)!.html.replace(' class="' + changed[0]!.cssClass + '"', "")).toBe(body);
  expect(current.styles[0]!.css).toContain("color:orange!important"); expect(session.history().undoSteps).toBe(1);
  expect(openDocument(session.serialize(), { kind: "workbook" }).snapshot()).toEqual(current);
  session.undo(); expect(session.snapshot().metadata).toEqual(before.metadata); expect(session.snapshot().styles).toEqual(before.styles);
  expect(session.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("range formatting rejects foreign formats, unsafe appearance and oversized expansion without history or source changes", async () => {
  const session = example(), before = session.snapshot(), range = { from: { row: 0, column: 0 }, to: { row: 300, column: 0 } };
  for (const operation of [
    { op: "formatWorkbookRanges" as const, sheetId: "sheet-1", ranges: [range], properties: { color: "red" }, numberFormatId: "missing" },
    { op: "formatWorkbookRanges" as const, sheetId: "sheet-1", ranges: [range], properties: { color: "url(https://private.invalid)" } },
    { op: "formatWorkbookRanges" as const, sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16383 } }], properties: { color: "red" } },
    ...[{ "--office-cell-fill-color": "invalid-color" }, { "--office-cell-text-rotation": "181" }, { "--office-cell-gradient-stops": "100000" }, { "--office-cell-border-left-kind": "unknown" },
      { "--office-cell-border-unknown-kind": "thin" }, { "--office-cell-border-vertical-kind": "thin" }, { "--office-cell-locked": "false" }].map(properties => ({ op: "formatWorkbookRanges" as const, sheetId: "sheet-1", ranges: [range], properties })),
  ]) { expect(() => apply(session, [operation])).toThrow(); expect(session.snapshot()).toEqual(before); }
  const abort = { aborted: false }, manifest = before.manifest;
  await expect(session.prepareAsync({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [{ op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [range], properties: { color: "red" } }] }, {},
    { signal: abort, yield: async () => { abort.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(session.snapshot()).toEqual(before); expect(session.history().undoSteps).toBe(0);
});

test("local aliases preserve class reuse on different elements and repeated formatting keeps selector specificity bounded", () => {
  const session = example("td.follows.follows span.follows{font-style:italic}"), range = { from: { row: 1, column: 0 }, to: { row: 1, column: 0 } };
  apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 1, column: 0, source: { kind: "richText", html: '<span class="follows">禰󠄀😀</span>' } }] }]);
  const original = session.getCell("sheet-1", 1, 0)!;
  for (const color of ["cyan", "magenta", "yellow"]) {
    apply(session, [{ op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [range], properties: { "background-color": color } }]);
    const cell = session.getCell("sheet-1", 1, 0)!, currentClass = cell.cssClass!, css = session.snapshot().styles[0]!.css;
    expect(css).toContain("td." + currentClass + "." + currentClass + " span.follows");
    expect(cell.html.slice(cell.html.indexOf(">") + 1)).toBe(original.html.slice(original.html.indexOf(">") + 1));
    const selector = "td" + ("." + currentClass).repeat(4);
    expect(declarations(css, selector)).toMatchObject({ "background-color": color + "!important" });
    expect(css).not.toContain(selector + "." + currentClass + "{");
  }
});
