import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { openSdkHost } from "./sdk-host.js";
import type * as fixture from "./fixtures/workbook-input.js";
import type * as editorFixture from "./fixtures/workbook-editor.js";
import { scalarWorkbook } from "../../packages/core/tests/workbook-scalar-fixture.js";
import { encodeTestPackage } from "./package-transport.js";
import { createDocumentSession, openDocument } from "@doffi/core";
import { importXlsx, writeDocumentZip } from "@doffi/office";
import type * as remoteFixture from "./fixtures/workbook-remote-text-host.js";
declare global { interface Window { workbookInput: typeof fixture } }
declare global { interface Window { workbookEditor: typeof editorFixture } }
declare global { interface Window { workbookRemoteText: typeof remoteFixture } }
let code = "", editorCode = "";
let remoteCode = "", remoteWorker = "";
test.beforeAll(async () => {
  const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/workbook-input.ts", import.meta.url)), formats: ["iife"], name: "WorkbookInput" } } });
  const output = Array.isArray(result) ? result.flatMap(part => part.output) : "output" in result ? result.output : [];
  code = output.flatMap(part => part.type === "chunk" ? [part.code] : []).join("\n");
  const editor = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/workbook-editor.ts", import.meta.url)), formats: ["iife"], name: "WorkbookEditorFixture" } } });
  const chunks = Array.isArray(editor) ? editor.flatMap(part => part.output) : "output" in editor ? editor.output : [];
  editorCode = chunks.flatMap(part => part.type === "chunk" ? [part.code] : []).join("\n");
  for (const name of ['workbook-remote-text-host', 'workbook-session-worker']) {
    const built = await build({ configFile: false, logLevel: 'silent', build: { write: false, target: 'es2022', lib: { entry: fileURLToPath(new URL(`./fixtures/${name}.ts`, import.meta.url)), formats: ['iife'], name: 'RemoteInputFixture' } } });
    const parts = Array.isArray(built) ? built.flatMap(part => part.output) : 'output' in built ? built.output : [];
    const value = parts.flatMap(part => part.type === 'chunk' ? [part.code] : []).join('\n');
    if (name === 'workbook-session-worker') remoteWorker = value; else remoteCode = value;
  }
});
function remoteSource() { return Buffer.from(writeDocumentZip(scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 1 } }]).serialize())).toString('base64'); }
test('Worker native text input waits for its transaction before navigation and retains literal types and history', async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: remoteCode });
  await page.evaluate(({ code, source }) => window.workbookRemoteText.mountRemoteText(code, source), { code: remoteWorker, source: remoteSource() });
  const input = page.getByRole('textbox', { name: 'Worker 单元格输入', exact: true });
  await input.pressSequentially('00123'); expect(await page.evaluate(() => window.workbookRemoteText.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, input: { dirty: true } });
  await input.press('Enter'); await page.evaluate(() => window.workbookRemoteText.settled());
  expect(await page.evaluate(() => window.workbookRemoteText.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, source: { text: '00123' }, input: { dirty: false }, navigation: ['down'], errors: [] });
  await page.evaluate(() => window.workbookRemoteText.undo()); await expect(input).toHaveValue('1');
  await page.evaluate(() => window.workbookRemoteText.setMode('text')); await input.fill('12'); await input.press('Control+Enter'); await page.evaluate(() => window.workbookRemoteText.settled());
  await expect(input).toHaveValue("'12"); expect(await page.evaluate(() => window.workbookRemoteText.inspect())).toMatchObject({ source: { text: "'12" }, errors: [] });
  expect(await page.evaluate(() => window.workbookRemoteText.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, hostUsable: { type: 'number', value: 42 } });
});
test('Worker native editing preserves original CRLF and variation sequences', async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: remoteCode });
  await page.evaluate(({ code, source }) => window.workbookRemoteText.mountRemoteText(code, source), { code: remoteWorker, source: remoteSource() });
  await page.evaluate(() => window.workbookRemoteText.external('ABC\r\n中文禰󠄀'));
  const input = page.getByRole('textbox', { name: 'Worker 单元格输入', exact: true }); await input.press('Control+End'); await input.pressSequentially('尾'); await input.press('Control+Enter');
  await page.evaluate(() => window.workbookRemoteText.settled());
  expect(await page.evaluate(() => window.workbookRemoteText.inspect())).toMatchObject({ source: { text: 'ABC\r\n中文禰󠄀尾' }, errors: [] });
  await page.evaluate(() => window.workbookRemoteText.releaseBinding()); await input.fill('ignored'); await input.press('Enter');
  expect(await page.evaluate(() => window.workbookRemoteText.hostStillUsable())).toBe(true);
  expect(await page.evaluate(() => window.workbookRemoteText.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true });
});
test('Worker real IME retains a conflicting draft until explicit cancellation and reload', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'CDP composition is Chromium-specific.');
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: remoteCode });
  await page.evaluate(({ code, source }) => window.workbookRemoteText.mountRemoteText(code, source), { code: remoteWorker, source: remoteSource() });
  const input = page.getByRole('textbox', { name: 'Worker 单元格输入', exact: true }), cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', { text: '中文', selectionStart: 2, selectionEnd: 2 });
  await page.evaluate(() => window.workbookRemoteText.external('external'));
  await cdp.send('Input.insertText', { text: '中文😀' }); await input.press('Enter'); await page.evaluate(() => window.workbookRemoteText.settled());
  expect(await page.evaluate(() => window.workbookRemoteText.inspect())).toMatchObject({ revision: 1, source: { text: 'external' }, input: { text: '中文😀', dirty: true, selection: { baseRevision: 0 } }, errors: [{ code: 'RevisionConflict' }], navigation: [] });
  await input.press('Escape'); await page.evaluate(() => window.workbookRemoteText.settled()); await expect(input).toHaveValue('external');
  await page.evaluate(() => window.workbookRemoteText.finish());
});
test('Worker typing while awaiting a real commit receipt preserves the next draft and delays navigation', async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: remoteCode });
  await page.evaluate(({ code, source }) => window.workbookRemoteText.mountRemoteText(code, source), { code: remoteWorker, source: remoteSource() });
  const input = page.getByRole('textbox', { name: 'Worker 单元格输入', exact: true });
  await input.fill('first'); await page.evaluate(() => window.workbookRemoteText.holdCommit()); await input.press('Enter');
  await expect.poll(() => page.evaluate(() => window.workbookRemoteText.commitHeld())).toBe(true); await input.pressSequentially('tail');
  expect(await page.evaluate(() => window.workbookRemoteText.inspect())).toMatchObject({ revision: 1, source: { text: 'first' }, input: { text: 'firsttail', dirty: true, applying: true }, navigation: [] });
  await page.evaluate(() => window.workbookRemoteText.releaseCommit()); await page.evaluate(() => window.workbookRemoteText.settled());
  expect(await page.evaluate(() => window.workbookRemoteText.inspect())).toMatchObject({ history: { undoSteps: 1 }, input: { text: 'firsttail', dirty: true, applying: false, selection: { baseRevision: 1 } }, navigation: [] });
  await input.press('Enter'); await page.evaluate(() => window.workbookRemoteText.settled());
  expect(await page.evaluate(() => window.workbookRemoteText.inspect())).toMatchObject({ revision: 2, source: { text: 'firsttail' }, history: { undoSteps: 2 }, input: { dirty: false }, navigation: ['down'], errors: [] });
  await page.evaluate(() => window.workbookRemoteText.finish());
});
function editorSource(rtl = false, structure = true) {
  return encodeTestPackage(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 1 } }, { row: 0, column: 1, value: { type: "number", value: 999 }, formula: "=A1+1" }], metadata => {
    const sheet = metadata.workbook.sheets[0]!; sheet.rightToLeft = rtl; sheet.defaultRowHeightCustom = true;
    if (structure) {
      sheet.rowDimensions = [{ from: 1, to: 1000, hidden: true }]; sheet.columnDimensions = [{ from: 1, to: 10, unit: "characters", hidden: true }];
      sheet.merges = [{ from: { row: 1001, column: 0 }, to: { row: 1002, column: 2 } }];
    }
    metadata.workbook.sheets.push({ ...sheet, sheetId: "second", name: "Second", rowDimensions: [], columnDimensions: [], merges: [] });
  }).serialize());
}
test("mounted workbook formula input, calculation, address selection and toolbar history share source", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource());
  await page.evaluate(() => window.workbookEditor.select(0)); const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
  await input.pressSequentially("5"); expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 } });
  await input.press("Enter"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, cleanSource: true, errors: [],
    calculated: { state: "valid", value: { type: "number", value: 6 } }, selection: { anchor: { ranges: [{ from: { row: 1001, column: 0 }, to: { row: 1002, column: 2 } }] } } });
  const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }); await address.fill("A1"); await address.press("Enter"); await page.evaluate(() => window.workbookEditor.settled()); await expect(input).toHaveValue("5");
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 0, redoSteps: 1 }, calculated: { value: { type: "number", value: 2 } }, errors: [] });
  await page.getByRole("button", { name: "重做", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled()); await expect(input).toHaveValue("5");
});

test("mounted direct range appearance edits native rich text with current colors, advanced fill and borders, and restores source in one undo", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const source = await importXlsx(readFileSync("examples/xlsx/01-values-text.xlsx"), { documentId: "native-format-ui" });
  await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(source)); await page.evaluate(() => window.workbookEditor.select(5));
  const original = await page.evaluate(() => window.workbookEditor.cell("sheet-1", 5, 0));
  await page.getByText("设置单元格格式", { exact: true }).click();
  await page.getByLabel("文字颜色", { exact: true }).fill("#cc2200"); await page.getByLabel("字重", { exact: true }).selectOption("normal");
  await page.getByLabel("斜体", { exact: true }).selectOption("normal"); await page.getByLabel("字号（磅）", { exact: true }).fill("16");
  await page.getByLabel("中文字体", { exact: true }).fill("Noto Sans CJK SC"); await page.getByLabel("水平对齐", { exact: true }).selectOption("center");
  await page.getByLabel("自动换行", { exact: true }).selectOption("true"); await page.getByLabel("下划线", { exact: true }).selectOption("doubleAccounting");
  await page.getByLabel("填充方式", { exact: true }).selectOption("linear"); await page.getByLabel("填充前景色", { exact: true }).fill("#cceeff");
  await page.getByLabel("填充背景色", { exact: true }).fill("#ffffff"); await page.getByLabel("渐变角度", { exact: true }).fill("45");
  await page.getByLabel("边框位置", { exact: true }).selectOption("diagonalBoth"); await page.getByLabel("边框线型", { exact: true }).selectOption("double");
  await page.getByRole("button", { name: "应用单元格格式", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  const changed = await page.evaluate(() => window.workbookEditor.cell("sheet-1", 5, 0)); expect(changed?.text).toBe(original?.text);
  expect(changed?.html.slice(changed.html.indexOf(">") + 1)).toBe(original?.html.slice(original.html.indexOf(">") + 1));
  const styles = await page.evaluate(() => window.workbookEditor.displayedTextStyles(5, 0));
  expect(styles.filter(style => style.text).some(style => style.color === "rgb(204, 34, 0)" && style.weight === "400" && style.italic === "normal")).toBe(true);
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 5, 0))).toEqual(original);
});

test("mounted direct selected text fonts map CRLF and literal-prefix offsets, preserve other characters, selection and one undo", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const s = createDocumentSession(scalarWorkbook([])), manifest = s.snapshot().manifest;
  s.commit(s.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 0, column: 0, source: { kind: "richText", html: '<span class="cell-default">=A&#13;<br>禰󠄀😀Z</span>' } },
  ] }] })); const entries = encodeTestPackage(s.serialize()); s.dispose();
  await page.evaluate(entries => window.workbookEditor.mountObserved(entries), entries); await page.evaluate(() => window.workbookEditor.select(0));
  const original = await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 0)), input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
  await page.getByText("设置单元格格式", { exact: true }).click(); await page.getByLabel("文字颜色", { exact: true }).fill("#cc2200");
  await page.getByLabel("字重", { exact: true }).selectOption("bold");
  await input.evaluate(element => { const field = element as HTMLTextAreaElement, from = field.value.indexOf("禰"); field.setSelectionRange(from, from + 5); field.dispatchEvent(new Event("select", { bubbles: true })); });
  await page.getByRole("button", { name: "应用选中文字字体", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await input.evaluate(element => { const field = element as HTMLTextAreaElement; return field.value.slice(field.selectionStart, field.selectionEnd); })).toBe("禰󠄀😀");
  const changed = await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 0)); expect(changed?.text).toBe("=A\r\n禰󠄀😀Z"); expect(changed?.html).toContain("=A&#13;<br><span class=\"workbook-run-");
  expect((await page.evaluate(() => window.workbookEditor.displayedTextStyles(0, 0))).some(style => style.text?.includes("禰") && style.color === "rgb(204, 34, 0)" && style.weight === "700")).toBe(true);
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 0))).toEqual(original);
});

test("mounted direct range appearance preserves nested selector styling while repeatedly changing an inherited named cell style", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const files = new Map(scalarWorkbook([], metadata => {
    metadata.workbook.cellStyles = { definitions: [{ styleId: "base", cssClass: "cell-default", numberFormatId: "general",
      apply: { font: true, fill: true, border: true, alignment: true, numberFormat: true, protection: true } }], names: [], bindings: [] };
  }).serialize());
  files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + "td.cell-default.cell-default span.cell-default{font-style:italic}"));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" })), manifest = session.snapshot().manifest;
  session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 0, column: 0, source: { kind: "richText", html: '<span class="cell-default">继承禰󠄀😀</span>' } },
  ] }] })); const entries = encodeTestPackage(session.serialize()); session.dispose();
  await page.evaluate(entries => window.workbookEditor.mountObserved(entries), entries); await page.evaluate(() => window.workbookEditor.select(0));
  expect((await page.evaluate(() => window.workbookEditor.displayedTextStyles(0, 0))).some(style => style.italic === "italic")).toBe(true);
  await page.getByText("设置单元格格式", { exact: true }).click(); await page.getByLabel("填充方式", { exact: true }).selectOption("solid");
  for (const color of ["cyan", "magenta", "yellow"]) {
    await page.getByLabel("填充前景色", { exact: true }).fill(color);
    await page.getByRole("button", { name: "应用单元格格式", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
    expect((await page.evaluate(() => window.workbookEditor.displayedTextStyles(0, 0))).some(style => style.italic === "italic")).toBe(true);
  }
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 3 }, errors: [] });
});

test("mounted direct private clipboard delays cutting, explicitly overwrites and restores formulas in one undo", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(scalarWorkbook([
    { row: 0, column: 0, value: { type: "number", value: 7 } }, { row: 0, column: 1, value: { type: "number", value: 999 }, formula: "=A1+1" },
    { row: 2, column: 0, value: { type: "text", value: "occupied" } }, { row: 0, column: 3, value: { type: "number", value: 999 }, formula: "=SUM(A1:B1)" },
  ]).serialize()));
  await page.evaluate(() => window.workbookEditor.select(0, 0, 0, 1));
  await page.getByRole("button", { name: "剪切选区", exact: true }).click();
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 } });
  await expect(page.getByLabel("区域粘贴方式", { exact: true })).toBeDisabled();
  await page.evaluate(() => window.workbookEditor.select(2));
  await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, errors: ["InvalidOperation"] });
  await page.getByLabel("剪切时允许覆盖已有内容", { exact: true }).check();
  await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => [window.workbookEditor.cell("sheet-1", 0, 0), window.workbookEditor.cell("sheet-1", 2, 0),
    window.workbookEditor.cell("sheet-1", 2, 1)?.formula?.expression, window.workbookEditor.cell("sheet-1", 0, 3)?.formula?.expression]))
    .toMatchObject([undefined, { value: { type: "number", value: 7 } }, "=(A3+1)", "=SUM(A3:B3)"]);
  await expect(page.getByRole("button", { name: "粘贴选区", exact: true })).toBeDisabled();
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 2, 0))).toMatchObject({ value: { type: "text", value: "occupied" } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 3)?.formula?.expression)).toBe("=SUM(A1:B1)");
  await page.evaluate(() => window.workbookEditor.select(0, 0, 0, 1)); await page.getByRole("button", { name: "复制选区", exact: true }).click();
  await page.evaluate(() => window.workbookEditor.select(4)); await page.getByLabel("区域粘贴方式", { exact: true }).selectOption("values");
  await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 4, 1))).toMatchObject({ value: { type: "number", value: 8 } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 4, 1)?.formula)).toBeUndefined();
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await input.fill("draft");
  await expect(page.getByRole("button", { name: "复制选区", exact: true })).toBeDisabled();
  await input.press("Escape"); expect(await page.evaluate(() => window.workbookEditor.dispose())).toMatchObject({ usable: true });
});

test("mounted direct sorting keeps headers/hidden records, copies current-key records and restores one history step", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const source = scalarWorkbook([
    { row: 0, column: 0, value: { type: "text", value: "Key" } }, { row: 0, column: 1, value: { type: "text", value: "Label" } },
    ...[3, 100, 1, 2].map((value, index) => ({ row: index + 1, column: 0, value: { type: "number" as const, value: 999 }, formula: "=" + value })),
    ...["three", "hidden", "one", "two"].map((value, index) => ({ row: index + 1, column: 1, value: { type: "text" as const, value } })),
    { row: 5, column: 3, value: { type: "number", value: 999 }, formula: "=A2" },
  ], metadata => { const sheet = metadata.workbook.sheets[0]!; sheet.defaultRowHeightCustom = true; sheet.rowDimensions = [{ from: 2, to: 2, hidden: true }]; });
  await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(source.serialize()));
  await page.evaluate(() => window.workbookEditor.select(0, 0, 4, 1)); await page.getByLabel("首行为表头", { exact: true }).check();
  await page.getByRole("button", { name: "升序排序", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => [0, 1, 2, 3, 4].map(row => window.workbookEditor.cell("sheet-1", row, 1)?.text))).toEqual(["Label", "one", "hidden", "two", "three"]);
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 5, 3)?.formula?.expression)).toBe("=A2");
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [], cleanSource: true });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => [1, 2, 3, 4].map(row => window.workbookEditor.cell("sheet-1", row, 1)?.text))).toEqual(["three", "hidden", "one", "two"]);
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await input.fill("draft"); await expect(page.getByRole("button", { name: "降序排序", exact: true })).toBeDisabled();
  await input.press("Escape"); await page.evaluate(() => window.workbookEditor.settled()); await page.getByRole("button", { name: "降序排序", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => [1, 2, 3, 4].map(row => window.workbookEditor.cell("sheet-1", row, 1)?.text))).toEqual(["three", "hidden", "two", "one"]);
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1, redoSteps: 0 }, errors: [] });
});
test("mounted workbook merge requires an explicit clearing choice and unmerge/undo retain the source", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const entries = encodeTestPackage(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 7 } }, { row: 0, column: 1, value: { type: "number", value: 3 } }]).serialize());
  await page.evaluate(entries => window.workbookEditor.mount(entries), entries); await page.evaluate(() => window.workbookEditor.select(0, 0, 0, 1));
  const area = { from: { row: 0, column: 0 }, to: { row: 0, column: 1 } };
  await page.getByRole("button", { name: "合并单元格", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await expect(page.getByRole("alert")).toContainText("合并范围包含其他内容");
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, merges: [], errors: ["InvalidOperation"] });
  await page.getByRole("button", { name: "合并并清除覆盖内容", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, merges: [area] });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 1))).toBeUndefined();
  await page.getByRole("button", { name: "取消合并", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 2, merges: [] });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ merges: [area] });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ merges: [], history: { undoSteps: 0, redoSteps: 2 }, cells: [{ value: { type: "number", value: 7 } }, { value: { type: "number", value: 3 } }] });
});

test("mounted direct merged clipboard and both fill axes retain complete templates and undo", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const area = (row: number, column: number, toRow: number, toColumn: number) => ({ from: { row, column }, to: { row: toRow, column: toColumn } });
  const source = scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "合并😀禰󠄀" } }], metadata => { metadata.workbook.sheets[0]!.merges = [area(0, 0, 1, 1)]; });
  await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(source.serialize()));
  await page.evaluate(() => window.workbookEditor.select(0, 0, 1, 1));
  await page.getByRole("button", { name: "复制选区", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await page.evaluate(() => window.workbookEditor.select(0, 3));
  await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ merges: [area(0, 0, 1, 1), area(0, 3, 1, 4)], history: { undoSteps: 1 }, errors: [] });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 3))).toMatchObject({ text: "合并😀禰󠄀" });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  for (const axis of ["row", "column"] as const) {
    await page.evaluate(axis => window.workbookEditor.select(0, 0, axis === "row" ? 5 : 1, axis === "column" ? 5 : 1), axis);
    await page.getByRole("button", { name: axis === "row" ? "向下填充" : "向右填充", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
    expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
    for (const offset of [2, 4]) {
      expect(await page.evaluate(({ row, column }) => window.workbookEditor.cell("sheet-1", row, column), { row: axis === "row" ? offset : 0, column: axis === "column" ? offset : 0 })).toMatchObject({ text: "合并😀禰󠄀" });
    }
    await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
    expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ merges: [area(0, 0, 1, 1)], history: { undoSteps: 0 }, cleanSource: true });
  }
});

test("mounted direct native whole-table cut keeps its definition across sheets and restores it in one undo", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/workbook-axis-tables.xlsx'), { documentId: 'mounted-native-table-cut' }), { kind: 'workbook' });
  const [source, target] = model.snapshot().metadata.workbook.sheets, table = source!.tables![0]!;
  await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(model.serialize()));
  await page.evaluate(() => window.workbookEditor.select(0, 0, 5, 3)); await page.getByRole('button', { name: '剪切选区', exact: true }).click();
  await page.getByRole('tab', { name: target!.name, exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true }); await address.fill('H260'); await address.press('Enter'); await page.evaluate(() => window.workbookEditor.settled());
  await page.getByRole('button', { name: '粘贴选区', exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [], tables: [
    { sheetId: source!.sheetId, tables: [] }, { sheetId: target!.sheetId, tables: [{ tableId: table.tableId, range: { from: { row: 259, column: 7 }, to: { row: 264, column: 10 } }, columns: table.columns }] },
  ] });
  expect(await page.evaluate(sheetId => window.workbookEditor.cell(sheetId, 260, 10), target!.sheetId)).toMatchObject({ value: { type: 'number', value: 20 } });
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 0, redoSteps: 1 }, errors: [], tables: [{ tables: [table] }, { tables: [] }] });
});

test("mounted direct native global-name cut shows the global value despite a target local shadow", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/workbook-name-prefix.xlsx'), { documentId: 'mounted-name-cut' }), { kind: 'workbook' }), target = model.snapshot().metadata.workbook.sheets[1]!;
  await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(model.serialize()));
  await page.evaluate(() => window.workbookEditor.select(6, 1)); await page.getByRole('button', { name: '剪切选区', exact: true }).click();
  await page.getByRole('tab', { name: target.name, exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true }); await address.fill('D1'); await address.press('Enter'); await page.evaluate(() => window.workbookEditor.settled());
  await page.getByRole('button', { name: '粘贴选区', exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(sheetId => window.workbookEditor.cell(sheetId, 0, 3)?.formula?.expression, target.sheetId)).toBe('=[0]!Tax');
  expect(await page.evaluate(() => window.workbookEditor.displayedCell(0, 3))).toBe('7');
  await expect(page.getByRole('textbox', { name: '单元格或公式输入', exact: true })).toHaveValue('=[0]!Tax');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 0, redoSteps: 1 }, errors: [] });
});

test("mounted direct native local-name copy creates the missing target scope and shows current values in one undo", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/workbook-name-prefix.xlsx'), { documentId: 'mounted-name-copy' }), { kind: 'workbook' }), target = model.snapshot().metadata.workbook.sheets[1]!;
  await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(model.serialize()));
  await page.evaluate(() => window.workbookEditor.select(4, 1, 7, 1)); await page.getByRole('button', { name: '复制选区', exact: true }).click();
  await page.getByRole('tab', { name: target.name, exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  const address = page.getByRole('textbox', { name: '选择单元格或区域', exact: true }); await address.fill('B5'); await address.press('Enter'); await page.evaluate(() => window.workbookEditor.settled());
  await page.getByRole('button', { name: '粘贴选区', exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => [4, 5, 6, 7].map(row => window.workbookEditor.displayedCell(row, 1)))).toEqual(['107', '700', '7', '7']);
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [], names: expect.arrayContaining([expect.objectContaining({ name: 'Rate', sheetId: target.sheetId })]) });
  expect(await page.evaluate(sheetId => window.workbookEditor.cell(sheetId, 6, 1)?.formula?.expression, target.sheetId)).toBe('=[0]!Tax');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 0, redoSteps: 1 }, errors: [], names: model.snapshot().metadata.workbook.names });
});

test("mounted direct name manager preserves scopes, drafts, IME and exact name consumers through one-step history", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/workbook-name-prefix.xlsx'), { documentId: 'mounted-name-manager' }), { kind: 'workbook' });
  await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(model.serialize()));
  const { exerciseWorkbookNameControls } = await import('./workbook-name-controls-cases.js');
  await exerciseWorkbookNameControls(page, false, model.snapshot().metadata.workbook.names!, model.snapshot().metadata.workbook.sheets[1]!.sheetId);
});

test("mounted direct hyperlink controls edit exact ranges, preserve source text and guard drafts and history", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/workbook-copy-sheet-hyperlinks.xlsx'), { documentId: 'mounted-hyperlink-controls' }), { kind: 'workbook' });
  await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(model.serialize()));
  const { exerciseWorkbookHyperlinkControls } = await import('./workbook-hyperlink-controls-cases.js');
  await exerciseWorkbookHyperlinkControls(page, false, model.snapshot().metadata.workbook.sheets);
});

test("mounted direct worksheet manager copies names and current formulas, changes sheet properties and retains stale drafts", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/workbook-name-prefix.xlsx'), { documentId: 'mounted-sheet-manager' }), { kind: 'workbook' });
  await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(model.serialize()));
  const { exerciseWorkbookSheetControls } = await import('./workbook-sheet-controls-cases.js');
  await exerciseWorkbookSheetControls(page, false, model.snapshot().metadata.workbook.sheets);
});

test("mounted direct worksheet manager resolves copied name conflicts explicitly in one candidate", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  const { workbookCopyNameConflictSource, exerciseWorkbookCopyNameConflict } = await import('./workbook-sheet-controls-cases.js');
  await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(workbookCopyNameConflictSource().serialize()));
  await exerciseWorkbookCopyNameConflict(page, false);
});

test("mounted workbook axis controls preserve cross-chunk cells, current formulas and four undo steps", async ({ page }) => {
  await page.addScriptTag({ content: editorCode });
  const model = scalarWorkbook([{ row: 1, column: 31, value: { type: "number", value: 7 } }, { row: 2, column: 32, value: { type: "number", value: 8 } },
    { row: 0, column: 1, value: { type: "number", value: 999 }, formula: "=SUM(AF2:AG3)" }]);
  await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(model.serialize()));
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }), address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true });
  await input.fill("草稿"); await expect(page.getByRole("button", { name: "插入选区行", exact: true })).toBeDisabled();
  await input.press("Escape"); await page.evaluate(() => window.workbookEditor.settled());
  await page.getByLabel("插入行列的格式", { exact: true }).selectOption("none");
  await address.fill("AF2:AG3"); await address.press("Enter"); await page.evaluate(() => window.workbookEditor.settled());
  await page.getByRole("button", { name: "插入选区行", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await expect(input).toBeFocused();
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 1, 31))).toBeUndefined();
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 3, 31))).toMatchObject({ value: { type: "number", value: 7 } });
  await address.fill("AF4:AG5"); await address.press("Enter"); await page.evaluate(() => window.workbookEditor.settled());
  await page.getByRole("button", { name: "插入选区列", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await expect(input).toBeFocused();
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 4, 34))).toMatchObject({ value: { type: "number", value: 8 } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 1))).toMatchObject({ formula: { expression: "=SUM(AH4:AI5)" } });
  await page.getByRole("button", { name: "删除选区列", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await address.fill("AF2:AG3"); await address.press("Enter"); await page.evaluate(() => window.workbookEditor.settled());
  await page.getByRole("button", { name: "删除选区行", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 4, history: { undoSteps: 4 }, calculated: { state: "valid", value: { type: "number", value: 15 } }, cleanSource: true, errors: [] });
  for (let index = 0; index < 4; index++) { await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled()); }
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 8, history: { undoSteps: 0, redoSteps: 4 }, calculated: { value: { type: "number", value: 15 } }, cleanSource: true, errors: [] });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 1, 31))).toMatchObject({ value: { type: "number", value: 7 } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 2, 32))).toMatchObject({ value: { type: "number", value: 8 } });
});

test("mounted workbook freezes before its selected cell and restores the same source through undo", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 7 } }]).serialize()));
  await page.evaluate(() => window.workbookEditor.select(10000, 1000));
  await page.getByRole("button", { name: "冻结选区之前", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await expect(page.getByRole("alert")).toContainText("冻结区域将占满当前视口");
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, frozen: { rows: 0, columns: 0 }, history: { undoSteps: 0 }, errors: ["InvalidOperation"] });
  await page.evaluate(() => window.workbookEditor.select(1, 1));
  await page.getByRole("button", { name: "冻结选区之前", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 1, frozen: { rows: 1, columns: 1 }, history: { undoSteps: 1 }, errors: ["InvalidOperation"] });
  await page.getByRole("button", { name: "取消冻结", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 2, frozen: { rows: 0, columns: 0 } });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ frozen: { rows: 1, columns: 1 } });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ frozen: { rows: 0, columns: 0 }, history: { undoSteps: 0, redoSteps: 2 }, cells: [{ value: { type: "number", value: 7 } }], errors: ["InvalidOperation"] });
});

for (const rtl of [false, true]) test(`mounted workbook keyboard extends source ranges across hidden dimensions in ${rtl ? "RTL" : "LTR"}`, async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource(rtl)); await page.evaluate(() => window.workbookEditor.select(0));
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await input.press(rtl ? "ArrowLeft" : "ArrowRight"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, selection: { anchor: { ranges: [{ from: { row: 0, column: 11 }, to: { row: 0, column: 11 } }] } } });
  await input.press(rtl ? "ArrowRight" : "ArrowLeft"); await page.evaluate(() => window.workbookEditor.settled()); await input.press("ArrowDown"); await page.evaluate(() => window.workbookEditor.settled());
  await input.press("Shift+ArrowDown"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, errors: [], selection: { anchor: { ranges: [{ from: { row: 1001, column: 0 }, to: { row: 1003, column: 2 } }] } } });
  await input.press("Shift+ArrowUp"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ selection: { anchor: { ranges: [{ from: { row: 1001, column: 0 }, to: { row: 1002, column: 2 } }] } } });
  await page.evaluate(() => window.workbookEditor.select(0));
  await input.press(rtl ? "ArrowLeft" : "ArrowRight"); await input.press(rtl ? "ArrowLeft" : "ArrowRight"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, selection: { anchor: { ranges: [{ from: { row: 0, column: 12 }, to: { row: 0, column: 12 } }] } } });
});
test("mounted workbook mouse and sheet switching retain invalid drafts until explicit cancellation", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource(false, false)); await page.evaluate(() => window.workbookEditor.select(1));
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await input.fill("=");
  const point = await page.evaluate(() => window.workbookEditor.point(0, "A")); await page.mouse.click(point.x, point.y); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, input: { text: "=", dirty: true }, selection: { anchor: { ranges: [{ from: { row: 1, column: 0 }, to: { row: 1, column: 0 } }] } }, errors: ["InvalidOperation"] });
  await page.getByRole("tab", { name: "Second", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled()); await expect(input).toHaveValue("=");
  await input.press("Escape"); await expect(input).toHaveValue(""); await page.getByRole("tab", { name: "Second", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ input: { text: "" }, selection: { anchor: { sheetId: "second" } }, revision: 0 });
  await input.pressSequentially("new sheet"); await input.press("Control+Enter"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("second", 0, 0))).toMatchObject({ value: { type: "text", value: "new sheet" } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 0))).toMatchObject({ value: { type: "number", value: 1 } });
});
test("mounted workbook TSV paste and rectangle clear use single cross-chunk history steps", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource(false, false)); await page.evaluate(() => window.workbookEditor.select(255, 31));
  expect(await page.evaluate(() => window.workbookEditor.paste("00123\t=AF256+1\n中文😀禰󠄀\tTRUE"))).toBe(true); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, cleanSource: true, errors: [], selection: { anchor: { ranges: [{ from: { row: 255, column: 31 }, to: { row: 256, column: 32 } }] } } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 31))).toMatchObject({ value: { type: "text", value: "中文😀禰󠄀" } });
  await page.getByRole("textbox", { name: "单元格或公式输入", exact: true }).press("Delete"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 2 }, errors: [] });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 31))).toBeUndefined();
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 31))).toMatchObject({ value: { type: "text", value: "中文😀禰󠄀" } });
});
test("mounted find follows exact occurrences and preserves selection scope through atomic replacements", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  await page.evaluate(entries => window.workbookEditor.mount(entries), encodeTestPackage(scalarWorkbook([
    { row: 255, column: 31, value: { type: "text", value: "猫😀禰󠄀 猫" } }, { row: 256, column: 32, value: { type: "text", value: "猫" } },
    { row: 0, column: 0, value: { type: "text", value: "猫 outside" } },
  ], metadata => { metadata.workbook.sheets[0]!.defaultRowHeightCustom = true; }).serialize()));
  await page.evaluate(() => window.workbookEditor.select(255, 31, 256, 32));
  const find = page.getByRole("textbox", { name: "查找内容", exact: true }), replacement = page.getByRole("textbox", { name: "替换为", exact: true });
  await find.fill("猫"); await page.getByRole("combobox", { name: "查找范围", exact: true }).selectOption("selection");
  await find.press("Enter"); await page.evaluate(() => window.workbookEditor.settled());
  await expect(page.getByRole("status", { name: "工作簿编辑状态", exact: true })).toContainText("匹配 1 / 3");
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
  expect(await input.evaluate((element: HTMLTextAreaElement) => [element.selectionStart, element.selectionEnd])).toEqual([0, 1]);
  await page.getByRole("button", { name: "查找下一个", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await expect(page.getByRole("status", { name: "工作簿编辑状态", exact: true })).toContainText("匹配 2 / 3");
  expect(await input.evaluate((element: HTMLTextAreaElement) => [element.selectionStart, element.selectionEnd])).toEqual([7, 8]);
  await replacement.fill("=literal"); await page.getByRole("button", { name: "替换当前项", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 255, 31))).toMatchObject({ value: { type: "text", value: "猫😀禰󠄀 =literal" } });
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  await replacement.fill("$&"); await page.getByRole("button", { name: "全部替换", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 255, 31))).toMatchObject({ value: { type: "text", value: "$&😀禰󠄀 =literal" } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 32))).toMatchObject({ value: { type: "text", value: "$&" } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 0))).toMatchObject({ value: { type: "text", value: "猫 outside" } });
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 2 }, cleanSource: true, errors: [] });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 32))).toMatchObject({ value: { type: "text", value: "猫" } });
});
test("mounted display search uses current calculation and cancellable replacement leaves source intact", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode });
  await page.evaluate(entries => window.workbookEditor.mountObserved(entries), encodeTestPackage(scalarWorkbook([
    { row: 0, column: 0, value: { type: "number", value: 1 } }, { row: 0, column: 1, value: { type: "number", value: 999 }, formula: "=A1+1" },
  ], metadata => { metadata.workbook.calculationMode = "manual"; metadata.workbook.sheets[0]!.defaultRowHeightCustom = true; }).serialize()));
  expect(await page.evaluate(() => window.workbookEditor.displayedCell(0, 1))).toBe("999");
  const find = page.getByRole("textbox", { name: "查找内容", exact: true }); await find.fill("2");
  await page.getByRole("combobox", { name: "查找字段", exact: true }).selectOption("display"); await find.press("Enter"); await page.evaluate(() => window.workbookEditor.settled());
  await expect(page.getByRole("status", { name: "工作簿编辑状态", exact: true })).toContainText("匹配 1 / 1");
  expect(await page.evaluate(() => window.workbookEditor.displayedCell(0, 1))).toBe("2");
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 1))).toMatchObject({ value: { type: "number", value: 999 }, formula: { cacheState: "imported" } });
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, selection: { anchor: { ranges: [{ from: { row: 0, column: 1 } }] } } });
  await expect(page.getByRole("button", { name: "全部替换", exact: true })).toBeDisabled();
  await page.getByRole("combobox", { name: "查找字段", exact: true }).selectOption("source"); await find.fill("1");
  await page.getByRole("textbox", { name: "替换为", exact: true }).fill("3"); await page.evaluate(() => window.workbookEditor.cancelSearchAtProgress());
  await page.getByRole("button", { name: "全部替换", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await expect(page.getByRole("status", { name: "工作簿编辑状态", exact: true })).toContainText("操作已取消");
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, errors: [] });
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await input.focus(); await input.press("Control+f"); await expect(find).toBeFocused();
});
test("workbook search identifies hidden-sheet matches without changing visibility or source", async ({ page }) => {
  const seed = createDocumentSession(openDocument(scalarWorkbook([], metadata => {
    metadata.workbook.sheets.push({ ...metadata.workbook.sheets[0]!, sheetId: "hidden", name: "Hidden", visibility: "hidden" });
  }).serialize(), { kind: "workbook" }));
  seed.commit(seed.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0,
    operations: [{ op: "setWorkbookCells", sheetId: "hidden", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "needle" } } }] }] }));
  const entries = encodeTestPackage(seed.serialize()); seed.dispose();
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), entries);
  const find = page.getByRole("textbox", { name: "查找内容", exact: true }); await find.fill("needle");
  await page.getByRole("combobox", { name: "查找范围", exact: true }).selectOption("workbook"); await find.press("Enter"); await page.evaluate(() => window.workbookEditor.settled());
  await expect(page.getByRole("status", { name: "工作簿编辑状态", exact: true })).toContainText("Hidden（隐藏）!A1");
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 0 }, errors: [] });
  await page.getByRole("textbox", { name: "替换为", exact: true }).fill("updated"); await page.getByRole("button", { name: "替换当前项", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("hidden", 0, 0))).toMatchObject({ value: { type: "text", value: "updated" } });
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  await expect(page.getByRole("tab", { name: "Hidden", exact: true })).toHaveCount(0);
});
test("mounted workbook fill shifts formulas, extends numeric seeds and cancels before source publication", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource(false, false));
  await page.evaluate(() => window.workbookEditor.select(0, 0, 2, 1)); await page.getByRole("button", { name: "向下填充", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 2, 0))).toMatchObject({ value: { type: "number", value: 1 } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 2, 1))).toMatchObject({ formula: { expression: "=(A3+1)" } });
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await page.evaluate(() => window.workbookEditor.select(0, 0, 4)); await page.getByRole("button", { name: "数列填充", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 4, 0))).toMatchObject({ value: { type: "number", value: 5 } });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await page.evaluate(() => window.workbookEditor.select(0, 0, 512)); await page.evaluate(() => window.workbookEditor.cancelFillAtProgress());
  await page.getByRole("button", { name: "向下填充", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 0 }, cleanSource: true, errors: [] });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 0))).toBeUndefined();
  await expect(page.getByRole("button", { name: "取消操作", exact: true })).toBeDisabled();
  // Let command generation finish, then cancel while core is preparing source.
  await page.evaluate(() => window.workbookEditor.cancelPreparationAtProgress());
  await page.getByRole("button", { name: "向下填充", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 0 }, errors: [] });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 0))).toBeUndefined();
  await expect(page.getByRole("status", { name: "工作簿编辑状态", exact: true })).toContainText("操作已取消");
});
test("mounted rich cell input retains font spans and line endings, with explicit type conversion", async ({ page }) => {
  const files = new Map(scalarWorkbook([]).serialize());
  files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + ".emphasis{font-style:italic}"));
  const seed = createDocumentSession(openDocument(files, { kind: "workbook" }));
  seed.commit(seed.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0,
    source: { kind: "richText", html: '<span class="cell-default">ABC&#13;</span><span class="emphasis">中文禰󠄀</span>' } }] }] }));
  const entries = encodeTestPackage(seed.serialize()); seed.dispose();
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), entries); await page.evaluate(() => window.workbookEditor.select(0));
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await expect(input).toHaveValue("ABC\n中文禰󠄀");
  await input.press("F2"); await input.pressSequentially("尾"); await input.press("Control+Enter"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 0))).toMatchObject({ text: "ABC\r中文禰󠄀尾", html: expect.stringContaining('<span class="cell-default">ABC&#13;</span><span class="emphasis">中文禰󠄀尾</span>') });
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  await page.getByRole("combobox", { name: "输入类型", exact: true }).selectOption("text"); await input.fill("12"); await page.getByRole("button", { name: "确认输入", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled()); await expect(input).toHaveValue("'12");
  await page.getByRole("combobox", { name: "输入类型", exact: true }).selectOption("number"); await expect(input).toHaveValue("12");
  await page.getByRole("button", { name: "确认输入", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 0))).toMatchObject({ value: { type: "number", value: 12 } });
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 2 }, errors: [], cleanSource: true });
});
test("mounted workbook real IME preserves a conflicting draft and disposal leaves its host session usable", async ({ page, browserName }) => {
  test.skip(browserName !== "chromium", "CDP composition is Chromium-specific.");
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource()); await page.evaluate(() => window.workbookEditor.select(0));
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }), cdp = await page.context().newCDPSession(page);
  await cdp.send("Input.imeSetComposition", { text: "中文", selectionStart: 2, selectionEnd: 2 }); expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, input: { composing: true } });
  await page.evaluate(() => window.workbookEditor.external("external")); await page.evaluate(() => window.workbookEditor.settled());
  await cdp.send("Input.insertText", { text: "中文😀" }); await input.press("Enter");
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, input: { text: "中文😀", dirty: true, selection: { baseRevision: 0 } }, cells: [{ value: { type: "text", value: "external" } }, expect.anything()], errors: ["RevisionConflict"] });
  await input.press("Escape"); await expect(input).toHaveValue("external");
  expect(await page.evaluate(() => window.workbookEditor.dispose())).toEqual({ revision: 1, usable: true });
});
test.beforeEach(async ({ page }) => { await openSdkHost(page); await page.addScriptTag({ content: code }); await page.evaluate(() => window.workbookInput.mount()); });
test("native typing remains transient until one Enter transaction, with shared undo and redo", async ({ page }) => {
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true });
  await input.pressSequentially("00123"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, state: { text: "00123", dirty: true } });
  await input.press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, state: { dirty: false }, cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "00123" } }) })]), navigation: [{ direction: "down", extend: false }], errors: [] });
  await input.press("Control+z"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 0, redoSteps: 1 }, state: { text: "" }, errors: [] });
  await input.press("Control+y"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, state: { text: "00123" }, errors: [] });
  await input.press("F2"); await input.press("End"); await input.pressSequentially("X"); await input.press("Escape"); await expect(input).toHaveValue("00123");
  await input.press("Shift+ArrowRight"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, navigation: [{ direction: "down", extend: false }, { direction: "right", extend: true }] });
});
test("native textarea confirmation and edits retain untouched source CR, CRLF and variation sequences", async ({ page }) => {
  const original = "甲\r\n乙\r丙\n丁😀禰󠄀";
  await page.evaluate(text => window.workbookInput.seedText(text), original);
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true }); await expect(input).toHaveValue(original.replace(/\r\n?/g, "\n"));
  await input.press("Control+Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 0 }, state: { text: original, dirty: false }, errors: [] });
  await input.press("Control+End"); await input.pressSequentially("尾"); await input.press("Control+Enter");
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 2, history: { undoSteps: 1 }, state: { text: original + "尾" }, errors: [] });
  await input.press("Control+z"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 0 }, state: { text: original } });
});
test("native editing preserves formula-like literal text and explicit whole-cell formula replacement", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.seedText("=1")); const input = page.getByRole("textbox", { name: "工作簿输入", exact: true }); await expect(input).toHaveValue("'=1");
  await input.press("F2"); await input.pressSequentially("X"); await input.press("Control+Enter");
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, state: { text: "'=1X" }, cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "=1X" } }) })]), errors: [] });
  await input.press("Control+a"); await input.pressSequentially("=1+1"); await input.press("Control+Enter");
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 2 }, cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ formula: expect.objectContaining({ expression: "=1+1" }) }) })]), errors: [] });
});
test("native formula and multiline editing preserve draft failures and source types", async ({ page }) => {
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true });
  await input.pressSequentially("="); await input.press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, state: { text: "=", dirty: true }, errors: [{ code: "InvalidOperation", text: "=" }] });
  await input.pressSequentially("1+2"); await input.press("Control+Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, navigation: [], cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ formula: expect.objectContaining({ expression: "=1+2" }) }) })]) });
  await input.press("Control+a"); await input.pressSequentially("中文"); await input.press("Alt+Enter"); await input.pressSequentially("混排"); await input.press("Tab");
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 2 }, state: { text: "中文\n混排" }, navigation: [{ direction: "right", extend: false }], cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "中文\n混排" } }) })]) });
});
test("host-authorized TSV paste crosses chunks atomically, rejects stale input and clears a rectangle", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.mount(255, 31, 256, 32));
  expect(await page.evaluate(() => window.workbookInput.paste("00123\t=AF256+1\n中文😀禰󠄀\tTRUE"))).toBe(true);
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [], cells: [expect.anything(),
    expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "00123" } }) }),
    expect.objectContaining({ cell: expect.objectContaining({ formula: expect.objectContaining({ expression: "=AF256+1" }) }) }),
    expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "中文😀禰󠄀" } }) }),
    expect.objectContaining({ cell: expect.objectContaining({ value: { type: "boolean", value: true } }) })] });
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true }); await input.press("Delete");
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 2 }, cells: Array.from({ length: 5 }, () => expect.not.objectContaining({ cell: expect.anything() })) });
  await input.press("Control+z"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, state: { text: "00123" } });
  await page.evaluate(() => window.workbookInput.external("external")); await page.evaluate(() => window.workbookInput.paste("stale"));
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 2 }, state: { selection: { baseRevision: 3 } }, errors: [{ code: "RevisionConflict", text: "00123" }] });
});
test("real Chromium IME commits only on confirmation and retains text after a source conflict", async ({ page, browserName }) => {
  test.skip(browserName !== "chromium", "CDP composition is Chromium-specific; all engines cover native text input and history.");
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true }), cdp = await page.context().newCDPSession(page);
  await cdp.send("Input.imeSetComposition", { text: "中", selectionStart: 1, selectionEnd: 1 });
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, state: { composing: true } });
  await cdp.send("Input.imeSetComposition", { text: "中文", selectionStart: 2, selectionEnd: 2 }); await cdp.send("Input.insertText", { text: "中文😀禰󠄀" });
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, state: { composing: false, text: "中文😀禰󠄀" } });
  await input.press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, state: { text: "中文😀禰󠄀" } });
  await cdp.send("Input.imeSetComposition", { text: "迟到", selectionStart: 2, selectionEnd: 2 }); await page.evaluate(() => window.workbookInput.external("external")); await cdp.send("Input.insertText", { text: "迟到" }); await input.press("Enter");
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 2 }, state: { text: "迟到", dirty: true }, cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "external" } }) })]), errors: [{ code: "RevisionConflict", text: "迟到" }] });
});
test("binding disposal releases DOM listeners and leaves its host session usable", async ({ page }) => {
  await page.getByRole("textbox", { name: "工作簿输入", exact: true }).pressSequentially("draft"); await page.evaluate(() => window.workbookInput.releaseBinding());
  await page.getByRole("textbox", { name: "工作簿输入", exact: true }).press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, state: { text: "draft" } });
  await page.evaluate(() => window.workbookInput.dispose()); expect(await page.evaluate(() => window.workbookInput.hostStillUsable())).toBe(true);
});
test("an oversized native draft cannot commit an older controller value and can be corrected", async ({ page }) => {
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true }); await input.fill("x".repeat(32768));
  expect(await page.evaluate(() => window.workbookInput.paste("must remain a text draft"))).toBe(false); await input.press("Enter");
  expect(await page.evaluate(() => { const state = window.workbookInput.inspect(); return { revision: state.revision, length: state.value.length, errors: [...new Set(state.errors.map(error => error.code))] }; }))
    .toEqual({ revision: 0, length: 32768, errors: ["ResourceLimitExceeded"] });
  await input.fill("corrected"); await input.press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, state: { text: "corrected", dirty: false } });
});
