import { describe, expect, it } from "vitest";
import { indexWorkbookChunk } from "../src/workbook/chunk-index.js";
import type { ChunkValidation } from "../src/workbook/chunk-index.js";
import { createFormulaTemplatePool } from "../src/workbook/formula.js";

const encoder = new TextEncoder();
const position = { sheetId: "sheet-1", row: 256, column: 32 };
function source(content: string): Uint8Array {
  return encoder.encode('<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="256" data-column="32"><tbody>' + content + '</tbody></table></body></html>');
}
const content = '<tr data-row="256"><td data-column="32" data-value-type="text"><span class="east-asian">中😀禰󠄀</span><br>字</td><td data-column="34" data-value-type="number" data-value="3.5" data-formula="=AG257+1" data-cache-state="imported">3.50</td></tr><tr data-row="511"><td data-column="63" data-value-type="boolean" data-value="false">FALSE</td></tr>';
const plainContent = '<tr data-row="256"><td data-column="32" data-value-type="text" class="east-asian">中文</td><td data-column="34" data-value-type="number" data-value="3.5" data-formula="=AG257+1" data-cache-state="imported">3.50</td></tr><tr data-row="511"><td data-column="63" data-value-type="blank"></td></tr>';

function observed(body: string, fallback: boolean, validation?: ChunkValidation) {
  const input = source(body), bytes = fallback ? encoder.encode(new TextDecoder().decode(input).replace('<!doctype html>', '<!DOCTYPE html>')) : input;
  let index: ReturnType<typeof indexWorkbookChunk>;
  try {
    index = indexWorkbookChunk(bytes, position, 8 * 1024 * 1024, { ...validation, formulas: createFormulaTemplatePool(20, 10000) });
  } catch (error) { return { code: (error as { code?: string }).code }; }
  expect(index.serialize()).toEqual(bytes); bytes.fill(0);
  const cells = index.readWindow({ row: 256, column: 32, rows: 256, columns: 32 });
  for (const cell of cells) expect(index.getCell(cell.row, cell.column)).toEqual(cell);
  return { cellCount: index.cellCount, formulaCount: index.formulaCount, indexBytes: index.indexBytes, cells, formulas: [...index.formulaLocations()] };
}

describe('canonical plain chunk validation agrees with the full HTML parser', () => {
  it('keeps the sparse cells and first-party formula index', () => {
    expect(observed(plainContent, false)).toMatchObject({ cellCount: 3, formulaCount: 1, formulas: [{ row: 256, column: 34 }] });
    expect(observed(plainContent.replace('中文', '<script>bad()</script>'), false)).toMatchObject({ code: 'UnsupportedFeature' });
  });
  it.each([
    plainContent, '', '<tr data-row="256"></tr>',
    plainContent.replace('中文', 'plain text\nwith a tab\t'), plainContent.replace('中文', 'x'.repeat(32767)),
    plainContent.replace('中文', 'x'.repeat(32768)), plainContent.replace('中文', '中😀禰󠄀'),
    plainContent.replace('中文', '&lt;&amp;&#13;'), plainContent.replace('中文', '\u0001'), plainContent.replace('中文', '\ufdd0'),
    plainContent.replace('中文', '<span class="east-asian">rich</span>'), plainContent.replace('中文', '<!-- hidden -->'),
    plainContent.replace('中文', '<script>bad()</script>'), plainContent.replace('中文', '<img src="https://example.test">'),
    plainContent.replace('data-column="34"', 'data-column="32"'), plainContent.replace('data-column="34"', 'data-column="64"'),
    plainContent.replace('data-row="511"', 'data-row="512"'), plainContent.replace('data-row="511"', 'data-row="256"'),
    plainContent.replace('data-value="3.5"', 'data-value="Infinity"'), plainContent.replace('data-value="3.5"', 'data-value="3.5" data-value="4"'),
    plainContent.replace('data-cache-state="imported"', 'data-cache-state="invented"'), plainContent.replace('data-formula="=AG257+1"', 'data-formula="invalid"'),
    plainContent.replace('data-value-type="blank"', 'data-value-type="blank" data-value="1"'),
    plainContent.replace('class="east-asian"', 'onclick="bad()"'), plainContent.replace('</td>', ''),
  ])('keeps exact cell slices and rejects the same invalid source (%#)', body => {
    expect(observed(body, false)).toEqual(observed(body, true));
  });
  it.each(Array.from({ length: 18 }, (_, maxNodes) => maxNodes + 1))('retains the complete node budget %i', maxNodes => {
    expect(observed(plainContent, false, { maxNodes })).toEqual(observed(plainContent, true, { maxNodes }));
  });
  it.each(Array.from({ length: 12 }, (_, maxDepth) => maxDepth + 1))('retains the depth budget %i', maxDepth => {
    expect(observed(plainContent, false, { maxDepth })).toEqual(observed(plainContent, true, { maxDepth }));
  });
  it.each([
    { cssClasses: new Set<string>() }, { cssClasses: new Set(['east-asian']) },
    { numberFormats: new Set<string>() }, { numberFormats: new Set(['fmt']) },
    { cellStyleClasses: new Set(['east-asian', 'another']) },
  ])('retains current class, named style and format bindings (%#)', validation => {
    for (const body of [plainContent, plainContent.replace('class="east-asian"', 'class="east-asian another"'), plainContent.replace('class="east-asian"', 'data-format-id="fmt"')])
      expect(observed(body, false, validation)).toEqual(observed(body, true, validation));
  });
});

describe("private workbook chunk storage baseline", () => {
  it("indexes sparse coordinates, preserves UTF-16 source and isolates bytes", () => {
    const bytes = source(content), copy = bytes.slice(), index = indexWorkbookChunk(bytes, position);
    expect(index.cellCount).toBe(3); expect(index.formulaCount).toBe(1); expect(index.indexBytes).toBe(30);
    bytes.fill(0); index.serialize().fill(0);
    expect(index.serialize()).toEqual(copy);
    const window = index.readWindow({ row: 256, column: 32, rows: 1, columns: 2 });
    expect(window).toEqual([{ row: 256, column: 32, html: '<td data-column="32" data-value-type="text"><span class="east-asian">中😀禰󠄀</span><br>字</td>' }]);
    expect(Object.isFrozen(window[0])).toBe(true);
    expect(index.readWindow({ row: 511, column: 63, rows: 1, columns: 1 })[0]?.html).toContain('data-value="false"');
    expect(index.readWindow({ row: 0, column: 0, rows: 1, columns: 1 })).toEqual([]);
  });
  it.each([
    content.replace('data-column="34"', 'data-column="32"'),
    content.replace('data-column="34"', 'data-column="31"'),
    content.replace('data-row="511"', 'data-row="512"'),
    content.replace('data-row="511"', 'data-row="256"'),
    content.replace('data-value="3.5"', 'data-value="Infinity"'),
    content.replace('data-value="false"', 'data-value="0"'),
    content.replace('data-cache-state="imported"', ''),
    content.replace('<br>', '<br onclick="alert(1)">'),
    content.replace('<br>', '<script>alert(1)</script>'),
    content.replace('<br>', '<img src="https://example.com/a.png">'),
    content.replace('<br>', '<!-- unmodeled -->'),
    content.replace('class="east-asian"', 'class="east-asian" style="color:red"'),
    content.replace('</td>', ''),
    content.replace('<span', '<table'),
  ])("rejects malformed, unordered or unmodeled source before indexing", input => {
    expect(() => indexWorkbookChunk(source(input), position)).toThrow();
  });
  it("enforces source budgets and bounds before keeping data", () => {
    expect(() => indexWorkbookChunk(source(content), position, 10)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(() => indexWorkbookChunk(new Uint8Array([0xff]), position)).toThrow();
    expect(() => indexWorkbookChunk(source(content), { ...position, row: 257 })).toThrow();
    expect(() => indexWorkbookChunk(source(content), { ...position, sheetId: "another" })).toThrow();
    const index = indexWorkbookChunk(source(content), position);
    for (const window of [
      { row: -1, column: 0, rows: 1, columns: 1 },
      { row: 0, column: 0, rows: 0, columns: 1 },
      { row: 0, column: 0, rows: 1_048_577, columns: 1 },
      { row: 0, column: 16_384, rows: 1, columns: 1 },
    ]) expect(() => index.readWindow(window)).toThrow();
  });
});
