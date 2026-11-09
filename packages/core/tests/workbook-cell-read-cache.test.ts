import { expect, it } from 'vitest';
import { WorkbookCellReadCache } from '../src/workbook/cell-read-cache.js';
import { WorkbookChunkTextCache } from '../src/workbook/chunk-text-cache.js';
import { createUtf8TextCopy } from '../src/document/package.js';
const source = (column: number, value = '😀禰󠄀') => ({ row: 0, column, html: `<td data-column="${column}" data-value-type="text">${value}</td>` });
it('copies repeated Unicode text without retaining decoder streaming state or stripping a literal BOM', () => {
  const copy = createUtf8TextCopy();
  for (const text of ['😀禰󠄀', '\ufeffsource', 'plain', '\r\n', '\ufeffagain', '']) expect(copy(text)).toBe(text);
});
it('bounds overlapping immutable reads by entry count with actual LRU eviction and exact UTF-16', () => {
  const cache = new WorkbookCellReadCache(2), first = cache.read('sheet', source(0)), second = cache.read('sheet', source(1));
  expect(cache.read('sheet', source(0))).toBe(first); cache.read('sheet', source(2));
  expect(cache.read('sheet', source(0))).toBe(first);
  expect(cache.read('sheet', source(1))).not.toBe(second);
  expect(first.value).toEqual({ type: 'text', value: '😀禰󠄀' }); expect(first.text.length).toBe(5);
  expect(Object.isFrozen(first)).toBe(true); expect(Object.isFrozen(first.value)).toBe(true);
});
it('point cache hits update the same LRU and never expose implicit blanks or another sheet', () => {
  const cache = new WorkbookCellReadCache(2), first = cache.read('s', source(0)); cache.read('s', source(1));
  expect(cache.get('s', 0, 0)).toBe(first); cache.read('s', source(2));
  expect(cache.get('s', 0, 1)).toBeUndefined(); expect(cache.get('other', 0, 0)).toBeUndefined(); expect(cache.get('s', 50, 0)).toBeUndefined();
  expect(cache.get('s', 0, 0)).toBe(first);
});
it('keeps exact bounded LRU identities across repeated middle hits, misses and evictions', () => {
  const cache = new WorkbookCellReadCache(7), expected = new Map<string, ReturnType<WorkbookCellReadCache['read']>>(); let seed = 20261110;
  for (let index = 0; index < 1000; index++) {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    const column = seed % 19, sheet = index % 3 ? 'first' : 'second', key = sheet + ':' + column, previous = expected.get(key);
    const cell = cache.read(sheet, source(column));
    if (previous) expect(cell).toBe(previous);
    expected.delete(key); if (expected.size === 7) expected.delete(expected.keys().next().value!); expected.set(key, cell);
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    const probe = seed % 19, probeKey = sheet + ':' + probe, known = expected.get(probeKey);
    expect(cache.get(sheet, 0, probe)).toBe(known);
    if (known) { expected.delete(probeKey); expected.set(probeKey, known); }
  }
});
it('enforces cumulative text storage and does not retain an individual oversized value', () => {
  const cache = new WorkbookCellReadCache(5000, 180), first = cache.read('s', source(0)); cache.read('s', source(1)); cache.read('s', source(2));
  expect(cache.read('s', source(0))).not.toBe(first);
  const large = source(3, 'x'.repeat(100)), old = cache.read('s', large); expect(cache.read('s', large)).not.toBe(old);
  expect(cache.read('s', large)).toEqual(old);
});
it('keeps sheet and model namespaces separate and freezes nested formula caches', () => {
  const cache = new WorkbookCellReadCache(2), first = cache.read('first', source(0)), other = cache.read('other', source(0, 'other'));
  expect(other.value).toEqual({ type: 'text', value: 'other' }); expect(cache.read('first', source(0))).toBe(first);
  const isolated = new WorkbookCellReadCache(2); expect(isolated.read('first', source(0, 'private'))).not.toBe(first);
  const formula = isolated.read('first', { row: 0, column: 1, html: '<td data-column="1" data-value-type="number" data-value="99" data-formula="=1+1" data-cache-state="imported">99</td>' });
  expect(formula.formula).toEqual({ expression: '=1+1', cacheState: 'imported' }); expect(Object.isFrozen(formula.formula)).toBe(true);
});
it('bounds source text across chunks, retains exact Unicode after eviction and rejects invalid UTF-8', () => {
  const cache = new WorkbookChunkTextCache(2, 10), encode = (text: string) => new TextEncoder().encode(text);
  const a = encode('😀禰󠄀'), b = encode('abc'), c = encode('def');
  expect(cache.read(a)).toBe('😀禰󠄀'); expect(cache.read(b)).toBe('abc'); cache.read(a); cache.read(c);
  expect(cache.retainedChunks).toBe(2); expect(cache.retainedCodeUnits).toBe(8);
  expect(cache.read(b)).toBe('abc'); expect(cache.retainedCodeUnits).toBe(6);
  const large = encode('x'.repeat(11)); expect(cache.read(large)).toBe('x'.repeat(11)); expect(cache.retainedCodeUnits).toBe(6);
  expect(() => cache.read(new Uint8Array([255]))).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
  expect(cache.retainedCodeUnits).toBe(6);
});
