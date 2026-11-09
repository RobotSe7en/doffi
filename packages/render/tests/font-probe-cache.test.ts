import { expect, it } from 'vitest';
import { FontProbeCache } from '../src/font-probe-cache.js';
import { WorkbookCellTextCache } from '../src/workbook/cell-text-cache.js';
const fonts = () => Object.assign(new Set<FontFace>(), { status: 'loaded' }) as unknown as FontFaceSet;
it('controller cleanup preserves borrowed surface probes and owning cleanup releases private probes', () => {
  const probes = new FontProbeCache(), borrowed = new WorkbookCellTextCache(2, 10, 1024, probes), own = new WorkbookCellTextCache(2, 10, 1024);
  probes.glyph('same-frame', () => 7); borrowed.clear();
  expect(probes.glyph('same-frame', () => { throw Error('A live surface probe was cleared.'); })).toBe(7);
  own.probes.glyph('private', () => 9); own.clear(); expect(own.probes.glyph('private', () => 10)).toBe(10);
  probes.clear(); expect(probes.glyph('same-frame', () => 11)).toBe(11);
});
it('retains zero/false font probes, bounds both LRUs and isolates font-set owners', () => {
  const cache = new FontProbeCache(2, 100), owner = fonts(); let calls = 0;
  cache.watch(owner);
  expect(cache.available('absent', () => { calls++; return false; })).toBe(false);
  expect(cache.available('absent', () => { throw Error('cached false'); })).toBe(false);
  expect(cache.glyph('zero', () => { calls++; return 0; })).toBe(0);
  expect(cache.glyph('zero', () => { throw Error('cached zero'); })).toBe(0);
  cache.glyph('new', () => 1); cache.available('absent', () => { calls++; return false; }); expect(calls).toBe(3);
  cache.watch(fonts()); cache.available('absent', () => { calls++; return false; }); expect(calls).toBe(4);
  cache.clear(); cache.watch(owner); cache.available('absent', () => { calls++; return false; }); expect(calls).toBe(5);
  const bounded = new FontProbeCache(10, 5); bounded.glyph('long-key', () => { calls++; return 1; }); bounded.glyph('long-key', () => { calls++; return 1; }); expect(calls).toBe(7);
});
it('invalidates on a completed face change, descriptor change and in-progress load without forcing fonts.ready', () => {
  const cache = new FontProbeCache(), owner = fonts(), set = owner as unknown as Set<FontFace>, face = { family: 'private', status: 'loaded' } as FontFace;
  Object.defineProperty(owner, 'ready', { get() { throw Error('must not force font layout'); } });
  const initial = cache.watch(owner); cache.glyph('x', () => 1); set.add(face);
  expect(cache.watch(owner).epoch).toBeGreaterThan(initial.epoch); expect(cache.glyph('x', () => 2)).toBe(2);
  face.family = 'updated'; const changed = cache.watch(owner); expect(cache.glyph('x', () => 3)).toBe(3);
  Object.assign(owner, { status: 'loading' }); expect(cache.watch(owner)).toMatchObject({ ready: false });
  Object.assign(owner, { status: 'loaded' }); expect(cache.watch(owner).epoch).toBeGreaterThan(changed.epoch);
});
