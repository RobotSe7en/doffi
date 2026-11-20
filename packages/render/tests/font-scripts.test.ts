import { describe, expect, it } from "vitest";
import { officeFontSegments } from "../src/font-scripts.js";

describe("shared Office font script segmentation", () => {
  it("retains supplementary characters, IVS, combining marks and exact UTF-16 offsets", () => {
    const text = 'A\u0301é漢禰\u{e0100}😀ש', result = officeFontSegments(text);
    expect(result).toEqual([
      { role: 'ascii', text: 'A\u0301', from: 0, to: 2 }, { role: 'h-ansi', text: 'é', from: 2, to: 3 },
      { role: 'east-asia', text: '漢禰\u{e0100}', from: 3, to: 7 }, { role: 'h-ansi', text: '😀', from: 7, to: 9 }, { role: 'complex-script', text: 'ש', from: 9, to: 10 },
    ]);
    expect(result.map(run => text.slice(run.from, run.to)).join('')).toBe(text); expect(Object.isFrozen(result[0])).toBe(true);
  });
  it("keeps East Asian scripts and punctuation together, and can carry a base role across source runs", () => {
    expect(officeFontSegments('ひカ한，。Ａ')).toEqual([{ role: 'east-asia', text: 'ひカ한，。Ａ', from: 0, to: 6 }]);
    expect(officeFontSegments('\u{e0100}', 'east-asia')).toEqual([{ role: 'east-asia', text: '\u{e0100}', from: 0, to: 2 }]);
    expect(officeFontSegments('العربية')).toEqual([{ role: 'complex-script', text: 'العربية', from: 0, to: 7 }]);
    expect(officeFontSegments('')).toEqual([]);
  });
  it("routes East Asian ambiguous punctuation and enclosed numbers to east-asia only in East Asian context", () => {
    expect(officeFontSegments('“旗帜”—①', 'ascii', false, true)).toEqual([
      { role: 'east-asia', text: '“旗帜”—①', from: 0, to: 6 },
    ]);
    expect(officeFontSegments('“quote”—', 'ascii', false, false)).toEqual([
      { role: 'h-ansi', text: '“', from: 0, to: 1 },
      { role: 'ascii', text: 'quote', from: 1, to: 6 },
      { role: 'h-ansi', text: '”—', from: 6, to: 8 },
    ]);
  });
});
