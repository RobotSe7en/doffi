import type { FontResolution } from "./contracts.js";
import type { FontProbeCache } from "./font-probe-cache.js";
import { ownViewStyle } from "./view-nodes.js";
export type { FontResolution } from "./contracts.js";

/** Font bytes supplied by an application, independent of the document package. */
export interface WordWebFont {
  readonly family: string;
  readonly aliases: readonly string[];
  readonly load: () => Promise<ArrayBuffer>;
  readonly weight?: string;
  readonly style?: string;
  /** The aliases are names of this same native face, for example 宋体 and SimSun. */
  readonly equivalentAliases?: boolean;
  /** Use this face only when the requested face is unavailable. */
  readonly fallbackOnly?: boolean;
}

type HostedWordFont = Pick<WordWebFont, "family" | "equivalentAliases" | "fallbackOnly">;
export type HostedWordFonts = ReadonlyMap<string, string | HostedWordFont>;

const loadedDefinitions = new WeakMap<Document, Set<WordWebFont>>();
const loadedAliases = new WeakMap<Document, Map<string, HostedWordFont>>();

function extractCssFontNames(css: string, into: Set<string>): void {
  const decl = /(?:font-family|--word-font-(?:ascii|hAnsi|eastAsia|cs)|--office-(?:(?:major|minor|font)-(?:ascii|h-ansi|east-asia|complex-script|symbol|script-[A-Za-z]{4})|bullet-font))\s*:\s*([^;{}<>]+)/gi;
  for (const match of css.matchAll(decl)) {
    for (const part of families(match[1]!)) {
      const lower = part.toLowerCase();
      if (!lower || lower.startsWith("var(") || generic.has(lower) || lower === "inherit" || lower === "initial" || lower === "unset") continue;
      into.add(lower);
    }
  }
}

/** Collect lowercased font family names referenced by a document snapshot. */
export function collectDocumentFontNames(snapshot: {
  readonly html: string;
  readonly styles: readonly { readonly css: string }[];
  readonly metadata?: unknown;
}): Set<string> {
  const names = new Set<string>();
  for (const sheet of snapshot.styles) extractCssFontNames(sheet.css, names);
  for (const match of snapshot.html.matchAll(/\bstyle="([^"]*)"/gi)) {
    const raw = match[1]!.replaceAll("&quot;", '"').replaceAll("&#34;", '"').replaceAll("&#39;", "'").replaceAll("&apos;", "'");
    extractCssFontNames(raw, names);
  }
  const meta = snapshot.metadata as { word?: { numbering?: readonly { levels?: readonly { font?: string }[] }[]; charts?: readonly { textStyle?: { fonts?: Record<string, string | undefined> } }[]; equations?: readonly unknown[] } } | undefined;
  if (meta?.word) {
    for (const def of meta.word.numbering ?? []) {
      for (const level of def.levels ?? []) if (level.font) for (const part of families(level.font)) names.add(part.toLowerCase());
    }
    for (const chart of meta.word.charts ?? []) {
      for (const face of Object.values(chart.textStyle?.fonts ?? {})) if (face) names.add(face.toLowerCase());
    }
    if (meta.word.equations?.length) { names.add("cambria"); names.add("times new roman"); }
  }
  return names;
}

function selectWordWebFonts(fonts: readonly WordWebFont[], requested?: ReadonlySet<string>): readonly WordWebFont[] {
  if (!requested || !requested.size) return fonts;
  const byName = new Map<string, WordWebFont[]>();
  for (const definition of fonts) {
    for (const name of [definition.family, ...definition.aliases]) {
      const key = name.toLowerCase();
      const list = byName.get(key) ?? [];
      list.push(definition);
      byName.set(key, list);
    }
  }
  const chosen = new Set<WordWebFont>();
  const addDirect = (key: string): boolean => {
    const hits = byName.get(key);
    if (!hits?.length) return false;
    const primary = hits.filter((item) => !item.fallbackOnly);
    for (const item of primary.length ? primary : hits) chosen.add(item);
    return true;
  };
  let needFangsong = false, needKai = false, needSerif = false, needSans = false;
  for (const name of requested) {
    if (addDirect(name)) continue;
    if (/仿宋|fangsong/i.test(name)) needFangsong = true;
    else if (/楷|kaiti|simkai/i.test(name)) needKai = true;
    else if (/宋|明|simsun|nsimsun|ming|mincho|serif/i.test(name)) needSerif = true;
    else needSans = true;
  }
  const addCategory = (preferred: readonly string[], fallback: string): void => {
    for (const name of preferred) {
      const hits = byName.get(name)?.filter((item) => !item.fallbackOnly);
      if (hits?.length) { for (const item of hits) chosen.add(item); return; }
    }
    addDirect(fallback);
  };
  if (needFangsong) addCategory(["fangsong"], "doffi fangsong sc");
  if (needKai) addCategory(["kaiti"], "doffi kai sc");
  if (needSerif) addCategory(["simsun"], "doffi serif sc");
  if (needSans) {
    addCategory(["microsoft yahei", "dengxian"], "doffi sans sc");
    addCategory(["calibri", "aptos"], "doffi sans sc");
  }
  return [...chosen];
}

export function hasUnloadedWordWebFonts(
  document: Document,
  fonts: readonly WordWebFont[],
  requested?: ReadonlySet<string>,
): boolean {
  const seen = loadedDefinitions.get(document);
  const selected = selectWordWebFonts(fonts, requested);
  return selected.some((definition) => !seen?.has(definition));
}

export async function loadWordWebFonts(
  document: Document,
  fonts: readonly WordWebFont[],
  requested?: ReadonlySet<string>,
): Promise<HostedWordFonts> {
  let seen = loadedDefinitions.get(document);
  if (!seen) { seen = new Set(); loadedDefinitions.set(document, seen); }
  let aliases = loadedAliases.get(document);
  if (!aliases) { aliases = new Map(); loadedAliases.set(document, aliases); }
  const pending = selectWordWebFonts(fonts, requested).filter((definition) => !seen!.has(definition));
  const loaded = await Promise.allSettled(pending.map(async (definition) => {
    const face = new document.defaultView!.FontFace(definition.family, await definition.load(), {
      weight: definition.weight ?? "normal",
      style: definition.style ?? "normal",
    });
    await face.load();
    return { definition, face };
  }));
  for (const result of loaded) {
    if (result.status !== "fulfilled") continue;
    const { definition, face } = result.value;
    seen.add(definition);
    document.fonts.add(face);
    for (const name of [definition.family, ...definition.aliases]) {
      const key = name.toLowerCase();
      if (!definition.fallbackOnly || !aliases.has(key)) aliases.set(key, definition);
    }
  }
  return aliases;
}
const generic = new Set([
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
]);
function families(value: string): string[] {
  return (value.match(/"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|[^,]+/g) ?? [])
    .map((name) =>
      name
        .trim()
        .replace(/^['"]|['"]$/g, "")
        .replace(/\\(["'\\])/g, "$1"),
    )
    .filter(Boolean);
}

/** Resolve an SDK default before it enters the document's font diagnostics. */
export function resolvePreviewDefaultFont(
  document: Document,
  stack: string,
  hosted?: HostedWordFonts,
  probes?: FontProbeCache,
): string {
  const scope = document.createElement("div"),
    probe = document.createElement("span");
  scope.hidden = true;
  probe.style.setProperty("font-family", stack, "important");
  scope.append(probe);
  document.body.append(scope);
  try {
    const [font] = resolvePreviewFontsInScope(document, scope, "span", probes, hosted);
    if (!font || font.selectedFamily === font.requestedFamily) return stack;
    if (generic.has(font.selectedFamily.toLowerCase())) return font.selectedFamily;
    return JSON.stringify(font.selectedFamily) + "," + stack;
  } finally {
    scope.remove();
  }
}

/** Measure against multiple fallback metrics; this neither requests permission nor loads fonts. */
export function resolvePreviewFonts(
  document: Document,
  selector: string = "[data-node-id]",
  hosted?: HostedWordFonts,
  probes?: FontProbeCache,
): readonly FontResolution[] {
  return resolvePreviewFontsInScope(document, document, selector, probes, hosted);
}
/** Limit private probes/window diagnostics to their owning surface. A selector
 * in another mounted document must not participate in font substitution. */
export function resolvePreviewFontsInScope(document: Document, scope: ParentNode, selector: string, probes?: FontProbeCache, hosted?: HostedWordFonts): readonly FontResolution[] {
  hosted ??= loadedAliases.get(document);
  probes?.watch(document.fonts);
  let context: CanvasRenderingContext2D | undefined;
  const probe = (): CanvasRenderingContext2D => {
    const create = () => {
      const canvas = document.createElement("canvas");
      canvas.width = canvas.height = 128;
      return canvas.getContext("2d", { willReadFrequently: true })!;
    };
    return context ??= probes ? probes.canvas(create) : create();
  };
  const cache = new Map<string, boolean>();
  const resolutions = new Map<
    string,
    {
      requestedFamily: string;
      selectedFamily: string;
      substituted: boolean;
      missingCharacters: Set<string>;
      glyphAuditComplete: boolean;
    }
  >();
  const selected = new WeakMap<HTMLElement, string>();
  const available = (family: string): boolean => {
    if (generic.has(family.toLowerCase())) return true;
    const known = cache.get(family);
    if (known !== undefined) return known;
    const measure = (): boolean => {
      const context = probe();
      const quoted = JSON.stringify(family), sample = "汉字中文，。AaWim0123"; let found = false;
      for (const base of ["monospace", "serif", "sans-serif"]) {
        context.font = "72px " + base; const fallback = context.measureText(sample).width;
        context.font = "72px " + quoted + "," + base;
        if (Math.abs(context.measureText(sample).width - fallback) > 0.01) found = true;
      }
      return found;
    };
    const found = probes ? probes.available(family, measure) : measure();
    cache.set(family, found);
    return found;
  };
  const nodes = [...scope.querySelectorAll<HTMLElement>(selector)];
  // Resolve parents before children so inherited fallback and native east-Asian overrides remain distinct.
  for (const node of nodes) {
    const sourceFamily = document.defaultView!.getComputedStyle(node).fontFamily;
    const requested = families(sourceFamily);
    if (!requested.length) continue;
    const first = requested[0]!;
    const candidate = hosted?.get(first.toLowerCase());
    const configured = typeof candidate === "string" ? { family: candidate } : candidate;
    const preferred = configured?.fallbackOnly && available(first) ? undefined : configured;
    let rendered = preferred?.family ?? requested.find(available);
    if (!rendered) {
      const hostedFallback = (names: readonly string[]): string | undefined => {
        if (!hosted) return undefined;
        for (const name of names) {
          const entry = hosted.get(name.toLowerCase());
          const item = typeof entry === "string" ? { family: entry } : entry;
          if (item && !item.fallbackOnly) return item.family;
        }
        return undefined;
      };
      const isFangsong = /仿宋|FangSong/i.test(first),
        isKai = /楷|KaiTi|SimKai/i.test(first),
        isSerif = /宋|明|SimSun|NSimSun|Ming|Mincho|Serif/i.test(first),
        isLatinScript = (node.dataset.officeScript === "ascii"
          || node.dataset.officeScript === "h-ansi"
          || node.dataset.wordScript === "ascii"
          || node.dataset.wordScript === "hAnsi")
          && !node.parentElement?.querySelector('[data-office-script="east-asia"],[data-word-script="eastAsia"]');
      const candidates = isFangsong
        ? ["FandolFang", "STFangsong", "Noto Serif CJK SC", "serif"]
        : isKai
          ? ["FandolKai", "LXGW WenKai", "Kaiti SC", "STKaiti", "Noto Serif CJK SC", "serif"]
          : isSerif
            ? ["Noto Serif CJK SC", "FandolSong", "Songti SC", "STSong", "Noto Sans CJK SC", "serif"]
            : ["Noto Sans CJK SC", "PingFang SC", "Noto Serif CJK SC", "sans-serif"];
      rendered =
        hostedFallback(
          isFangsong
            ? ["FangSong"]
            : isKai
              ? ["KaiTi"]
              : isSerif
                ? ["SimSun"]
                : isLatinScript
                  ? ["Calibri", "Aptos", "Arial", "Microsoft YaHei", "DengXian"]
                  : ["Microsoft YaHei", "DengXian"],
        ) ?? candidates.find(available)!;
    }
    if (rendered !== first) {
      if (preferred) ownViewStyle(node, "font-family");
      const cssFamily = generic.has(rendered.toLowerCase()) ? rendered : JSON.stringify(rendered);
      node.style.fontFamily = cssFamily + (preferred ? "," + sourceFamily : "");
    }
    selected.set(node, first);
    if (!resolutions.has(first))
      resolutions.set(first, {
        requestedFamily: first,
        selectedFamily: rendered,
        substituted: rendered !== first && !preferred?.equivalentAliases,
        missingCharacters: new Set(),
        glyphAuditComplete: true,
      });
  }
  // A local raster comparison also catches the generic missing-glyph box. This
  // is a preview diagnostic, not a promise of platform-independent font coverage.
  const hashes = new Map<string, number>();
  const unavailableAudits = new Set<string>();
  const auditedCharacters = new Set<string>();
  const glyphHash = (font: string, character: string, small = false): number => {
    const key = font + "\0" + character + (small ? "\0small" : "");
    const cached = hashes.get(key);
    if (cached !== undefined) return cached;
    const measure = (): number => {
      const context = probe();
      const extent = small ? 24 : 128;
      context.clearRect(0, 0, extent, extent);
      context.font = (small ? "8px " : "64px ") + font;
      context.fillText(character, small ? 2 : 16, small ? 18 : 90);
      const pixels = context.getImageData(0, 0, extent, extent).data; let hash = 2166136261;
      for (const value of pixels) hash = Math.imul(hash ^ value, 16777619);
      return hash >>> 0;
    };
    const hash = probes ? probes.glyph(key, measure) : measure(); hashes.set(key, hash); return hash;
  };
  const missingGlyph = (font: string, character: string): boolean => {
    const supplementary = glyphHash(font, "\u{10ffff}"), bmp = glyphHash(font, "\uffff");
    if (supplementary === bmp) return glyphHash(font, character) === bmp;

    // Some engines put the code point inside their missing-glyph box. At a
    // small size they draw only the border. Compare two controls from the same
    // Unicode range so the box has the same width, and require visible ink.
    const [first, second] = character.length === 2 ? ["\u{10fffe}", "\u{10ffff}"] : ["\ufffe", "\uffff"];
    const outline = glyphHash(font, first, true);
    return outline === glyphHash(font, second, true) &&
      outline !== glyphHash(font, "", true) &&
      outline === glyphHash(font, character, true);
  };
  for (const node of nodes.filter((item) =>
    item.matches('[data-word-script="eastAsia"],[data-office-script="east-asia"],.sdk-number-marker,.sdk-ppt-marker'),
  )) {
    const resolution = resolutions.get(selected.get(node) ?? "");
    if (!resolution) continue;
    const font = document.defaultView!.getComputedStyle(node).fontFamily;
    if (unavailableAudits.has(font)) { resolution.glyphAuditComplete = false; continue; }
    for (const character of new Set(node.textContent ?? "")) {
      // Whitespace intentionally has no ink. Comparing its empty raster with
      // a font's blank missing-glyph control cannot establish a missing glyph.
      if (/\p{White_Space}/u.test(character)) continue;
      const key = font + "\0" + character;
      if (!auditedCharacters.has(key) && auditedCharacters.size >= 4096) {
        resolution.glyphAuditComplete = false;
        continue;
      }
      auditedCharacters.add(key);
      try {
        if (missingGlyph(font, character)) resolution.missingCharacters.add(character);
      } catch (error) {
        // Platform/privacy raster failures cannot establish missing glyphs.
        // Keep rendering with an explicit incomplete audit, and stop retrying
        // this face. Source font names and already established results survive.
        if (!(error instanceof document.defaultView!.DOMException)) throw error;
        unavailableAudits.add(font); resolution.glyphAuditComplete = false; break;
      }
    }
  }
  return Object.freeze(
    [...resolutions.values()].map((item) =>
      Object.freeze({
        ...item,
        missingCharacters: Object.freeze([...item.missingCharacters]),
      }),
    ),
  );
}
