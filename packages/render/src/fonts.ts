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

export async function loadWordWebFonts(document: Document, fonts: readonly WordWebFont[]): Promise<HostedWordFonts> {
  const aliases = new Map<string, HostedWordFont>();
  const loaded = await Promise.allSettled(fonts.map(async (definition) => {
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
): string {
  const scope = document.createElement("div"),
    probe = document.createElement("span");
  scope.hidden = true;
  probe.style.setProperty("font-family", stack, "important");
  scope.append(probe);
  document.body.append(scope);
  try {
    const [font] = resolvePreviewFontsInScope(document, scope, "span", undefined, hosted);
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
): readonly FontResolution[] {
  return resolvePreviewFontsInScope(document, document, selector, undefined, hosted);
}
/** Limit private probes/window diagnostics to their owning surface. A selector
 * in another mounted document must not participate in font substitution. */
export function resolvePreviewFontsInScope(document: Document, scope: ParentNode, selector: string, probes?: FontProbeCache, hosted?: HostedWordFonts): readonly FontResolution[] {
  probes?.watch(document.fonts);
  let canvas: HTMLCanvasElement | undefined, context: CanvasRenderingContext2D | undefined, glyphPhase = false;
  const probe = (): CanvasRenderingContext2D => {
    if (!context) { canvas = document.createElement("canvas"); if (glyphPhase) { canvas.width = 128; canvas.height = 128; } context = canvas.getContext("2d")!; }
    return context;
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
      const candidates = /仿宋|FangSong/i.test(first)
        ? ["FandolFang", "Noto Serif CJK SC", "serif"]
        : /楷|KaiTi|SimKai/i.test(first)
          ? ["FandolKai", "LXGW WenKai", "Noto Serif CJK SC", "serif"]
          : /宋|明|SimSun|NSimSun|Ming|Mincho|Serif/i.test(first)
            ? ["Noto Serif CJK SC", "FandolSong", "Noto Sans CJK SC", "serif"]
            : ["Noto Sans CJK SC", "Noto Serif CJK SC", "sans-serif"];
      rendered = candidates.find(available)!;
    }
    if (rendered !== first) {
      if (preferred) ownViewStyle(node, "font-family");
      node.style.fontFamily = JSON.stringify(rendered) + (preferred ? "," + sourceFamily : "");
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
  glyphPhase = true;
  if (canvas) { canvas.width = 128; canvas.height = 128; }
  const hashes = new Map<string, number>();
  const unavailableAudits = new Set<string>();
  let audited = 0;
  const glyphHash = (font: string, character: string): number => {
    const key = font + "\0" + character;
    const cached = hashes.get(key);
    if (cached !== undefined) return cached;
    const measure = (): number => {
      const context = probe();
      context.clearRect(0, 0, 128, 128); context.font = "64px " + font; context.fillText(character, 16, 90);
      const pixels = context.getImageData(0, 0, 128, 128).data; let hash = 2166136261;
      for (const value of pixels) hash = Math.imul(hash ^ value, 16777619);
      return hash >>> 0;
    };
    const hash = probes ? probes.glyph(key, measure) : measure(); hashes.set(key, hash); return hash;
  };
  for (const node of nodes.filter((item) =>
    item.matches('[data-word-script="eastAsia"],[data-office-script="east-asia"],.sdk-number-marker,.sdk-ppt-marker'),
  )) {
    const resolution = resolutions.get(selected.get(node) ?? "");
    if (!resolution) continue;
    const font = document.defaultView!.getComputedStyle(node).fontFamily;
    if (unavailableAudits.has(font)) { resolution.glyphAuditComplete = false; continue; }
    for (const character of new Set(node.textContent ?? "")) {
      const key = font + "\0" + character;
      if (!hashes.has(key) && audited >= 4096) {
        resolution.glyphAuditComplete = false;
        continue;
      }
      if (!hashes.has(key)) audited++;
      try {
        if (glyphHash(font, character) === glyphHash(font, "\u{10ffff}") && glyphHash(font, character) === glyphHash(font, "\uffff")) resolution.missingCharacters.add(character);
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
