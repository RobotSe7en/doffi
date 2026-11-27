import { DocumentError } from "@doffi/core";
import type {
  DeepReadonly,
  ContentDocumentMetadata,
  ContentDocumentModel,
  DocumentTransaction,
  WordMetadata,
} from "@doffi/core";
import { evaluateWordTableFormula, formatWordDatePicture } from "./word-table-formulas.js";
import {
  applyWordFieldSwitches,
  evaluateBibliographyField,
  evaluateCitationField,
  evaluateIfOrCompare,
  evaluateSymbolField,
  formatIndexField,
  formatToaField,
  parseTaInstruction,
  parseXeInstruction,
  tokenizeWordFieldArgs,
  type ParsedTaEntry,
  type ParsedXeEntry,
} from "./word-fields-eval.js";
import type {
  DocumentRenderReport,
  LayoutDiagnostic,
  LayoutFieldResult,
  RenderedPage,
} from "./contracts.js";
import { hideFromViewSelection, markViewNode, measureViewContent } from "./view-nodes.js";
import type { PreviewSource } from "./source-index.js";

interface FieldView {
  readonly id: string;
  readonly kind: LayoutFieldResult["kind"];
  readonly instruction: string;
  readonly results: readonly string[];
  readonly display: HTMLElement;
}
export interface LayoutFields {
  readonly changed: boolean;
  evaluate(
    pages: readonly RenderedPage[],
    nodePages: Readonly<Record<string, readonly string[]>>,
  ): readonly LayoutFieldResult[];
  repeat(pages: readonly RenderedPage[]): void;
  readonly diagnostics: readonly LayoutDiagnostic[];
}
function displayText(document: Document, target: HTMLElement, value: string): void {
  target.replaceChildren();
  let role = "ascii",
    part = "";
  const append = (): void => {
    if (!part) return;
    const span = document.createElement("span");
    span.dataset.wordScript = role;
    span.style.fontFamily = `var(--word-font-${role},var(--word-font-ascii,inherit))`;
    span.textContent = part;
    target.append(span);
    part = "";
  };
  for (const character of value) {
    const next = /\p{Mark}/u.test(character)
      ? role
      : /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(character)
        ? "eastAsia"
        : /[\p{Script=Arabic}\p{Script=Hebrew}]/u.test(character)
          ? "cs"
          : character.codePointAt(0)! <= 0x7f
            ? "ascii"
            : "hAnsi";
    if (next !== role) {
      append();
      role = next;
    }
    part += character;
  }
  append();
}

function displayFieldValue(
  document: Document,
  target: HTMLElement,
  value: string,
  kind: FieldView["kind"],
): void {
  target.dataset.sdkFieldValue = value;
  if (kind === "TOA" && value.includes("\t")) {
    target.replaceChildren();
    const lines = value.split("\n");
    for (const line of lines) {
      const tabIdx = line.lastIndexOf("\t");
      if (tabIdx >= 0) {
        const row = document.createElement("div");
        row.className = "sdk-toa-entry";
        row.style.cssText =
          "display:flex;align-items:baseline;width:100%;padding-left:18pt;box-sizing:border-box;";
        const label = document.createElement("span");
        label.className = "sdk-toa-label";
        displayText(document, label, line.slice(0, tabIdx));
        const leader = document.createElement("span");
        leader.className = "sdk-toa-leader";
        leader.setAttribute("aria-hidden", "true");
        leader.style.cssText =
          "flex:1;min-width:12pt;border-bottom:1pt dotted currentColor;margin:0 4pt;opacity:0.75;font-size:0;line-height:0;overflow:hidden;";
        leader.textContent = "\t";
        const pages = document.createElement("span");
        pages.className = "sdk-toa-pages";
        pages.style.whiteSpace = "nowrap";
        displayText(document, pages, line.slice(tabIdx + 1));
        row.append(label, leader, pages);
        target.append(row);
      } else {
        const heading = document.createElement("div");
        heading.className = "sdk-toa-heading";
        heading.style.cssText = "margin-top:2pt;margin-bottom:4pt;";
        displayText(document, heading, line);
        target.append(heading);
      }
    }
    return;
  }
  displayText(document, target, value);
}

/** Results are a reversible display projection. Explicit source cache updates are
 * returned as ordinary core operations by createLayoutFieldTransaction below. */
export function layoutFields(
  document: Document,
  source: PreviewSource,
  word: DeepReadonly<WordMetadata> | undefined,
  metadata?: DeepReadonly<ContentDocumentMetadata>,
): LayoutFields {
  const views: FieldView[] = [],
    diagnostics: LayoutDiagnostic[] = [];
  const bookmarkNames = new Set(word?.bookmarks.map((bookmark) => bookmark.name));
  for (const link of word?.hyperlinks ?? [])
    if (link.bookmark && !link.target && !bookmarkNames.has(link.bookmark))
      diagnostics.push({
        code: "UnresolvedReference",
        nodeId: link.nodeId,
        message: "The internal hyperlink target is not present in this document.",
      });
  let changed = false;
  const all = word?.fields ?? [];
  const fieldsById = new Map(all.map((field) => [field.nodeId, field]));
  const instructionOnly = new Set(
    all.flatMap((field) => field.instructionFields?.map((item) => item.fieldNodeId) ?? []),
  );
  const styles = new Map(word?.styles.map((style) => [style.styleId, style]));
  const headings = [...source.root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")].flatMap(
    (element) => {
      let level =
        element.dataset.wordOutlineLevel !== undefined
          ? Number(element.dataset.wordOutlineLevel)
          : undefined;
      for (
        let style = styles.get(element.dataset.wordStyle ?? "");
        level === undefined && style;
        style = style.basedOn ? styles.get(style.basedOn) : undefined
      )
        level = style.outlineLevel;
      level ??= /^h[1-6]$/.test(element.localName) ? Number(element.localName[1]) - 1 : undefined;
      return level !== undefined && level < 9
        ? [
            {
              id: element.dataset.nodeId!,
              level: level + 1,
              text: source.elements.get(element.dataset.nodeId!)?.element.textContent?.trim() ?? "",
            },
          ]
        : [];
    },
  );

  // Mark XE / TA / SET invisible fields as hidden markers (visible when Show/Hide ¶ is enabled).
  for (const field of all) {
    const rawKind = field.instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? "";
    if (rawKind === "XE" || rawKind === "TA" || rawKind === "SET") {
      const host = source.elements.get(field.nodeId)?.element;
      if (host) {
        host.dataset.sdkWordHidden = "true";
        if (!host.textContent?.trim()) {
          const badge = markViewNode(document.createElement("span"));
          hideFromViewSelection(badge);
          badge.className = "sdk-hidden-field-marker";
          badge.textContent = `{ ${field.instruction.trim()} }`;
          host.append(badge);
        }
      }
    }
  }

  const supportedKinds = new Set([
    "PAGE",
    "NUMPAGES",
    "TOC",
    "TOF",
    "INDEX",
    "TOA",
    "CITATION",
    "BIBLIOGRAPHY",
    "REF",
    "PAGEREF",
    "NOTEREF",
    "STYLEREF",
    "SEQ",
    "IF",
    "COMPARE",
    "SET",
    "QUOTE",
    "SYMBOL",
    "DATE",
    "TIME",
    "CREATEDATE",
    "SAVEDATE",
    "PRINTDATE",
    "EDITTIME",
    "AUTHOR",
    "TITLE",
    "SUBJECT",
    "KEYWORDS",
    "COMMENTS",
    "LASTSAVEDBY",
    "REVNUM",
    "DOCPROPERTY",
    "FILENAME",
    "TEMPLATE",
    "NUMWORDS",
    "NUMCHARS",
    "SECTION",
    "SECTIONPAGES",
    "FORMULA",
  ]);
  for (const field of all) {
    const trimmedInstr = field.instruction.trim();
    const rawKind = trimmedInstr.startsWith("=")
      ? "FORMULA"
      : trimmedInstr.split(/\s+/)[0]!.toUpperCase();
    if (rawKind === "XE" || rawKind === "TA") continue;
    const isTof = rawKind === "TOC" && /\\c\s+(?:"[^"]+"|\S+)/i.test(field.instruction);
    const kind = isTof ? "TOF" : rawKind;
    if (field.locked) {
      diagnostics.push({
        code: "UnsupportedField",
        nodeId: field.nodeId,
        message: "Locked field result remains cached.",
      });
      continue;
    }
    if (!supportedKinds.has(kind) || instructionOnly.has(field.nodeId)) {
      if (!instructionOnly.has(field.nodeId))
        diagnostics.push({
          code: "UnsupportedField",
          nodeId: field.nodeId,
          message: "Field result remains cached: " + rawKind,
        });
      continue;
    }
    const results = field.resultNodeIds ?? [],
      first =
        (results[0] ? source.elements.get(results[0])?.element : undefined) ??
        source.elements.get(field.nodeId)?.element;
    if (!first) {
      diagnostics.push({
        code: "UnsupportedField",
        nodeId: field.nodeId,
        message: "Field has no editable cached result scope.",
      });
      continue;
    }
    if (kind === "TOC" && /\\(?:f|a|b|c|s|t)\b/i.test(field.instruction)) {
      diagnostics.push({
        code: "UnsupportedField",
        nodeId: field.nodeId,
        message: "This TOC switch combination remains cached.",
      });
      continue;
    }
    const display = markViewNode(document.createElement("span"));
    measureViewContent(display);
    display.className = "sdk-field-display";
    display.dataset.sdkFieldNodeId = field.nodeId;
    display.style.whiteSpace = "pre-wrap";
    // The cache can contain a cleared/formatted run below its result scope.
    // Preserve that first run's appearance when replacing only the displayed
    // value; the source cache and its formatting remain independently editable.
    const formatted = "[data-word-reset-text],[data-word-current-format],[data-word-run-overlay]";
    if (first.matches(formatted) || first.closest(formatted) || first.querySelector(formatted)) {
      const walker = document.createTreeWalker(first, 4),
        text = walker.nextNode(),
        owner = text?.parentElement;
      if (owner) {
        const css = document.defaultView!.getComputedStyle(owner);
        for (const name of [
          "font-family",
          "font-size",
          "font-weight",
          "font-style",
          "font-variant",
          "text-transform",
          "color",
          "background-color",
          "vertical-align",
          "letter-spacing",
          "line-height",
          "text-decoration-line",
          "text-decoration-style",
          "text-decoration-color",
          ...[...css].filter((name) =>
            /^--word-(?:font-|bold-|italic-|language-|kerning-)/.test(name),
          ),
        ])
          display.style.setProperty(name, css.getPropertyValue(name));
      }
    }
    if (
      kind === "TOC" ||
      kind === "TOF" ||
      kind === "INDEX" ||
      kind === "TOA" ||
      kind === "BIBLIOGRAPHY"
    ) {
      display.style.display = "block";
      if (kind === "INDEX") {
        const colMatch = /\\c\s+"?([1-4])"?/i.exec(field.instruction);
        const cols = colMatch ? Number(colMatch[1]) : 2;
        if (cols > 1) {
          display.style.columnCount = String(cols);
          display.style.columnGap = "24pt";
        }
      }
    }
    for (const id of results) {
      const result = source.elements.get(id)?.element;
      if (!result) continue;
      const cache = markViewNode(document.createElement("span"), "unwrap");
      hideFromViewSelection(cache);
      cache.style.cssText = "position:absolute;visibility:hidden;pointer-events:none";
      cache.append(...result.childNodes);
      result.append(cache);
      result.contentEditable = "false";
      source.setFieldView(id, display);
    }
    first.append(display);
    views.push({
      id: field.nodeId,
      kind: kind as FieldView["kind"],
      instruction: field.instruction,
      results,
      display,
    });
  }

  function resolveInstruction(
    fieldId: string | undefined,
    rawInstruction: string,
    pages: readonly RenderedPage[],
    nodePages: Readonly<Record<string, readonly string[]>>,
    page: RenderedPage | undefined,
    depth: number,
  ): string {
    if (depth >= 8) return rawInstruction;
    let resolved = rawInstruction;
    const def = fieldId ? fieldsById.get(fieldId) : undefined;
    if (def?.instructionFields?.length) {
      const sorted = [...def.instructionFields].sort((a, b) => b.offset - a.offset);
      for (const nested of sorted) {
        const childField = fieldsById.get(nested.fieldNodeId);
        if (!childField) continue;
        const val = evaluateInstruction(
          childField.nodeId,
          childField.instruction,
          pages,
          nodePages,
          page,
          depth + 1,
        );
        resolved = resolved.slice(0, nested.offset) + val + resolved.slice(nested.offset);
      }
    }
    // Also support inline nested `{ ... }` subfields authored directly in instructions.
    let guard = 0;
    while (resolved.includes("{") && guard++ < 16) {
      const inner = /\{([^{}]+)\}/.exec(resolved);
      if (!inner) break;
      const subVal = evaluateInstruction(
        undefined,
        inner[1]!.trim(),
        pages,
        nodePages,
        page,
        depth + 1,
      );
      resolved =
        resolved.slice(0, inner.index) + subVal + resolved.slice(inner.index + inner[0].length);
    }
    return resolved;
  }

  function computeDocumentCounts(): { words: number; chars: number } {
    const mainStory = word?.stories.find((story) => story.kind === "main");
    const rootEl =
      (mainStory ? source.elements.get(mainStory.rootNodeId)?.element : undefined) ?? source.root;
    const text = (rootEl.textContent ?? "").trim();
    if (!text) return { words: 0, chars: 0 };
    const cjk = (text.match(/[\u3400-\u9fff\uf900-\ufaff]/g) ?? []).length;
    const latin = (text.replace(/[\u3400-\u9fff\uf900-\ufaff]/g, " ").match(/\S+/g) ?? []).length;
    return { words: cjk + latin, chars: text.replace(/\s+/g, "").length };
  }

  function collectSetVariables(
    pages: readonly RenderedPage[],
    nodePages: Readonly<Record<string, readonly string[]>>,
    page: RenderedPage | undefined,
    depth: number,
  ): Map<string, string> {
    const vars = new Map<string, string>();
    for (const field of all) {
      if (!/^\s*SET\s+/i.test(field.instruction)) continue;
      const resolved = resolveInstruction(
        field.nodeId,
        field.instruction,
        pages,
        nodePages,
        page,
        depth + 1,
      );
      const args = tokenizeWordFieldArgs(resolved.replace(/^\s*SET\s+/i, ""));
      if (args.length >= 2) vars.set(args[0]!, args[1]!);
    }
    return vars;
  }

  function collectCitations(): { order: Map<string, number>; tags: string[] } {
    const ordered = all
      .filter((f) => /^\s*CITATION\s+/i.test(f.instruction))
      .sort(
        (a, b) =>
          (source.elements.get(a.nodeId)?.from ?? 0) - (source.elements.get(b.nodeId)?.from ?? 0),
      );
    const order = new Map<string, number>();
    const tags: string[] = [];
    for (const f of ordered) {
      const m = /^\s*CITATION\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(f.instruction);
      const tag = m?.[1] ?? m?.[2];
      if (tag && !order.has(tag)) {
        order.set(tag, order.size + 1);
        tags.push(tag);
      }
    }
    return { order, tags };
  }

  function evaluateInstruction(
    fieldId: string | undefined,
    rawInstruction: string,
    pages: readonly RenderedPage[],
    nodePages: Readonly<Record<string, readonly string[]>>,
    page: RenderedPage | undefined,
    depth: number,
  ): string {
    const instruction = resolveInstruction(fieldId, rawInstruction, pages, nodePages, page, depth);
    const trimmed = instruction.trim();
    const rawKind = trimmed.startsWith("=") ? "FORMULA" : (trimmed.split(/\s+/)[0]?.toUpperCase() ?? "");
    const kind = rawKind === "TOC" && /\\c\s+(?:"[^"]+"|\S+)/i.test(instruction) ? "TOF" : rawKind;
    const isZh = (metadata?.locale ?? "zh-CN").toLowerCase().startsWith("zh");

    if (kind === "SET") return "";
    if (kind === "IF" || kind === "COMPARE") {
      return applyWordFieldSwitches(evaluateIfOrCompare(instruction, kind === "COMPARE"), instruction);
    }
    if (kind === "QUOTE") {
      const args = tokenizeWordFieldArgs(instruction.replace(/^\s*QUOTE\s+/i, ""));
      const textParts: string[] = [];
      for (let i = 0; i < args.length; i++) {
        if (/^\\[*#@a-z]$/i.test(args[i]!)) {
          i++;
          continue;
        }
        textParts.push(args[i]!);
      }
      return applyWordFieldSwitches(textParts.join(" "), instruction);
    }
    if (kind === "SYMBOL") return evaluateSymbolField(instruction);
    if (kind === "CITATION") {
      const { order } = collectCitations();
      return applyWordFieldSwitches(evaluateCitationField(instruction, word, order), instruction);
    }
    if (kind === "BIBLIOGRAPHY") {
      const { tags } = collectCitations();
      return evaluateBibliographyField(word, tags);
    }
    if (kind === "INDEX") {
      const entries: ParsedXeEntry[] = [];
      for (const f of all) {
        if (!/^\s*XE\s+/i.test(f.instruction)) continue;
        const parsed = parseXeInstruction(f.instruction);
        if (!parsed) continue;
        const p = pages.find((item) => nodePages[f.nodeId]?.includes(item.pageId)) ?? pages[0];
        entries.push({
          ...parsed,
          pageNumber: p?.number ?? 1,
          pageLabel: p?.label ?? "1",
        });
      }
      return formatIndexField(entries, instruction);
    }
    if (kind === "TOA") {
      const entries: ParsedTaEntry[] = [];
      for (const f of all) {
        if (!/^\s*TA\s+/i.test(f.instruction)) continue;
        const parsed = parseTaInstruction(f.instruction);
        if (!parsed) continue;
        const p = pages.find((item) => nodePages[f.nodeId]?.includes(item.pageId)) ?? pages[0];
        entries.push({
          ...parsed,
          pageNumber: p?.number ?? 1,
          pageLabel: p?.label ?? "1",
        });
      }
      return formatToaField(entries, instruction);
    }
    if (kind === "STYLEREF") {
      const m = /^\s*STYLEREF\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(instruction);
      const targetStyle = (m?.[1] ?? m?.[2] ?? "").trim();
      const searchFromBottom = /\\l\b/i.test(instruction);
      const wantNumber = /\\[nrw]\b/i.test(instruction);
      const normalizedTarget = targetStyle.toLowerCase().replace(/\s+/g, "");
      const headingAlias = /^标题([1-9])$/.exec(normalizedTarget)?.[1];
      const matchesStyle = (el: HTMLElement): boolean => {
        const sid = el.dataset.wordStyle ?? "";
        const styleDef = styles.get(sid);
        const candidates = [sid, styleDef?.name ?? "", styleDef?.aliases ?? ""]
          .map((s) => s.toLowerCase().replace(/\s+/g, ""))
          .filter(Boolean);
        if (candidates.includes(normalizedTarget)) return true;
        if (headingAlias && candidates.includes("heading" + headingAlias)) return true;
        if (/^[1-9]$/.test(normalizedTarget) && candidates.includes("heading" + normalizedTarget))
          return true;
        return false;
      };
      const allParas = [...source.root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")].filter(
        matchesStyle,
      );
      const currentPage =
        page ??
        (fieldId ? pages.find((item) => nodePages[fieldId]?.includes(item.pageId)) : undefined) ??
        pages[0];
      const pageIdx = currentPage ? pages.indexOf(currentPage) : 0;
      const onCurrentPage = currentPage
        ? allParas.filter((p) => nodePages[p.dataset.nodeId ?? ""]?.includes(currentPage.pageId))
        : [];
      let chosen: HTMLElement | undefined = searchFromBottom
        ? onCurrentPage.at(-1)
        : onCurrentPage[0];
      if (!chosen && currentPage) {
        const prior = allParas.filter((p) => {
          const pIds = nodePages[p.dataset.nodeId ?? ""] ?? [];
          const idx = pages.findIndex((pg) => pIds.includes(pg.pageId));
          return idx >= 0 && idx < pageIdx;
        });
        chosen = prior.at(-1) ?? allParas[0];
      }
      if (!chosen) return "";
      const raw = wantNumber
        ? (chosen.querySelector(".sdk-number-marker")?.textContent?.trim() ?? "")
        : (chosen.textContent?.replace(
            chosen.querySelector(".sdk-number-marker")?.textContent ?? "",
            "",
          ).trim() ?? "");
      return applyWordFieldSwitches(raw, instruction);
    }
    if (["REF", "PAGEREF", "NOTEREF"].includes(kind)) {
      const match = /^\s*(?:REF|PAGEREF|NOTEREF)\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(instruction),
        name = match?.[1] ?? match?.[2];
      if (kind === "REF" && name) {
        const setVars = collectSetVariables(pages, nodePages, page, depth);
        if (setVars.has(name)) return applyWordFieldSwitches(setVars.get(name)!, instruction);
      }
      const bookmark = word?.bookmarks.find((bookmark) => bookmark.name === name),
        start = bookmark ? source.elements.get(bookmark.startNodeId) : undefined,
        end = bookmark ? source.elements.get(bookmark.endNodeId) : undefined;
      if (!start || !end) return "错误！未找到引用源。";
      const selfFrom = fieldId ? (source.elements.get(fieldId)?.from ?? 0) : 0;
      const relPos = /\\p\b/i.test(instruction)
        ? start.from < selfFrom
          ? isZh
            ? "上文"
            : "above"
          : isZh
            ? "下文"
            : "below"
        : "";
      if (kind === "PAGEREF") {
        const pg =
          pages.find((item) => nodePages[bookmark!.startNodeId]?.includes(item.pageId))?.label ??
          "1";
        return applyWordFieldSwitches(relPos ? `${pg} (${relPos})` : pg, instruction);
      }
      if (kind === "NOTEREF") {
        const reference = word?.noteReferences?.find((reference) => {
          const node = source.elements.get(reference.nodeId);
          return node && node.from >= start.from && node.to <= end.from;
        });
        if (!reference) return "错误！未找到注释。";
        const marker = source.elements.get(reference.nodeId)?.element;
        const val = marker?.lastElementChild?.textContent ?? reference.customMark ?? "1";
        return applyWordFieldSwitches(relPos ? `${val} ${relPos}` : val, instruction);
      }
      if (/\\[nrw]\b/i.test(instruction)) {
        const para = start.element.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li");
        const markerText = para?.querySelector(".sdk-number-marker")?.textContent?.trim() ?? "";
        return applyWordFieldSwitches(
          relPos ? `${markerText} ${relPos}`.trim() : markerText,
          instruction,
        );
      }
      let previous: HTMLElement | undefined;
      const bodyText = source.tokens
        .filter((token) => token.to > start.from && token.from < end.from)
        .map((token) => {
          const paragraph =
              token.owner.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li") ?? token.owner,
            prefix = previous && previous !== paragraph ? "\n" : "";
          previous = paragraph;
          return (
            prefix +
            token.text.slice(
              Math.max(0, start.from - token.from),
              Math.min(token.text.length, end.from - token.from),
            )
          );
        })
        .join("");
      return applyWordFieldSwitches(relPos ? `${bodyText} ${relPos}`.trim() : bodyText, instruction);
    }
    if (kind === "SEQ") {
      const match = /^\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(instruction),
        name = match?.[1] ?? match?.[2],
        seqFields = all
          .filter((field) => /^\s*SEQ\s+/i.test(field.instruction))
          .sort(
            (a, b) =>
              (source.elements.get(a.nodeId)?.from ?? 0) -
              (source.elements.get(b.nodeId)?.from ?? 0),
          );
      let number = 0;
      for (const field of seqFields) {
        const key = /^\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(field.instruction);
        if ((key?.[1] ?? key?.[2]) !== name) continue;
        const restart = /\\r\s+(\d+)/i.exec(field.instruction);
        number = restart
          ? Number(restart[1])
          : /\\c\b/i.test(field.instruction)
            ? number
            : number + 1;
        if (field.nodeId === fieldId) break;
      }
      if (/\\h\b/i.test(instruction) && !/\\\*\s+/i.test(instruction)) return "";
      return applyWordFieldSwitches(String(number), instruction);
    }
    if (kind === "DATE" || kind === "TIME" || kind === "CREATEDATE" || kind === "SAVEDATE" || kind === "PRINTDATE")
      return formatWordDatePicture(new Date(), instruction, kind === "TIME");
    if (kind === "EDITTIME") return applyWordFieldSwitches("1", instruction);
    if (kind === "AUTHOR") return applyWordFieldSwitches(word?.properties?.creator ?? "", instruction);
    if (kind === "LASTSAVEDBY") return applyWordFieldSwitches(word?.properties?.creator ?? "", instruction);
    if (kind === "REVNUM") return applyWordFieldSwitches("1", instruction);
    if (kind === "TITLE") return applyWordFieldSwitches(metadata?.title ?? "", instruction);
    if (kind === "SUBJECT") return applyWordFieldSwitches(word?.properties?.subject ?? "", instruction);
    if (kind === "KEYWORDS") return applyWordFieldSwitches(word?.properties?.keywords ?? "", instruction);
    if (kind === "COMMENTS") return applyWordFieldSwitches(word?.properties?.description ?? "", instruction);
    if (kind === "FILENAME") {
      const name = (metadata?.title?.trim() || "文档") + ".docx";
      return applyWordFieldSwitches(name, instruction);
    }
    if (kind === "TEMPLATE") {
      const name = word?.template?.name || word?.template?.target || "Normal.dotm";
      return applyWordFieldSwitches(name, instruction);
    }
    if (kind === "NUMWORDS") return applyWordFieldSwitches(String(computeDocumentCounts().words), instruction);
    if (kind === "NUMCHARS") return applyWordFieldSwitches(String(computeDocumentCounts().chars), instruction);
    if (kind === "DOCPROPERTY") {
      const m = /^\s*DOCPROPERTY\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(instruction);
      const prop = (m?.[1] ?? m?.[2] ?? "").toLowerCase();
      let raw = "";
      if (prop === "author") raw = word?.properties?.creator ?? "";
      else if (prop === "title") raw = metadata?.title ?? "";
      else if (prop === "subject") raw = word?.properties?.subject ?? "";
      else if (prop === "keywords") raw = word?.properties?.keywords ?? "";
      else if (prop === "comments" || prop === "description") raw = word?.properties?.description ?? "";
      else if (prop === "category") raw = word?.properties?.category ?? "";
      else if (prop === "template") raw = word?.template?.name ?? "Normal.dotm";
      else if (prop === "pages") raw = String(pages.length);
      else if (prop === "words") raw = String(computeDocumentCounts().words);
      else if (prop === "characters") raw = String(computeDocumentCounts().chars);
      else if (prop === "lastsavedby") raw = word?.properties?.creator ?? "";
      else if (prop === "revnum" || prop === "revisionnumber") raw = "1";
      return applyWordFieldSwitches(raw, instruction);
    }
    if (kind === "SECTION" || kind === "SECTIONPAGES") {
      const current =
        page ??
        (fieldId ? pages.find((item) => nodePages[fieldId]?.includes(item.pageId)) : undefined) ??
        pages[0];
      if (!current) return "1";
      if (kind === "SECTION") {
        const idx = metadata?.sections?.findIndex((s) => s.sectionId === current.sectionId) ?? 0;
        return applyWordFieldSwitches(String(Math.max(1, idx + 1)), instruction);
      }
      return applyWordFieldSwitches(
        String(pages.filter((p) => p.sectionId === current.sectionId).length || 1),
        instruction,
      );
    }
    if (kind === "FORMULA") {
      const host = (fieldId ? source.elements.get(fieldId)?.element : undefined) ?? source.root;
      return applyWordFieldSwitches(evaluateWordTableFormula(instruction, host), instruction);
    }
    if (kind === "TOF") {
      const match = /\\c\s+(?:"([^"]+)"|(\S+))/i.exec(instruction);
      const label = match?.[1] ?? match?.[2] ?? "图";
      const seqIds = new Set(
        all
          .filter((f) => {
            const m = /^\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(f.instruction);
            return (m?.[1] ?? m?.[2]) === label;
          })
          .map((f) => f.nodeId),
      );
      const entries: string[] = [];
      for (const p of source.root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")) {
        const hasSeq = [...p.querySelectorAll<HTMLElement>("[data-node-id]")].some((el) =>
          seqIds.has(el.dataset.nodeId!),
        );
        if (!hasSeq) continue;
        const pId = p.dataset.nodeId!;
        const pPage = pages.find((item) => nodePages[pId]?.includes(item.pageId));
        entries.push((p.textContent?.trim() ?? label) + "\t" + (pPage?.label ?? "1"));
      }
      return entries.join("\n");
    }
    if (kind === "NUMPAGES") return applyWordFieldSwitches(String(pages.length), instruction);
    if (kind === "PAGE") {
      const current =
        page ??
        (fieldId ? pages.find((item) => nodePages[fieldId]?.includes(item.pageId)) : undefined) ??
        pages[0];
      if (!current) return "1";
      if (/\\\*\s+/.test(instruction)) {
        return applyWordFieldSwitches(String(current.number), instruction);
      }
      return current.label;
    }
    if (kind === "TOC") {
      const levels = /\\o\s+"([1-6])-([1-6])"/i.exec(instruction),
        min = Number(levels?.[1] ?? 1),
        max = Number(levels?.[2] ?? 3);
      return headings
        .filter((item) => item.level >= min && item.level <= max)
        .map((heading) => {
          const current = pages.find((item) => nodePages[heading.id]?.includes(item.pageId));
          return heading.text + "\t" + (current?.label ?? "1");
        })
        .join("\n");
    }
    // Fallback for bare bookmark reference inside nested `{ bookmarkName }`
    const setVars = collectSetVariables(pages, nodePages, page, depth);
    if (setVars.has(trimmed)) return setVars.get(trimmed)!;
    return "";
  }

  function value(
    view: FieldView,
    pages: readonly RenderedPage[],
    nodePages: Readonly<Record<string, readonly string[]>>,
    page?: RenderedPage,
  ): string {
    return evaluateInstruction(view.id, view.instruction, pages, nodePages, page, 0);
  }
  return {
    diagnostics,
    get changed() {
      return changed;
    },
    evaluate(pages, nodePages) {
      changed = false;
      const drawingFields = new Map<string, HTMLElement[]>();
      for (const node of document.querySelectorAll<HTMLElement>(
        ".sdk-word-drawing-text .sdk-field-display[data-sdk-field-node-id]",
      )) {
        if (node.closest(".sdk-paper")) continue;
        const id = node.dataset.sdkFieldNodeId!,
          instances = drawingFields.get(id) ?? [];
        instances.push(node);
        drawingFields.set(id, instances);
      }
      return Object.freeze(
        views.map((view) => {
          const result = value(view, pages, nodePages);
          if (view.display.dataset.sdkFieldValue !== result) {
            displayFieldValue(document, view.display, result, view.kind);
            changed = true;
          }
          for (const node of drawingFields.get(view.id) ?? []) {
            const drawing = node.closest<HTMLElement>("[data-sdk-word-object]"),
              ownerPages = nodePages[drawing?.dataset.sdkWordObject ?? ""],
              ownerPage = ownerPages && pages.find((page) => ownerPages.includes(page.pageId)),
              result = value(view, pages, nodePages, ownerPage);
            if (node.dataset.sdkFieldValue !== result) {
              displayFieldValue(document, node, result, view.kind);
              changed = true;
            }
          }
          return Object.freeze({
            fieldNodeId: view.id,
            resultNodeIds: view.results,
            kind: view.kind,
            value: result,
          });
        }),
      );
    },
    repeat(pages) {
      for (const page of pages) {
        const paper = document.querySelector<HTMLElement>(
          '.sdk-paper[data-page-id="' + CSS.escape(page.pageId) + '"]',
        );
        const repeated = document.querySelectorAll<HTMLElement>(
          '.sdk-repeat-header[data-sdk-repeat-page-id="' +
            CSS.escape(page.pageId) +
            '"] .sdk-field-display[data-sdk-field-node-id]',
        );
        for (const node of [
          ...(paper?.querySelectorAll<HTMLElement>(".sdk-field-display[data-sdk-field-node-id]") ??
            []),
          ...repeated,
        ]) {
          const view = views.find((item) => item.id === node.dataset.sdkFieldNodeId);
          if (!view) continue;
          displayFieldValue(document, node, value(view, pages, {}, page), view.kind);
        }
      }
    },
  };
}

export function createLayoutFieldTransaction(
  model: ContentDocumentModel,
  report: DocumentRenderReport,
): DocumentTransaction {
  const snapshot = model.snapshot();
  if (report.layout !== "pages" || !report.pages.length)
    throw new DocumentError("InvalidOperation", "Layout fields require a completed pages layout.");
  if (
    report.documentId !== snapshot.manifest.documentId ||
    report.revision !== snapshot.manifest.revision
  )
    throw new DocumentError(
      "RevisionConflict",
      "Layout fields belong to a different document revision.",
    );
  const operations: DocumentTransaction["operations"] = [];
  const seen = new Set<string>();
  for (const field of report.fields) {
    const definition = snapshot.metadata.word?.fields.find(
      (item) => item.nodeId === field.fieldNodeId,
    );
    if (
      !definition ||
      JSON.stringify(definition.resultNodeIds ?? []) !== JSON.stringify(field.resultNodeIds)
    )
      throw new DocumentError(
        "InvalidOperation",
        "Layout field scopes differ from the document source.",
      );
    for (const [index, nodeId] of field.resultNodeIds.entries()) {
      if (seen.has(nodeId))
        throw new DocumentError("InvalidOperation", "Layout field result scopes overlap.");
      seen.add(nodeId);
      const node = model.getNode(nodeId);
      if (!node)
        throw new DocumentError("InvalidOperation", "Layout field result scope is missing.");
      const text = index === 0 ? field.value : "";
      if (node.text !== text)
        operations.push({
          op: "replaceText",
          nodeId,
          from: 0,
          to: node.text.length,
          expectedText: node.text,
          text,
        });
    }
  }
  return {
    protocolVersion: 1,
    documentId: snapshot.manifest.documentId,
    baseRevision: snapshot.manifest.revision,
    operations,
  };
}
