import { z } from "zod";
import { IdentifierSchema } from "./contracts/primitives.js";
import type { DeepReadonly } from "./runtime.js";

export const WordCitationSourceTypeSchema = z.enum([
  "Book",
  "BookSection",
  "JournalArticle",
  "ArticleInAPeriodical",
  "ConferenceProceedings",
  "Report",
  "InternetSite",
  "DocumentFromInternetSite",
  "ElectronicSource",
  "Misc",
]);
export type WordCitationSourceType = z.infer<typeof WordCitationSourceTypeSchema>;

export const WordCitationAuthorSchema = z.strictObject({
  first: z.string().max(128).optional(),
  middle: z.string().max(128).optional(),
  last: z.string().max(128).optional(),
  corporate: z.string().max(256).optional(),
});
export type WordCitationAuthor = z.infer<typeof WordCitationAuthorSchema>;

export const WordCitationSourceSchema = z.strictObject({
  tag: IdentifierSchema,
  sourceType: WordCitationSourceTypeSchema,
  title: z.string().min(1).max(512),
  authors: z.array(WordCitationAuthorSchema).max(64).optional(),
  year: z.string().max(32).optional(),
  city: z.string().max(128).optional(),
  publisher: z.string().max(256).optional(),
  journalName: z.string().max(256).optional(),
  volume: z.string().max(64).optional(),
  issue: z.string().max(64).optional(),
  pages: z.string().max(64).optional(),
  url: z.string().max(1024).optional(),
  standardNumber: z.string().max(128).optional(),
  guid: z.string().max(64).optional(),
  lcid: z.string().max(16).optional(),
});
export type WordCitationSource = z.infer<typeof WordCitationSourceSchema>;

export const WordBibliographyStyleSchema = z.enum([
  "APA",
  "MLA",
  "GB7714",
  "Chicago",
  "IEEE",
]);
export type WordBibliographyStyle = z.infer<typeof WordBibliographyStyleSchema>;

export const WordBibliographySchema = z.strictObject({
  style: WordBibliographyStyleSchema.optional(),
  sources: z.array(WordCitationSourceSchema).max(1000),
});
export type WordBibliography = z.infer<typeof WordBibliographySchema>;

export interface WordCitationFormatOptions {
  readonly pages?: string;
  readonly suppressAuthor?: boolean;
  readonly suppressYear?: boolean;
  readonly suppressTitle?: boolean;
  readonly ordinal?: number;
}

const isCjkText = (value: string): boolean => /[\u3400-\u9fff\uf900-\ufaff]/.test(value);

function authorShortName(author: WordCitationAuthor): string {
  if (author.corporate?.trim()) return author.corporate.trim();
  const last = author.last?.trim() ?? "";
  const first = author.first?.trim() ?? "";
  if (isCjkText(last + first)) return (last + first).trim();
  return last || first || "Unknown";
}

function initials(value: string | undefined): string {
  if (!value?.trim()) return "";
  return value
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part[0]!.toUpperCase() + ".")
    .join(" ");
}

function formatAuthorsInline(
  authors: readonly DeepReadonly<WordCitationAuthor>[] | undefined,
  style: WordBibliographyStyle,
): string {
  if (!authors?.length) return "";
  const names = authors.map(authorShortName).filter(Boolean);
  if (!names.length) return "";
  if (names.length === 1) return names[0]!;
  if (names.length === 2) {
    const joiner = style === "APA" ? " & " : " and ";
    return names[0]! + joiner + names[1]!;
  }
  return names[0]! + (isCjkText(names[0]!) ? "等" : " et al.");
}

function formatAuthorEntry(author: WordCitationAuthor, style: WordBibliographyStyle): string {
  if (author.corporate?.trim()) return author.corporate.trim();
  const last = author.last?.trim() ?? "";
  const first = author.first?.trim() ?? "";
  const middle = author.middle?.trim() ?? "";
  if (isCjkText(last + first + middle)) return (last + first + middle).trim();
  if (style === "APA") {
    const init = [initials(first), initials(middle)].filter(Boolean).join(" ");
    return last && init ? `${last}, ${init}` : last || first || "Unknown";
  }
  if (style === "IEEE") {
    const init = [initials(first), initials(middle)].filter(Boolean).join(" ");
    return init && last ? `${init} ${last}` : last || first || "Unknown";
  }
  if (style === "GB7714") {
    const given = [first, middle].filter(Boolean).join(" ");
    const init = given
      ? given
          .split(/\s+/)
          .filter(Boolean)
          .map((part) => part[0]!.toUpperCase())
          .join(" ")
      : "";
    return last && init ? `${last.toUpperCase()} ${init}` : last || first || "Unknown";
  }
  // MLA / Chicago
  const given = [first, middle].filter(Boolean).join(" ");
  return last && given ? `${last}, ${given}` : last || given || "Unknown";
}

function gb7714DocCode(type: WordCitationSourceType): string {
  switch (type) {
    case "JournalArticle":
    case "ArticleInAPeriodical":
      return "J";
    case "ConferenceProceedings":
      return "C";
    case "Report":
      return "R";
    case "InternetSite":
    case "DocumentFromInternetSite":
    case "ElectronicSource":
      return "EB/OL";
    case "Book":
    case "BookSection":
      return "M";
    default:
      return "Z";
  }
}

export function formatWordCitation(
  source: DeepReadonly<WordCitationSource>,
  style: WordBibliographyStyle = "APA",
  options: WordCitationFormatOptions = {},
): string {
  const pages = options.pages?.trim();
  if (style === "IEEE" || style === "GB7714") {
    if (options.ordinal !== undefined && options.ordinal > 0) {
      if (style === "IEEE") return pages ? `[${options.ordinal}, p. ${pages}]` : `[${options.ordinal}]`;
      return pages ? `[${options.ordinal}]:${pages}` : `[${options.ordinal}]`;
    }
  }
  const author = options.suppressAuthor ? "" : formatAuthorsInline(source.authors, style);
  const year = options.suppressYear ? "" : (source.year?.trim() ?? "");
  const title = options.suppressTitle ? "" : source.title.trim();
  const lead = author || title || source.tag;

  if (style === "MLA") {
    const parts = [lead, pages].filter(Boolean);
    return `(${parts.join(" ")})`;
  }
  if (style === "Chicago") {
    const main = [lead, year].filter(Boolean).join(" ");
    return pages ? `(${main}, ${pages})` : `(${main})`;
  }
  // APA / fallback
  const parts = [lead, year, pages ? `p. ${pages}` : ""].filter(Boolean);
  return `(${parts.join(", ")})`;
}

export function formatWordBibliographyEntry(
  source: DeepReadonly<WordCitationSource>,
  style: WordBibliographyStyle = "APA",
  ordinal = 1,
): string {
  const formattedAuthors = (source.authors ?? [])
    .map((author) => formatAuthorEntry(author, style))
    .filter(Boolean);
  const authorText =
    formattedAuthors.length === 0
      ? ""
      : formattedAuthors.length === 1
        ? formattedAuthors[0]!
        : style === "APA"
          ? formattedAuthors.slice(0, -1).join(", ") + ", & " + formattedAuthors.at(-1)!
          : style === "GB7714"
            ? formattedAuthors.slice(0, 3).join(", ") +
              (formattedAuthors.length > 3
                ? isCjkText(formattedAuthors[0]!)
                  ? ", 等"
                  : ", et al."
                : "")
            : formattedAuthors.slice(0, -1).join(", ") + ", and " + formattedAuthors.at(-1)!;
  const year = source.year?.trim() ?? "";
  const title = source.title.trim();
  const city = source.city?.trim() ?? "";
  const publisher = source.publisher?.trim() ?? "";
  const journal = source.journalName?.trim() ?? "";
  const volume = source.volume?.trim() ?? "";
  const issue = source.issue?.trim() ?? "";
  const pages = source.pages?.trim() ?? "";
  const url = source.url?.trim() ?? "";

  if (style === "GB7714") {
    const code = gb7714DocCode(source.sourceType);
    const pieces: string[] = [];
    if (authorText) pieces.push(authorText + ".");
    pieces.push(`${title}[${code}].`);
    if (journal) {
      const volIssue = [
        year,
        volume ? (issue ? `${volume}(${issue})` : volume) : issue ? `(${issue})` : "",
      ]
        .filter(Boolean)
        .join(", ");
      pieces.push(
        `${journal}${volIssue ? `, ${volIssue}` : ""}${pages ? `: ${pages}` : ""}.`,
      );
    } else {
      const pub = [city && publisher ? `${city}: ${publisher}` : city || publisher, year]
        .filter(Boolean)
        .join(", ");
      if (pub) pieces.push(pub + (pages ? `: ${pages}` : "") + ".");
    }
    if (url) pieces.push(url + ".");
    return `[${ordinal}] ${pieces.join(" ")}`;
  }

  if (style === "IEEE") {
    const pieces: string[] = [];
    if (authorText) pieces.push(authorText + ",");
    if (journal) {
      pieces.push(`"${title},"`);
      pieces.push(journal + ",");
      if (volume) pieces.push(`vol. ${volume},`);
      if (issue) pieces.push(`no. ${issue},`);
      if (pages) pieces.push(`pp. ${pages},`);
      if (year) pieces.push(year + ".");
    } else {
      pieces.push(title + ".");
      const pub = [city && publisher ? `${city}: ${publisher}` : city || publisher, year]
        .filter(Boolean)
        .join(", ");
      if (pub) pieces.push(pub + ".");
    }
    if (url) pieces.push(url);
    return `[${ordinal}] ${pieces.join(" ").replace(/,\s*$/, ".")}`;
  }

  if (style === "MLA") {
    const pieces: string[] = [];
    if (authorText) pieces.push(authorText + ".");
    if (journal) {
      pieces.push(`"${title}."`);
      const detail = [
        journal,
        volume ? `vol. ${volume}` : "",
        issue ? `no. ${issue}` : "",
        year,
        pages ? `pp. ${pages}` : "",
      ]
        .filter(Boolean)
        .join(", ");
      if (detail) pieces.push(detail + ".");
    } else {
      pieces.push(title + ".");
      const pub = [publisher, year].filter(Boolean).join(", ");
      if (pub) pieces.push(pub + ".");
    }
    if (url) pieces.push(url + ".");
    return pieces.join(" ");
  }

  if (style === "Chicago") {
    const pieces: string[] = [];
    if (authorText) pieces.push(authorText + ".");
    if (year) pieces.push(year + ".");
    if (journal) {
      pieces.push(`"${title}."`);
      const vol = [volume, issue ? `no. ${issue}` : ""].filter(Boolean).join(", ");
      pieces.push(`${journal}${vol ? ` ${vol}` : ""}${pages ? `: ${pages}` : ""}.`);
    } else {
      pieces.push(title + ".");
      const pub = city && publisher ? `${city}: ${publisher}` : city || publisher;
      if (pub) pieces.push(pub + ".");
    }
    if (url) pieces.push(url + ".");
    return pieces.join(" ");
  }

  // APA 7th
  const pieces: string[] = [];
  if (authorText) pieces.push(authorText);
  if (year) pieces.push(`(${year}).`);
  pieces.push(title + ".");
  if (journal) {
    const vol = volume ? (issue ? `${volume}(${issue})` : volume) : issue ? `(${issue})` : "";
    const detail = [journal, vol, pages].filter(Boolean).join(", ");
    if (detail) pieces.push(detail + ".");
  } else if (publisher) {
    pieces.push(publisher + ".");
  }
  if (url) pieces.push(url);
  return pieces.join(" ");
}

export function formatWordBibliography(
  sources: readonly DeepReadonly<WordCitationSource>[],
  style: WordBibliographyStyle = "APA",
  citedTags?: readonly string[],
): { readonly tag: string; readonly text: string }[] {
  const byTag = new Map(sources.map((source) => [source.tag, source]));
  let ordered: DeepReadonly<WordCitationSource>[];
  if (style === "IEEE" || style === "GB7714") {
    const seen = new Set<string>();
    ordered = [];
    for (const tag of citedTags ?? []) {
      const source = byTag.get(tag);
      if (source && !seen.has(tag)) {
        seen.add(tag);
        ordered.push(source);
      }
    }
    for (const source of sources) {
      if (!seen.has(source.tag)) {
        seen.add(source.tag);
        ordered.push(source);
      }
    }
  } else {
    ordered = [...sources].sort((a, b) => {
      const keyA = (formatAuthorsInline(a.authors, style) || a.title).toLowerCase();
      const keyB = (formatAuthorsInline(b.authors, style) || b.title).toLowerCase();
      return keyA.localeCompare(keyB, undefined, { sensitivity: "base" });
    });
  }
  return ordered.map((source, index) => ({
    tag: source.tag,
    text: formatWordBibliographyEntry(source, style, index + 1),
  }));
}
