import type {
  DeepReadonly,
  WordBibliography,
  WordBibliographyStyle,
  WordCitationAuthor,
  WordCitationSource,
  WordCitationSourceType,
} from "@doffi/core";
import { WordCitationSourceTypeSchema } from "@doffi/core";
import type { OpcPackage } from "./ooxml/package.js";
import { attr, child, elements, text } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import type { OutputOpc } from "./export/opc.js";
import { declaration, escapeXml } from "./export/xml.js";

export const BIB = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
const CUSTOM_XML_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
const BIB_STORE_ITEM_ID = "{B5A2C9D4-7B8F-4E2A-9D1E-4C3B2A1F0E9D}";

const styleToXsl: Record<WordBibliographyStyle, string> = {
  APA: "\\APASixthEditionOfficeOnline.xsl",
  MLA: "\\MLASeventhEditionOfficeOnline.xsl",
  GB7714: "\\GB.XSL",
  Chicago: "\\CHICAGO.XSL",
  IEEE: "\\IEEE2006OfficeOnline.xsl",
};

function parseStyle(raw: string | undefined): WordBibliographyStyle | undefined {
  if (!raw) return undefined;
  const upper = raw.toUpperCase();
  if (upper.includes("GB")) return "GB7714";
  if (upper.includes("MLA")) return "MLA";
  if (upper.includes("CHICAGO")) return "Chicago";
  if (upper.includes("IEEE")) return "IEEE";
  if (upper.includes("APA")) return "APA";
  return undefined;
}

function readBibText(parent: XmlNode | undefined, name: string): string | undefined {
  const node = child(parent, name, BIB);
  if (!node) return undefined;
  const value = text(node).trim();
  return value || undefined;
}

function readAuthors(source: XmlNode): WordCitationAuthor[] {
  const authors: WordCitationAuthor[] = [];
  const outer = child(source, "Author", BIB);
  const inner = child(outer, "Author", BIB) ?? outer;
  if (!inner) return authors;
  const corporate = readBibText(inner, "Corporate");
  if (corporate) authors.push({ corporate });
  const nameList = child(inner, "NameList", BIB);
  for (const person of elements(nameList, "Person", BIB)) {
    const first = readBibText(person, "First");
    const middle = readBibText(person, "Middle");
    const last = readBibText(person, "Last");
    if (first || middle || last) {
      authors.push({
        ...(first ? { first } : {}),
        ...(middle ? { middle } : {}),
        ...(last ? { last } : {}),
      });
    }
  }
  return authors;
}

export function readWordBibliography(opc: OpcPackage, main: string): WordBibliography | undefined {
  for (const relation of opc.relationships(main).values()) {
    if (relation.external || !relation.type.endsWith("/customXml")) continue;
    let root: XmlNode;
    try {
      root = opc.read(relation.target);
    } catch {
      continue;
    }
    if (root.uri !== BIB || root.name !== "Sources") continue;
    const style =
      parseStyle(attr(root, "SelectedStyle", "") ?? attr(root, "StyleName", "")) ?? undefined;
    const sources: WordCitationSource[] = [];
    const seen = new Set<string>();
    for (const node of elements(root, "Source", BIB)) {
      const tag = readBibText(node, "Tag");
      const title = readBibText(node, "Title");
      if (!tag || !title || seen.has(tag)) continue;
      seen.add(tag);
      const rawType = readBibText(node, "SourceType") ?? "Misc";
      const parsedType = WordCitationSourceTypeSchema.safeParse(rawType);
      const sourceType: WordCitationSourceType = parsedType.success ? parsedType.data : "Misc";
      const authors = readAuthors(node);
      const year = readBibText(node, "Year");
      const city = readBibText(node, "City");
      const publisher = readBibText(node, "Publisher");
      const journalName = readBibText(node, "JournalName");
      const volume = readBibText(node, "Volume");
      const issue = readBibText(node, "Issue");
      const pages = readBibText(node, "Pages");
      const url = readBibText(node, "URL");
      const standardNumber = readBibText(node, "StandardNumber");
      const guid = readBibText(node, "Guid");
      const lcid = readBibText(node, "LCID");
      sources.push({
        tag,
        sourceType,
        title,
        ...(authors.length ? { authors } : {}),
        ...(year ? { year } : {}),
        ...(city ? { city } : {}),
        ...(publisher ? { publisher } : {}),
        ...(journalName ? { journalName } : {}),
        ...(volume ? { volume } : {}),
        ...(issue ? { issue } : {}),
        ...(pages ? { pages } : {}),
        ...(url ? { url } : {}),
        ...(standardNumber ? { standardNumber } : {}),
        ...(guid ? { guid } : {}),
        ...(lcid ? { lcid } : {}),
      });
    }
    if (sources.length || style) return { ...(style ? { style } : {}), sources };
  }
  return undefined;
}

function bibTag(name: string, value: string | undefined): string {
  return value ? `<b:${name}>${escapeXml(value)}</b:${name}>` : "";
}

export function writeWordBibliography(
  opc: OutputOpc,
  bibliography: DeepReadonly<WordBibliography> | undefined,
): void {
  if (!bibliography || !bibliography.sources.length) return;
  let slot = 1;
  while (opc.parts.files.has(`customXml/item${slot}.xml`)) slot++;
  const itemPath = `customXml/item${slot}.xml`;
  const propsPath = `customXml/itemProps${slot}.xml`;
  const style = bibliography.style ?? "APA";

  const out = opc.part(itemPath, "application/xml");
  out.add(
    declaration +
      `<b:Sources SelectedStyle="${escapeXml(styleToXsl[style])}" StyleName="${escapeXml(style)}" xmlns:b="${BIB}" xmlns="${BIB}">`,
  );
  for (const source of bibliography.sources) {
    out.add("<b:Source>");
    out.add(bibTag("Tag", source.tag));
    out.add(bibTag("SourceType", source.sourceType));
    if (source.guid) out.add(bibTag("Guid", source.guid));
    if (source.lcid) out.add(bibTag("LCID", source.lcid));
    if (source.authors?.length) {
      const corp = source.authors.find((author) => author.corporate?.trim());
      const persons = source.authors.filter(
        (author) => !author.corporate?.trim() && (author.last || author.first || author.middle),
      );
      out.add("<b:Author><b:Author>");
      if (corp?.corporate) out.add(bibTag("Corporate", corp.corporate));
      if (persons.length) {
        out.add("<b:NameList>");
        for (const person of persons) {
          out.add(
            "<b:Person>" +
              bibTag("Last", person.last) +
              bibTag("First", person.first) +
              bibTag("Middle", person.middle) +
              "</b:Person>",
          );
        }
        out.add("</b:NameList>");
      }
      out.add("</b:Author></b:Author>");
    }
    out.add(bibTag("Title", source.title));
    out.add(bibTag("Year", source.year));
    out.add(bibTag("City", source.city));
    out.add(bibTag("Publisher", source.publisher));
    out.add(bibTag("JournalName", source.journalName));
    out.add(bibTag("Volume", source.volume));
    out.add(bibTag("Issue", source.issue));
    out.add(bibTag("Pages", source.pages));
    out.add(bibTag("URL", source.url));
    out.add(bibTag("StandardNumber", source.standardNumber));
    out.add("</b:Source>");
  }
  out.add("</b:Sources>");
  out.finish();

  const props = opc.part(
    propsPath,
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
  );
  props.add(
    declaration +
      `<ds:datastoreItem ds:itemID="${BIB_STORE_ITEM_ID}" xmlns:ds="${CUSTOM_XML_PROPS}"><ds:schemaRefs><ds:schemaRef ds:uri="${BIB}"/></ds:schemaRefs></ds:datastoreItem>`,
  );
  props.finish();

  opc.relate("word/document.xml", "customXml", `../${itemPath}`);
  opc.relate(itemPath, "customXmlProps", `itemProps${slot}.xml`);
}
