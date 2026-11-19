import { body, mutateDocx, styles } from "./regression-fixtures.js";
export const nativeRangeFamilies = [
  "bookmark",
  "perm",
  "moveFromRange",
  "moveToRange",
  "customXmlInsRange",
  "customXmlDelRange",
  "customXmlMoveFromRange",
  "customXmlMoveToRange",
];
export function nativeRangeFixture(family: string, inside: "start" | "end" | "both" = "end") {
  const extra =
      family === "bookmark"
        ? ' w:name="CrossNative"'
        : family === "perm"
          ? ' w:edGrp="everyone"'
          : (family.startsWith("move") ? ' w:name="CrossNative"' : "") +
            ' w:author="Original" w:date="2026-11-16T00:00:00Z"',
    start = `<w:${family}Start w:id="9"${extra}/>`,
    end = `<w:${family}End w:id="9"/>`,
    native = (content: string) =>
      `<w:customXml w:element="payload" w:uri="urn:test"><w:p>${content}</w:p></w:customXml>`;
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/document.xml": (source) =>
      body(
        source,
        inside === "both"
          ? native(start + "<w:r><w:t>PRIVATE-A</w:t></w:r>") +
              "<w:p><w:r><w:t>EDITABLE</w:t></w:r></w:p>" +
              native("<w:r><w:t>PRIVATE-B</w:t></w:r>" + end)
          : "<w:p>" +
              (inside === "end" ? start : "") +
              "<w:r><w:t>BEFORE</w:t></w:r></w:p>" +
              native(
                (inside === "start" ? start : "") +
                  "<w:r><w:t>PRIVATE</w:t></w:r>" +
                  (inside === "end" ? end : ""),
              ) +
              "<w:p><w:r><w:t>AFTER</w:t></w:r>" +
              (inside === "start" ? end : "") +
              "</w:p>",
      ),
  });
}
