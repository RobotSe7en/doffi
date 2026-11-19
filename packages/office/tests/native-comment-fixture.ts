import { body, mutateDocx, styles } from "./regression-fixtures.js";
export function nativeCommentFixture(
  placement: "inside" | "start" | "end" | "point" = "inside",
  inline = false,
): Uint8Array {
  const start = '<w:commentRangeStart w:id="7"/>',
    end = '<w:commentRangeEnd w:id="7"/>',
    reference = '<w:r><w:commentReference w:id="7"/></w:r>',
    raw =
      (placement === "inside" || placement === "start" ? start : "") +
      "<w:r><w:t>PRIVATE</w:t></w:r>" +
      (placement === "inside" || placement === "end" ? end : "") +
      reference,
    object =
      '<w:customXml w:element="payload" w:uri="urn:doffi:test">' +
      (inline ? raw : "<w:p>" + raw + "</w:p>") +
      "</w:customXml>",
    before = (placement === "end" ? start : "") + "<w:r><w:t>BEFORE</w:t></w:r>",
    after = "<w:r><w:t>AFTER</w:t></w:r>" + (placement === "start" ? end : "");
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/document.xml": (source) =>
      body(
        source,
        inline
          ? "<w:p>" + before + object + after + "</w:p>"
          : "<w:p>" + before + "</w:p>" + object + "<w:p>" + after + "</w:p>",
      ),
    "word/comments.xml": () =>
      '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="7" w:author="Original"><w:p><w:r><w:t>REVIEW</w:t></w:r></w:p></w:comment></w:comments>',
    "word/_rels/document.xml.rels": (source) =>
      source.replace(
        "</Relationships>",
        '<Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
      ),
    "[Content_Types].xml": (source) =>
      source.replace(
        "</Types>",
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
      ),
  });
}
