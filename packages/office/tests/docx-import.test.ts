import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import type { WordDocumentMetadata } from "@doffi/core";
import { DocumentError } from "@doffi/core";
import { importDocx as importNativeDocx, readDocumentZip, writeDocumentZip } from "@doffi/office";
// These historical rejection cases exercise the explicitly retained strict policy.
const importDocx: typeof importNativeDocx = (bytes, options) =>
  importNativeDocx(bytes, { ...options, unsupportedContent: "reject" });
import { zipSync, unzipSync, strFromU8, strToU8 } from "fflate";

const directory = new URL("../../../examples/docx/", import.meta.url);
const cases = JSON.parse(readFileSync(new URL("expected.json", directory), "utf8")) as {
  file: string;
  requiredText: string[];
}[];
const fixture = (name = cases[0]!.file): Uint8Array =>
  new Uint8Array(readFileSync(new URL(name, directory)));
function mutate(name: string, part: string, change: (source: string) => string): Uint8Array {
  const files = unzipSync(fixture(name));
  files[part] = strToU8(change(strFromU8(files[part]!)));
  return zipSync(files);
}
for (const item of cases) {
  test("imports independent DOCX and reopens its package: " + item.file, () => {
    const input = fixture(item.file);
    const before = new Uint8Array(input);
    const files = importDocx(input, { documentId: "import-test" });
    const model = openDocument(files);
    expect(input).toEqual(before);
    expect(model.snapshot().manifest).toMatchObject({
      documentId: "import-test",
      revision: 0,
      kind: "document",
    });
    const snapshot = model.snapshot();
    for (const expected of item.requiredText)
      expect(snapshot.html.replace(/<[^>]*>/g, "")).toContain(expected);
    expect(openDocument(model.serialize()).snapshot()).toEqual(snapshot);
    expect(openDocument(readDocumentZip(writeDocumentZip(files))).snapshot()).toEqual(snapshot);
    expect(importDocx(input, { documentId: "import-test" })).toEqual(files);
  });
}

test("preserves style inheritance, direct overrides, sections and shared stories", () => {
  const snapshot = openDocument(
    importDocx(fixture("04-styles-sections.docx"), { documentId: "styles" }),
  ).snapshot();
  expect(snapshot.metadata.sections).toHaveLength(2);
  expect(snapshot.metadata.sections![1]!.headers).toEqual(snapshot.metadata.sections![0]!.headers);
  expect(snapshot.metadata.word!.stories.filter((x) => x.kind === "header")).toHaveLength(1);
  expect(snapshot.styles.map((item) => item.css).join("\n")).toContain("#123456");
  expect(snapshot.html).toContain("#cc0000");
  expect(snapshot.html + snapshot.styles.map((item) => item.css).join("\n")).toContain(
    "font-weight:bold",
  );
  expect(snapshot.html + snapshot.styles.map((item) => item.css).join("\n")).toContain(
    "font-style:italic",
  );
  expect(snapshot.metadata.pageSetup!.width).toBe(595.3);
});
test("preserves native rich comments, UTF-16 anchors and insertion/deletion records", () => {
  const model = openDocument(
    importDocx(fixture("08-comments-revisions.docx"), { documentId: "review" }),
  );
  const snapshot = model.snapshot();
  expect(snapshot.annotations).toHaveLength(1);
  const comment = snapshot.annotations[0]!;
  expect(comment.instruction).toBeUndefined();
  expect(comment.author?.name).toBe("审阅者");
  expect(model.getNode(comment.bodyNodeId!)!.text).toBe("请修改时间安排");
  expect(comment.anchor.kind).toBe("text-set");
  if (comment.anchor.kind === "text-set")
    expect(comment.anchor.ranges.map((range) => range.quote).join("")).toBe("第三季度完成项目");
  expect(snapshot.metadata.word!.revisions.map((x) => x.kind)).toEqual(["delete", "insert"]);
});
test("retains complex, nested and multi-paragraph field codes and original caches", () => {
  const snapshot = openDocument(
    importDocx(fixture("07-fields.docx"), { documentId: "fields" }),
  ).snapshot();
  expect(snapshot.metadata.word!.fields).toHaveLength(5);
  const conditional = snapshot.metadata.word!.fields.find((x) => x.instruction.includes(" IF "))!;
  expect(conditional.instructionFields).toHaveLength(1);
  const toc = snapshot.metadata.word!.fields.find((x) => x.instruction.includes(" TOC "))!;
  expect(toc.endNodeId).toBeDefined();
  expect(toc.resultNodeIds!.length).toBeGreaterThan(1);
  expect(snapshot.metadata.word!.fields.every((x) => x.resultState === "cached")).toBe(true);
});
test("retains horizontal and vertical table merges and immutable image bytes", () => {
  const table = openDocument(
    importDocx(fixture("02-table-merges.docx"), { documentId: "table" }),
  ).snapshot();
  expect(table.html).toContain('colspan="2"');
  expect(table.html).toContain('rowspan="2"');
  const input = fixture("03-images.docx");
  const packageFiles = importDocx(input, { documentId: "images" });
  const model = openDocument(packageFiles);
  expect(model.snapshot().manifest.assets).toHaveLength(3);
  const original = unzipSync(input);
  for (const [index, ext] of ["png", "jpg", "webp"].entries())
    expect(packageFiles.get("assets/image-" + (index + 1) + "." + ext)).toEqual(
      original["word/media/sample." + ext],
    );
});

test.each([
  ["unsupported shape", "<w:p><w:r><w:pict/></w:r></w:p>", "UnsupportedFeature"],
  ["malformed format revision", "<w:p><w:pPr><w:pPrChange/></w:pPr></w:p>", "InvalidDocument"],
  ["malformed move revision", "<w:p><w:moveFrom/></w:p>", "InvalidDocument"],
  ["hanging punctuation", "<w:p><w:pPr><w:overflowPunct/></w:pPr></w:p>", "UnsupportedFeature"],
  ["unknown namespace", '<x:content xmlns:x="urn:unknown"/>', "UnsupportedFeature"],
] as const)("rejects %s without partial output", (_name, content, code) => {
  const bytes = mutate(cases[0]!.file, "word/document.xml", (source) =>
    source.replace("<w:body>", "<w:body>" + content),
  );
  expect(() => importDocx(bytes, { documentId: "bad" })).toThrowError(
    expect.objectContaining({ code }),
  );
});
test("rejects DTD, malformed XML, missing relations, style cycles and budget overflow", () => {
  const dtd = mutate(cases[0]!.file, "word/document.xml", (source) =>
    source.replace(
      "<w:document",
      "<!DOCTYPE x [<!ENTITY boom SYSTEM 'file:///etc/passwd'>]><w:document",
    ),
  );
  expect(() => importDocx(dtd, { documentId: "dtd" })).toThrowError(
    expect.objectContaining({ code: "InvalidDocument" }),
  );
  const malformed = mutate(cases[0]!.file, "word/document.xml", (source) =>
    source.replace("</w:body>", ""),
  );
  expect(() => importDocx(malformed, { documentId: "xml" })).toThrowError(DocumentError);
  const missing = mutate(cases[0]!.file, "word/_rels/document.xml.rels", (source) =>
    source.replace('Target="styles.xml"', 'Target="missing.xml"'),
  );
  expect(() => importDocx(missing, { documentId: "rels" })).toThrowError(DocumentError);
  const cyclic = mutate(cases[0]!.file, "word/styles.xml", (source) =>
    source.replace(
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>',
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:basedOn w:val="Derived"/>',
    ),
  );
  expect(() => importDocx(cyclic, { documentId: "style" })).toThrowError(DocumentError);
  expect(() =>
    importDocx(fixture(), { documentId: "budget", limits: { maxNodes: 20 } }),
  ).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
});
test("rejects executable hyperlink targets and unresolved bookmark targets", () => {
  for (const target of ["javascript:alert(1)", "data:text/html,bad"]) {
    const bytes = mutate("06-links-tabs.docx", "word/_rels/document.xml.rels", (source) =>
      source.replace("https://example.com/document", target),
    );
    expect(() => importDocx(bytes, { documentId: "link" })).toThrowError(DocumentError);
  }
  const bytes = mutate("06-links-tabs.docx", "word/document.xml", (source) =>
    source.replace('w:anchor="目标"', 'w:anchor="missing"'),
  );
  expect(() => importDocx(bytes, { documentId: "bookmark" })).toThrowError(DocumentError);
});
test("uses relationship targets and namespace URIs rather than hard-coded prefixes and paths", () => {
  const files = unzipSync(fixture());
  files["word/custom.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!)
      .replaceAll("xmlns:w=", "xmlns:doc=")
      .replaceAll("<w:", "<doc:")
      .replaceAll("</w:", "</doc:")
      .replaceAll(" w:", " doc:"),
  );
  delete files["word/document.xml"];
  files["word/_rels/custom.xml.rels"] = files["word/_rels/document.xml.rels"]!;
  delete files["word/_rels/document.xml.rels"];
  for (const part of ["_rels/.rels", "[Content_Types].xml"])
    files[part] = strToU8(
      strFromU8(files[part]!).replaceAll("word/document.xml", "word/custom.xml"),
    );
  expect(
    openDocument(importDocx(zipSync(files), { documentId: "paths" })).snapshot().html,
  ).toContain("实施报告");
});

test("resolves conditional table styles with direct cell overrides", () => {
  const files = unzipSync(fixture("02-table-merges.docx"));
  const style =
    '<w:style w:type="table" w:styleId="Banded"><w:name w:val="Banded"/><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="ddeeff"/></w:tcPr></w:tblStylePr></w:style>';
  files["word/styles.xml"] = strToU8(
    strFromU8(files["word/styles.xml"]!).replace("</w:styles>", style + "</w:styles>"),
  );
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace(
      "<w:tblPr>",
      '<w:tblPr><w:tblStyle w:val="Banded"/><w:tblLook w:firstRow="1" w:noHBand="1" w:noVBand="1"/>',
    ),
  );
  const snapshot = openDocument(importDocx(zipSync(files), { documentId: "condition" })).snapshot();
  expect(snapshot.metadata.word!.styles.find((x) => x.styleId === "Banded")!.conditions).toEqual([
    {
      region: "firstRow",
      cssClass: expect.any(String),
      components: {
        p: expect.any(String),
        r: expect.any(String),
        cell: expect.any(String),
        row: expect.any(String),
        table: expect.any(String),
      },
    },
  ]);
  expect(snapshot.styles.map((item) => item.css).join("\n")).toContain("background-color:#ddeeff");
  expect(snapshot.metadata.word!.tables[0]!.lookMask).toBe(1568);
  expect(snapshot.html + snapshot.styles.map((item) => item.css).join("\n")).toContain(
    "font-weight:bold",
  );
});
test("uses strict namespaces and preserves native style toggles over paragraph formatting", () => {
  const files = unzipSync(fixture("04-styles-sections.docx"));
  for (const path of Object.keys(files))
    if (path.endsWith(".xml") || path.endsWith(".rels"))
      files[path] = strToU8(
        strFromU8(files[path]!)
          .replaceAll(
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
          )
          .replaceAll(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
          ),
      );
  const model = openDocument(importDocx(zipSync(files), { documentId: "strict" }));
  expect(model.snapshot().html).toContain("继承样式");
  const toggled = mutate("04-styles-sections.docx", "word/document.xml", (source) =>
    source
      .replace("直接覆盖颜色</w:t>", "直接覆盖颜色</w:t>")
      .replace('<w:color w:val="cc0000"/>', '<w:color w:val="cc0000"/><w:b w:val="0"/>'),
  );
  expect(openDocument(importDocx(toggled, { documentId: "toggle" })).snapshot().html).toContain(
    "font-weight:normal",
  );
});
test("uses a supported compatibility fallback without guessing an unknown branch", () => {
  const bytes = mutate(cases[0]!.file, "word/document.xml", (source) =>
    source.replace(
      "<w:body>",
      '<w:body><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:z="urn:unsupported"><mc:Choice Requires="z"><z:unknown/></mc:Choice><mc:Fallback><w:p><w:r><w:t>可用后备</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>',
    ),
  );
  expect(openDocument(importDocx(bytes, { documentId: "fallback" })).snapshot().html).toContain(
    "可用后备",
  );
});

test("retains east-Asian font roles independently of Latin fonts and UTF-16 text", () => {
  const model = openDocument(importDocx(fixture("09-chinese-fonts.docx"), { documentId: "fonts" }));
  const source = model.snapshot().html;
  expect(source).toContain("Liberation Serif");
  expect(source).toContain("Noto Serif CJK SC");
  expect(source).toContain("--word-font-eastAsia");
  expect(source).toContain("𠀀");
  expect(source).toContain("Dagent Office Missing CJK Test Font");
});

test("validates semantic field, story and rich-comment references on reopening", () => {
  const base = importDocx(fixture("07-fields.docx"), {
    documentId: "semantics",
  });
  for (const change of [
    (data: WordDocumentMetadata) => {
      data.word!.fields[1]!.endNodeId = "absent";
    },
    (data: WordDocumentMetadata) => {
      data.word!.fields[2]!.instructionFields = [
        { offset: 0, fieldNodeId: data.word!.fields[2]!.nodeId },
      ];
    },
    (data: WordDocumentMetadata) => {
      data.word!.stories.push({
        ...data.word!.stories[0]!,
        storyId: "duplicate-root",
      });
    },
  ]) {
    const files = new Map(base);
    const data = JSON.parse(
      new TextDecoder().decode(files.get("metadata.json")),
    ) as WordDocumentMetadata;
    change(data);
    files.set("metadata.json", new TextEncoder().encode(JSON.stringify(data)));
    expect(() => openDocument(files)).toThrowError(
      expect.objectContaining({ code: "InvalidDocument" }),
    );
  }
});

test("maps Word theme tints and Chinese theme fonts using source language", () => {
  const files = unzipSync(fixture());
  const theme =
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="fixture"><a:accent2><a:srgbClr val="4F81BD"/></a:accent2></a:clrScheme><a:fontScheme name="fixture"><a:minorFont><a:latin typeface="Liberation Serif"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Japanese fixture face"/><a:font script="Hans" typeface="Noto Serif CJK SC"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>';
  files["word/theme/theme.xml"] = strToU8(theme);
  files["word/_rels/document.xml.rels"] = strToU8(
    strFromU8(files["word/_rels/document.xml.rels"]!).replace(
      "</Relationships>",
      '<Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme.xml"/></Relationships>',
    ),
  );
  files["word/styles.xml"] = strToU8(
    strFromU8(files["word/styles.xml"]!).replace(
      '<w:sz w:val="24"/>',
      '<w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/>',
    ),
  );
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace(
      "<w:b/>",
      '<w:b/><w:rFonts w:ascii="OldLatin" w:asciiTheme="minorAscii" w:hAnsi="OldHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsia="OldFace" w:eastAsiaTheme="minorEastAsia"/><w:color w:themeColor="accent2" w:themeTint="99"/>',
    ),
  );
  const source = openDocument(importDocx(zipSync(files), { documentId: "theme" })).snapshot().html;
  expect(source).toContain("#95b3d7");
  expect(source).toContain("Noto Serif CJK SC");
  expect(source).toContain("Liberation Serif");
  expect(source).not.toContain("OldLatin");
  expect(source).not.toContain("OldHAnsi");
  expect(source).not.toContain("OldFace");
  expect(source).not.toContain("Japanese fixture face");
});

for (const item of cases) {
  test(
    "reads a LibreOffice-produced DOCX and preserves its current text and package: " + item.file,
    () => {
      const bytes = new Uint8Array(readFileSync(new URL("producer/" + item.file, directory)));
      const model = openDocument(importDocx(bytes, { documentId: "producer" }));
      expect(model.snapshot().metadata.word!.automaticHyphenation).toBe(true);
      if (item.file === "08-comments-revisions.docx")
        expect(model.snapshot().html).toContain("--word-overflow-punctuation:0");
      const original = unzipSync(bytes);
      // This independent extraction does not use the SDK's OOXML or HTML mapping.
      const documentXml = strFromU8(original["word/document.xml"]!);
      const expected = [...documentXml.matchAll(/<w:(t|delText)(?:\s[^>]*)?>([\s\S]*?)<\/w:\1>/g)]
        .map((match) =>
          match[2]!
            .replace(/&#x([0-9a-f]+);/gi, (_whole, value: string) =>
              String.fromCodePoint(parseInt(value, 16)),
            )
            .replace(/&#(\d+);/g, (_whole, value: string) => String.fromCodePoint(Number(value)))
            .replace(/&lt;/g, "<")
            .replace(/&gt;/g, ">")
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
            .replace(/&amp;/g, "&"),
        )
        .join("");
      const main = model.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!;
      expect(model.getNode(main.rootNodeId)!.text.replace(/[\n\t\ufffc]/g, "")).toBe(expected);
      expect(openDocument(model.serialize()).snapshot()).toEqual(model.snapshot());
    },
  );
}

test("preserves field lock and dirty flags without updating the source cache", () => {
  const bytes = mutate("07-fields.docx", "word/document.xml", (source) =>
    source.replace(
      '<w:fldSimple w:instr=" PAGE "',
      '<w:fldSimple w:fldLock="1" w:dirty="1" w:instr=" PAGE "',
    ),
  );
  const model = openDocument(importDocx(bytes, { documentId: "locked" }));
  expect(model.snapshot().metadata.word!.fields[0]).toMatchObject({
    locked: true,
    dirty: true,
    resultState: "cached",
  });
  expect(model.getNode(model.snapshot().metadata.word!.fields[0]!.nodeId)!.text).toBe("3");
});

test("keeps native comment offsets through a tab and a surrogate pair in one run", () => {
  const bytes = mutate("08-comments-revisions.docx", "word/document.xml", (source) =>
    source
      .replace("第三季度</w:t>", "<w:t-placeholder/>")
      .replace("<w:t-placeholder/>", '第三季度</w:t><w:tab/><w:t xml:space="preserve">😀</w:t>'),
  );
  const model = openDocument(importDocx(bytes, { documentId: "offsets" }));
  const anchor = model.snapshot().annotations[0]!.anchor;
  expect(anchor.kind).toBe("text-set");
  if (anchor.kind === "text-set") {
    expect(anchor.ranges.map((range) => range.quote).join("")).toBe("第三季度\t😀完成项目");
    for (const range of anchor.ranges)
      expect(model.getNode(range.nodeId)!.text.slice(range.from, range.to)).toBe(range.quote);
  }
});
