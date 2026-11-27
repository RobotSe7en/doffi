import { body, mutateDocx, styles } from "./regression-fixtures.js";

export const tableDefaultCases = [
  { kind: "none", padding: 0 },
  { kind: "custom", padding: 0 },
  { kind: "custom-default", padding: 0 },
  { kind: "stock", padding: 5.4 },
  { kind: "stock-unused", padding: 0 },
  { kind: "stock-lowercase", padding: 5.4 },
  { kind: "stock-localized", padding: 5.4 },
  { kind: "stock-custom", padding: 0 },
  { kind: "id-only", padding: 0 },
  { kind: "derived", padding: 5.4 },
] as const;
export type TableDefaultKind = typeof tableDefaultCases[number]["kind"];

/** Empty native declarations isolate Word's stock defaults from ordinary styles. */
export function tableDefaultFixture(kind: TableDefaultKind, options: {
  mode?: number; override?: boolean;
} = {}): Uint8Array {
  const stock = kind.startsWith("stock") || kind === "derived",
    id = kind === "id-only" ? "TableNormal" : "ArbitraryIdentity",
    name = kind === "stock-localized" ? "普通表格" : kind === "stock-lowercase" ? "normal table" : stock ? "Normal Table" : "Independent Custom",
    definition = kind === "none" ? "" : '<w:style w:type="table" w:styleId="' + id + '"'
      + (kind === "stock-custom" ? ' w:customStyle="1"' : "")
      + (kind !== "custom" && kind !== "stock-unused" && kind !== "derived" ? ' w:default="1"' : "")
      + '><w:name w:val="' + name + '"/></w:style>'
      + (kind === "derived" ? '<w:style w:type="table" w:styleId="Leaf"><w:name w:val="Leaf"/><w:basedOn w:val="' + id + '"/></w:style>' : ""),
    binding = kind === "custom" || kind === "derived" ? '<w:tblStyle w:val="' + (kind === "derived" ? "Leaf" : id) + '"/>' : "";
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () => styles('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Noto Sans CJK SC" w:hAnsi="Noto Sans CJK SC" w:eastAsia="Noto Sans CJK SC"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' + definition),
    "word/settings.xml": xml => xml.replace("</w:settings>", options.mode === undefined ? "</w:settings>"
      : '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="' + options.mode + '"/></w:compat></w:settings>'),
    "word/document.xml": xml => body(xml, '<w:p><w:r><w:t>REFERENCE</w:t></w:r></w:p><w:tbl><w:tblPr>' + binding
      + (options.override ? '<w:tblCellMar><w:left w:w="140" w:type="dxa"/></w:tblCellMar>' : "")
      + '<w:tblW w:w="4800" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr>'
      + (options.override ? '<w:tcMar><w:right w:w="180" w:type="dxa"/></w:tcMar>' : "")
      + '</w:tcPr><w:p><w:r><w:t>LEADING</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>TRAILING</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'),
  });
}

export function fontLineFixture(font = "Noto Sans CJK SC"): Uint8Array {
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () => styles('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="' + font + '" w:hAnsi="' + font + '" w:eastAsia="' + font + '"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'),
    "word/document.xml": xml => body(xml, '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>'
      + [8, 12, 18].flatMap(size => ["Abcd", "汉字"].map(text => '<w:tr><w:tc><w:p><w:pPr><w:rPr><w:sz w:val="' + size * 2 + '"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="' + size * 2 + '"/></w:rPr><w:t>' + text + '</w:t><w:br/><w:t>' + text + '</w:t></w:r></w:p></w:tc></w:tr>')).join("") + '</w:tbl>'),
  });
}

export const rowHeightCases = [
  { rule: "auto", height: 100, top: 0, bottom: 0 },
  { rule: "atLeast", height: 60, top: 0, bottom: 0 },
  { rule: "atLeast", height: 60, top: 7, bottom: 11 },
  { rule: "exact", height: 40, top: 0, bottom: 0 },
  { rule: "exact", height: 40, top: 4, bottom: 0 },
  { rule: "exact", height: 40, top: 0, bottom: 4 },
  { rule: "exact", height: 40, top: 7, bottom: 11 },
  { rule: "atLeast", height: 60, top: 3, bottom: 11, otherTop: 7, otherBottom: 2 },
  { rule: "exact", height: 40, top: 3, bottom: 11, otherTop: 7, otherBottom: 2 },
  { rule: "auto", height: 100, top: 3, bottom: 11, otherTop: 7, otherBottom: 2 },
] as const;
export function rowHeightFixture(mode: 14 | 15 = 14): Uint8Array {
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () => styles('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Noto Serif CJK SC" w:hAnsi="Noto Serif CJK SC" w:eastAsia="Noto Serif CJK SC"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'),
    "word/settings.xml": xml => xml.replace('</w:settings>', '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="' + mode + '"/></w:compat></w:settings>'),
    "word/document.xml": xml => body(xml, '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>'
      + ['top','bottom','left','right','insideH','insideV'].map(side => '<w:' + side + ' w:val="single" w:sz="8" w:color="334455"/>').join('')
      + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
      + rowHeightCases.map(spec => { const { rule, height, top, bottom } = spec; return '<w:tr><w:trPr><w:trHeight w:val="' + height * 20 + '" w:hRule="' + rule + '"/></w:trPr>'
        + [0,1].map(column => '<w:tc><w:tcPr><w:tcMar><w:top w:w="' + ((column && "otherTop" in spec) ? spec.otherTop : top) * 20 + '" w:type="dxa"/><w:bottom w:w="' + ((column && "otherBottom" in spec) ? spec.otherBottom : bottom) * 20 + '" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:t>行高Row</w:t><w:br/><w:t>行高Row</w:t></w:r></w:p></w:tc>').join('') + '</w:tr>'; }).join('') + '</w:tbl>'),
  });
}
