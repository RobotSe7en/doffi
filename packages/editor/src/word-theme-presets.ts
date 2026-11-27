import type { WordTheme } from "@doffi/core";

export interface WordColorSchemePreset {
  readonly id: string;
  readonly label: string;
  readonly colors: NonNullable<WordTheme["colors"]>;
}

export interface WordFontSchemePreset {
  readonly id: string;
  readonly label: string;
  readonly majorLatin: string;
  readonly minorLatin: string;
  readonly majorEastAsia: string;
  readonly minorEastAsia: string;
}

export interface WordThemePreset {
  readonly id: string;
  readonly label: string;
  readonly theme: WordTheme;
}

export interface WordStyleSetPreset {
  readonly id: string;
  readonly label: string;
  readonly accent: string;
  readonly headingFont: string;
  readonly bodyFont: string;
  readonly headingColor: string;
  readonly lineHeight: string;
  readonly spaceAfter: string;
  readonly headingItalic?: boolean;
  readonly headingUppercase?: boolean;
}

export const WORD_COLOR_SCHEMES: readonly WordColorSchemePreset[] = [
  {
    id: "Office",
    label: "Office",
    colors: {
      dk1: "000000", lt1: "FFFFFF", dk2: "44546A", lt2: "E7E6E6",
      accent1: "4472C4", accent2: "ED7D31", accent3: "A5A5A5", accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47",
      hlink: "0563C1", folHlink: "954F72",
    },
  },
  {
    id: "BlueGreen",
    label: "蓝绿",
    colors: {
      dk1: "000000", lt1: "FFFFFF", dk2: "263238", lt2: "ECEFF1",
      accent1: "1F7A8C", accent2: "2E8B57", accent3: "4EA8DE", accent4: "56CFE1", accent5: "48BFE3", accent6: "64DFDF",
      hlink: "0B6E99", folHlink: "5E548E",
    },
  },
  {
    id: "Warm",
    label: "暖橙",
    colors: {
      dk1: "1C1917", lt1: "FFFFFF", dk2: "57534E", lt2: "F5F5F4",
      accent1: "D95D39", accent2: "F0A202", accent3: "C44536", accent4: "E09F3E", accent5: "9E2A2B", accent6: "774936",
      hlink: "B45309", folHlink: "7C2D12",
    },
  },
  {
    id: "Green",
    label: "翠绿",
    colors: {
      dk1: "14201A", lt1: "FFFFFF", dk2: "2D4A3E", lt2: "EEF4F0",
      accent1: "1E7E34", accent2: "38A169", accent3: "68D391", accent4: "2F855A", accent5: "276749", accent6: "48BB78",
      hlink: "15803D", folHlink: "166534",
    },
  },
  {
    id: "Grayscale",
    label: "灰阶",
    colors: {
      dk1: "000000", lt1: "FFFFFF", dk2: "3F3F3F", lt2: "EAEAEA",
      accent1: "595959", accent2: "7F7F7F", accent3: "A6A6A6", accent4: "404040", accent5: "8C8C8C", accent6: "BFBFBF",
      hlink: "333333", folHlink: "666666",
    },
  },
];

export const WORD_FONT_SCHEMES: readonly WordFontSchemePreset[] = [
  { id: "Office", label: "Office（等线 / Calibri）", majorLatin: "Calibri Light", minorLatin: "Calibri", majorEastAsia: "等线 Light", minorEastAsia: "等线" },
  { id: "YaHei", label: "现代无衬线（微软雅黑 / Segoe UI）", majorLatin: "Segoe UI", minorLatin: "Segoe UI", majorEastAsia: "微软雅黑", minorEastAsia: "微软雅黑" },
  { id: "Classic", label: "经典公文（宋体 / Times New Roman）", majorLatin: "Times New Roman", minorLatin: "Times New Roman", majorEastAsia: "黑体", minorEastAsia: "宋体" },
  { id: "Editorial", label: "书刊典雅（楷体 / Georgia）", majorLatin: "Georgia", minorLatin: "Georgia", majorEastAsia: "楷体", minorEastAsia: "宋体" },
  { id: "Noto", label: "开源思源（Noto Sans / Serif CJK）", majorLatin: "Liberation Sans", minorLatin: "Liberation Serif", majorEastAsia: "Noto Sans CJK SC", minorEastAsia: "Noto Serif CJK SC" },
];

export const WORD_THEMES: readonly WordThemePreset[] = [
  {
    id: "Office",
    label: "Office 主题",
    theme: {
      name: "Office 主题",
      colorSchemeName: "Office",
      colors: WORD_COLOR_SCHEMES[0]!.colors,
      fontSchemeName: "Office",
      majorLatin: "Calibri Light",
      minorLatin: "Calibri",
      majorEastAsia: "等线 Light",
      minorEastAsia: "等线",
    },
  },
  {
    id: "Facet",
    label: "流畅",
    theme: {
      name: "流畅",
      colorSchemeName: "BlueGreen",
      colors: WORD_COLOR_SCHEMES[1]!.colors,
      fontSchemeName: "YaHei",
      majorLatin: "Segoe UI",
      minorLatin: "Segoe UI",
      majorEastAsia: "微软雅黑",
      minorEastAsia: "微软雅黑",
    },
  },
  {
    id: "Integral",
    label: "东方典雅",
    theme: {
      name: "东方典雅",
      colorSchemeName: "Warm",
      colors: WORD_COLOR_SCHEMES[2]!.colors,
      fontSchemeName: "Classic",
      majorLatin: "Times New Roman",
      minorLatin: "Times New Roman",
      majorEastAsia: "黑体",
      minorEastAsia: "宋体",
    },
  },
  {
    id: "Organic",
    label: "有机翠绿",
    theme: {
      name: "有机翠绿",
      colorSchemeName: "Green",
      colors: WORD_COLOR_SCHEMES[3]!.colors,
      fontSchemeName: "Editorial",
      majorLatin: "Georgia",
      minorLatin: "Georgia",
      majorEastAsia: "楷体",
      minorEastAsia: "宋体",
    },
  },
  {
    id: "Ion",
    label: "极简灰阶",
    theme: {
      name: "极简灰阶",
      colorSchemeName: "Grayscale",
      colors: WORD_COLOR_SCHEMES[4]!.colors,
      fontSchemeName: "Noto",
      majorLatin: "Liberation Sans",
      minorLatin: "Liberation Serif",
      majorEastAsia: "Noto Sans CJK SC",
      minorEastAsia: "Noto Serif CJK SC",
    },
  },
];

export const WORD_STYLE_SETS: readonly WordStyleSetPreset[] = [
  { id: "文档", label: "文档", accent: "#4472c4", headingFont: "等线 Light", bodyFont: "等线", headingColor: "#2f5496", lineHeight: "1.15", spaceAfter: "8pt" },
  { id: "独特", label: "独特", accent: "#b7472a", headingFont: "微软雅黑", bodyFont: "宋体", headingColor: "#9c2b1e", lineHeight: "1.2", spaceAfter: "6pt" },
  { id: "典雅型", label: "典雅型", accent: "#595959", headingFont: "宋体", bodyFont: "宋体", headingColor: "#262626", lineHeight: "1.25", spaceAfter: "8pt", headingUppercase: true },
  { id: "流行", label: "流行", accent: "#1f7a8c", headingFont: "微软雅黑", bodyFont: "等线", headingColor: "#1f7a8c", lineHeight: "1.15", spaceAfter: "6pt" },
  { id: "正式", label: "正式", accent: "#2b579a", headingFont: "黑体", bodyFont: "仿宋", headingColor: "#1f3864", lineHeight: "1.5", spaceAfter: "10pt" },
  { id: "手稿", label: "手稿", accent: "#6b4f3a", headingFont: "楷体", bodyFont: "宋体", headingColor: "#4a3728", lineHeight: "1.5", spaceAfter: "12pt", headingItalic: true },
  { id: "现代", label: "现代", accent: "#107c41", headingFont: "微软雅黑", bodyFont: "微软雅黑", headingColor: "#107c41", lineHeight: "1.2", spaceAfter: "8pt" },
  { id: "新闻纸", label: "新闻纸", accent: "#444444", headingFont: "黑体", bodyFont: "宋体", headingColor: "#111111", lineHeight: "1.05", spaceAfter: "4pt" },
  { id: "透视", label: "透视", accent: "#44546a", headingFont: "等线", bodyFont: "等线", headingColor: "#44546a", lineHeight: "1.3", spaceAfter: "8pt" },
  { id: "简单", label: "简单", accent: "#000000", headingFont: "宋体", bodyFont: "宋体", headingColor: "#000000", lineHeight: "1.15", spaceAfter: "6pt" },
];
