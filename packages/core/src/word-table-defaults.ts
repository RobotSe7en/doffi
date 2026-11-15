import type { WordMetadata } from "./contracts/word.js";
/** Fresh definitions keep new-document style state isolated between sessions. */
export function defaultWordTableStyles(): WordMetadata["styles"] {
  return [
    {
      styleId: "TableNormal",
      name: "普通表格",
      kind: "table",
      default: true,
      cssClass: "word-table-normal",
      paragraphProperties: [],
      runProperties: [],
      components: {
        p: "word-table-normal-p",
        r: "word-table-normal-r",
        table: "word-table-normal-table",
        cell: "word-table-normal-cell",
        row: "word-table-normal-row",
      },
    },
    {
      styleId: "TableGrid",
      name: "网格表格",
      kind: "table",
      basedOn: "TableNormal",
      cssClass: "word-table-grid",
      paragraphProperties: [],
      runProperties: [],
      components: {
        p: "word-table-grid-p",
        r: "word-table-grid-r",
        table: "word-table-grid-table",
        cell: "word-table-grid-cell",
        row: "word-table-grid-row",
      },
    },
  ];
}
const grid =
  "border-collapse:collapse;border:0.5pt solid #000000;--word-border-inside-h:0.5pt solid #000000;--word-border-inside-v:0.5pt solid #000000";
export const defaultWordTableCss =
  ".word-table-normal{border-collapse:collapse}.word-table-normal-p{}.word-table-normal-r{}.word-table-normal-table{border-collapse:collapse}.word-table-normal-cell{padding-left:5.4pt;padding-right:5.4pt;vertical-align:top}.word-table-normal-row{}.word-table-grid{" +
  grid +
  "}.word-table-grid-p{}.word-table-grid-r{}.word-table-grid-table{" +
  grid +
  "}.word-table-grid-cell{padding:4pt;vertical-align:top}.word-table-grid-row{}";
