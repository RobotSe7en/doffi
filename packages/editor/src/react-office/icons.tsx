import type { ReactElement } from "react";

const paths = {
  newSlide: "M3 5h19v16H3Z M7 10h11 M7 14h5",
  word: "M7 2h14v20H7Z",
  excel: "M7 2h14v20H7Z",
  powerPoint: "M7 2h14v20H7Z",
  saveAs: "M4 3h13l4 4v14H3V3Z M7 3v6h10V3 M7 21v-8h10v8",
  conditionalFormat: "M2 3h20v18H2Z M2 9h20 M2 15h20 M9 3v18 M16 3v18",
  tableFormat: "M2 3h20v18H2Z M2 9h20 M2 15h20 M9 3v18 M16 3v18",
  cellStyles: "M2 2h20v20H2Z M2 8h20 M8 2v20",
  sortFilter: "M3 3h6l-6 8h6 M3 15h6 M6 13v9 M15 3v18 M12 18l3 3 3-3 M18 3h5l-2 5v5",
  arrange: "M2 2h13v13H2Z M9 9h13v13H9Z",
  quickStyles: "M3 2h17v19H3Z M6 16l4-10 4 10 M8 12h4 m3 4 5-9 3 2-5 9Z",
  insertCells: "M3 3h18v18H3Z M3 9h18 M9 3v18 M3 15h18 M15 3v18",
  deleteCells: "M3 3h18v18H3Z M3 9h18 M9 3v18 M3 15h18 M15 3v18",
  save: "M4 3h13l4 4v14H3V3Z M7 3v6h10V3 M7 21v-8h10v8",
  undo: "M9 5 4 10l5 5 M4 10h10a6 6 0 0 1 6 6v3",
  redo: "m15 5 5 5-5 5 M20 10H10a6 6 0 0 0-6 6v3",
  edit: "m16 3 5 5-12 12-6 1 1-6Z M13 6l5 5",
  paste: "M8 5H5v16h14V5h-3 M8 3h8v5H8Z M8 12h8 M8 16h5",
  cut: "m8 8 12 12 M8 16 20 4 M9 6a3 3 0 1 1-6 0 3 3 0 0 1 6 0 M9 18a3 3 0 1 1-6 0 3 3 0 0 1 6 0",
  copy: "M8 8h13v13H8Z M16 8V3H3v13h5",
  brush: "M5 3h14v8H5Z M12 11v10 M8 5v4 M16 5v4",
  clearFormat: "M3 18 8 3l5 15 M5 12h6 M13 16l5-6 5 4-5 6h-3Z",
  pinyin: "M3 6h5 M5 3v6 M12 6h5 M14 3v6 M4 18h16 M7 12v10 M17 12v10",
  fontBorder: "M3 3h18v18H3Z M7 18 12 6l5 12 M9 14h6",
  textEffects: "M4 21 10 3h4l6 18h-4l-1-4H9l-1 4Z M10 13h4l-2-6Z",
  enclose: "M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0 M7 12h10 M12 7v10",
  eraser: "m14 3 7 7-11 11H7l-5-5Z M8 10l7 7 M10 21h12",
  highlight: "m15 3 6 6-10 10-6-6Z M4 14l6 6-7 1Z",
  left: "M3 5h18 M3 10h11 M3 15h18 M3 20h11",
  center: "M3 5h18 M6 10h12 M3 15h18 M6 20h12",
  right: "M3 5h18 M10 10h11 M3 15h18 M10 20h11",
  justify: "M3 5h18 M3 10h18 M3 15h18 M3 20h18",
  list: "M8 5h13 M8 12h13 M8 19h13 M3 5h.1 M3 12h.1 M3 19h.1",
  numbering: "M9 5h12 M9 12h12 M9 19h12 M3 3v5 M2 10h3v2l-3 3h3 M2 18h3v4H2 M2 20h3",
  indent: "M3 4h18 M9 9h12 M9 14h12 M3 19h18 M2 8l4 4-4 4",
  outdent: "M3 4h18 M9 9h12 M9 14h12 M3 19h18 M6 8l-4 4 4 4",
  distributed: "M3 5h18 M3 10h18 M3 15h18 M3 20h18 M3 8v9 M21 8v9 M9 12l-3 2 3 2 M15 12l3 2-3 2",
  verticalTop: "M3 3h18 M5 8h14 M5 12h14 M5 16h14",
  verticalCenter: "M3 3h18 M5 8h14 M5 12h14 M5 16h14 M3 21h18",
  verticalBottom: "M3 21h18 M5 8h14 M5 12h14 M5 16h14",
  textDirection: "M3 8h12 M6 3l-3 5 3 5 M12 3l3 5-3 5 M7 17h12 M13 13v8 M16 13v8",
  sort: "M3 5h6l-6 8h6 M3 17h6 M15 3v18 M12 18l3 3 3-3 M20 5v8",
  merge: "M3 4h18v16H3Z M3 9h5 M3 15h5 M16 9h5 M16 15h5 M8 12h8 M6 10l2 2-2 2 M18 10l-2 2 2 2",
  mic: "M9 3a3 3 0 0 1 6 0v9a3 3 0 0 1-6 0Z M6 9v3a6 6 0 0 0 12 0V9 M12 18v4 M8 22h8",
  addins: "M3 3h7v7H3Z M14 3h7v7h-7Z M3 14h7v7H3Z M14 14h7v7h-7Z",
  editor: "m15 2 6 6-13 13-6 1 1-6Z M12 5l6 6 M5 15l4 4 M13 21h9",
  styles: "M3 3h18v18H3Z M7 7h8 M7 11h6 M7 15h4 m4 5 4-7 3 2-4 7Z",
  reset: "M3 4h18v16H3Z M16 9a5 5 0 1 0 1 6 M17 6v4h-4",
  section: "M3 4h18 M3 9h18 M3 15h18 M3 20h18 M6 2v9 M6 13v9",
  lineSpacing: "M9 5h12 M9 12h12 M9 19h12 M4 3v18 M1 6l3-3 3 3 M1 18l3 3 3-3",
  fill: "m8 3 10 10-7 7-9-9Z M5 2l7 7 M19 16s4 4 2 6-5-1-2-6",
  table: "M3 3h18v18H3Z M3 9h18 M3 15h18 M9 3v18 M15 3v18",
  cursor: "M4 2v19l5-6 4 7 3-2-4-7h8Z",
  replace: "M3 4h10 M8 2v6 M2 9l5-3 5 3 M17 10v11 M13 17l4 4 4-4",
  image: "M3 3h18v18H3Z M3 17l6-7 4 5 3-3 5 7 M15 7h.1",
  shapes: "M3 3h10v10H3Z M16 12a5 5 0 1 1-4 4 M8 21l6-10 7 10Z",
  diagram: "M9 2h6v5H9Z M2 17h6v5H2Z M16 17h6v5h-6Z M12 7v5 M5 17v-5h14v5",
  chart: "M3 3v18h18 M7 16V9h3v7 M13 16V5h3v11 M19 16v-5h3",
  link: "m10 14 4-4 M8 15l-2 2a3 3 0 0 1-4-4l5-5a3 3 0 0 1 4 0 M13 16a3 3 0 0 0 4 0l5-5a3 3 0 0 0-4-4l-2 2",
  bookmark: "M6 3h12v18l-6-5-6 5Z",
  header: "M4 2h16v20H4Z M4 7h16 M8 4h8",
  footer: "M4 2h16v20H4Z M4 17h16 M8 20h8",
  textBox: "M3 3h18v18H3Z M7 8h10 M12 8v10 M9 18h6",
  pageBreak: "M5 3h14v6 M5 21h14v-6 M2 12h4 M9 12h6 M18 12h4",
  pdf: "M4 2h11l5 5v15H4Z M14 2v6h6 M7 13h10 M7 17h7",
  print: "M6 8V2h12v6 M6 17H3V8h18v9h-3 M6 14h12v8H6Z M17 11h1",
  proof: "m4 13 4 4 9-13 M3 21h18 M16 10l5 7h-9Z",
  book: "M12 5c-3-3-7-3-10-2v17c4-1 7-1 10 2 3-3 6-3 10-2V3c-4-1-7-1-10 2Z M12 5v17",
  globe: "M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0 M3 12h18 M12 3c-6 5-6 13 0 18 6-5 6-13 0-18",
  close: "M5 5l14 14 M19 5 5 19",
  margins: "M4 2h16v20H4Z M7 6h10v12H7Z",
  orientation: "M3 2h11v7 M3 2v20h18V9H10 M10 9l4-4 M10 9l4 4",
  columns: "M3 3h18v18H3Z M10 6v12 M14 6v12 M6 6v12 M18 6v12",
  refresh: "M20 8a8 8 0 1 0 0 8 M20 2v6h-6",
  pen: "m15 3 6 6-12 12-7 1 1-7Z M12 6l6 6 M3 15l6 6",
  palette: "M12 3a9 9 0 1 0 0 18h2c3 0 3-4 0-4-2 0-2-3 1-3h3c5 0 4-11-6-11 M7 8h.1 M12 6h.1 M17 8h.1 M6 13h.1",
  arrow: "M5 12h14 m-5-5 5 5-5 5",
  external: "M14 3h7v7 M21 3 10 14 M10 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-5",
  folder: "M3 7V5a2 2 0 0 1 2-2h5l3 3h6a2 2 0 0 1 2 2v1 M3 7h18l-2 13H3Z",
  shield: "m12 3 8 3v6c0 5-8 9-8 9s-8-4-8-9V6Z m-4 9 3 3 5-5",
  panel: "M3 4h18v16H3Z M9 4v16",
  home: "m3 11 9-8 9 8 M5 10v11h14V10 M9 21v-8h6v8",
  file: "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z M14 2v6h6 M8 13h8 M8 17h5",
  slides: "M3 4h18v13H3z M12 17v4 M8 21h8 M7 13l3-3 3 2 4-5",
  upload: "M12 16V3 M7 8l5-5 5 5 M4 15v5a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-5",
  download: "M12 3v13 M7 11l5 5 5-5 M4 16v4a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-4",
  plus: "M12 5v14 M5 12h14",
  search: "M21 21l-5-5 M18 10a8 8 0 1 1-16 0 8 8 0 0 1 16 0",
  chevron: "m9 5 7 7-7 7",
  comment: "M21 11.5a8.5 8.5 0 0 1-8.5 8.5H3l2-5A8.5 8.5 0 1 1 21 11.5Z M8 9h8 M8 13h5",
  sparkles: "m12 3 2.5 6.5L21 12l-6.5 2.5L12 21l-2.5-6.5L3 12l6.5-2.5Z M20 2v4 M18 4h4",
  grid: "M3 3h7v7H3z M14 3h7v7h-7z M3 14h7v7H3z M14 14h7v7h-7z",
  check: "m5 12 4 4L19 6",
} as const;

/** First-party vector artwork. Colored details share the native ribbon's visual roles. */
export function OfficeIcon({ name, size = 18 }: { name: keyof typeof paths; size?: number }): ReactElement {
  let drawing: ReactElement;
  switch (name) {
    case "newSlide": drawing = <><path d="M4 4h18v17H4Z M7 8h12 M7 12h5 M15 12h4 M7 16h5 M15 16h4" /><path fill="#fff" stroke="none" d="M0 0h11v11H0Z" /><path stroke="#d49a20" d="M5 0v10 M0 5h10 M1.5 1.5l7 7 M1.5 8.5l7-7" /></>; break;
    case "word": case "excel": case "powerPoint": {
      const color = name === "word" ? "#185abd" : name === "excel" ? "#107c41" : "#b7472a";
      drawing = <><path fill="#fff" stroke={color} d="M7 2h14v20H7Z" /><path stroke={color} opacity=".5" d="M10 6h8 M10 10h8 M10 14h8 M10 18h8" /><rect x="1" y="6" width="13" height="14" rx="1" fill={color} stroke="none" /><text x="7.5" y="16.5" fill="#fff" stroke="none" textAnchor="middle" fontFamily="Arial,sans-serif" fontWeight="700" fontSize="10">{name === "word" ? "W" : name === "excel" ? "X" : "P"}</text></>; break;
    }
    case "saveAs": drawing = <><path d="M3 2h13l3 3v8 M3 2v19h9 M6 2v6h9V2 M6 21v-8h7" /><path fill="#fff" d="m19 11 4 3-6 8-5 1 1-5Z" /></>; break;
    case "conditionalFormat": drawing = <><path fill="#f0f0f0" d="M2 3h20v18H2Z" /><path fill="#ffb7b7" stroke="none" d="M2 3h7v6H2Z" /><path fill="#f9dc73" stroke="none" d="M9 9h7v6H9Z" /><path fill="#97ca9d" stroke="none" d="M16 15h6v6h-6Z" /><path d={paths.conditionalFormat} /></>; break;
    case "tableFormat": drawing = <><path fill="#edf3fa" d="M2 3h20v18H2Z" /><path fill="#658eb5" stroke="none" d="M2 3h20v6H2Z" /><path fill="#b4cbe0" stroke="none" d="M2 15h20v6H2Z" /><path stroke="#45698c" d={paths.tableFormat} /></>; break;
    case "cellStyles": drawing = <><path fill="#eaf2e1" d="M2 2h20v20H2Z" /><path fill="#ffe0ac" stroke="none" d="M2 2h20v6H2Z" /><path stroke="#929292" d={paths.cellStyles} /><text x="10" y="18" fill="#485749" stroke="none" fontFamily="serif" fontSize="12">A</text></>; break;
    case "sortFilter": drawing = <><path stroke="#2786c4" d="M2 10 5 2h1l3 8 M3 7h5 M2 14h7l-7 8h7 M12 3v18 M10 19l2 2 2-2" /><path d="M16 3h7l-2.5 5v7h-2V8Z" /></>; break;
    case "arrange": drawing = <><path fill="#d9e7f5" stroke="#5680b0" d="M2 2h13v13H2Z" /><path fill="#fff" stroke="#777" d="M9 9h13v13H9Z" /><path fill="#d9e7f5" stroke="#5680b0" d="M5 12h12v10H5Z" /></>; break;
    case "quickStyles": drawing = <><path fill="#fff" d="M3 2h17v19H3Z M6 16l4-10 4 10 M8 12h4" /><path fill="#79b9e8" stroke="#4a85ad" d="m16 15 4-8 3 2-4 8-4 3Z" /></>; break;
    case "insertCells": case "deleteCells": drawing = <><path d="M3 7h18v14H3Z M3 12h18 M3 17h18 M9 7v14 M15 7v14" /><path fill="#fff" stroke="none" d="M10 0h14v11H10Z" /><path stroke={name === "insertCells" ? "#107c41" : "#c4312f"} d={name === "insertCells" ? "M17 1v9 M13 5.5h8" : "m13 2 8 8 M21 2l-8 8"} /></>; break;
    case "header": case "footer": drawing = <><path d="M4 2h16v20H4Z" /><path fill="#8eb6db" stroke="none" d={name === "header" ? "M5 3h14v4H5Z" : "M5 17h14v4H5Z"} /><path stroke="#999" d="M8 10h8 M8 13h8" /></>; break;
    case "paste": drawing = <><path stroke="#c87d32" d="M8 3H5v18h12V3h-2 M8 2h7v5H8Z" /><path fill="#fff" stroke="#787878" d="M10 9h10v13H10Z" /><path stroke="#969696" d="M12 12h6 M12 15h6 M12 18h4" /></>; break;
    case "brush": drawing = <><path fill="#fff" stroke="#777" d="M4 2h16v8H4Z M7 2v6 M17 2v6" /><path stroke="#c87728" d="M4 10h16v3H4Z M10 13h4v9h-4Z" /></>; break;
    case "mic": drawing = <><rect x="8" y="1" width="8" height="15" rx="4" fill="#62b4ec" stroke="#0078cf" /><path stroke="#737373" d="M5 10v3a7 7 0 0 0 14 0v-3 M12 20v3 M8 23h8" /></>; break;
    case "editor": drawing = <><path fill="#185abd" stroke="#185abd" d="m16 1 5 4-12 15-6 1 1-6Z" /><path stroke="#fff" d="m14 4 5 4 M5 15l4 4" /><path stroke="#185abd" d="M12 22h10" /></>; break;
    case "styles": drawing = <><path d="M3 2h16v18H3Z M6 5h10 M6 8h8 M6 11h7 M6 14h5" /><path fill="#fff" d="m17 13 5 3-5 8-4-3Z" /><path d="m14 20 4 2" /></>; break;
    case "list": drawing = <><path d="M8 4h13 M8 11h13 M8 18h13" /><path stroke="#2475bb" strokeWidth="2.6" d="M3 4h.01 M3 11h.01 M3 18h.01" /></>; break;
    case "numbering": drawing = <><path d="M9 4h12 M9 11h12 M9 18h12" /><path stroke="#2475bb" d="M3 1v5 M2 8h3v2l-3 3h3 M2 16h3v5H2 M2 18.5h3" /></>; break;
    case "indent": case "outdent": drawing = <><path d="M3 3h18 M10 8h11 M10 13h11 M3 18h18" /><path stroke="#2475bb" d={name === "indent" ? "M2 10h5 M4 7l3 3-3 3" : "M2 10h5 M5 7l-3 3 3 3"} /></>; break;
    case "lineSpacing": drawing = <><path d="M10 4h11 M10 11h11 M10 18h11" /><path stroke="#2475bb" d="M4 1v21 M1 4l3-3 3 3 M1 19l3 3 3-3" /></>; break;
    case "distributed": drawing = <><path d="M3 3h18 M3 8h18 M3 19h18" /><path stroke="#2475bb" d="M3 11v5 M21 11v5 M10 13.5H3 m3-2-3 2 3 2 M14 13.5h7 m-3-2 3 2-3 2" /></>; break;
    case "addins": drawing = <path stroke="#f65b24" d="M2 2h8v8H2Z M14 2h8v8h-8Z M2 14h8v8H2Z M14 14h8v8h-8Z" />; break;
    case "chart": drawing = <><path d="M2 1v21h21" /><path fill="#5b9bd5" stroke="none" d="M5 11h4v9H5Z" /><path fill="#ed7d31" stroke="none" d="M11 4h4v16h-4Z" /><path fill="#a5a5a5" stroke="none" d="M17 8h4v12h-4Z" /></>; break;
    case "image": drawing = <><path d="M2 3h20v18H2Z" /><path stroke="#477d56" d="m3 18 6-8 5 6 3-4 4 7" /><circle cx="16" cy="7" r="1.5" fill="#deb64b" stroke="none" /></>; break;
    case "shapes": drawing = <><path fill="#d9e7f5" stroke="#5485b3" d="M2 2h11v11H2Z" /><circle cx="17" cy="12" r="5" fill="#fff" stroke="#5485b3" /><path fill="#d9e7f5" stroke="#5485b3" d="m9 11 7 11H2Z" /></>; break;
    case "clearFormat": drawing = <><text x="1" y="17" fill="#333" stroke="none" fontFamily="Times New Roman,serif" fontSize="20">A</text><path fill="#e4b4d8" stroke="#8e548e" d="m17 8 6 5-6 7-6-5Z" /><path stroke="#777" d="m11 15 6 5 M17 21h6" /></>; break;
    case "pinyin": drawing = <><text x="1" y="8" fill="#777" stroke="none" fontFamily="Arial,sans-serif" fontSize="8">abc</text><path stroke="#185abd" d="m17 4 2 2 3-4" /><text x="4" y="21" fill="#333" stroke="none" fontFamily="serif" fontSize="15">文</text></>; break;
    case "textEffects": drawing = <path d={paths.textEffects} stroke="#248ee3" fill="#c9e7ff" />; break;
    case "highlight": drawing = <><path d="m14 2 7 6-9 10-6-6Z M6 12l6 6-4 1-3-3Z" /><path fill="#ffd800" stroke="none" d="m2 20 8-3 3 4-11 2Z" /></>; break;
    default: drawing = <path d={paths[name]} />;
  }
  return <svg data-office-icon={name} width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.25" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{drawing}</svg>;
}
