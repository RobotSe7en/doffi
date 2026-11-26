import { integer, bool } from "./drawingml/values.js";
import { A, C, PIC, R, W, WP, fail, unsupported } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
type Rule = { attrs: readonly string[]; children: readonly string[] };
const rules = new Map<string, Rule>();
const names = (uri: string, locals: string): string[] =>
  locals
    .split(" ")
    .filter(Boolean)
    .map((name) => uri + " " + name);
function rule(uri: string, locals: string, attrs = "", children: string[] = []): void {
  for (const name of names(uri, locals)) rules.set(name, { attrs: names("", attrs), children });
}
rule(W, "drawing", "", names(WP, "inline anchor"));
rule(WP, "inline", "distT distB distL distR", [
  ...names(WP, "extent effectExtent docPr cNvGraphicFramePr"),
  ...names(A, "graphic"),
]);
rule(
  WP,
  "anchor",
  "distT distB distL distR simplePos relativeHeight behindDoc locked layoutInCell allowOverlap hidden",
  [
    ...names(
      WP,
      "simplePos positionH positionV extent effectExtent wrapNone wrapSquare wrapTight wrapThrough wrapTopAndBottom docPr cNvGraphicFramePr",
    ),
    ...names(A, "graphic"),
    ...names(WP14, "sizeRelH sizeRelV"),
  ],
);
rule(WP, "simplePos start lineTo", "x y");
rule(WP, "positionH positionV", "relativeFrom", names(WP, "align posOffset"));
rule(WP, "align posOffset");
rule(WP, "extent", "cx cy");
rule(WP, "effectExtent", "l t r b");
rule(WP, "docPr", "id name descr hidden");
rule(WP, "cNvGraphicFramePr", "", names(A, "graphicFrameLocks"));
// Word emits this empty default container on resave; nonempty restrictions
// still retain the native drawing through the attribute checks below.
rule(A, "graphicFrameLocks");
rule(WP, "wrapNone");
rule(WP, "wrapTopAndBottom", "distT distB");
rule(WP, "wrapSquare", "wrapText distT distB distL distR");
rule(WP, "wrapTight wrapThrough", "wrapText distL distR", names(WP, "wrapPolygon"));
rule(WP, "wrapPolygon", "edited", names(WP, "start lineTo"));
rule(WP14, "sizeRelH", "relativeFrom", names(WP14, "pctWidth"));
rule(WP14, "sizeRelV", "relativeFrom", names(WP14, "pctHeight"));
rule(WP14, "pctWidth pctHeight");
rule(A, "graphic", "", names(A, "graphicData"));
rule(A, "graphicData", "uri", [
  ...names(PIC, "pic"),
  ...names(WPS, "wsp"),
  ...names(WPG, "wgp grpSp"),
  ...names(C, "chart"),
]);
rule(A, "xfrm", "rot flipH flipV", names(A, "off ext chOff chExt"));
rule(A, "off chOff", "x y");
rule(A, "ext chExt", "cx cy");
rule(A, "prstGeom", "prst", names(A, "avLst"));
rule(A, "avLst", "", names(A, "gd"));
rule(A, "gd", "name fmla");
rule(A, "noFill round bevel");
rule(A, "solidFill gs fillRef lnRef effectRef", "", names(A, "srgbClr schemeClr sysClr"));
rule(A, "gs", "pos", names(A, "srgbClr schemeClr sysClr"));
rule(A, "fillRef lnRef effectRef", "idx", names(A, "srgbClr schemeClr sysClr"));
rule(A, "fontRef", "idx", names(A, "schemeClr"));
rule(A, "srgbClr schemeClr", "val", names(A, "alpha lumMod lumOff tint shade"));
rule(A, "sysClr", "val lastClr", names(A, "alpha lumMod lumOff tint shade"));
rule(A, "alpha lumMod lumOff tint shade prstDash", "val");
rule(A, "gradFill", "rotWithShape", names(A, "gsLst lin tileRect"));
rule(A, "gsLst", "", names(A, "gs"));
rule(A, "lin", "ang scaled");
rule(A, "tileRect srcRect fillRect", "l t r b");
rule(
  A,
  "ln",
  "w cap cmpd algn",
  names(A, "noFill solidFill prstDash round bevel miter headEnd tailEnd"),
);
rule(A, "miter", "lim");
rule(A, "headEnd tailEnd", "type w len");
rule(A, "effectStyle", "", names(A, "effectLst"));
rule(A, "effectLst", "", names(A, "outerShdw"));
rule(
  A,
  "outerShdw",
  "blurRad dist dir sx sy kx ky algn rotWithShape",
  names(A, "srgbClr schemeClr sysClr"),
);
rule(A, "stretch", "", names(A, "fillRect"));
rule(A, "blip");
rules.set(A + " blip", { attrs: [R + " embed"], children: [] });
rule(A, "stCxn endCxn", "id idx");
rule(A, "noAutofit spAutoFit");
rule(A, "normAutofit", "fontScale lnSpcReduction");
rule(PIC, "pic", "", names(PIC, "nvPicPr blipFill spPr"));
rule(PIC, "nvPicPr", "", names(PIC, "cNvPr cNvPicPr"));
rule(PIC, "cNvPr", "id name descr");
rule(PIC, "cNvPicPr");
rule(PIC, "blipFill", "", names(A, "blip srcRect stretch"));
rule(PIC, "spPr", "", names(A, "xfrm prstGeom noFill ln effectLst"));
rule(WPS, "wsp", "", names(WPS, "cNvPr cNvSpPr cNvCnPr spPr style txbx bodyPr"));
rule(WPS, "cNvPr", "id name descr");
rule(WPS, "cNvSpPr", "txBox", names(A, "spLocks"));
rule(WPS, "cNvCnPr", "", names(A, "cxnSpLocks stCxn endCxn"));
rule(A, "spLocks cxnSpLocks", "noChangeArrowheads noChangeShapeType");
rule(WPS, "spPr", "bwMode", names(A, "xfrm prstGeom solidFill noFill gradFill ln effectLst"));
rule(WPS, "style", "", names(A, "fillRef lnRef effectRef fontRef"));
rule(WPS, "txbx", "", names(W, "txbxContent"));
rule(
  WPS,
  "bodyPr",
  "lIns tIns rIns bIns anchor vert wrap numCol rtlCol fromWordArt rot spcFirstLastPara vertOverflow horzOverflow spcCol anchorCtr forceAA compatLnSpc upright",
  names(A, "noAutofit normAutofit spAutoFit"),
);
rule(WPG, "wgp grpSp", "", [
  ...names(WPG, "cNvPr cNvGrpSpPr grpSpPr grpSp graphicFrame"),
  ...names(WPS, "wsp"),
  ...names(PIC, "pic"),
]);
rule(WPG, "cNvPr", "id name descr");
rule(WPG, "cNvGrpSpPr cNvFrPr");
rule(WPG, "grpSpPr", "", names(A, "xfrm"));
rule(WPG, "graphicFrame", "", [...names(WPG, "cNvPr cNvFrPr xfrm"), ...names(A, "graphic")]);
rule(WPG, "xfrm", "rot flipH flipV", names(A, "off ext"));
rules.set(C + " chart", { attrs: [R + " id"], children: [] });

/** Every accepted DrawingML slot must be consumed by the editable mapping.
 * Unknown attributes/children retain the entire native drawing, including its
 * relationships and preview. Text stories use the ordinary Word reader. */
export function checkWordDrawingCoverage(source: XmlNode): void {
  const scan = (node: XmlNode): void => {
    if (node.uri === W && node.name === "txbxContent") return;
    const known = rules.get(node.uri + " " + node.name);
    if (!known)
      unsupported(node, "This nested drawing feature remains preserved with its native drawing.");
    if (node.uri === WPS && node.name === "spPr" && [...node.attrs].some(([key, value]) => key === " bwMode" && value !== "auto"))
      unsupported(node, "This black and white drawing mode remains native.");
    for (const [key, value] of node.attrs) {
      if (
        node.uri === WP &&
        (node.name === "anchor" || node.name === "inline") &&
        (key === WP14 + " anchorId" || key === WP14 + " editId")
      ) {
        if (!/^[0-9a-f]{8}$/i.test(value)) fail(node, "A drawing edit identity is invalid.");
        continue;
      }
      if (
        !key.startsWith("http://www.w3.org/2000/xmlns/ ") &&
        key !== "http://schemas.openxmlformats.org/markup-compatibility/2006 Ignorable" &&
        !known.attrs.includes(key)
      )
        unsupported(node, "This drawing attribute remains preserved with its native drawing.", key);
    }
    if (
      node.uri === WP &&
      node.name === "effectExtent" &&
      ["l", "t", "r", "b"].some((n) => integer(node, n, 0) < 0)
    )
      unsupported(node, "Negative drawing effect extents remain preserved with their native drawing.");
    if (node.uri === WP && node.name === "anchor" && bool(node, "simplePos", false))
      unsupported(node, "Absolute simple-position frames remain preserved as native drawings.");
    if (
      node.uri === WP &&
      node.name === "simplePos" &&
      ["x", "y"].some((n) => integer(node, n, 0) !== 0)
    )
      unsupported(node, "Native simple-position coordinates remain preserved with their drawing.");
    if (
      node.uri === A &&
      node.name === "fillRect" &&
      ["l", "t", "r", "b"].some((n) => integer(node, n, 0) !== 0)
    )
      unsupported(node, "Offset picture fill rectangles remain preserved with their drawing.");
    for (const child of node.content)
      if (typeof child !== "string") {
        if (!known.children.includes(child.uri + " " + child.name))
          unsupported(child, "This drawing child remains preserved with its native drawing.");
        scan(child);
      } else if (
        child.trim() &&
        !(node.uri === WP && ["align", "posOffset"].includes(node.name)) &&
        !(node.uri === WP14 && ["pctWidth", "pctHeight"].includes(node.name))
      )
        fail(node, "A drawing container has unexpected text.");
  };
  scan(source);
}
