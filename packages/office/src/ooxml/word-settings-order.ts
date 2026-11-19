import { W } from "./xml.js";
import type { XmlNode } from "./xml.js";
/** CT_Settings sequence from the repository's pinned official Transitional XSD.
 * Unknown extension nodes retain their relative slots during native recovery. */
export const wordSettingsOrder: readonly string[] =
  "writeProtection view zoom removePersonalInformation removeDateAndTime doNotDisplayPageBoundaries displayBackgroundShape printPostScriptOverText printFractionalCharacterWidth printFormsData embedTrueTypeFonts embedSystemFonts saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop hideSpellingErrors hideGrammaticalErrors activeWritingStyle proofState formsDesign attachedTemplate linkStyles stylePaneFormatFilter stylePaneSortMethod documentType mailMerge revisionView trackRevisions doNotTrackMoves doNotTrackFormatting documentProtection autoFormatOverride styleLockTheme styleLockQFSet defaultTabStop autoHyphenation consecutiveHyphenLimit hyphenationZone doNotHyphenateCaps showEnvelope summaryLength clickAndTypeStyle defaultTableStyle evenAndOddHeaders bookFoldRevPrinting bookFoldPrinting bookFoldPrintingSheets drawingGridHorizontalSpacing drawingGridVerticalSpacing displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery doNotUseMarginsForDrawingGridOrigin drawingGridHorizontalOrigin drawingGridVerticalOrigin doNotShadeFormData noPunctuationKerning characterSpacingControl printTwoOnOne strictFirstAndLastChars noLineBreaksAfter noLineBreaksBefore savePreviewPicture doNotValidateAgainstSchema saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText doNotDemarcateInvalidXml saveXmlDataOnly useXSLTWhenSaving saveThroughXslt showXMLTags alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr endnotePr compat docVars rsids m:mathPr attachedSchema themeFontLang clrSchemeMapping doNotIncludeSubdocsInStats doNotAutoCompressPictures forceUpgrade captions readModeInkLockDown smartTagType sl:schemaLibrary shapeDefaults doNotEmbedSmartTags decimalSymbol listSeparator".split(
    " ",
  );
const aliases: Readonly<Record<string, string>> = {
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  sl: "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
};
const ranks = new Map(
  wordSettingsOrder.map((name, index) => {
    const colon = name.indexOf(":");
    return [(colon < 0 ? W : aliases[name.slice(0, colon)]) + " " + name.slice(colon + 1), index];
  }),
);
export function sortWordSettings(root: XmlNode): void {
  if (root.uri !== W || root.name !== "settings") return;
  sortWordCompat(root);
  const known = root.content
    .filter(
      (node): node is XmlNode => typeof node !== "string" && ranks.has(node.uri + " " + node.name),
    )
    .sort((a, b) => ranks.get(a.uri + " " + a.name)! - ranks.get(b.uri + " " + b.name)!);
  let index = 0;
  for (let slot = 0; slot < root.content.length; slot++) {
    const node = root.content[slot]!;
    if (typeof node !== "string" && ranks.has(node.uri + " " + node.name))
      root.content[slot] = known[index++]!;
  }
}

/** CT_Compat uses a sequence too; mapped flags and retained flags share it. */
const compatOrder =
  "useSingleBorderforContiguousCells wpJustification noTabHangInd noLeading spaceForUL noColumnBalance balanceSingleByteDoubleByteWidth noExtraLineSpacing doNotLeaveBackslashAlone ulTrailSpace doNotExpandShiftReturn spacingInWholePoints lineWrapLikeWord6 printBodyTextBeforeHeader printColBlack wpSpaceWidth showBreaksInFrames subFontBySize suppressBottomSpacing suppressTopSpacing suppressSpacingAtTopOfPage suppressTopSpacingWP suppressSpBfAfterPgBrk swapBordersFacingPages convMailMergeEsc truncateFontHeightsLikeWP6 mwSmallCaps usePrinterMetrics doNotSuppressParagraphBorders wrapTrailSpaces footnoteLayoutLikeWW8 shapeLayoutLikeWW8 alignTablesRowByRow forgetLastTabAlignment adjustLineHeightInTable autoSpaceLikeWord95 noSpaceRaiseLower doNotUseHTMLParagraphAutoSpacing layoutRawTableWidth layoutTableRowsApart useWord97LineBreakRules doNotBreakWrappedTables doNotSnapToGridInCell selectFldWithFirstOrLastChar applyBreakingRules doNotWrapTextWithPunct doNotUseEastAsianBreakRules useWord2002TableStyleRules growAutofit useFELayout useNormalStyleForList doNotUseIndentAsNumberingTabStop useAltKinsokuLineBreakRules allowSpaceOfSameStyleInTable doNotSuppressIndentation doNotAutofitConstrainedTables autofitToFirstFixedWidthCell underlineTabInNumList displayHangulFixedWidth splitPgBreakAndParaMark doNotVertAlignCellWithSp doNotBreakConstrainedForcedTable doNotVertAlignInTxbx useAnsiKerningPairs cachedColBalance compatSetting".split(
    " ",
  );
const compatRanks = new Map(compatOrder.map((name, index) => [W + " " + name, index]));
function sortWordCompat(root: XmlNode): void {
  for (const node of root.content)
    if (typeof node !== "string" && node.uri === W && node.name === "compat") {
      const known = node.content
        .filter(
          (child): child is XmlNode =>
            typeof child !== "string" && compatRanks.has(child.uri + " " + child.name),
        )
        .sort(
          (a, b) => compatRanks.get(a.uri + " " + a.name)! - compatRanks.get(b.uri + " " + b.name)!,
        );
      let index = 0;
      for (let slot = 0; slot < node.content.length; slot++) {
        const child = node.content[slot]!;
        if (typeof child !== "string" && compatRanks.has(child.uri + " " + child.name))
          node.content[slot] = known[index++]!;
      }
    }
}
