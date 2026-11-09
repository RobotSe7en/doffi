import { DocumentError } from "@doffi/core";
import type { SourceStyleResolver, WorkbookDisplayFormat, WorkbookNumberDisplay, WorkbookNumberFormatter, WorkbookSnapshot, WorkbookValue } from "@doffi/core";
import type { FontResolution } from "../contracts.js";
import type { FontProbeCache } from "../font-probe-cache.js";
import type { WorkbookSourceCell } from "./cell-source.js";
import { createWorkbookTextView } from "./text-view.js";
import { createWorkbookNumberText } from "./number-text.js";
import { layoutWorkbookTexts } from "./text-layout.js";
import type { WorkbookTextLayout, WorkbookTextLayoutInput } from "./text-layout.js";
import { layoutWorkbookNumbers } from "./number-text-layout.js";
import type { WorkbookNumberLayoutInput } from "./number-text-layout.js";
import { layoutWorkbookFormattedTexts } from "./formatted-text.js";
import type { WorkbookFormattedTextInput } from "./formatted-text.js";
import { createSourceStyleResolver } from "../source-style-probe.js";

export interface WorkbookCellTextInput {
  readonly source: WorkbookSourceCell;
  readonly value: WorkbookValue | undefined;
  readonly formula: boolean;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly bordered: boolean;
}
export interface WorkbookCellTextLayouts {
  readonly nodes: number;
  readonly units: number;
  readonly fonts: readonly FontResolution[];
  readonly usedFallbackLineMetrics: boolean;
  readonly cells: readonly { readonly text?: WorkbookTextLayout; readonly numberOverflow?: Extract<WorkbookNumberDisplay, { kind: "overflow" }>["reason"] }[];
}
/** One display and natural-height pipeline for visible cells and stable row
 * measurement. Both use source styles, numeric/text formats and local fonts. */
export function layoutWorkbookCellTexts(scope: HTMLElement, snapshot: WorkbookSnapshot, inputs: readonly WorkbookCellTextInput[], formatter: WorkbookNumberFormatter,
  options: { readonly normalSpaceWidthPx: number }, maxNodes: number, maxTextUnits: number, probes?: FontProbeCache): WorkbookCellTextLayouts {
  const formats = new Map(snapshot.metadata.workbook.numberFormats.map(format => [format.formatId, format]));
  let sourceStyleResolver: SourceStyleResolver | undefined;
  const sourceStyles = (): SourceStyleResolver => sourceStyleResolver ??= createSourceStyleResolver(scope.ownerDocument);
  const texts: WorkbookTextLayoutInput[] = [], numbers: WorkbookNumberLayoutInput[] = [], formatted: WorkbookFormattedTextInput[] = [];
  const locations = new Map<WorkbookCellTextInput, { kind: "text" | "number" | "formatted"; index: number }>();
  let textNodes = 0, units = 0;
  const remaining = (): number => {
    if (textNodes >= maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook window text exceeds its cumulative node budget.");
    return maxNodes - textNodes;
  };
  for (const input of inputs) {
    const { source, value, formula, widthPt, heightPt, bordered } = input;
    if (!value || value.type === "blank") continue;
    const format: WorkbookDisplayFormat | undefined = source.format.numberFormatId ? formats.get(source.format.numberFormatId) : undefined;
    const size = { widthPt, heightPt, bordered };
    if (value.type === "number") {
      locations.set(input, { kind: "number", index: numbers.length }); numbers.push({ source, value: value.value, ...(format ? { format } : {}), ...size });
      continue;
    }
    if (value.type === "text" && format) {
      const display = formatter.format(value.value, format);
      if (display.kind !== "parts" || display.color || display.parts.some(part => part.kind !== "text")
        || display.parts.length !== 1 || display.parts[0]?.kind !== "text" || !display.parts[0].source) {
        // Text formats apply to both plain and rich values. @ fields retain
        // source runs; affixes are derived and a format color overrides runs.
        locations.set(input, { kind: "formatted", index: formatted.length }); formatted.push({ source, value: value.value, format, derived: formula, ...size }); continue;
      }
    }
    const text = value.type === "text" && !formula ? createWorkbookTextView(source, remaining(), sourceStyles)
      : createWorkbookNumberText(source, [{ kind: "text", text: value.type === "boolean" ? value.value ? "TRUE" : "FALSE" : value.value }], undefined, remaining());
    textNodes += text.nodes + 1; units += text.element.textContent?.length ?? 0;
    if (units > maxTextUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook window text exceeds its cumulative source-text budget.");
    locations.set(input, { kind: "text", index: texts.length }); texts.push({ source, text, valueType: value.type, ...size });
  }
  const rich = layoutWorkbookTexts(scope, texts, options.normalSpaceWidthPx, maxNodes, maxTextUnits, probes);
  textNodes = rich.cells.reduce((sum, cell) => sum + cell.nodes, 0);
  units = rich.cells.reduce((sum, cell) => sum + cell.codeUnits, 0);
  const formattedText = layoutWorkbookFormattedTexts(scope, formatted, formatter, options.normalSpaceWidthPx, remaining(), Math.max(1, maxTextUnits - units), probes, sourceStyles);
  textNodes += formattedText.cells.reduce((sum, cell) => sum + cell.nodes, 0);
  units += formattedText.cells.reduce((sum, cell) => sum + cell.codeUnits, 0);
  if (units > maxTextUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook window exceeds its cumulative displayed-text budget.");
  const numeric = layoutWorkbookNumbers(scope, numbers, formatter, options.normalSpaceWidthPx, remaining(), Math.max(1, maxTextUnits - units), probes);
  textNodes += numeric.cells.reduce((sum, cell) => sum + cell.nodes, 0);
  units += numeric.cells.reduce((sum, cell) => sum + cell.codeUnits, 0);
  if (units > maxTextUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook window exceeds its cumulative displayed-text budget.");
  if (textNodes >= maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook text exceeds its cumulative node budget.");
  return Object.freeze({ nodes: textNodes, units,
    fonts: Object.freeze([...rich.fonts, ...formattedText.fonts, ...numeric.fonts]),
    usedFallbackLineMetrics: rich.usedFallbackLineMetrics || formattedText.usedFallbackLineMetrics || numeric.usedFallbackLineMetrics,
    cells: Object.freeze(inputs.map(input => {
      const location = locations.get(input);
      const text = location ? (location.kind === "number" ? numeric : location.kind === "formatted" ? formattedText : rich).cells[location.index] : undefined;
      const overflow = location?.kind === "number" ? numeric.overflow[location.index] : undefined;
      return Object.freeze({ ...(text ? { text } : {}), ...(overflow ? { numberOverflow: overflow } : {}) });
    })),
  });
}
