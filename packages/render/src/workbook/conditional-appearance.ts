import { createWorkbookConditionalFormatResolver, DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookConditionalRule, WorkbookRuleReadResponse, WorkbookSheet } from "@doffi/core";
import type { WorkbookSourceCell } from "./cell-source.js";
import { resolvedWorkbookColor } from "./color.js";
import { WorkbookStyleLayers, workbookStylePropertyGroup } from "./style-layers.js";

/** A matched higher-priority rule claims only its declared appearance groups.
 * An unavailable condition claims the same groups without painting them, so a
 * lower rule cannot present an uncertain result as a known match. */
export class WorkbookConditionalAppearance {
  private readonly styles: WorkbookStyleLayers;
  private readonly definitions: ReadonlyMap<string, DeepReadonly<WorkbookConditionalRule>>;
  private readonly values: ReadonlyMap<number, WorkbookRuleReadResponse["cells"][number]>;
  private units = 0;
  private readonly format: ReturnType<typeof createWorkbookConditionalFormatResolver>;
  constructor(private readonly document: Document, sheet: DeepReadonly<WorkbookSheet>, response: WorkbookRuleReadResponse, private readonly account: (amount: number) => void) {
    this.styles = new WorkbookStyleLayers(document, account);
    this.definitions = new Map(sheet.conditionalRules?.map(rule => [rule.ruleId, rule]));
    this.values = new Map(response.cells.map(cell => [cell.row * 16_384 + cell.column, cell]));
    this.format = createWorkbookConditionalFormatResolver(sheet);
  }
  apply(key: number, source: WorkbookSourceCell): WorkbookSourceCell {
    const cell = this.values.get(key); if (!cell?.conditional.length) return source;
    const claimed = new Set<string>(), patch = new Map<string, string>();
    for (const result of cell.conditional) {
      if (result.state === "miss") continue;
      const rule = this.definitions.get(result.ruleId)!;
      let properties: ReadonlyMap<string, string>;
      if (rule.kind === "colorScale") {
        if (claimed.has("fill")) continue;
        if (result.state === "unavailable") { claimed.add("fill"); claimed.add("background-color"); continue; }
        const scale = result.scale!, stops = [rule.stops[scale.fromStop]!, rule.stops[scale.toStop]!];
        const colors = stops.map(stop => {
          this.account(1);
          const probe = source.element.cloneNode(false) as HTMLTableCellElement; probe.removeAttribute("style"); probe.className = stop.cssClass;
          source.element.parentElement!.append(probe);
          try {
            const color = resolvedWorkbookColor(this.document, this.document.defaultView!.getComputedStyle(probe), "--office-scale");
            const channels = color && /^rgba\((\d+),(\d+),(\d+),([\d.]+)\)$/.exec(color);
            if (!channels) throw new DocumentError("InvalidDocument", "A workbook color scale requires a resolved sRGB stop.");
            return channels.slice(1).map(Number);
          } finally { probe.remove(); this.account(-1); }
        });
        const first = colors[0]!, last = colors[1]!, rgb = first.slice(0, 3).map((value, index) => value + Math.trunc((last[index]! - value) * scale.ratio));
        const alpha = first[3]! * (1 - scale.ratio) + last[3]! * scale.ratio;
        properties = new Map([["--office-cell-fill-kind", "initial"], ["background-color", `rgba(${rgb.join(",")},${alpha})`]]);
      } else {
        properties = rule.cssClass ? this.styles.read(source.element, rule.cssClass, true) : new Map();
      }
      const groups = new Set([...properties.keys()].map(workbookStylePropertyGroup));
      if (result.state === "match") for (const [name, value] of properties) if (!claimed.has(workbookStylePropertyGroup(name))) patch.set(name, value);
      for (const group of groups) claimed.add(group);
    }
    for (const [name, value] of patch) {
      // A conditional font property is above an individual rich run as well
      // as the cell's ordinary font. Unspecified run properties stay local.
      const font = name.startsWith("font-") || name.startsWith("--office-font-") || name.startsWith("--office-cell-foreground-")
        || name === "color" || name === "line-height" || name === "letter-spacing" || name === "word-spacing" || name === "vertical-align" || name.startsWith("text-decoration-");
      for (const element of [source.element, ...font ? source.element.querySelectorAll<HTMLElement>("span") : []]) {
        this.units += name.length + value.length + 2;
        if (this.units > 8 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "Conditional appearance exceeds its window style budget.");
        element.style.setProperty(name, value, "important");
      }
    }
    const format = this.format({ row: Math.floor(key / 16_384), column: key % 16_384 }, cell.conditional, source.format).format;
    return format === source.format ? source : Object.freeze({ element: source.element, format });
  }
}
