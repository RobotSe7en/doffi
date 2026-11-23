import { copyWorkbookFormula } from "@doffi/core";
import type { DeepReadonly, WorkbookColorThreshold, WorkbookConditionalRule, WorkbookCoordinate, WorkbookDataValidation, WorkbookRange } from "@doffi/core";
import { empty } from "../export/xml.js";
import { colorTag } from "./style-values.js";
import type { WorkbookStylesWriter } from "./styles.js";
import { bool, element, formulaText, rangeRef, xstring } from "./xml.js";

function ruleOrigin(ranges: readonly WorkbookRange[]): WorkbookCoordinate {
  return ranges.reduce((point, area) => area.from.column < point.column ||
    (area.from.column === point.column && area.from.row < point.row) ? area.from : point, ranges[0]!.from);
}
const references = (ranges: readonly WorkbookRange[]): string => ranges.map(range => rangeRef(range)).join(" ");

export function conditionalRule(rule: DeepReadonly<WorkbookConditionalRule>, styles: WorkbookStylesWriter): string {
  const origin = ruleOrigin(rule.ranges);
  const formula = (expression: string): string => formulaText(copyWorkbookFormula(expression, rule.origin, origin));
  let body: string, dxfId: number | undefined, operator: string | undefined;
  if (rule.kind === "colorScale") {
    const threshold = (stop: DeepReadonly<WorkbookColorThreshold>): string =>
      empty("cfvo", { type: stop.kind === "number" ? "num" : stop.kind,
        val: "expression" in stop ? copyWorkbookFormula(stop.expression, rule.origin, origin).slice(1) : "value" in stop ? stop.value : undefined });
    const colors = rule.stops.map(stop => colorTag("color", styles.raw(stop.cssClass), styles.css(stop.cssClass), "--office-scale", "#000000"));
    body = element("colorScale", rule.stops.map(threshold).join("") + colors.join(""));
  } else {
    dxfId = styles.differential(rule.cssClass, rule.numberFormatId);
    operator = rule.kind === "cellIs" ? rule.operator : undefined;
    body = (rule.kind === "cellIs" ? rule.formulas : [rule.expression]).map(expression => element("formula", formula(expression))).join("");
  }
  return element("conditionalFormatting", element("cfRule", body, {
    type: rule.kind, priority: rule.priority, stopIfTrue: bool(rule.stopIfTrue), operator, dxfId,
  }), { sqref: references(rule.ranges) });
}

export function dataValidation(rule: DeepReadonly<WorkbookDataValidation>): string {
  const origin = ruleOrigin(rule.ranges);
  const expressions = "formulas" in rule ? rule.formulas : "expression" in rule ? [rule.expression] : [];
  const formulas = expressions.map((expression, index) =>
    element("formula" + (index + 1), formulaText(copyWorkbookFormula(expression, rule.origin, origin)))).join("");
  return element("dataValidation", formulas, {
    type: rule.kind, operator: "operator" in rule ? rule.operator : undefined, allowBlank: bool(rule.allowBlank),
    showDropDown: rule.kind === "list" ? bool(!rule.dropdown) : undefined,
    showInputMessage: bool(rule.showInputMessage), showErrorMessage: bool(rule.showErrorMessage),
    errorStyle: rule.errorStyle, errorTitle: xstring(rule.errorTitle), error: xstring(rule.error),
    promptTitle: xstring(rule.promptTitle), prompt: xstring(rule.prompt), sqref: references(rule.ranges),
  });
}
