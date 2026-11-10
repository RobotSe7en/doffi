import { createWorkbookNumberFormatter, DocumentError, WorkbookConditionalRuleSchema, WorkbookDataValidationSchema } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, WorkbookConditionalRule, WorkbookDataValidation } from "@doffi/core";
import type { WorkbookControlsHost, WorkbookControlsIdentity } from "./controls-host.js";
import { workbookControlAddress, workbookControlRange, workbookControlRangeText } from "./controls-range.js";
import { reconcileWorkbookInputText } from "./text-projection.js";
import { createWorkbookRuleAppearanceControls } from "./rule-appearance-controls.js";

type Family = "conditional" | "validation";
interface Draft {
  readonly identity: WorkbookControlsIdentity; readonly sheetId: string; readonly family: Family; readonly ruleId: string;
  readonly original?: DeepReadonly<WorkbookConditionalRule | WorkbookDataValidation>;
  readonly stopClasses: readonly string[];
  texts: Record<string, string>; dirty: boolean;
}
/** Native definitions, appearance and explicit warning choices form one source
 * transaction. Navigation never transfers a draft to a different worksheet. */
export function createWorkbookRulesControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void; hasPendingInput(): boolean } {
  const panel = document.createElement("details"), summary = document.createElement("summary"); summary.textContent = "条件格式与数据验证"; panel.append(summary); parent.append(panel);
  const fields: (HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement)[] = [], raw = new Map<string, HTMLInputElement | HTMLTextAreaElement>();
  const field = <T extends HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(owner: HTMLElement, label: string, input: T): T => { const wrapper = document.createElement("label"); wrapper.style.cssText = "display:flex;flex-wrap:wrap;align-items:center;gap:4px;max-width:100%"; input.style.maxWidth = "100%"; input.style.minWidth = "0"; input.setAttribute("aria-label", label); wrapper.append(label, input); owner.append(wrapper); fields.push(input); return input; };
  const text = (owner: HTMLElement, label: string, key: string, maximum: number, multiline = true) => { const input = field(owner, label, document.createElement(multiline ? "textarea" : "input")); input.maxLength = maximum; raw.set(key, input); return input; };
  const option = (select: HTMLSelectElement, value: string, label: string) => { const item = document.createElement("option"); item.value = value; const prefix = [...label.slice(0, 96)].slice(0, 48).join(""); item.textContent = prefix + (prefix.length < label.length ? "…" : ""); item.title = label; select.append(item); };
  const select = (owner: HTMLElement, label: string, options: readonly (readonly [string, string])[]) => { const input = field(owner, label, document.createElement("select")); for (const [value, label] of options) option(input, value, label); return input; };
  const check = (owner: HTMLElement, label: string) => { const input = field(owner, label, document.createElement("input")); input.type = "checkbox"; return input; };
  const number = (owner: HTMLElement, label: string, minimum: number, maximum: number) => { const input = field(owner, label, document.createElement("input")); input.type = "number"; input.min = String(minimum); input.max = String(maximum); return input; };
  const choice = select(panel, "选择条件格式或验证规则", []); text(panel, "规则单元格区域（用分号分隔）", "ranges", 32767); text(panel, "规则公式基准单元格", "origin", 16, false);
  const conditional = document.createElement("div"), validation = document.createElement("div"); panel.append(conditional, validation);
  const cfKind = select(conditional, "条件格式类型", [["cellIs", "数值比较"], ["expression", "标量公式"], ["colorScale", "色阶"]]), priority = number(conditional, "条件格式优先级", 1, 2147483647), stop = check(conditional, "条件满足后停止其他规则"), numberFormat = select(conditional, "条件格式数字格式", []);
  const dvKind = select(validation, "数据验证类型", [["none", "任意值"], ["list", "列表"], ["whole", "整数"], ["decimal", "小数"], ["date", "日期序号"], ["time", "时间序号"], ["textLength", "文本长度"], ["custom", "自定义公式"]]);
  const allowBlank = check(validation, "验证允许空白"), dropdown = check(validation, "显示验证下拉列表"), showInput = check(validation, "显示验证输入提示"), showError = check(validation, "显示验证错误提示"), errorStyle = select(validation, "验证错误方式", [["stop", "停止"], ["warning", "警告"], ["information", "信息"]]);
  text(validation, "验证输入提示标题", "promptTitle", 32); text(validation, "验证输入提示文字", "prompt", 255); text(validation, "验证错误提示标题", "errorTitle", 32); text(validation, "验证错误提示文字", "error", 225);
  const acceptWarnings = check(validation, "确认本次规则的验证警告");
  const formulas = document.createElement("div"); panel.append(formulas);
  const operator = select(formulas, "规则比较方式", [["between", "介于"], ["notBetween", "不介于"], ["equal", "等于"], ["notEqual", "不等于"], ["lessThan", "小于"], ["lessThanOrEqual", "小于或等于"], ["greaterThan", "大于"], ["greaterThanOrEqual", "大于或等于"]]);
  text(formulas, "规则公式一", "firstFormula", 8192); const secondFormula = text(formulas, "规则公式二", "secondFormula", 8192);
  const scale = document.createElement("div"); conditional.append(scale); const count = select(scale, "色阶颜色数量", [["2", "双色"], ["3", "三色"]]);
  let draft: Draft | undefined, selected = "new-conditional", key = "", formatsKey = "", serial = 0, saving = false, composing: HTMLElement | undefined;
  function changed(): void { if (!draft) return; for (const [name, input] of raw) draft.texts[name] = reconcileWorkbookInputText(draft.texts[name] ?? "", input.value); draft.dirty = true; refresh(); host.draftChanged?.(); }
  const appearance = createWorkbookRuleAppearanceControls(document, conditional, host.signal, "条件规则", true, changed);
  const stops = ["色阶起点", "色阶中点", "色阶终点"].map((label, index) => {
    const element = document.createElement("div"); scale.append(element);
    const kind = select(element, label + "类型", [["min", "最小值"], ["max", "最大值"], ["number", "数值"], ["percent", "百分比"], ["percentile", "百分位"], ["formula", "公式"]]);
    const value = number(element, label + "数值", -Number.MAX_VALUE, Number.MAX_VALUE), expression = text(element, label + "公式", "stopFormula" + index, 8192), paint = createWorkbookRuleAppearanceControls(document, element, host.signal, label, false, changed);
    return { element, kind, value, expression, paint };
  });
  const status = document.createElement("p"); status.setAttribute("role", "status"); status.setAttribute("aria-label", "规则编辑状态"); panel.append(status);
  const sources = () => host.snapshot()?.metadata.workbook.sheets.flatMap(sheet => [
    ...[...(sheet.conditionalRules ?? [])].sort((left, right) => left.priority - right.priority).map(rule => ({ value: JSON.stringify(["conditional", sheet.sheetId, rule.ruleId]), sheet, rule, family: "conditional" as const })),
    ...(sheet.dataValidations ?? []).map(rule => ({ value: JSON.stringify(["validation", sheet.sheetId, rule.ruleId]), sheet, rule, family: "validation" as const })),
  ]) ?? [];
  const stale = () => !!draft && (host.snapshot()?.manifest.documentId !== draft.identity.documentId || host.snapshot()?.manifest.revision !== draft.identity.revision);
  function load(): void {
    const snapshot = host.snapshot(), source = sources().find(source => source.value === selected), selection = host.selection(), sheet = source?.sheet ?? snapshot?.metadata.workbook.sheets.find(sheet => sheet.sheetId === selection?.sheetId); draft = undefined;
    if (!snapshot || !sheet) return;
    const family = source?.family ?? (selected === "new-validation" ? "validation" : "conditional"), original = source?.rule, at = selection?.range ?? { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } };
    let ruleId = original?.ruleId; if (!ruleId) do { ruleId = "editor-rule-" + ++serial; } while (sources().some(source => source.rule.ruleId === ruleId));
    const originalScale = original?.kind === "colorScale" ? original : undefined, sourceStops = [originalScale?.stops[0], originalScale?.stops.length === 3 ? originalScale.stops[1] : undefined, originalScale?.stops.at(-1)];
    const texts: Record<string, string> = { ranges: (original?.ranges ?? [at]).map(workbookControlRangeText).join(";"), origin: workbookControlAddress(original?.origin ?? at.from), firstFormula: original && "expression" in original ? original.expression : original && "formulas" in original ? original.formulas[0] ?? "=0" : "=0", secondFormula: original && "formulas" in original ? original.formulas[1] ?? "=100" : "=100", promptTitle: "", prompt: "", errorTitle: "", error: "" };
    if (original && "errorStyle" in original) for (const name of ["promptTitle", "prompt", "errorTitle", "error"] as const) texts[name] = original[name];
    for (const [index, stop] of sourceStops.entries()) texts["stopFormula" + index] = stop?.kind === "formula" ? stop.expression : "=0";
    draft = { identity: { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision }, sheetId: sheet.sheetId, family, ruleId, ...(original ? { original } : {}), stopClasses: sourceStops.map(stop => stop?.cssClass ?? sheet.defaultCssClass), texts, dirty: false };
    for (const [name, input] of raw) input.value = texts[name] ?? "";
    cfKind.value = original && ["cellIs", "expression", "colorScale"].includes(original.kind) ? original.kind : "cellIs"; dvKind.value = original && "errorStyle" in original ? original.kind : "whole";
    operator.value = original && "operator" in original ? original.operator : "between"; priority.value = String(original && "priority" in original ? original.priority : Math.max(0, ...sheet.conditionalRules?.map(rule => rule.priority) ?? []) + 1); stop.checked = !!original && "stopIfTrue" in original && original.stopIfTrue;
    numberFormat.value = original && "numberFormatId" in original ? original.numberFormatId ?? "" : "";
    const dv = original && "errorStyle" in original ? original : undefined; allowBlank.checked = dv?.allowBlank ?? true; dropdown.checked = dv?.kind === "list" ? dv.dropdown : true; showInput.checked = dv?.showInputMessage ?? true; showError.checked = dv?.showErrorMessage ?? true; errorStyle.value = dv?.errorStyle ?? "stop"; acceptWarnings.checked = false;
    appearance.load(snapshot, original && "cssClass" in original ? original.cssClass : undefined, !original ? "#ffeb9c" : undefined);
    count.value = String(originalScale?.stops.length ?? 2);
    for (const [index, slot] of stops.entries()) { const source = sourceStops[index]; slot.kind.value = source?.kind ?? (index === 0 ? "min" : index === 1 ? "percentile" : "max"); slot.value.value = String(source && "value" in source ? source.value : index === 1 ? 50 : 0); slot.paint.load(snapshot, source?.cssClass, source ? undefined : ["#f8696b", "#ffeb84", "#63be7b"][index]); }
  }
  function commands(): DocumentOperation[] {
    const current = draft!, parsedOrigin = workbookControlRange(current.texts.origin!), target = { ruleId: current.ruleId, ranges: current.texts.ranges!.split(";").filter(text => text.trim()).map(workbookControlRange), origin: parsedOrigin.from };
    if (parsedOrigin.from.row !== parsedOrigin.to.row || parsedOrigin.from.column !== parsedOrigin.to.column) throw new DocumentError("InvalidOperation", "规则基准须为一个单元格。");
    const pair = [current.texts.firstFormula!, ...["between", "notBetween"].includes(operator.value) ? [current.texts.secondFormula!] : []];
    if (current.family === "validation") {
      const common = { ...target, allowBlank: allowBlank.checked, showInputMessage: showInput.checked, showErrorMessage: showError.checked, errorStyle: errorStyle.value, errorTitle: current.texts.errorTitle!, error: current.texts.error!, promptTitle: current.texts.promptTitle!, prompt: current.texts.prompt! };
      const rule = WorkbookDataValidationSchema.parse({ ...common, kind: dvKind.value, ...dvKind.value === "none" ? {} : dvKind.value === "list" ? { expression: current.texts.firstFormula!, dropdown: dropdown.checked } : dvKind.value === "custom" ? { expression: current.texts.firstFormula! } : { operator: operator.value, formulas: pair } });
      return [{ op: "setWorkbookDataValidation", sheetId: current.sheetId, rule }];
    }
    const common = { ...target, priority: priority.valueAsNumber, stopIfTrue: cfKind.value === "colorScale" ? false : stop.checked }, operations: DocumentOperation[] = [];
    if (cfKind.value === "colorScale") {
      const selected = Number(count.value) === 3 ? [0, 1, 2] : [0, 2], thresholds = selected.map(index => { const slot = stops[index]!, kind = slot.kind.value, base = { kind, cssClass: current.stopClasses[index] }; return kind === "formula" ? { ...base, expression: current.texts["stopFormula" + index] } : kind === "min" || kind === "max" ? base : { ...base, value: slot.value.valueAsNumber }; });
      const rule = WorkbookConditionalRuleSchema.parse({ ...common, kind: "colorScale", stops: thresholds }); operations.push({ op: "setWorkbookConditionalRule", sheetId: current.sheetId, rule });
      for (const [stopIndex, index] of selected.entries()) { const properties = stops[index]!.paint.properties(); if (Object.keys(properties).length) operations.push({ op: "formatWorkbookConditionalRule", sheetId: current.sheetId, ruleId: current.ruleId, stopIndex, properties }); }
    } else {
      const original = current.original, rule = WorkbookConditionalRuleSchema.parse({ ...common, kind: cfKind.value, ...cfKind.value === "expression" ? { expression: current.texts.firstFormula } : { operator: operator.value, formulas: pair }, ...(original && "cssClass" in original && original.cssClass ? { cssClass: original.cssClass } : {}), ...(numberFormat.value ? { numberFormatId: numberFormat.value } : {}) });
      operations.push({ op: "setWorkbookConditionalRule", sheetId: current.sheetId, rule }); const properties = appearance.properties(); if (Object.keys(properties).length) operations.push({ op: "formatWorkbookConditionalRule", sheetId: current.sheetId, ruleId: current.ruleId, properties });
    }
    return operations;
  }
  const button = (label: string, action: () => Promise<void>, requiresCurrent = true) => { const button = document.createElement("button"); button.type = "button"; button.textContent = label; panel.append(button); button.addEventListener("click", () => host.run(async () => {
    if (!host.idle() || saving || composing) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。"); if (requiresCurrent && (!draft || stale())) throw new DocumentError("RevisionConflict", "规则草稿属于旧修订，请重新读取。");
    saving = true; try { const submitted = action(); refresh(); host.draftChanged?.(); await submitted; } finally { saving = false; refresh(); host.draftChanged?.(); }
  }), { signal: host.signal }); return button; };
  const save = button("保存条件格式或验证规则", async () => {
    const current = draft!, operations = commands(), defaults = current.family === "validation" && acceptWarnings.checked ? host.prepareOptions?.() : undefined;
    const options = current.family === "validation" && acceptWarnings.checked ? { ...defaults, validation: { acceptedRuleIds: [...new Set([...(defaults?.validation?.acceptedRuleIds ?? []), current.ruleId])] } } : undefined;
    await host.perform(operations, options, current.identity); selected = JSON.stringify([current.family, current.sheetId, current.ruleId]); draft = undefined; key = "";
  });
  const remove = button("删除当前规则", async () => { const current = draft!; await host.perform([{ op: current.family === "conditional" ? "removeWorkbookConditionalRule" : "removeWorkbookDataValidation", sheetId: current.sheetId, ruleId: current.ruleId }], undefined, current.identity); selected = current.family === "conditional" ? "new-conditional" : "new-validation"; draft = undefined; key = ""; });
  function priorityPeer(direction: number) {
    if (!draft?.original || draft.family !== "conditional") return;
    const rules = [...host.snapshot()?.metadata.workbook.sheets.find(sheet => sheet.sheetId === draft!.sheetId)?.conditionalRules ?? []].sort((left, right) => left.priority - right.priority), index = rules.findIndex(rule => rule.ruleId === draft!.ruleId);
    return index < 0 ? undefined : rules[index + direction];
  }
  async function movePriority(direction: number): Promise<void> {
    if (draft?.dirty) throw new DocumentError("InvalidOperation", "请先保存或放弃当前规则草稿。");
    const peer = priorityPeer(direction), current = draft?.original; if (!draft || !peer || !current || !("priority" in current)) return;
    await host.perform([{ op: "setWorkbookConditionalRule", sheetId: draft.sheetId, rule: { ...current, priority: peer.priority } as WorkbookConditionalRule }, { op: "setWorkbookConditionalRule", sheetId: draft.sheetId, rule: { ...peer, priority: current.priority } as WorkbookConditionalRule }], undefined, draft.identity); draft = undefined;
  }
  const up = button("条件规则上移", () => movePriority(-1)), down = button("条件规则下移", () => movePriority(1));
  const reload = button("放弃规则草稿并重新读取", async () => { if (!sources().some(source => source.value === selected) && !selected.startsWith("new-")) selected = draft?.family === "validation" ? "new-validation" : "new-conditional"; draft = undefined; key = ""; refresh(); load(); }, false);
  choice.addEventListener("change", () => { if (draft?.dirty || composing || saving) { choice.value = selected; return; } selected = choice.value; draft = undefined; load(); refresh(); }, { signal: host.signal });
  for (const field of fields.filter(field => field !== choice)) field.addEventListener("input", changed, { signal: host.signal });
  panel.addEventListener("compositionstart", event => { if (!(event.target instanceof document.defaultView!.HTMLElement)) return; composing = event.target; changed(); }, { capture: true, signal: host.signal }); panel.addEventListener("compositionend", () => { composing = undefined; changed(); }, { capture: true, signal: host.signal });
  function refresh(): void {
    const snapshot = host.snapshot(), all = sources(), next = JSON.stringify([snapshot?.manifest.documentId, all.map(source => [source.value, source.sheet.name, source.rule.kind, source.rule.ranges, "priority" in source.rule ? source.rule.priority : undefined])]);
    if (next !== key) { key = next; choice.replaceChildren(); option(choice, "new-conditional", "新建条件格式"); option(choice, "new-validation", "新建数据验证"); const labels = { cellIs: "数值比较", expression: "标量公式", colorScale: "色阶", none: "任意值", list: "列表", whole: "整数", decimal: "小数", date: "日期", time: "时间", textLength: "文本长度", custom: "自定义公式" }; for (const source of all) option(choice, source.value, source.sheet.name + " · " + labels[source.rule.kind] + " " + source.rule.ranges.map(workbookControlRangeText).join(";").slice(0, 96) + ("priority" in source.rule ? " · 优先级 " + source.rule.priority : "")); if (!selected.startsWith("new-") && !all.some(source => source.value === selected)) option(choice, selected, "已移除的规则 · 草稿保留"); choice.value = selected; }
    const formats = snapshot?.metadata.workbook.numberFormats ?? [], nextFormats = JSON.stringify([formats, snapshot?.metadata.workbook.dateSystem, snapshot?.metadata.locale, host.locale]); if (nextFormats !== formatsKey && snapshot) {
      const previous = numberFormat.value; formatsKey = nextFormats; numberFormat.replaceChildren(); option(numberFormat, "", "沿用单元格格式");
      const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem, ...(snapshot.metadata.locale ? { documentLocale: snapshot.metadata.locale } : {}), ...(host.locale ? { hostLocale: host.locale } : {}) });
      try { for (const format of formats) {
        let label: string;
        if ("code" in format) label = format.code;
        else try { const display = formatter.format(1234.5, format); label = "格式示例：" + (display.kind === "parts" ? display.parts.flatMap(part => "text" in part ? [part.text] : []).join("") : "超出显示范围"); }
        catch (error) { if (!(error instanceof DocumentError) || error.code !== "UnsupportedFeature") throw error; label = "来源数字格式（暂无示例）"; }
        option(numberFormat, format.formatId, label);
      } } finally { formatter.dispose(); }
      numberFormat.value = previous;
    }
    const selection = host.selection(), followsSelection = draft && !draft.original && selection && (draft.sheetId !== selection.sheetId || draft.texts.ranges !== workbookControlRangeText(selection.range));
    if (!saving && !composing && (!draft || !draft.dirty && (stale() || followsSelection))) load();
    const blocked = !host.idle() || saving || !!composing, isConditional = draft?.family === "conditional", isScale = isConditional && cfKind.value === "colorScale", comparisons = isConditional ? cfKind.value === "cellIs" : !["none", "list", "custom"].includes(dvKind.value);
    for (const input of fields) input.disabled = (blocked || !draft) && input !== composing; choice.disabled = blocked || !!draft?.dirty;
    conditional.hidden = !isConditional; validation.hidden = !!isConditional; formulas.hidden = !!isScale || !isConditional && dvKind.value === "none"; operator.parentElement!.hidden = !comparisons; secondFormula.parentElement!.hidden = !comparisons || !["between", "notBetween"].includes(operator.value);
    scale.hidden = !isScale; appearance.element.hidden = !!isScale; appearance.refresh(blocked || !draft, composing); numberFormat.disabled ||= !!isScale; stop.disabled ||= !!isScale; dropdown.disabled ||= dvKind.value !== "list"; acceptWarnings.disabled ||= errorStyle.value === "stop";
    for (const [index, slot] of stops.entries()) { slot.element.hidden = index === 1 && count.value !== "3"; slot.expression.parentElement!.hidden = slot.kind.value !== "formula"; slot.value.parentElement!.hidden = ["min", "max", "formula"].includes(slot.kind.value); slot.paint.refresh(blocked || !draft, composing); }
    save.disabled = blocked || !draft || stale(); remove.disabled = save.disabled || !draft?.original; reload.disabled = blocked;
    up.disabled = save.disabled || !!draft?.dirty || !priorityPeer(-1); down.disabled = save.disabled || !!draft?.dirty || !priorityPeer(1);
    const sheet = snapshot?.metadata.workbook.sheets.find(sheet => sheet.sheetId === draft?.sheetId); status.textContent = stale() ? "文档已变化。规则草稿已保留，请重新读取后编辑。" : !draft ? "请选择单元格区域。" : (sheet?.name ?? draft.sheetId) + " · " + draft.texts.ranges + (draft.dirty ? " · 规则草稿尚未保存。" : "");
  }
  refresh(); return { element: panel, refresh, hasPendingInput: () => !!draft?.dirty || saving || !!composing };
}
