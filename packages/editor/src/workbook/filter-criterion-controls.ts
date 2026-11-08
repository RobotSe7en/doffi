import type { WorkbookFilterCriterion, WorkbookFilterDate } from "@doffi/core";
import { reconcileWorkbookInputText } from "./text-projection.js";

type Field = HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement;
/** Criterion drafts are local form state. Lists are paged so long native value
 * and date vectors do not create an unbounded select or duplicate source. */
export function createWorkbookFilterCriterionControls(document: Document, parent: ParentNode, signal: AbortSignal, changed: () => void): { readonly element: HTMLDivElement; readonly fields: readonly Field[]; load(source: WorkbookFilterCriterion | undefined): void; value(): WorkbookFilterCriterion | undefined; refresh(disabled: boolean, active?: HTMLElement): void } {
  const element = document.createElement("div"); parent.append(element);
  const fields: Field[] = [], buttons: HTMLButtonElement[] = [], sections = new Map<string, HTMLElement>();
  const field = <T extends Field>(owner: HTMLElement, label: string, input: T): T => { const wrapper = document.createElement("label"); input.setAttribute("aria-label", label); wrapper.append(label, input); owner.append(wrapper); fields.push(input); return input; };
  const select = (owner: HTMLElement, label: string, options: readonly (readonly [string, string])[]) => { const input = field(owner, label, document.createElement("select")); for (const [value, label] of options) { const item = document.createElement("option"); item.value = value; item.textContent = label; input.append(item); } return input; };
  const text = (owner: HTMLElement, label: string) => { const input = field(owner, label, document.createElement("textarea")); input.maxLength = 32767; return input; };
  const check = (owner: HTMLElement, label: string) => { const input = field(owner, label, document.createElement("input")); input.type = "checkbox"; return input; };
  const number = (owner: HTMLElement, label: string, minimum: number, maximum: number) => { const input = field(owner, label, document.createElement("input")); input.type = "number"; input.min = String(minimum); input.max = String(maximum); return input; };
  const section = (kind: string) => { const part = document.createElement("div"); element.append(part); sections.set(kind, part); return part; };
  const kind = select(element, "筛选条件类型", [["none", "全部"], ["values", "值或日期列表"], ["custom", "自定义比较"], ["top", "排名"], ["dynamic", "动态日期或平均值"]]);
  const valuesSection = section("values"), blank = check(valuesSection, "筛选包含空白"), valuesChoice = select(valuesSection, "筛选值列表项", []), valueText = text(valuesSection, "筛选值文字");
  const dateChoice = select(valuesSection, "筛选日期列表项", []), dateGroup = select(valuesSection, "筛选日期精度", [["year", "年"], ["month", "月"], ["day", "日"], ["hour", "时"], ["minute", "分"], ["second", "秒"]]);
  const dateFields = [number(valuesSection, "筛选日期年", 1, 9999), number(valuesSection, "筛选日期月", 1, 12), number(valuesSection, "筛选日期日", 1, 31), number(valuesSection, "筛选日期时", 0, 23), number(valuesSection, "筛选日期分", 0, 59), number(valuesSection, "筛选日期秒", 0, 59)], dateKeys = ["year", "month", "day", "hour", "minute", "second"] as const;
  const customSection = section("custom"), operators = [["equal", "等于"], ["notEqual", "不等于"], ["lessThan", "小于"], ["lessThanOrEqual", "小于或等于"], ["greaterThan", "大于"], ["greaterThanOrEqual", "大于或等于"]] as const;
  const firstOperator = select(customSection, "筛选第一个比较", operators), firstText = text(customSection, "筛选第一个比较值"), second = check(customSection, "筛选使用第二个条件"), conjunction = select(customSection, "筛选条件组合", [["and", "同时满足"], ["or", "任一满足"]]), secondOperator = select(customSection, "筛选第二个比较", operators), secondText = text(customSection, "筛选第二个比较值");
  const topSection = section("top"), direction = select(topSection, "筛选排名方向", [["top", "最大"], ["bottom", "最小"]]), percent = check(topSection, "筛选排名使用百分比"), count = number(topSection, "筛选排名数量", 1, 500);
  const dynamicSection = section("dynamic"), dynamicOptions: [string, string][] = [["aboveAverage", "高于平均值"], ["belowAverage", "低于平均值"], ["tomorrow", "明天"], ["today", "今天"], ["yesterday", "昨天"], ["nextWeek", "下周"], ["thisWeek", "本周"], ["lastWeek", "上周"], ["nextMonth", "下月"], ["thisMonth", "本月"], ["lastMonth", "上月"], ["nextQuarter", "下季度"], ["thisQuarter", "本季度"], ["lastQuarter", "上季度"], ["nextYear", "明年"], ["thisYear", "今年"], ["lastYear", "去年"], ["yearToDate", "本年累计"]];
  for (let index = 1; index <= 4; index++) dynamicOptions.push(["Q" + index, "第 " + index + " 季度"]);
  for (let index = 1; index <= 12; index++) dynamicOptions.push(["M" + index, index + " 月"]);
  const dynamic = select(dynamicSection, "筛选动态条件", dynamicOptions);
  let values: string[] = [], dates: WorkbookFilterDate[] = [], valueIndex = -1, dateIndex = -1, firstRaw = "", secondRaw = "", blocked = false, composing: HTMLElement | undefined;
  const button = (label: string, action: () => void, mutates = true) => { const button = document.createElement("button"); button.type = "button"; button.textContent = label; valuesSection.append(button); buttons.push(button); button.addEventListener("click", () => { if (blocked) return; action(); render(); if (mutates) changed(); }, { signal }); return button; };
  const previousValue = button("筛选值上一页", () => { valueIndex = Math.max(0, Math.floor(valueIndex / 50) * 50 - 50); readValue(); }, false), nextValue = button("筛选值下一页", () => { valueIndex = Math.min(values.length - 1, Math.floor(valueIndex / 50) * 50 + 50); readValue(); }, false);
  const addValue = button("添加筛选值", () => { values.push(""); valueIndex = values.length - 1; readValue(); }), removeValue = button("移除筛选值", () => { if (valueIndex < 0) return; values.splice(valueIndex, 1); valueIndex = Math.min(valueIndex, values.length - 1); readValue(); });
  const previousDate = button("筛选日期上一页", () => { dateIndex = Math.max(0, Math.floor(dateIndex / 50) * 50 - 50); readDate(); }, false), nextDate = button("筛选日期下一页", () => { dateIndex = Math.min(dates.length - 1, Math.floor(dateIndex / 50) * 50 + 50); readDate(); }, false);
  const addDate = button("添加筛选日期", () => { dates.push({ grouping: "year", year: 2000 }); dateIndex = dates.length - 1; readDate(); }), removeDate = button("移除筛选日期", () => { if (dateIndex < 0) return; dates.splice(dateIndex, 1); dateIndex = Math.min(dateIndex, dates.length - 1); readDate(); });
  function readValue(): void { valueText.value = values[valueIndex] ?? ""; }
  function readDate(): void { const value = dates[dateIndex]; dateGroup.value = value?.grouping ?? "year"; for (const [index, input] of dateFields.entries()) input.value = value?.[dateKeys[index]!] === undefined ? "" : String(value[dateKeys[index]!]); }
  function writeDate(): void {
    if (dateIndex < 0) return; const grouping = dateGroup.value as WorkbookFilterDate["grouping"], last = dateKeys.indexOf(grouping);
    const date = { grouping, year: dateFields[0]!.valueAsNumber } as WorkbookFilterDate;
    for (let index = 1; index <= last; index++) Object.assign(date, { [dateKeys[index]!]: dateFields[index]!.valueAsNumber }); dates[dateIndex] = date;
  }
  const pages = new WeakMap<HTMLSelectElement, string>();
  function page<T>(select: HTMLSelectElement, entries: readonly T[], at: number, describe: (entry: T) => string): void {
    const from = Math.max(0, Math.floor(at / 50) * 50), labels: [number, string][] = [];
    for (let index = from; index < Math.min(entries.length, from + 50); index++) labels.push([index, (index + 1) + " · " + [...describe(entries[index]!).slice(0, 96)].slice(0, 48).join("")]);
    const key = JSON.stringify(labels);
    if (pages.get(select) !== key) {
      pages.set(select, key); select.replaceChildren();
      for (const [index, label] of labels) { const item = document.createElement("option"); item.value = String(index); item.textContent = label; select.append(item); }
    }
    select.value = String(at);
  }
  function render(): void {
    for (const [name, part] of sections) part.hidden = name !== kind.value;
    page(valuesChoice, values, valueIndex, value => value || "（空文字）"); page(dateChoice, dates, dateIndex, value => dateKeys.flatMap(key => value[key] === undefined ? [] : [value[key]]).join("-"));
    for (const input of fields) input.disabled = blocked && input !== composing;
    for (const button of buttons) button.disabled = blocked;
    valueText.disabled ||= valueIndex < 0; valuesChoice.disabled ||= !values.length; dateChoice.disabled ||= !dates.length; dateGroup.disabled ||= dateIndex < 0;
    for (const [index, input] of dateFields.entries()) input.disabled ||= dateIndex < 0 || index > dateKeys.indexOf(dateGroup.value as WorkbookFilterDate["grouping"]);
    secondOperator.disabled ||= !second.checked; secondText.disabled ||= !second.checked; conjunction.disabled ||= !second.checked;
    previousValue.disabled ||= valueIndex < 50; nextValue.disabled ||= Math.floor(valueIndex / 50) * 50 + 50 >= values.length; removeValue.disabled ||= valueIndex < 0;
    previousDate.disabled ||= dateIndex < 50; nextDate.disabled ||= Math.floor(dateIndex / 50) * 50 + 50 >= dates.length; removeDate.disabled ||= dateIndex < 0;
    // Keep explicit names for capability decisions; they do not borrow another
    // input's transient disabled state when an action captures its fields.
    addValue.disabled = addDate.disabled = blocked;
  }
  for (const input of fields.filter(input => input !== valuesChoice && input !== dateChoice)) input.addEventListener("input", () => {
    if (input === valueText && valueIndex >= 0) values[valueIndex] = reconcileWorkbookInputText(values[valueIndex]!, valueText.value);
    if (input === firstText) firstRaw = reconcileWorkbookInputText(firstRaw, firstText.value);
    if (input === secondText) secondRaw = reconcileWorkbookInputText(secondRaw, secondText.value);
    if (input === dateGroup || dateFields.includes(input as HTMLInputElement)) writeDate(); render(); changed();
  }, { signal });
  valuesChoice.addEventListener("change", () => { valueIndex = Number(valuesChoice.value); readValue(); render(); }, { signal });
  dateChoice.addEventListener("change", () => { dateIndex = Number(dateChoice.value); readDate(); render(); }, { signal });
  function load(source: WorkbookFilterCriterion | undefined): void {
    kind.value = source?.kind ?? "none"; values = source?.kind === "values" ? [...source.values] : []; dates = source?.kind === "values" ? source.dates.map(value => ({ ...value })) : [];
    blank.checked = source?.kind === "values" && source.blank; valueIndex = values.length ? 0 : -1; dateIndex = dates.length ? 0 : -1; readValue(); readDate();
    const custom = source?.kind === "custom" ? source : undefined; firstOperator.value = custom?.tests[0]?.operator ?? "equal"; firstRaw = custom?.tests[0]?.value ?? ""; firstText.value = firstRaw; second.checked = custom?.tests.length === 2; secondOperator.value = custom?.tests[1]?.operator ?? "equal"; secondRaw = custom?.tests[1]?.value ?? ""; secondText.value = secondRaw; conjunction.value = custom?.and === false ? "or" : "and";
    direction.value = source?.kind === "top" && !source.top ? "bottom" : "top"; percent.checked = source?.kind === "top" && source.percent; count.value = String(source?.kind === "top" ? source.count : 10); dynamic.value = source?.kind === "dynamic" ? source.type : "aboveAverage"; render();
  }
  function value(): WorkbookFilterCriterion | undefined {
    switch (kind.value) {
      case "values": return { kind: "values", blank: blank.checked, values: [...values], dates: dates.map(value => ({ ...value })) };
      case "custom": return { kind: "custom", and: conjunction.value === "and", tests: [{ operator: firstOperator.value as Extract<WorkbookFilterCriterion, { kind: "custom" }>["tests"][number]["operator"], value: firstRaw }, ...second.checked ? [{ operator: secondOperator.value as Extract<WorkbookFilterCriterion, { kind: "custom" }>["tests"][number]["operator"], value: secondRaw }] : []] };
      case "top": return { kind: "top", top: direction.value === "top", percent: percent.checked, count: count.valueAsNumber };
      case "dynamic": return { kind: "dynamic", type: dynamic.value as Extract<WorkbookFilterCriterion, { kind: "dynamic" }>["type"] };
      default: return undefined;
    }
  }
  load(undefined); return { element, fields, load, value, refresh(disabled: boolean, active?: HTMLElement): void { blocked = disabled; composing = active; render(); } };
}
