import { createWorkbookNumberFormatter, DocumentError } from "@doffi/core";
import type { WorkbookChart, DocumentOperation } from "@doffi/core";

type Target = Extract<DocumentOperation, { op: "setWorkbookChartStyle" }>["target"];
type Field = HTMLInputElement | HTMLSelectElement;
interface Factory {
  input(name: string, type?: string): HTMLInputElement;
  choice(name: string, choices: readonly (readonly [string, string])[]): HTMLSelectElement;
  value(key: string, field: Field): Field;
}
const keep = ["", "保持原设置"] as const, reset = ["initial", "恢复默认"] as const;
const axes: readonly Target[] = ["categoryAxis", "valueAxis"], labels: readonly Target[] = ["label", "seriesLabel", "pointLabel"];
/** These fields share the owning part's revision-bound draft and transaction.
 * Empty numeric inputs keep source values; explicit defaults remove the local
 * native setting without copying resolved values from another role. */
export function createWorkbookChartDetailFields(factory: Factory) {
  const fields: { field: Field; enabled(target: Target, kind: WorkbookChart["kind"]): boolean }[] = [];
  const register = (key: string, field: Field, targets: readonly Target[]) => {
    factory.value(key, field); fields.push({ field, enabled: (target, kind) => targets.includes(target) && (kind !== "pie" || !axes.includes(target)) }); return field;
  };
  const select = (key: string, text: string, choices: readonly (readonly [string, string])[], targets = axes) => register(key, factory.choice(text, choices), targets) as HTMLSelectElement;
  const numeric = (key: string, text: string, targets = axes) => { const field = register(key, factory.input(text, "number"), targets) as HTMLInputElement; field.step = "any"; return field; };
  const crossing = select("crossing", "坐标轴交叉方式", [keep, reset, ["autoZero", "自动过零"], ["min", "另一轴最小端"], ["max", "另一轴最大端"], ["at", "指定另一轴上的值"]]);
  const crossingAt = numeric("crossingAt", "坐标轴在另一轴上的交叉值");
  const majorTick = select("majorTick", "坐标轴主刻度线", [keep, reset, ["none", "无"], ["in", "内部"], ["out", "外部"], ["cross", "交叉"]]);
  const minorTick = select("minorTick", "坐标轴次刻度线", [keep, reset, ["none", "无"], ["in", "内部"], ["out", "外部"], ["cross", "交叉"]]);
  const axisLabel = select("axisLabel", "坐标轴标签位置", [keep, reset, ["nextTo", "紧靠轴"], ["low", "低端"], ["high", "高端"], ["none", "无标签"]]);
  const alignment = select("axisAlign", "分类轴标签对齐", [keep, reset, ["l", "左对齐"], ["ctr", "居中"], ["r", "右对齐"]], ["categoryAxis"]);
  const offset = numeric("axisOffset", "分类轴标签距离（0 至 1000）", ["categoryAxis"]);
  const labelSkip = numeric("labelSkip", "分类轴标签间隔（0 自动）", ["categoryAxis"]);
  const tickSkip = numeric("tickSkip", "分类轴刻度间隔（0 默认）", ["categoryAxis"]);
  const between = select("crossBetween", "分类区间交叉", [keep, reset, ["between", "分类之间"], ["midCat", "分类中间"]], ["valueAxis"]);
  const minorUnit = numeric("minorUnit", "数值轴次刻度单位（0 自动）", ["valueAxis"]);
  const formatMode = select("partFormatMode", "当前轴或标签数值格式来源", [keep, ["source", "关联源单元格"], ["custom", "自定义格式"], ["default", "恢复默认"]], [...axes, ...labels]);
  const format = register("partFormat", factory.input("当前轴或标签自定义数值格式"), [...axes, ...labels]);
  const labelPosition = select("labelPosition", "数据标签位置", [keep, reset, ["ctr", "居中"], ["inBase", "底端内部"], ["inEnd", "顶端内部"], ["outEnd", "外侧末端"], ["l", "左侧"], ["r", "右侧"], ["t", "上方"], ["b", "下方"], ["bestFit", "自动最佳位置"]], labels);
  const separatorMode = select("separatorMode", "数据标签分隔符设置", [keep, ["custom", "指定文字（可为空）"], ["default", "恢复默认"]], labels);
  const separator = register("separator", factory.input("数据标签分隔文字"), labels);
  const leaders = select("leaders", "数据标签引导线", [keep, reset, ["true", "显示"], ["false", "隐藏"]], labels);
  const leaderColor = register("leaderColor", factory.input("数据标签引导线颜色"), labels);
  const leaderWidth = numeric("leaderWidth", "数据标签引导线宽度（磅）", labels);
  const leaderKind = select("leaderKind", "数据标签引导线外观", [keep, reset, ["solid", "显示线条"], ["none", "无线条"]], labels);
  const leaderDash = select("leaderDash", "数据标签引导线线型", [keep, reset, ["solid", "实线"], ["dash", "虚线"], ["dot", "点线"], ["dashDot", "点划线"]], labels);
  const positions = (kind: WorkbookChart["kind"]) => kind === "column" ? ["ctr", "inBase", "inEnd", "outEnd"] : kind === "line" ? ["ctr", "l", "r", "t", "b"] : ["ctr", "inEnd", "outEnd", "bestFit"];
  const finite = (field: HTMLInputElement, low = -Number.MAX_VALUE, high = Number.MAX_VALUE, integer = false): number => {
    const value = Number(field.value); if (field.validity.badInput || !field.value.trim() || !Number.isFinite(value) || value < low || value > high || integer && !Number.isInteger(value)) throw new DocumentError("InvalidOperation", "请输入有效的坐标轴或标签数值。"); return value;
  };
  const quote = (value: string) => { if (value.length > 8192 || /[\u0000-\u001f\u007f]/.test(value)) throw new DocumentError("InvalidOperation", "格式或分隔符超出支持的文字范围。"); return JSON.stringify(value); };
  return {
    refresh(target: Target, kind: WorkbookChart["kind"], composing?: HTMLElement): void {
      for (const item of fields) { const enabled = item.enabled(target, kind); item.field.disabled ||= !enabled && item.field !== composing; item.field.parentElement!.hidden = !enabled && item.field !== composing; }
      crossingAt.disabled ||= crossing.value !== "at" && crossingAt !== composing;
      format.disabled ||= formatMode.value !== "custom" && format !== composing;
      separator.disabled ||= separatorMode.value !== "custom" && separator !== composing;
      for (const option of labelPosition.options) option.disabled = !!option.value && option.value !== "initial" && !positions(kind).includes(option.value);
    },
    properties(target: Target, kind: WorkbookChart["kind"]): Record<string, string> {
      const result: Record<string, string> = {};
      const assign = (field: Field, property: string) => { if (field.value) result[property] = field.value; };
      const hasNumber = (field: HTMLInputElement) => !!field.value.trim() || field.validity.badInput;
      if (axes.includes(target) && kind !== "pie") {
        for (const [field, name] of [[majorTick, "major-tick"], [minorTick, "minor-tick"], [axisLabel, "label-position"]] as const) assign(field, "--office-axis-" + name);
        if (crossing.value === "at") result["--office-axis-crosses-at"] = String(finite(crossingAt, -Number.MAX_VALUE, Number.MAX_VALUE, target === "valueAxis"));
        else if (crossing.value) { result["--office-axis-crosses"] = crossing.value; result["--office-axis-crosses-at"] = "initial"; }
        if (target === "categoryAxis") {
          assign(alignment, "--office-axis-label-align");
          if (hasNumber(offset)) result["--office-axis-label-offset"] = String(finite(offset, 0, 1000, true));
          for (const [field, name] of [[labelSkip, "label-skip"], [tickSkip, "tick-skip"]] as const) if (hasNumber(field)) result["--office-axis-" + name] = String(finite(field, 0, 0x7fff_ffff, true) || "initial");
        } else {
          assign(between, "--office-axis-cross-between");
          if (hasNumber(minorUnit)) result["--office-axis-minor-unit"] = String(finite(minorUnit, 0) || "initial");
        }
      }
      if (labels.includes(target)) {
        if (labelPosition.value && labelPosition.value !== "initial" && !positions(kind).includes(labelPosition.value)) throw new DocumentError("InvalidOperation", "该图表类型不支持所选标签位置。");
        assign(labelPosition, "--office-label-position"); assign(leaders, "--office-label-show-leader-lines");
        if (leaderColor.value.trim()) result["--office-label-leader-line-color"] = leaderColor.value.trim();
        if (hasNumber(leaderWidth)) result["--office-label-leader-line-width"] = finite(leaderWidth, 0, 1584) + "pt";
        assign(leaderDash, "--office-label-leader-line-dash");
        if (leaderKind.value || leaderColor.value.trim() || hasNumber(leaderWidth) || leaderDash.value) result["--office-label-leader-line-kind"] = leaderKind.value || "solid";
        if (separatorMode.value) result["--office-label-separator"] = separatorMode.value === "custom" ? quote(separator.value) : "initial";
      }
      if ((axes.includes(target) && kind !== "pie" || labels.includes(target)) && formatMode.value) {
        const prefix = axes.includes(target) ? "--office-axis-number-format" : "--office-label-number-format";
        if (formatMode.value === "custom") {
          const code = quote(format.value), formatter = createWorkbookNumberFormatter({ dateSystem: "1900" });
          try { formatter.format(0, { code: format.value }); } finally { formatter.dispose(); }
          result[prefix] = code; result[prefix + "-source-linked"] = "false";
        } else { result[prefix + "-source-linked"] = formatMode.value === "source" ? "true" : "initial"; if (formatMode.value === "default") result[prefix] = "initial"; }
      }
      return result;
    },
  };
}
