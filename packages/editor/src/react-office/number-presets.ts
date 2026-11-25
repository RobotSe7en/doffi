import type { TextFormatState } from "../text-format-state.js";
export const officeNumberPresets = [
  ["常规", "numberGeneral"], ["数字", "numberNumber"], ["货币", "numberCurrency"], ["会计专用", "numberAccounting"],
  ["短日期", "numberShortDate"], ["长日期", "numberLongDate"], ["时间", "numberTime"], ["百分比", "numberPercent"],
  ["分数", "numberFraction"], ["科学记数", "numberScientific"], ["文本", "numberText"],
] as const;
export function officeNumberLabel(format: TextFormatState): string {
  const id = format.numberFormatBuiltinId, code = (format.numberFormatCode ?? "").replace(/"[^"]*"|\\.|\[[^\]]*\]/g, "");
  if (id === 0 || format.numberFormatCode?.toLowerCase() === "general") return "常规";
  if (id === 49 || code === "@") return "文本";
  if (id === 9 || id === 10 || code.includes("%")) return "百分比";
  if (id === 11 || id === 48 || /e[+-]0/i.test(code)) return "科学记数";
  if (id === 12 || id === 13 || /\?+\//.test(code)) return "分数";
  if (id !== undefined && id >= 14 && id <= 17 || /y|d/i.test(code)) return /年|dddd/.test(format.numberFormatCode ?? "") ? "长日期" : "短日期";
  if (id !== undefined && (id >= 18 && id <= 22 || id >= 45 && id <= 47) || /h|s/i.test(code)) return "时间";
  if (code.includes("*")) return "会计专用";
  if (/[$¥€£]/.test(code)) return "货币";
  return id !== undefined && id >= 1 && id <= 4 || code ? "数字" : "常规";
}
