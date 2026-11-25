/** Chinese Word labels and the native application's numeric menu presets. */
export const wordSizeNames: Readonly<Record<string, number>> = {
  初号: 42, 小初: 36, 一号: 26, 小一: 24, 二号: 22, 小二: 18, 三号: 16, 小三: 15,
  四号: 14, 小四: 12, 五号: 10.5, 小五: 9, 六号: 7.5, 小六: 6.5, 七号: 5.5, 八号: 5,
};
export const officeSizePresets: Readonly<Record<"document" | "presentation" | "workbook", readonly number[]>> = {
  document: [5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72],
  presentation: [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96],
  workbook: [6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72],
};
export function officeSizeLabel(kind: "document" | "presentation" | "workbook", size?: number): string {
  if (size === undefined) return "";
  return (kind === "document" ? Object.keys(wordSizeNames).find(name => wordSizeNames[name] === size) : undefined) ?? String(size);
}
