import type { WordWebFont } from "@doffi/render";
import sans from "./assets/word-fonts/sans-regular.woff2?url";
import serif from "./assets/word-fonts/serif-regular.woff2?url";
import fangsong from "./assets/word-fonts/fangsong-regular.woff2?url";
import kai from "./assets/word-fonts/kai-regular.woff2?url";

function webFont(family: string, aliases: readonly string[], url: string): WordWebFont {
  let pending: Promise<ArrayBuffer> | undefined;
  return {
    family,
    aliases,
    load() {
      return pending ??= fetch(url, {
        mode: "same-origin",
        signal: AbortSignal.timeout(20_000),
      }).then((response) => {
        if (!response.ok) throw new Error("The Playground font asset could not be loaded.");
        return response.arrayBuffer();
      }).catch((error) => {
        pending = undefined;
        throw error;
      });
    },
  };
}

const serifAliases = ["宋体", "新宋体", "SimSun", "NSimSun", "宋体-简", "Songti SC", "华文宋体", "STSong", "Noto Serif CJK SC", "Noto Serif SC", "思源宋体", "Source Han Serif SC"];
const sansAliases = ["黑体", "SimHei", "黑体-简", "Heiti SC", "STHeiti", "微软雅黑", "Microsoft YaHei", "Microsoft YaHei UI", "等线", "DengXian", "苹方-简", "PingFang SC", "Noto Sans CJK SC", "Noto Sans SC", "思源黑体", "Source Han Sans SC"];
const fangsongAliases = ["仿宋", "仿宋_GB2312", "FangSong", "FangSong_GB2312", "华文仿宋", "STFangsong", "朱雀仿宋", "Zhuque Fangsong"];
const kaiAliases = ["楷体", "楷体_GB2312", "KaiTi", "KaiTi_GB2312", "SimKai", "楷体-简", "Kaiti SC", "华文楷体", "STKaiti", "LXGW WenKai", "霞鹜文楷"];

export const wordWebFonts: readonly WordWebFont[] = [
  webFont("Doffi Serif SC", serifAliases, serif),
  webFont("Doffi Sans SC", sansAliases, sans),
  webFont("Doffi Fangsong SC", fangsongAliases, fangsong),
  webFont("Doffi Kai SC", kaiAliases, kai),
];
