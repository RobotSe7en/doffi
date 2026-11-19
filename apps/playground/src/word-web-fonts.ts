import type { WordWebFont } from "@doffi/render";
import sans from "./assets/word-fonts/sans-regular.woff2?url";
import serif from "./assets/word-fonts/serif-regular.woff2?url";
import fangsong from "./assets/word-fonts/fangsong-regular.woff2?url";
import kai from "./assets/word-fonts/kai-regular.woff2?url";

function webFont(definition: Omit<WordWebFont, "load">, url: string): WordWebFont {
  let pending: Promise<ArrayBuffer> | undefined;
  return {
    ...definition,
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

type InternalFont = Pick<WordWebFont, "family" | "aliases" | "weight" | "style"> & { file: string };
const internalAssets = import.meta.glob<string>("./assets/word-fonts/internal/*.woff", {
  eager: true, query: "?url", import: "default",
});
const internalManifests = import.meta.glob<{ fonts: InternalFont[] }>("./assets/word-fonts/internal/sources.json", {
  eager: true, import: "default",
});
const internalFonts = Object.values(internalManifests).flatMap(({ fonts }) =>
  fonts.map(({ file, ...definition }) => {
    const url = internalAssets["./assets/word-fonts/internal/" + file];
    if (!url) throw new Error("Missing internal Word font asset: " + file);
    return webFont({ ...definition, equivalentAliases: true }, url);
  }),
);

export const wordWebFonts: readonly WordWebFont[] = [
  webFont({ family: "Doffi Serif SC", aliases: serifAliases, fallbackOnly: true }, serif),
  webFont({ family: "Doffi Sans SC", aliases: sansAliases, fallbackOnly: true }, sans),
  webFont({ family: "Doffi Fangsong SC", aliases: fangsongAliases, fallbackOnly: true }, fangsong),
  webFont({ family: "Doffi Kai SC", aliases: kaiAliases, fallbackOnly: true }, kai),
  ...internalFonts,
];
