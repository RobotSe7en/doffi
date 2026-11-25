import { fileURLToPath } from "node:url";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import type { Plugin } from "vite";
import type { NodeWebFont } from "@doffi/node";
import { createDevelopmentConversionMiddleware } from "@doffi/node/dev";
import { docxDiagnostics } from "./docx-diagnostics.js";

function playgroundNodeWebFonts(): readonly NodeWebFont[] {
  const asset = (relative: string): string => fileURLToPath(new URL("./src/assets/word-fonts/" + relative, import.meta.url));
  const fonts: NodeWebFont[] = [
    { family: "Doffi Serif SC", aliases: ["宋体", "新宋体", "SimSun", "NSimSun", "宋体-简", "Songti SC", "华文宋体", "STSong", "Noto Serif CJK SC", "Noto Serif SC", "思源宋体", "Source Han Serif SC"], fallbackOnly: true, path: asset("serif-regular.woff2") },
    { family: "Doffi Sans SC", aliases: ["黑体", "SimHei", "黑体-简", "Heiti SC", "STHeiti", "微软雅黑", "Microsoft YaHei", "Microsoft YaHei UI", "等线", "DengXian", "苹方-简", "PingFang SC", "Noto Sans CJK SC", "Noto Sans SC", "思源黑体", "Source Han Sans SC"], fallbackOnly: true, path: asset("sans-regular.woff2") },
    { family: "Doffi Fangsong SC", aliases: ["仿宋", "仿宋_GB2312", "FangSong", "FangSong_GB2312", "华文仿宋", "STFangsong", "朱雀仿宋", "Zhuque Fangsong"], fallbackOnly: true, path: asset("fangsong-regular.woff2") },
    { family: "Doffi Kai SC", aliases: ["楷体", "楷体_GB2312", "KaiTi", "KaiTi_GB2312", "SimKai", "楷体-简", "Kaiti SC", "华文楷体", "STKaiti", "LXGW WenKai", "霞鹜文楷"], fallbackOnly: true, path: asset("kai-regular.woff2") },
  ];
  const localAsset = (relative: string): string => fileURLToPath(new URL("../../fonts/local/" + relative, import.meta.url));
  const manifestPath = localAsset("sources.json");
  if (existsSync(manifestPath)) {
    const manifest = JSON.parse(readFileSync(manifestPath, "utf8")) as {
      fonts?: readonly { family: string; aliases?: readonly string[]; weight?: string; style?: string; file: string }[];
    };
    for (const entry of manifest.fonts ?? []) {
      const path = localAsset(entry.file);
      if (!existsSync(path)) continue;
      fonts.push({
        family: entry.family,
        aliases: entry.aliases ?? [],
        ...(entry.weight !== undefined ? { weight: entry.weight } : {}),
        ...(entry.style !== undefined ? { style: entry.style } : {}),
        equivalentAliases: true,
        path,
      });
    }
  }
  return fonts;
}

function localConversion(): Plugin {
  const middleware = createDevelopmentConversionMiddleware({
    browserCacheDirectory: fileURLToPath(new URL("../../.playwright/", import.meta.url)),
    ...(process.env.DOFFI_CHROMIUM_EXECUTABLE ? { browserExecutable: process.env.DOFFI_CHROMIUM_EXECUTABLE } : {}),
    webFonts: playgroundNodeWebFonts(),
    maxMessageBytes: 32 * 1024 * 1024,
    allowedOrigins: (process.env.DOFFI_DEV_ALLOWED_ORIGINS ?? "").split(",").map((origin) => origin.trim()).filter(Boolean),
  });
  return { name: "local-office-conversion", configureServer(server) { server.middlewares.use(middleware); }, configurePreviewServer(server) { server.middlewares.use(middleware); } };
}

/** Public synthetic inputs belong to this development host, never to SDK runtime I/O. */
function collaborationExamples(): Plugin {
  const files = ["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"];
  const assets: Record<string, string | Uint8Array> = {
    "brand/doffi-logo-color.svg": readFileSync(new URL("../../docs/assets/doffi-logo-color.svg", import.meta.url)),
    "report-demo.json": JSON.stringify(files.map((path) => [path, readFileSync(new URL("../../examples/report/" + path, import.meta.url), "utf8")])),
    "presentation-demo.json": JSON.stringify([...files,"assets/sample.png"].map(path => [path,readFileSync(new URL("../../examples/presentation/"+path,import.meta.url)).toString("base64")])),
    "presentation.proposal.json": readFileSync(new URL("../../examples/presentation.proposal.json",import.meta.url),"utf8"),
    "report.proposal.json": readFileSync(new URL("../../examples/report.proposal.json", import.meta.url), "utf8"),
  };
  for (const name of ["NOTO-LICENSE.txt", "ZHUQUE-OFL.txt", "WENKAI-OFL.txt", "sources.json"])
    assets["word-fonts/" + name] = readFileSync(new URL("./src/assets/word-fonts/" + name, import.meta.url), "utf8");
  for (const name of readdirSync(new URL("../../examples/pptx/", import.meta.url)).filter(name => name.endsWith(".pptx"))) assets["pptx/"+name]=readFileSync(new URL("../../examples/pptx/"+name,import.meta.url));
  assets["xlsx/15-ten-thousand-rows.xlsx"] = readFileSync(new URL("../../examples/xlsx/15-ten-thousand-rows.xlsx", import.meta.url));
  return {
    name: "collaboration-examples",
    configureServer(server) {
      server.middlewares.use((request, response, next) => {
        const source = assets[(request.url ?? "").split("?")[0]!.replace(/^\//, "")];
        if (source === undefined) { next(); return; }
        const path = request.url?.split("?")[0];
        response.setHeader("Content-Type", path?.endsWith(".txt") ? "text/plain; charset=utf-8" : path?.endsWith(".svg") ? "image/svg+xml" : typeof source === "string" ? "application/json; charset=utf-8" : path?.endsWith(".xlsx") ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.openxmlformats-officedocument.presentationml.presentation"); response.end(source);
      });
    },
    generateBundle() {
      for (const [fileName, source] of Object.entries(assets)) this.emitFile({ type: "asset", fileName, source });
    },
  };
}

export default defineConfig({
  plugins: [react(), collaborationExamples(), localConversion(),
    process.env.DOFFI_DEV_DOCX_DIAGNOSTICS === "1" && docxDiagnostics()],
  publicDir: fileURLToPath(new URL("../../examples/docx/", import.meta.url)),
  build: { target: "es2022", rolldownOptions: { output: { banner: "/*! @license\n" + ["../../NOTICE", "../../packages/core/ECMA-376-GEOMETRY-LICENSE.txt"].map(path => readFileSync(new URL(path, import.meta.url), "utf8")).join("\n") + "\n*/", comments: { legal: true } } } },
});
