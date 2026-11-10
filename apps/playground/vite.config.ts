import { fileURLToPath } from "node:url";
import { readFileSync, readdirSync } from "node:fs";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import type { Plugin } from "vite";
import { createDevelopmentConversionMiddleware } from "@doffi/node/dev";

function localConversion(): Plugin {
  const middleware = createDevelopmentConversionMiddleware({
    browserCacheDirectory: fileURLToPath(new URL("../../.playwright/", import.meta.url)),
    ...(process.env.DOFFI_CHROMIUM_EXECUTABLE ? { browserExecutable: process.env.DOFFI_CHROMIUM_EXECUTABLE } : {}),
    maxMessageBytes: 16 * 1024 * 1024,
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
  for (const name of readdirSync(new URL("../../examples/pptx/", import.meta.url)).filter(name => name.endsWith(".pptx"))) assets["pptx/"+name]=readFileSync(new URL("../../examples/pptx/"+name,import.meta.url));
  assets["xlsx/15-ten-thousand-rows.xlsx"] = readFileSync(new URL("../../examples/xlsx/15-ten-thousand-rows.xlsx", import.meta.url));
  return {
    name: "collaboration-examples",
    configureServer(server) {
      server.middlewares.use((request, response, next) => {
        const source = assets[(request.url ?? "").split("?")[0]!.replace(/^\//, "")];
        if (source === undefined) { next(); return; }
        response.setHeader("Content-Type", request.url?.split("?")[0]?.endsWith(".svg") ? "image/svg+xml" : typeof source === "string" ? "application/json; charset=utf-8" : request.url?.split("?")[0]?.endsWith(".xlsx") ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.openxmlformats-officedocument.presentationml.presentation"); response.end(source);
      });
    },
    generateBundle() {
      for (const [fileName, source] of Object.entries(assets)) this.emitFile({ type: "asset", fileName, source });
    },
  };
}

export default defineConfig({
  plugins: [react(), collaborationExamples(), localConversion()],
  publicDir: fileURLToPath(new URL("../../examples/docx/", import.meta.url)),
  build: { target: "es2022", rolldownOptions: { output: { banner: "/*! @license\n" + ["../../NOTICE", "../../packages/core/ECMA-376-GEOMETRY-LICENSE.txt"].map(path => readFileSync(new URL(path, import.meta.url), "utf8")).join("\n") + "\n*/", comments: { legal: true } } } },
});
