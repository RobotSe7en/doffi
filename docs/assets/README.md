# Doffi 品牌素材

素材来自用户提供的 `Doffi_Brand_Kit_v1.zip`（v1.0，2026-11-07）。本次接入保留所用文件的原始字节、路径轮廓及配色。

- `doffi-logo-color.svg`：浅色背景字标，用于 README 与 Playground 页头。
- `doffi-logo-on-dark.svg`：深色背景字标，用于 README 的深色主题。
- `doffi-icon-color.svg`：独立图标，用于 Playground 窄屏页头。
- 浏览器图标位于 `apps/playground/src/assets/`，分别来自品牌包的 `web/favicon.svg`、`web/favicon.ico` 与 `web/apple-touch-icon.png`。

原始压缩包 SHA-256：`fb34f4a4930c6eeb66e534ee0a5b3f9169a912df64e473a24a4481fd3d320e76`。

SVG 字形均已转为路径，无需字体文件。README 通过 `picture` 选择深浅主题；Playground 通过 Vite 导入同一份字标和独立图标。
