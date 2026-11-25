# 中文字体与本地预览环境

文档源保留原字体名称和 ASCII、高位西文、东亚、复杂文字角色。支持保存名称不代表本机已经安装该字体；SDK 不在打开文档时下载字体，也不把替代字体改名为源字体。

R1-06 回归覆盖宋体、新宋体、`SimSun/NSimSun`，仿宋、仿宋_GB2312、华文仿宋、`FangSong/FangSong_GB2312/STFangsong`，楷体、楷体_GB2312、华文楷体、`KaiTi/KaiTi_GB2312/STKaiti`，以及下列开源字体。文字替换、历史和文档包重开保留这些名称、中西文角色、UTF-16 及变体选择符。

## 已安装的补充字体

本次开发环境按用户要求安装到 `/usr/local/share/fonts/dagent-office`，Fontconfig 缓存已刷新。字体二进制不提交仓库、不随 SDK 分发；版本、下载地址及逐文件 SHA-256 记录在[字体环境](../examples/fonts/environment.json)。许可与 README 保存在 `/usr/local/share/doc/dagent-office-fonts`。

| 项目 | 版本 | 已安装字形与字重 | 许可及来源 |
| --- | --- | --- | --- |
| Fandol | 0.3 | FandolSong、FandolHei 的 Regular/Bold，FandolFang、FandolKai 的 Regular | GPL v3 加字体例外，见发行包 README/COPYING；[CTAN 项目](https://ctan.org/pkg/fandol) |
| 霞鹜文楷 / LXGW WenKai | 1.530 | Light、Regular、Medium | SIL OFL 1.1；[官方发行](https://github.com/lxgw/LxgwWenKai/releases/tag/v1.530)、[许可](https://github.com/lxgw/LxgwWenKai/blob/v1.530/OFL.txt) |

系统原有 Noto Sans CJK SC、Noto Serif CJK SC 和 Liberation Serif 继续作为 DOCX 参考环境，既有[参考记录](../examples/docx/reference.json)没有因安装新字体而刷新。

Linux Office CI 通过独立[验收脚本](../scripts/install-acceptance-fonts.mjs)将 Noto Sans／Serif CJK Regular/Bold 与 Color Emoji 固定为既有 XLSX 参考字体的实际字节。开发者显式运行时会下载清单列出的两个 Debian 字体归档；[清单](../scripts/fixtures/acceptance-fonts.json)保存版本、来源、归档与逐文件摘要及许可。字体和随附版权文件只解到工作区缓存，Fontconfig 配置仅对显式使用它的进程生效，不替换系统字体或进入 SDK 包。CI 字体版本不同会修正验收环境，不改写视觉参考或放宽摘要检查。

## 缺失字体的预览

预览先尝试文档请求的字体列表。列表中均不可用时，仿宋类名称依次尝试 FandolFang、Noto Serif CJK SC、serif；楷体类名称依次尝试 FandolKai、LXGW WenKai、Noto Serif CJK SC、serif。宋体类优先使用 Noto Serif CJK SC，再尝试 FandolSong。字体不存在时仍明确报告请求名称与实际选用名称，源 CSS 保持原样。

这些是预览替代，不承诺与微软或华文字体相同的字形、字宽或断行。已有正版字体可由宿主安装后使用；当前未从第三方下载或安装微软宋体、仿宋、楷体等专有字体。

浏览器用固定字体用例验证 FandolSong/FandolFang/FandolKai/LXGW WenKai 实际可选，以及仿宋、楷体名称的显式替代、中西文分工和缺字诊断。测量诊断仍受[DOCX 读取说明](docx-import.md#中文字体)所列边界约束。

Word 的字体替代提示只报告文档指定字体的实际替换。使用预览默认字体的文字仍参与缺字检测。

## Playground 的 Word Web 字体

用户授权的本地字体放在项目根目录的 `fonts/local/`，以 `sources.json` 登记字体名称、别名、字重、样式和摘要。Playground 启动后自动接入，Git 和 SDK 包不收录这些文件；缺少此目录时使用开源替代字体。已有本地字体无需重复转换。如需从已授权的字体目录准备完整 WOFF，可显式运行 `python3 scripts/prepare-word-test-fonts.py --font-dir <字体目录>`。字体的网页嵌入和再分发权仍由其许可决定；`.gitignore` 只阻止提交，不赋予字体使用权。参见微软的[字体再分发说明](https://learn.microsoft.com/en-us/typography/fonts/font-faq)。本地应用构建会包含已配置的字体，发布前应按实际授权配置。

打开首份 Word 文档时由宿主从同一站点加载 Web 字体，预览和编辑复用缓存字节。内部包支持宋体、新宋体、黑体、仿宋、楷体、等线、微软雅黑及 Times New Roman，并注册现有的独立字重及斜体。中文名和英文名指向同一字体时视为命中原字体。原字体和系统字体均不可用时，才使用随 Playground 提供的替代字体并记录诊断。文档源与原生导出均保留请求的字体名。

替代字体覆盖 GB 2312 常用汉字和常见中西文标点。较罕见的字形按字体列表回退，缺少独立字重及斜体时可由浏览器合成。替代字体的 SIL OFL 1.1 许可、发行来源、版本及摘要见[字体清单](../apps/playground/src/assets/word-fonts/sources.json)。

Word、PPT、Excel 预览与编辑共用 `wordWebFonts` 宿主选项。浏览器宿主先读取已授权的字节，在隔离文档里注册字体，再开始测量；新增字体完成加载后，编辑器保留选区与输入法组合状态并重新绑定视图。预览内的 CSP 继续阻止从文档内容发起字体请求。Playground 的 PDF／PNG 服务默认读取同一份本地清单；独立 SDK 的字体由宿主配置。
