# 本地字体

将已授权使用的字体及 `sources.json` 放入 `fonts/local/`。默认开发服务启动时自动读取，浏览器为 Word、PPT、Excel 与字体菜单按需加载，无需额外开关。支持 WOFF、WOFF2、TTF 和 OTF。

清单格式：

```json
{
  "fonts": [
    {
      "family": "Example Sans",
      "aliases": ["示例黑体"],
      "weight": "400",
      "style": "normal",
      "file": "example-sans.woff2"
    }
  ]
}
```

`local/` 已被 Git 忽略，不随 SDK 分发。此目录缺失时使用开源替代字体。保留字体许可和来源记录；本地应用构建含已配置字体，其发布仍须具备相应授权。
