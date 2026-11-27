import { createHash } from "node:crypto";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx } from "@doffi/office";

const TINY_PNG = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

async function openWordPlayground(
  page: Page,
  docxBytes: Uint8Array,
  name = "wave2.docx",
  mode: "preview" | "edit" = "edit",
  layout: "pages" | "continuous" = "pages",
): Promise<void> {
  await page.addInitScript(
    ({ editingMode, viewLayout }) => {
      localStorage.setItem("doffi:library", "closed");
      localStorage.setItem("doffi:editing", editingMode);
      localStorage.setItem("doffi:layout", viewLayout);
      localStorage.removeItem("doffi.word.defaultTemplate.v1");
    },
    { editingMode: mode, viewLayout: layout },
  );
  await page.goto("/");
  await page.getByLabel("打开文件", { exact: true }).setInputFiles({
    name,
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from(docxBytes),
  });
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false");
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
}

function buildWave2ShowcaseDocx(): Uint8Array {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "wave2-showcase", title: "Wave 2 完整对齐测试" }), {
      kind: "document",
    }),
  );
  const apply = (
    ops: readonly DocumentOperation[],
    assets?: { path: string; mimeType: string; bytes: Uint8Array }[],
  ) => {
    const snap = session.snapshot();
    session.commit(
      session.prepare(
        {
          protocolVersion: 1,
          documentId: snap.manifest.documentId,
          baseRevision: snap.manifest.revision,
          operations: [...ops],
        },
        assets
          ? {
              assets: assets.map((item) => ({
                descriptor: {
                  path: item.path,
                  mimeType: item.mimeType,
                  sha256: createHash("sha256").update(item.bytes).digest("hex"),
                },
                bytes: item.bytes,
              })),
            }
          : undefined,
      ),
    );
  };

  const secId = session.snapshot().metadata.sections![0]!.rootNodeId;
  const mainStory = session.snapshot().metadata.word!.stories.find((st) => st.kind === "main")!;
  const articleId = session.getNode(mainStory.rootNodeId)!.parentNodeId ?? mainStory.rootNodeId;
  const p1 = session.getNode(secId)!.childNodeIds[0]!;

  apply(
    [
      {
        op: "setWordOptions",
        nodeId: articleId,
        properties: { creator: "王明" },
        template: { name: "CorpTemplate.dotx", target: "CorpTemplate.dotx", linkStyles: true },
        bibliography: {
          style: "GB7714",
          sources: [
            {
              tag: "Zha2025",
              sourceType: "JournalArticle",
              title: "浏览器端高保真文档布局引擎",
              authors: [
                { first: "伟", last: "张" },
                { first: "娜", last: "李" },
              ],
              year: "2025",
              journalName: "软件学报",
              volume: "36",
              issue: "4",
              pages: "100-115",
            },
          ],
        },
      },
      {
        op: "setWordSection",
        nodeId: secId,
        changes: {
          pageBorders: {
            offsetFrom: "page",
            display: "allPages",
            zOrder: "front",
            top: { style: "stars", width: 18, color: "2F5496", space: 18 },
            bottom: { style: "stars", width: 18, color: "2F5496", space: 18 },
            left: { style: "stars", width: 18, color: "2F5496", space: 18 },
            right: { style: "stars", width: 18, color: "2F5496", space: 18 },
          },
        },
      },
      {
        op: "setWordNumbering",
        nodeId: articleId,
        numberingId: "12",
        name: "Wave2PicBullet",
        multiLevelType: "hybridMultilevel",
        levels: [
          {
            level: 0,
            start: 1,
            format: "bullet",
            text: "•",
            suffix: "tab",
            cssClass: "wl-12-0",
            pictureBulletAssetPath: "assets/bullet-star.png",
            pictureBulletWidthPt: 11,
            pictureBulletHeightPt: 11,
            isLgl: true,
          },
        ],
      },
      {
        op: "replaceText",
        nodeId: p1,
        from: 0,
        to: 0,
        expectedText: "",
        text: "三维渐变艺术字标题",
      },
      {
        op: "formatWordText",
        nodeId: p1,
        from: 0,
        to: 9,
        expectedText: "三维渐变艺术字标题",
        properties: {
          "font-size": "24pt",
          "font-weight": "700",
          color: "#2f5496",
          "--word-text-fill-gradient": "135deg #ffc000 0% #ed7d31 55% #c00000 100%",
          "--word-text-bevel": "circle 3pt 3pt",
          "--word-text-3d": "orthographicFront",
        },
      },
      {
        op: "insertFragment",
        nodeId: p1,
        position: "after",
        html:
          "<p>图片项目符号条目</p>" +
          "<p>引文示例： ｜条件： ｜符号： ｜中文数字： ｜模板： </p>" +
          "<p></p>" +
          "<p></p>" +
          "<p></p>" +
          "<p>双浮动环绕测试段落，用于验证同一水平高度上多个左右浮动图形并发排斥时的正文文字自动绕排与折行效果，确保不会发生重叠遮挡。</p>",
      },
    ],
    [{ path: "assets/bullet-star.png", mimeType: "image/png", bytes: TINY_PNG }],
  );

  const paras = session
    .getNode(secId)!
    .childNodeIds.filter((id) => /^p$|^h[1-6]$/.test(session.getNode(id)!.tagName));
  const picPara = paras[1]!;
  const fieldPara = paras[2]!;
  const bibPara = paras[3]!;
  const idxPara = paras[4]!;
  const toaPara = paras[5]!;
  const wrapPara = paras[6]!;

  // Insert sibling fields from right to left at distinct offsets so none nest inside another
  apply([
    { op: "setWordList", nodeId: picPara, numberingId: "12", level: 0 },
    {
      op: "insertWordField",
      nodeId: fieldPara,
      offset: 28,
      instruction: " TEMPLATE ",
      cachedText: "tpl",
    },
    {
      op: "insertWordField",
      nodeId: fieldPara,
      offset: 23,
      instruction: " = 2026 \\* CHINESENUM3 ",
      cachedText: "0",
    },
    {
      op: "insertWordField",
      nodeId: fieldPara,
      offset: 16,
      instruction: " SYMBOL 0x2605 \\u ",
      cachedText: "?",
    },
    {
      op: "insertWordField",
      nodeId: fieldPara,
      offset: 11,
      instruction: ' IF 10 > 5 "达标" "未达标" ',
      cachedText: "待计算",
    },
    {
      op: "insertWordField",
      nodeId: fieldPara,
      offset: 6,
      instruction: ' TA \\l "民法典第577条" \\s "民法典577" \\c 2 ',
      cachedText: "",
    },
    {
      op: "insertWordField",
      nodeId: fieldPara,
      offset: 6,
      instruction: ' XE "排版引擎:分页算法" ',
      cachedText: "",
    },
    {
      op: "insertWordField",
      nodeId: fieldPara,
      offset: 5,
      instruction: " CITATION Zha2025 \\l 2052 ",
      cachedText: "[?]",
    },
    {
      op: "insertWordField",
      nodeId: bibPara,
      offset: 0,
      instruction: " BIBLIOGRAPHY ",
      cachedText: "",
    },
    {
      op: "insertWordField",
      nodeId: idxPara,
      offset: 0,
      instruction: ' INDEX \\h "A" \\c "2" ',
      cachedText: "",
    },
    {
      op: "insertWordField",
      nodeId: toaPara,
      offset: 0,
      instruction: ' TOA \\h \\c "2" ',
      cachedText: "",
    },
    {
      op: "insertWordDrawing",
      nodeId: wrapPara,
      offset: 0,
      geometry: { x: 0, y: 0, width: 72, height: 48, rotation: 0, flipH: false, flipV: false },
      layout: {
        mode: "floating",
        wrap: "square",
        horizontal: "column",
        vertical: "paragraph",
        distances: { top: 4, right: 6, bottom: 4, left: 4 },
      },
      object: { kind: "shape", preset: "rect", text: "左浮动" },
    },
    {
      op: "insertWordDrawing",
      nodeId: wrapPara,
      offset: 0,
      geometry: { x: 84, y: 0, width: 72, height: 48, rotation: 0, flipH: false, flipV: false },
      layout: {
        mode: "floating",
        wrap: "square",
        horizontal: "column",
        vertical: "paragraph",
        distances: { top: 4, right: 6, bottom: 4, left: 6 },
      },
      object: { kind: "shape", preset: "roundRect", text: "邻浮动" },
    },
    {
      op: "insertWordTable",
      nodeId: wrapPara,
      position: "after",
      rows: 2,
      columns: 3,
      cells: [
        ["列A", "列B", "列C"],
        ["10", "20", "30"],
      ],
    },
  ]);

  const tableId = session.snapshot().metadata.word!.tables[0]!.nodeId;
  apply([
    {
      op: "formatWordTable",
      nodeId: tableId,
      properties: { "table-layout": "auto", width: "100%", "--word-table-width-type": "pct" },
    },
  ]);

  return exportDocx(session).bytes;
}

test("Wave 2 rendering (art borders, WordArt gradient/3D, picture bullets, citations/bibliography, XE/INDEX, TA/TOA, advanced fields, autofit colgroup, multi-float wrap) works in both preview and edit", async ({
  page,
}) => {
  const bytes = buildWave2ShowcaseDocx();

  for (const mode of ["preview", "edit"] as const) {
    await openWordPlayground(page, bytes, "wave2-showcase.docx", mode, "pages");
    const frame = page.frameLocator(
      mode === "edit" ? 'iframe[title="正文编辑"]' : 'iframe[title="文档只读预览"]',
    );
    await expect(frame.locator(".sdk-paper").first()).toBeVisible();

    const state = await frame.locator(".sdk-paper").first().evaluate((paper) => {
      const doc = paper.ownerDocument;
      const cs = (el: Element) => doc.defaultView!.getComputedStyle(el);
      const artNodes = [...doc.querySelectorAll<HTMLElement>('[data-sdk-page-border][data-sdk-page-border-style="stars"]')];
      const artTiles = artNodes.filter((n) => Boolean(n.querySelector("svg pattern"))).length;
      const topArt = artNodes.find((n) => n.dataset.sdkPageBorder === "top");
      const leftArt = artNodes.find((n) => n.dataset.sdkPageBorder === "left");
      const artCornersNonOverlapping = Boolean(
        topArt && leftArt && parseFloat(leftArt.style.top) >= parseFloat(topArt.style.top) + parseFloat(topArt.style.height) - 0.1,
      );
      const picBullet = doc.querySelector<HTMLImageElement>("img.sdk-picture-bullet");
      const wordArtEl = [...doc.querySelectorAll<HTMLElement>(".sdk-word-run-decoration")].find((el) =>
        el.textContent?.includes("三维渐变艺术字标题"),
      );
      const indexColEl = [...doc.querySelectorAll<HTMLElement>(".sdk-field-display")].find(
        (el) => cs(el).columnCount === "2",
      );
      const hasToaLeader = Boolean(doc.querySelector(".sdk-toa-entry .sdk-toa-leader"));
      const autofitColCount = doc.querySelectorAll("table > colgroup > col").length;
      const allText = doc.body.innerText;
      return {
        artTiles,
        artCornersNonOverlapping,
        hasPicBullet: Boolean(picBullet && picBullet.getAttribute("src")),
        wordArtGradient: wordArtEl ? cs(wordArtEl).backgroundImage : "",
        wordArtFilter: wordArtEl ? cs(wordArtEl).filter : "",
        indexColumns: indexColEl ? cs(indexColEl).columnCount : "",
        hasToaLeader,
        autofitColCount,
        hasCitation: allText.includes("[1]"),
        hasBibliography: allText.includes("浏览器端高保真文档布局引擎") && allText.includes("[J]"),
        hasIndexEntry: allText.includes("排版引擎") && allText.includes("分页算法"),
        hasToaEntry: allText.includes("民法典第577条") && allText.includes("法规"),
        hasIfTrue: allText.includes("达标"),
        hasSymbolStar: allText.includes("★"),
        hasChineseNum: allText.includes("二千〇二十六"),
        hasTemplate: allText.includes("CorpTemplate.dotx"),
      };
    });

    expect(state.artTiles).toBeGreaterThanOrEqual(4);
    expect(state.artCornersNonOverlapping).toBe(true);
    expect(state.hasPicBullet).toBe(true);
    expect(state.wordArtGradient).toContain("linear-gradient");
    expect(state.wordArtFilter).toContain("drop-shadow");
    expect(state.indexColumns).toBe("2");
    expect(state.hasToaLeader).toBe(true);
    expect(state.autofitColCount).toBe(3);
    expect(state.hasCitation).toBe(true);
    expect(state.hasBibliography).toBe(true);
    expect(state.hasIndexEntry).toBe(true);
    expect(state.hasToaEntry).toBe(true);
    expect(state.hasIfTrue).toBe(true);
    expect(state.hasSymbolStar).toBe(true);
    expect(state.hasChineseNum).toBe(true);
    expect(state.hasTemplate).toBe(true);
  }
});

test("Wave 2 editing workflows: grouped font menu, reversible live hover preview, Paste Special, Cover Page, Quick Parts, WordArt, Academic dialogs, Templates, Tabbed dialogs, Draw Table and Alt Key Tips", async ({
  page,
}) => {
  const bytes = buildWave2ShowcaseDocx();
  await openWordPlayground(page, bytes, "wave2-ui.docx", "edit", "pages");
  const frame = page.frameLocator('iframe[title="正文编辑"]');
  await expect(frame.locator(".sdk-paper").first()).toBeVisible();

  // Place caret in paragraph 2
  const targetPara = frame.locator('[data-office-story="main"] p').nth(2);
  await targetPara.click();

  // 1. Grouped Font Menu & Reversible Live Hover Preview (must NOT increment data-revision!)
  const revBeforeHover = await page.locator(".office-app").getAttribute("data-revision");
  const fontTrigger = page.getByLabel("展开字体", { exact: true });
  await fontTrigger.click();
  await expect(page.getByText("主题字体", { exact: true })).toBeVisible();
  await expect(page.getByText("常用中文字体", { exact: true })).toBeVisible();
  const kaiOption = page.getByRole("option", { name: "楷体", exact: true });
  await kaiOption.hover();
  expect(await page.locator(".office-app").getAttribute("data-revision")).toBe(revBeforeHover);
  await page.keyboard.press("Escape");
  expect(await page.locator(".office-app").getAttribute("data-revision")).toBe(revBeforeHover);

  // 2. Paste Special dialog from Paste menu
  await page.getByRole("button", { name: "粘贴选项", exact: true }).click();
  await page.getByRole("menuitem", { name: "选择性粘贴…", exact: true }).click();
  const pasteDialog = page.getByRole("dialog", { name: "选择性粘贴", exact: true });
  await expect(pasteDialog).toBeVisible();
  await pasteDialog.getByRole("button", { name: "取消", exact: true }).click();
  await expect(pasteDialog).toBeHidden();

  // 3. Paragraph dialog -> 制表位… button opens Tab Stops dialog
  await page.getByRole("button", { name: "行和段落间距", exact: true }).click();
  await page.getByRole("menuitem", { name: "行距选项…", exact: true }).click();
  const paraDialog = page.getByRole("dialog", { name: "段落", exact: true });
  await expect(paraDialog).toBeVisible();
  await paraDialog.getByRole("button", { name: "制表位…", exact: true }).click();
  const tabsDialog = page.getByRole("dialog", { name: /制表位/ });
  await expect(tabsDialog).toBeVisible();
  await tabsDialog.getByRole("button", { name: /关闭|取消/ }).first().click();

  // 4. Insert tab -> Cover Page insert + remove + Undo, Quick Parts, WordArt, Insert Field
  await page.getByRole("tab", { name: "插入", exact: true }).click();
  await page.getByRole("button", { name: "封面", exact: true }).click();
  const coverDialog = page.getByRole("dialog", { name: "内置封面页库", exact: true });
  await expect(coverDialog).toBeVisible();
  await coverDialog.getByLabel("封面文档标题", { exact: true }).fill("2026 年度技术报告");
  await coverDialog.getByRole("button", { name: "插入封面", exact: true }).click();
  await expect(frame.locator('[data-office-story="main"]')).toContainText("2026 年度技术报告");

  // Remove cover page & undo
  await page.getByRole("button", { name: "封面", exact: true }).click();
  await page.getByRole("dialog", { name: "内置封面页库", exact: true }).getByRole("button", { name: "删除当前封面", exact: true }).click();
  await expect(frame.locator('[data-office-story="main"]')).not.toContainText("2026 年度技术报告");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(frame.locator('[data-office-story="main"]')).toContainText("2026 年度技术报告");

  // Quick Parts insert
  await frame.locator('[data-office-story="main"] p').last().click();
  await page.getByRole("button", { name: "文档部件", exact: true }).click();
  const qpDialog = page.getByRole("dialog", { name: "文档部件与自动图文集", exact: true });
  await expect(qpDialog).toBeVisible();
  await qpDialog.getByRole("button", { name: "插入所选部件", exact: true }).click();
  await expect(frame.locator('[data-office-story="main"]')).toContainText("敬礼");

  // WordArt gallery insert
  await page.getByRole("button", { name: "艺术字", exact: true }).click();
  const waDialog = page.getByRole("dialog", { name: "插入艺术字 (WordArt)", exact: true });
  await expect(waDialog).toBeVisible();
  await waDialog.getByLabel("艺术字文本", { exact: true }).fill("全新艺术字");
  await waDialog.getByRole("button", { name: "确定", exact: true }).click();
  await expect(frame.locator('[data-office-story="main"]')).toContainText("全新艺术字");

  // 5. References tab -> Manage Sources, Insert Citation, Mark Index Entry, Mark TOA
  await page.getByRole("tab", { name: "引用", exact: true }).click();
  await page.getByRole("button", { name: "引文", exact: true }).click();
  const srcDialog = page.getByRole("dialog", { name: "管理源（引文管理器）", exact: true });
  await expect(srcDialog).toBeVisible();
  await srcDialog.getByLabel("引文标记", { exact: true }).fill("Lin2026");
  await srcDialog.getByLabel("文献标题", { exact: true }).fill("现代办公软件架构");
  await srcDialog.getByRole("button", { name: "保存引文源", exact: true }).click();
  await srcDialog.getByRole("button", { name: "应用到文档", exact: true }).click();
  await expect(srcDialog).toBeHidden();

  // 6. Design tab -> Document Template dialog & Art Page Borders up to 31pt
  await page.getByRole("tab", { name: "设计", exact: true }).click();
  await page.getByRole("button", { name: "文档模板", exact: true }).click();
  const tplDialog = page.getByRole("dialog", { name: "文档模板和加载项", exact: true });
  await expect(tplDialog).toBeVisible();
  await tplDialog.getByRole("button", { name: "将当前主题与页面设为默认模板", exact: true }).click();
  await tplDialog.getByRole("button", { name: "确定", exact: true }).click();
  await expect(tplDialog).toBeHidden();

  // 7. Table Properties tabs + Draw Table mode toggle & Escape exit
  await frame.locator("table td").first().click();
  await page.getByRole("tab", { name: "表格布局", exact: true }).click();
  const drawBtn = page.getByRole("button", { name: "绘制表格", exact: true });
  await drawBtn.click();
  await expect(drawBtn).toHaveAttribute("aria-pressed", "true");
  await frame.locator('[data-office-story="main"]').press("Escape");
  await expect(drawBtn).toHaveAttribute("aria-pressed", "false");

  await page.getByRole("button", { name: "属性", exact: true }).click();
  const tableRegion = page.getByRole("region", { name: "表格设置", exact: true });
  await expect(tableRegion).toBeVisible();
  const cellTab = tableRegion.getByRole("tab", { name: "单元格", exact: true });
  await expect(cellTab).toBeVisible();
  await cellTab.click();
  await expect(tableRegion.getByLabel("单元格垂直对齐", { exact: true })).toBeVisible();
  await expect(tableRegion.getByLabel("表格布局算法", { exact: true })).toBeHidden();
  await tableRegion.getByRole("button", { name: "关闭表格设置", exact: true }).click();

  // 8. Word Alt Key Tips badges + keyboard tab navigation
  await page.getByRole("tab", { name: "开始", exact: true }).click();
  await page.keyboard.press("Alt");
  await expect(page.locator(".office-keytip").first()).toBeVisible();
  await page.keyboard.press("n");
  await expect(page.getByRole("tab", { name: "插入", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.locator(".office-keytip")).toHaveCount(0);
});
