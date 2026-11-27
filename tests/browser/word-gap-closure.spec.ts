import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import { exportDocx } from "@doffi/office";
import { body, mutateDocx, styles } from "../../packages/office/tests/regression-fixtures.js";

const withW14 = (xml: string): string =>
  xml.includes("xmlns:w14=")
    ? xml
    : xml.replace(/<(w:(?:document|styles))\b/, '<$1 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');

async function openWordPlayground(
  page: Page,
  docxBytes?: Uint8Array,
  name = "test.docx",
  mode: "preview" | "edit" = "edit",
  layout: "pages" | "continuous" = "pages",
): Promise<void> {
  await page.addInitScript(
    ({ editingMode, viewLayout }) => {
      localStorage.setItem("doffi:library", "closed");
      localStorage.setItem("doffi:editing", editingMode);
      localStorage.setItem("doffi:layout", viewLayout);
    },
    { editingMode: mode, viewLayout: layout },
  );
  await page.goto("/");
  if (docxBytes) {
    await page.getByLabel("打开文件", { exact: true }).setInputFiles({
      name,
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      buffer: Buffer.from(docxBytes),
    });
  } else {
    await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  }
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false");
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
}

function richTypographyDocx(): Uint8Array {
  return mutateDocx("01-text-numbering.docx", {
    "word/_rels/document.xml.rels": (source) =>
      source.replace(
        "</Relationships>",
        '<Relationship Id="rIdThemeTest" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>',
      ),
    "word/theme/theme1.xml": () =>
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="CustomTheme">' +
      '<a:themeElements><a:clrScheme name="Custom">' +
      '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="123456"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>' +
      '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>' +
      '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>' +
      '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>' +
      '</a:clrScheme><a:fontScheme name="Custom"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Custom"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>',
    "word/styles.xml": () =>
      styles(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:sz w:val="24"/></w:rPr></w:style>',
      ),
    "word/document.xml": (source) =>
      withW14(
        body(
          source,
          "<w:p>" +
            '<w:r><w:rPr><w:dstrike/><w:u w:val="single" w:color="CC0000"/></w:rPr><w:t>DoubleStrike</w:t></w:r>' +
            '<w:r><w:rPr><w:vanish/></w:rPr><w:t>HiddenSecret</w:t></w:r>' +
            '<w:r><w:rPr><w:position w:val="12"/></w:rPr><w:t>Raised</w:t></w:r>' +
            '<w:r><w:t>Base</w:t></w:r>' +
            '<w:r><w:rPr><w:em w:val="dot"/></w:rPr><w:t>着重点</w:t></w:r>' +
            '<w:r><w:rPr><w:bdr w:val="single" w:sz="6" w:space="0" w:color="224466"/></w:rPr><w:t>Boxed</w:t></w:r>' +
            '<w:r><w:rPr><w:u w:val="words"/></w:rPr><w:t>Alpha Beta</w:t></w:r>' +
            '<w:r><w:rPr><w14:ligatures w14:val="standardContextual"/><w14:numForm w14:val="oldStyle"/></w:rPr><w:t>fi 123</w:t></w:r>' +
            "</w:p>" +
            '<w:p><w:r><w:t>hello world. second sentence!</w:t></w:r></w:p>',
        ),
      ),
  });
}

test("WORD-03 renders double-strike+underline, hidden text, position, emphasis, character border, words-underline, and OpenType in both preview and editing", async ({ page }) => {
  const bytes = richTypographyDocx();

  for (const mode of ["preview", "edit"] as const) {
    await openWordPlayground(page, bytes, "word03-render.docx", mode, "continuous");
    const frame = page.frameLocator(mode === "edit" ? 'iframe[title="正文编辑"]' : 'iframe[title="文档只读预览"]');
    await expect(frame.locator(".sdk-word-run-decoration").first()).toBeVisible();

    const metrics = await frame.locator("p").first().evaluate((p) => {
      const findByText = (needle: string): HTMLElement | undefined => {
        const walker = p.ownerDocument.createTreeWalker(p, 4);
        for (let n = walker.nextNode(); n; n = walker.nextNode()) {
          if (n.textContent?.includes(needle)) return n.parentElement!;
        }
        return undefined;
      };
      const cs = (el: HTMLElement) => p.ownerDocument.defaultView!.getComputedStyle(el);
      const dstrikeEl = findByText("DoubleStrike")!;
      const hiddenEl = findByText("HiddenSecret")!;
      const raisedEl = findByText("Raised")!;
      const baseEl = findByText("Base")!;
      const emEl = findByText("着重点")!;
      const boxedEl = findByText("Boxed")!;
      const alphaEl = findByText("Alpha")!;
      const otEl = findByText("fi 123")!;
      const spaceAfterAlpha = alphaEl.nextElementSibling as HTMLElement | null;
      return {
        dstrikeBg: cs(dstrikeEl).backgroundImage,
        dstrikeUnderlineColor: cs(dstrikeEl).textDecorationColor,
        hiddenDisplay: cs(hiddenEl).display,
        hiddenWidth: hiddenEl.getBoundingClientRect().width,
        raisedTop: raisedEl.getBoundingClientRect().top,
        baseTop: baseEl.getBoundingClientRect().top,
        emStyle: cs(emEl).getPropertyValue("text-emphasis-style") || cs(emEl).getPropertyValue("-webkit-text-emphasis-style"),
        boxedBorder: cs(boxedEl).borderTopStyle,
        alphaDeco: cs(alphaEl).textDecorationLine,
        spaceDeco: spaceAfterAlpha ? cs(spaceAfterAlpha).textDecorationLine : "missing",
        otFeatures: cs(otEl).fontFeatureSettings,
      };
    });

    expect(metrics.dstrikeBg).toContain("linear-gradient");
    expect(metrics.dstrikeUnderlineColor).toBe("rgb(204, 0, 0)");
    expect(metrics.hiddenDisplay).toBe("none");
    expect(metrics.hiddenWidth).toBe(0);
    expect(metrics.raisedTop).toBeLessThan(metrics.baseTop - 2);
    expect(metrics.emStyle).toContain("dot");
    expect(metrics.boxedBorder).toBe("solid");
    expect(metrics.alphaDeco).toContain("underline");
    expect(metrics.spaceDeco).toBe("none");
    expect(metrics.otFeatures).toContain("liga");
    expect(metrics.otFeatures).toContain("onum");
  }
});

test("WORD-03 Font dialog, hidden marks toggle, character border, changeCase, and dynamic theme palette work end-to-end", async ({ page }) => {
  await openWordPlayground(page, richTypographyDocx(), "word03.docx");
  const frame = page.frameLocator('iframe[title="正文编辑"]');
  await expect(frame.locator("p").first()).toBeVisible();

  // Hidden text appears when Show Formatting Marks (¶) is toggled on.
  const hiddenSpan = frame.locator('[data-sdk-word-hidden="true"]').first();
  await expect(hiddenSpan).toHaveCSS("display", "none");
  await page.getByRole("button", { name: "显示编辑标记", exact: true }).click();
  await expect(hiddenSpan).toHaveCSS("display", "inline");
  await page.getByRole("button", { name: "显示编辑标记", exact: true }).click();
  await expect(hiddenSpan).toHaveCSS("display", "none");

  // Dynamic theme palette reflects document accent1 (#123456).
  await page.getByRole("button", { name: "字体颜色选项", exact: true }).click();
  await expect(page.getByRole("menuitem", { name: "#123456", exact: true })).toBeVisible();
  await page.keyboard.press("Escape");

  // Select second paragraph and test Change Case (all 5 modes) without breaking DOCX export.
  const p2 = frame.locator("p").nth(1);
  await p2.evaluate((el) => {
    const r = el.ownerDocument.createRange();
    r.selectNodeContents(el);
    const s = el.ownerDocument.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
    el.ownerDocument.dispatchEvent(new Event("selectionchange"));
  });
  await page.getByRole("button", { name: "更改大小写", exact: true }).click();
  await page.getByRole("menuitem", { name: "每个单词首字母大写", exact: true }).click();
  await expect(p2).toHaveText("Hello World. Second Sentence!");

  await p2.evaluate((el) => {
    const r = el.ownerDocument.createRange();
    r.selectNodeContents(el);
    const s = el.ownerDocument.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
    el.ownerDocument.dispatchEvent(new Event("selectionchange"));
  });
  await page.getByRole("button", { name: "更改大小写", exact: true }).click();
  await page.getByRole("menuitem", { name: "句首字母大写", exact: true }).click();
  await expect(p2).toHaveText("Hello world. Second sentence!");

  // Toggle character border on p2.
  await p2.evaluate((el) => {
    const r = el.ownerDocument.createRange();
    r.selectNodeContents(el);
    const s = el.ownerDocument.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
    el.ownerDocument.dispatchEvent(new Event("selectionchange"));
  });
  await page.getByRole("button", { name: "字符边框", exact: true }).click();
  await expect(p2.locator(".sdk-word-run-decoration").first()).toHaveCSS("border-top-style", "solid");

  // Open Font Dialog and verify new controls are enabled, commit cleanly, render visibly, and undo cleanly.
  await page.getByRole("button", { name: "展开字体", exact: true }).click();
  await page.getByRole("button", { name: "字体…", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "字体", exact: true });
  await expect(dialog).toBeVisible();
  const dstrike = dialog.getByLabel("双删除线", { exact: true });
  await expect(dstrike).toBeEnabled();
  await dstrike.check();
  await dialog.getByLabel("着重号", { exact: true }).selectOption("comma");
  await dialog.getByRole("tab", { name: "高级", exact: true }).click();
  await dialog.getByLabel("字符位置偏移", { exact: true }).selectOption("raised");
  await dialog.getByLabel("位置磅值", { exact: true }).fill("4");
  await dialog.getByLabel("连字", { exact: true }).selectOption("standard");
  await dialog.getByLabel("数字间距", { exact: true }).selectOption("tabular");
  await dialog.getByRole("button", { name: "确定", exact: true }).click();
  await expect(dialog).toBeHidden();

  const applied = await p2.locator(".sdk-word-run-decoration").first().evaluate((el) => {
    const cs = el.ownerDocument.defaultView!.getComputedStyle(el);
    return {
      bg: cs.backgroundImage,
      em: cs.getPropertyValue("text-emphasis-style") || cs.getPropertyValue("-webkit-text-emphasis-style"),
      va: cs.verticalAlign,
      ot: cs.fontFeatureSettings,
    };
  });
  expect(applied.bg).toContain("linear-gradient");
  expect(applied.em).toContain("sesame");
  expect(applied.va).not.toBe("baseline");
  expect(applied.ot).toContain("liga");
  expect(applied.ot).toContain("tnum");

  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(p2.locator(".sdk-word-run-decoration").first()).not.toHaveCSS("background-image", /linear-gradient/);
});

test("UI-01 and WORD-02 single/multiple/page-width zoom modes are mutually exclusive, and draft/focus views work cleanly", async ({ page }) => {
  await openWordPlayground(page);
  await page.getByRole("tab", { name: "视图", exact: true }).click();

  const pageWidthBtn = page.getByRole("button", { name: "页宽", exact: true });
  const singlePageBtn = page.getByRole("button", { name: "单页", exact: true });
  const multiPageBtn = page.getByRole("button", { name: "多页", exact: true });
  const zoom100Btn = page.getByRole("button", { name: "100%", exact: true });
  const zoomLabel = page.locator(".doffi-office-statusbar .zoom-number");

  // Initially none of the three zoom-fit buttons should be pressed simultaneously.
  await expect(pageWidthBtn).not.toHaveAttribute("aria-pressed", "true");
  await expect(singlePageBtn).not.toHaveAttribute("aria-pressed", "true");
  await expect(multiPageBtn).not.toHaveAttribute("aria-pressed", "true");

  await pageWidthBtn.click();
  await expect(pageWidthBtn).toHaveAttribute("aria-pressed", "true");
  await expect(singlePageBtn).toHaveAttribute("aria-pressed", "false");
  await expect(multiPageBtn).toHaveAttribute("aria-pressed", "false");
  const pwZoom = parseInt((await zoomLabel.textContent())!, 10);

  await singlePageBtn.click();
  await expect(singlePageBtn).toHaveAttribute("aria-pressed", "true");
  await expect(pageWidthBtn).toHaveAttribute("aria-pressed", "false");
  await expect(multiPageBtn).toHaveAttribute("aria-pressed", "false");
  const spZoom = parseInt((await zoomLabel.textContent())!, 10);

  await multiPageBtn.click();
  await expect(multiPageBtn).toHaveAttribute("aria-pressed", "true");
  await expect(singlePageBtn).toHaveAttribute("aria-pressed", "false");
  await expect(pageWidthBtn).toHaveAttribute("aria-pressed", "false");
  const mpZoom = parseInt((await zoomLabel.textContent())!, 10);

  expect(pwZoom).toBeGreaterThan(spZoom);
  expect(spZoom).toBeGreaterThan(mpZoom);

  await zoom100Btn.click();
  await expect(zoomLabel).toHaveText("100%");
  await expect(multiPageBtn).toHaveAttribute("aria-pressed", "false");

  // Draft View
  const draftBtn = page.getByRole("button", { name: "草稿", exact: true });
  await draftBtn.click();
  await expect(draftBtn).toHaveAttribute("aria-pressed", "true");
  await expect(page.locator(".doffi-office-ribbon").getByRole("button", { name: "页面视图", exact: true })).toHaveAttribute("aria-pressed", "false");
  await expect(page.frameLocator('iframe[title="正文编辑"]').locator("html")).toHaveAttribute("data-word-view", "draft");

  // Focus Mode + Escape exit
  await page.getByRole("button", { name: "专注", exact: true }).click();
  const exitFocus = page.getByRole("button", { name: "退出专注模式 (Esc)", exact: true });
  await expect(exitFocus).toBeVisible();
  await expect(page.locator(".office-ribbon")).toHaveCount(0);
  await page.keyboard.press("Escape");
  await expect(exitFocus).toHaveCount(0);
  await expect(page.locator(".office-ribbon")).toBeVisible();
});

test("UI-02 Selection Pane rename/hide/show/reorder, contextual Table tabs, context menu, sort, and text<->table work end-to-end", async ({ page }) => {
  // Create a synthetic DOCX with 2 shapes and 3 paragraphs for sorting / text-to-table.
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "ui-tools-doc" }), { kind: "document" }),
  );
  const sectionId = session.snapshot().metadata.sections![0]!.rootNodeId;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "ui-tools-doc",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: sectionId,
          html: "<p>香蕉,30</p><p>苹果,10</p><p>樱桃,20</p>",
        },
      ],
    }),
  );
  const firstP = session.getNode(sectionId)!.childNodeIds[0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "ui-tools-doc",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        {
          op: "insertWordDrawing",
          nodeId: firstP,
          offset: 0,
          geometry: { x: 40, y: 80, width: 80, height: 50, rotation: 0, flipH: false, flipV: false },
          layout: { mode: "floating", horizontal: "column", vertical: "paragraph", wrap: "none", zOrder: 10 },
          object: { kind: "shape", preset: "rect", text: "矩形甲" },
        },
        {
          op: "insertWordDrawing",
          nodeId: firstP,
          offset: 0,
          geometry: { x: 140, y: 80, width: 80, height: 50, rotation: 0, flipH: false, flipV: false },
          layout: { mode: "floating", horizontal: "column", vertical: "paragraph", wrap: "none", zOrder: 20 },
          object: { kind: "shape", preset: "ellipse", text: "椭圆乙" },
        },
      ],
    }),
  );

  await openWordPlayground(page, exportDocx(session).bytes, "ui-tools.docx");
  const frame = page.frameLocator('iframe[title="正文编辑"]');

  // Open Selection Pane via Layout tab.
  await page.getByRole("tab", { name: "布局", exact: true }).click();
  await page.getByRole("button", { name: "选择窗格", exact: true }).click();
  const pane = page.getByRole("region", { name: "选择窗格", exact: true });
  await expect(pane.getByRole("option")).toHaveCount(2);

  // Select first item, rename it, hide it (remains listed in Selection Pane!), then show all.
  await pane.getByRole("option").first().click();
  await pane.getByRole("button", { name: "重命名", exact: true }).click();
  const renameInput = pane.getByLabel("对象名称", { exact: true });
  await renameInput.fill("顶部徽标");
  await renameInput.press("Enter");
  await expect(pane.getByRole("option").first()).toContainText("顶部徽标");

  await pane.getByRole("button", { name: "隐藏 顶部徽标", exact: true }).click();
  await expect(pane.getByRole("option")).toHaveCount(2);
  await expect(pane.getByRole("option").first()).toContainText("顶部徽标（已隐藏）");

  await pane.getByRole("button", { name: "全部显示", exact: true }).click();
  await expect(pane.getByRole("option").first()).not.toContainText("已隐藏");

  // Reorder z-order via 下移一层.
  await pane.getByRole("option").first().click();
  await pane.getByRole("button", { name: "下移一层", exact: true }).click();
  await expect(pane.getByRole("option").nth(1)).toContainText("顶部徽标");
  await pane.getByRole("button", { name: "关闭选择窗格", exact: true }).click();

  // Sort paragraphs via Home -> 排序.
  const mainParas = frame.locator('[data-office-story="main"] section > p');
  await page.getByRole("tab", { name: "开始", exact: true }).click();
  await mainParas.first().click();
  await page.getByRole("button", { name: "排序", exact: true }).click();
  const sortDialog = page.getByRole("dialog", { name: "排序", exact: true });
  await expect(sortDialog).toBeVisible();
  await sortDialog.getByRole("button", { name: "确定", exact: true }).click();
  await expect(mainParas.first()).toContainText("苹果,10");

  // Select all 3 paragraphs and convert Text -> Table via Insert -> 表格选项 -> 将文字转换成表格….
  await mainParas.first().evaluate((first) => {
    const doc = first.ownerDocument;
    const all = doc.querySelectorAll('[data-office-story="main"] section > p');
    const r = doc.createRange();
    r.setStart(all[0]!, 0);
    r.setEnd(all[2]!, all[2]!.childNodes.length);
    const s = doc.getSelection()!;
    s.removeAllRanges();
    s.addRange(r);
    doc.dispatchEvent(new Event("selectionchange"));
  });
  await page.getByRole("tab", { name: "插入", exact: true }).click();
  await page.getByRole("button", { name: "表格选项", exact: true }).click();
  await page.getByRole("menuitem", { name: "将文字转换成表格…", exact: true }).click();
  const t2t = page.getByRole("dialog", { name: "将文字转换成表格", exact: true });
  await expect(t2t).toBeVisible();
  await t2t.getByRole("button", { name: "确定", exact: true }).click();
  await expect(frame.locator("table tr")).toHaveCount(3);

  // Clicking inside the table reveals contextual 表格设计 & 表格布局 tabs!
  await frame.locator("table td").first().click();
  await expect(page.getByRole("tab", { name: "表格设计", exact: true })).toBeVisible();
  await expect(page.getByRole("tab", { name: "表格布局", exact: true })).toBeVisible();

  // Right-click inside table cell opens Context Menu with table actions!
  await frame.locator("table td").first().click({ button: "right" });
  const menu = page.getByRole("menu", { name: "上下文菜单", exact: true });
  await expect(menu).toBeVisible();
  await menu.getByRole("menuitem", { name: "在下方插入行", exact: true }).click();
  await expect(frame.locator("table tr")).toHaveCount(4);

  // Convert Table -> Text via 表格布局 contextual tab.
  await page.getByRole("tab", { name: "表格布局", exact: true }).click();
  await page.getByRole("button", { name: "转换为文本", exact: true }).click();
  const tbl2txt = page.getByRole("dialog", { name: "表格转换成文本", exact: true });
  await tbl2txt.getByRole("button", { name: "确定", exact: true }).click();
  await expect(frame.locator("table")).toHaveCount(0);
});

test("Word paste modes menu, Insert blank page & date/time field, and Review comment prev/next/resolve/delete work end-to-end", async ({ page }) => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "review-tools-doc" }), { kind: "document" }),
  );
  const sectionId = session.snapshot().metadata.sections![0]!.rootNodeId;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "review-tools-doc",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: sectionId,
          html: "<p>第一段审阅文字</p><p>第二段审阅文字</p>",
        },
      ],
    }),
  );
  const [p1, p2] = session.getNode(sectionId)!.childNodeIds;
  const rev = session.snapshot().manifest.revision;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "review-tools-doc",
      baseRevision: rev,
      operations: [
        {
          op: "addAnnotation",
          annotation: {
            annotationId: "c1",
            documentId: "review-tools-doc",
            baseRevision: rev,
            status: "open",
            author: { name: "审阅甲" },
            instruction: "第一条批注内容",
            anchor: { kind: "text", nodeId: p1!, from: 0, to: 4, quote: "第一段审" },
          },
        },
        {
          op: "addAnnotation",
          annotation: {
            annotationId: "c2",
            documentId: "review-tools-doc",
            baseRevision: rev,
            status: "open",
            author: { name: "审阅乙" },
            instruction: "第二条批注内容",
            anchor: { kind: "text", nodeId: p2!, from: 0, to: 4, quote: "第二段审" },
          },
        },
      ],
    }),
  );

  await openWordPlayground(page, exportDocx(session).bytes, "review-tools.docx");
  const frame = page.frameLocator('iframe[title="正文编辑"]');
  await expect(frame.locator("p").first()).toBeVisible();

  // Verify split Paste dropdown exposes all three Word paste modes.
  await page.getByRole("button", { name: "粘贴选项", exact: true }).click();
  await expect(page.getByRole("menuitem", { name: "保留源格式", exact: true })).toBeVisible();
  await expect(page.getByRole("menuitem", { name: "合并格式", exact: true })).toBeVisible();
  await expect(page.getByRole("menuitem", { name: "仅保留文本", exact: true })).toBeVisible();
  await page.keyboard.press("Escape");

  // Insert Date & Time as a Word DATE field, then Insert Blank Page.
  await frame.locator("p").first().click();
  await page.getByRole("tab", { name: "插入", exact: true }).click();
  await page.getByRole("button", { name: "日期和时间", exact: true }).click();
  const dtDialog = page.getByRole("dialog", { name: "日期和时间", exact: true });
  await expect(dtDialog).toBeVisible();
  await dtDialog.getByLabel("自动更新", { exact: true }).check();
  await dtDialog.getByRole("button", { name: "确定", exact: true }).click();
  await expect(dtDialog).toBeHidden();
  await expect(frame.locator('[data-word-field="simple"]')).toHaveCount(1);

  await page.getByRole("button", { name: "空白页", exact: true }).click();
  await expect(frame.locator('br[data-word-break="page"]')).toHaveCount(2);

  // Switch to Review tab and exercise 下一条批注, 解决, 上一条批注, 删除.
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "显示批注", exact: true }).click();
  const commentPane = page.getByRole("region", { name: "文档批注", exact: true });
  await expect(commentPane).toContainText("2 条批注");

  await page.getByRole("button", { name: "下一条批注", exact: true }).click();
  await page.getByRole("button", { name: "解决", exact: true }).click();
  await expect(commentPane).toContainText("已处理");

  await page.getByRole("button", { name: "下一条批注", exact: true }).click();
  await page.getByRole("button", { name: "上一条批注", exact: true }).click();
  await page.getByRole("button", { name: "删除", exact: true }).click();
  await expect(commentPane).toContainText("1 条批注");
});
