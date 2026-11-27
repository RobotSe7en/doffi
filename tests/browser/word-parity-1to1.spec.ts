import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx } from "@doffi/office";

async function openWordPlayground(
  page: Page,
  docxBytes: Uint8Array,
  name = "parity.docx",
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
  await page.getByLabel("打开文件", { exact: true }).setInputFiles({
    name,
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from(docxBytes),
  });
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false");
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
}

function buildParityShowcaseDocx(): Uint8Array {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "showcase", title: "Word 1:1 对齐演示" }), {
      kind: "document",
    }),
  );
  const apply = (ops: readonly DocumentOperation[]) => {
    const snap = session.snapshot();
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snap.manifest.documentId,
        baseRevision: snap.manifest.revision,
        operations: [...ops],
      }),
    );
  };
  const secId = session.snapshot().metadata.sections![0]!.rootNodeId;
  const mainStory = session.snapshot().metadata.word!.stories.find((st) => st.kind === "main")!;
  const articleId = session.getNode(mainStory.rootNodeId)!.parentNodeId ?? mainStory.rootNodeId;
  const p1 = session.getNode(secId)!.childNodeIds[0]!;

  apply([
    {
      op: "setWordOptions",
      nodeId: articleId,
      pageBackground: { color: "FFF8E7" },
      automaticHyphenation: true,
    },
    {
      op: "setWordSection",
      nodeId: secId,
      changes: {
        verticalAlign: "center",
        lineNumbers: { countBy: 1, start: 1, distance: 18, restart: "newPage" },
      },
    },
    {
      op: "insertWordStory",
      nodeId: secId,
      kind: "header",
      variant: "default",
      html: "<p>页眉标题：季度对齐文档</p>",
    },
    {
      op: "replaceText",
      nodeId: p1,
      from: 0,
      to: 0,
      expectedText: "",
      text: "首字下沉与拼音指南以及发光效果示例段落，用于验证多行文字环绕首字下沉的视觉布局效果。",
    },
    {
      op: "setWordParagraph",
      nodeId: p1,
      properties: {
        "--word-drop-cap": "drop",
        "--word-drop-cap-lines": "3",
        "--word-drop-cap-space": "6pt",
      },
    },
    {
      op: "insertWordRuby",
      nodeId: p1,
      from: 5,
      to: 9,
      expectedText: "拼音指南",
      rubyText: "pīn yīn zhǐ nán",
    },
    {
      op: "formatWordText",
      nodeId: p1,
      from: 11,
      to: 15,
      expectedText: "发光效果",
      properties: {
        "--word-text-glow": "6pt #4472c4",
        "--word-run-native-shadow": "1",
        "--word-text-reflection": "1",
      },
    },
    {
      op: "formatWordText",
      nodeId: p1,
      from: 15,
      to: 19,
      expectedText: "示例段落",
      properties: {
        "--word-ea-combine": "round",
      },
    },
    {
      op: "insertWordField",
      nodeId: p1,
      offset: 10,
      instruction: " EQ \\o\\ac(○,甲) ",
      cachedText: "甲",
    },
    {
      op: "insertWordTable",
      nodeId: p1,
      position: "after",
      rows: 4,
      columns: 2,
      cells: [
        ["项目", "金额"],
        ["第一季度", "120"],
        ["第二季度", "380"],
        ["合计", ""],
      ],
    },
  ]);

  const tableId = session.snapshot().metadata.word!.tables[0]!.nodeId;
  const tbodyId = session.getNode(tableId)!.childNodeIds[0]!;
  const rowIds = session.getNode(tbodyId)!.childNodeIds;
  const topLeftCell = session.getNode(rowIds[0]!)!.childNodeIds[0]!;
  const sumCell = session.getNode(rowIds[3]!)!.childNodeIds[1]!;
  const sumPara = session.getNode(sumCell)!.childNodeIds[0]!;

  apply([
    {
      op: "formatWordCell",
      nodeId: topLeftCell,
      properties: { "--word-border-tl2br": "1pt solid #185abd" },
    },
    {
      op: "insertWordField",
      nodeId: sumPara,
      offset: 0,
      instruction: " =SUM(ABOVE) ",
      cachedText: "0",
    },
  ]);

  return exportDocx(session).bytes;
}

test("1:1 Word parity rendering (page color, line numbers, drop cap, ruby, enclosed, glow, diagonal border, table SUM formula) works in both preview and edit", async ({ page }) => {
  const bytes = buildParityShowcaseDocx();

  for (const mode of ["preview", "edit"] as const) {
    await openWordPlayground(page, bytes, "parity-showcase.docx", mode, "pages");
    const frame = page.frameLocator(mode === "edit" ? 'iframe[title="正文编辑"]' : 'iframe[title="文档只读预览"]');
    await expect(frame.locator(".sdk-paper").first()).toBeVisible();

    const result = await frame.locator(".sdk-paper").first().evaluate((paper) => {
      const doc = paper.ownerDocument;
      const cs = (el: Element) => doc.defaultView!.getComputedStyle(el);
      const lineNumbers = doc.querySelectorAll(".sdk-line-number").length;
      const dropCap = doc.querySelector<HTMLElement>(".sdk-word-drop-cap,[data-sdk-word-drop-cap]");
      const rubyRt = doc.querySelector(".sdk-word-ruby-rt");
      const enclosed = doc.querySelector(".sdk-word-enclosed");
      const diagonal = doc.querySelector(".sdk-word-diagonal-border");
      const tableCells = [...doc.querySelectorAll<HTMLElement>("td,th")].map((c) => c.innerText.trim());
      const dropCapFloat = dropCap ? cs(dropCap).float : "";
      const firstLetterSize = dropCap ? parseFloat(cs(dropCap).fontSize) : 0;
      const bodyPara = doc.querySelector<HTMLElement>("p:not([data-sdk-word-drop-cap])") ?? doc.querySelector<HTMLElement>("p");
      const bodySize = bodyPara ? parseFloat(cs(bodyPara).fontSize) : 0;
      const reflectionEl = [...doc.querySelectorAll<HTMLElement>(".sdk-word-run-decoration")].find((el) =>
        el.textContent?.includes("发光效果"),
      );
      const eaBrackets = [...doc.querySelectorAll<HTMLElement>(".sdk-word-ea-bracket")].map((b) => b.textContent);
      return {
        paperBg: cs(paper).backgroundColor,
        lineNumbers,
        dropCapFloat,
        bodySize,
        firstLetterSize,
        rubyText: rubyRt?.textContent ?? "",
        hasEnclosed: Boolean(enclosed),
        hasDiagonal: Boolean(diagonal),
        reflectionDisplay: reflectionEl ? cs(reflectionEl).display : "",
        eaBrackets,
        sumCellText: tableCells.at(-1) ?? "",
      };
    });

    expect(result.paperBg).toBe("rgb(255, 248, 231)");
    expect(result.lineNumbers).toBeGreaterThanOrEqual(2);
    expect(result.dropCapFloat).toBe("left");
    expect(result.bodySize).toBeLessThan(24);
    expect(result.firstLetterSize).toBeGreaterThan(result.bodySize * 1.8);
    expect(result.rubyText).toBe("pīn yīn zhǐ nán");
    expect(result.hasEnclosed).toBe(true);
    expect(result.hasDiagonal).toBe(true);
    expect(result.reflectionDisplay).toBe("inline-block");
    expect(result.eaBrackets).toEqual(["(", ")"]);
    expect(result.sumCellText).toBe("500");
  }
});

test("Word UI fixes, dialogs, contextual header/footer & table tabs, protection, outline, gridlines, immersive reader, split window and 400% zoom", async ({ page }) => {
  const bytes = buildParityShowcaseDocx();
  await openWordPlayground(page, bytes, "parity-ui.docx", "edit", "pages");
  const frame = page.frameLocator('iframe[title="正文编辑"]');
  await expect(frame.locator(".sdk-paper").first()).toBeVisible();

  // 1. Status bar zoom reaches 400%
  const zoomSlider = page.getByLabel("文档缩放", { exact: true });
  await expect(zoomSlider).toHaveAttribute("max", "400");

  // 2. Design tab -> Page Borders opens dedicated 页面边框 dialog (not paragraph borders!) & Style Sets are enabled
  await page.getByRole("tab", { name: "设计", exact: true }).click();
  await expect(page.getByRole("button", { name: "正式", exact: true })).toBeEnabled();
  await page.getByRole("button", { name: "页面边框", exact: true }).click();
  const pageBorderDialog = page.getByRole("dialog", { name: "页面边框", exact: true });
  await expect(pageBorderDialog).toBeVisible();
  await pageBorderDialog.getByLabel("页面边框设置", { exact: true }).selectOption("box");
  await pageBorderDialog.getByRole("button", { name: "确定", exact: true }).click();
  await expect(pageBorderDialog).toBeHidden();

  // 3. Double-click header in canvas enters Header & Footer contextual tab, Esc returns to main story
  const headerStory = frame.locator('[data-sdk-page-story-kind="header"]').first();
  await expect(headerStory).toBeVisible();
  await headerStory.dblclick({ force: true });
  await expect(page.getByRole("tab", { name: "页眉和页脚", exact: true })).toBeVisible();
  await page.getByRole("tab", { name: "页眉和页脚", exact: true }).click();
  await page.getByRole("button", { name: "关闭页眉和页脚", exact: true }).click();
  await expect(page.getByRole("tab", { name: "页眉和页脚", exact: true })).toHaveCount(0);

  // 4. View tab -> Gridlines, Outline Arranger, Immersive Reader, Split Window
  await page.getByRole("tab", { name: "视图", exact: true }).click();
  await page.getByRole("button", { name: "网格线", exact: true }).click();
  await expect(frame.locator(".sdk-paper").first()).toHaveAttribute("data-sdk-gridlines", "true");

  await page.getByRole("button", { name: "大纲", exact: true }).click();
  await expect(page.getByRole("region", { name: "大纲编排视图", exact: true })).toBeVisible();
  await page.getByRole("button", { name: "关闭大纲视图", exact: true }).click();

  await page.getByRole("button", { name: "沉浸式阅读器", exact: true }).click();
  const reader = page.getByRole("dialog", { name: "沉浸式阅读器", exact: true });
  await expect(reader).toBeVisible();
  await reader.getByRole("button", { name: /关闭沉浸式阅读器/ }).click();
  await expect(reader).toBeHidden();

  await page.getByRole("button", { name: "拆分", exact: true }).click();
  await expect(page.getByRole("region", { name: "拆分窗口", exact: true })).toBeVisible();
  await page.getByRole("button", { name: "取消拆分", exact: true }).click();

  // 5. Review tab -> 1-click Track Changes toggle & Protection dialog
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  const trackBtn = page.getByRole("button", { name: "修订", exact: true });
  await expect(trackBtn).toHaveAttribute("aria-pressed", "false");
  await trackBtn.click();
  await expect(trackBtn).toHaveAttribute("aria-pressed", "true");

  await page.getByRole("button", { name: "保护", exact: true }).click();
  const protDialog = page.getByRole("dialog", { name: "限制编辑 / 文档保护", exact: true });
  await expect(protDialog).toBeVisible();
  await protDialog.getByLabel("编辑限制", { exact: true }).selectOption("readOnly");
  await protDialog.getByLabel("启动强制保护", { exact: true }).check();
  await protDialog.getByRole("button", { name: "确定", exact: true }).click();
  await expect(page.getByRole("button", { name: "保护", exact: true })).toHaveAttribute("aria-pressed", "true");

  // Disable protection again cleanly
  await page.getByRole("button", { name: "保护", exact: true }).click();
  const protDialog2 = page.getByRole("dialog", { name: "限制编辑 / 文档保护", exact: true });
  await protDialog2.getByLabel("编辑限制", { exact: true }).selectOption("none");
  await protDialog2.getByRole("button", { name: "确定", exact: true }).click();
  await expect(page.getByRole("button", { name: "保护", exact: true })).toHaveAttribute("aria-pressed", "false");

  // 6. Insert tab -> Advanced Symbols dialog inserts ★
  await page.getByRole("tab", { name: "插入", exact: true }).click();
  await page.getByRole("button", { name: "高级符号", exact: true }).click();
  const symDialog = page.getByRole("dialog", { name: "符号", exact: true });
  await expect(symDialog).toBeVisible();
  await symDialog.getByRole("button", { name: "符号 ★", exact: true }).click();
  await symDialog.getByRole("button", { name: "插入", exact: true }).click();
  await expect(frame.locator('[data-office-story="main"]')).toContainText("★");
});
