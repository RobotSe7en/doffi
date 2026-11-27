import { expect, test } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import {
  compareWordDocuments,
  createDocumentSession,
  createWordDocument,
  findWordText,
  openDocument,
  wordTextReplacementOperations,
} from "@doffi/core";
import type { ContentDocumentModel, DocumentEditingSession, DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { formatWordDatePicture } from "@doffi/render";

function applyOps(
  session: DocumentEditingSession,
  operations: readonly DocumentOperation[],
) {
  const snap = session.snapshot();
  return session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: snap.manifest.documentId,
      baseRevision: snap.manifest.revision,
      operations: [...operations],
    }),
  );
}

function mainParagraphs(model: ContentDocumentModel) {
  const sections = model.snapshot().metadata.sections ?? [];
  const out: string[] = [];
  for (const sec of sections) {
    for (const id of model.getNode(sec.rootNodeId)?.childNodeIds ?? []) {
      const node = model.getNode(id);
      if (node && /^p$|^h[1-6]$/.test(node.tagName)) out.push(id);
    }
  }
  return out;
}

function rootId(model: ContentDocumentModel): string {
  const main = model.snapshot().metadata.word!.stories.find((s) => s.kind === "main")!;
  return model.getNode(main.rootNodeId)!.parentNodeId ?? main.rootNodeId;
}

test("page background, vertical alignment, line numbers, theme, document properties and protection roundtrip strictly twice with undo", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "word-parity-meta", title: "初始标题" }), {
      kind: "document",
    }),
  );
  const articleId = rootId(session);
  const sectionId = session.snapshot().metadata.sections![0]!.rootNodeId;
  const before = session.snapshot();

  applyOps(session, [
    {
      op: "setWordOptions",
      nodeId: articleId,
      title: "季度总结报告",
      pageBackground: { color: "FFF8E7" },
      protection: { edit: "trackedChanges", enforcement: true },
      automaticHyphenation: true,
      properties: {
        creator: "王明",
        subject: "年度规划",
        keywords: "Word, 排版, 对比",
        category: "内部文档",
        description: "1:1 Word 功能对齐测试文档",
      },
      theme: {
        name: "东方典雅",
        colorSchemeName: "暖橙",
        colors: {
          dk1: "1C1917",
          lt1: "FFFFFF",
          dk2: "57534E",
          lt2: "F5F5F4",
          accent1: "D95D39",
          accent2: "F0A202",
          accent3: "C44536",
          accent4: "E09F3E",
          accent5: "9E2A2B",
          accent6: "774936",
          hlink: "B45309",
          folHlink: "7C2D12",
        },
        fontSchemeName: "经典公文",
        majorLatin: "Times New Roman",
        minorLatin: "Times New Roman",
        majorEastAsia: "黑体",
        minorEastAsia: "宋体",
      },
    },
    {
      op: "setWordSection",
      nodeId: sectionId,
      changes: {
        verticalAlign: "center",
        lineNumbers: { countBy: 2, start: 1, distance: 18, restart: "newPage" },
      },
    },
  ]);

  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const parts = unzipSync(bytes);
    const docXml = strFromU8(parts["word/document.xml"]!);
    const settingsXml = strFromU8(parts["word/settings.xml"]!);
    const coreXml = strFromU8(parts["docProps/core.xml"]!);
    const themeXml = strFromU8(parts["word/theme/theme1.xml"]!);

    expect(docXml).toContain('<w:background w:color="FFF8E7"');
    expect(docXml).toMatch(/<w:lnNumType\b[^>]*w:countBy="2"[^>]*w:start="1"[^>]*w:restart="newPage"/);
    expect(docXml).toContain('<w:vAlign w:val="center"/>');
    expect(settingsXml).toContain('<w:displayBackgroundShape w:val="1"/>');
    expect(settingsXml).toMatch(/<w:documentProtection\b[^>]*w:edit="trackedChanges"[^>]*w:enforcement="1"/);
    expect(settingsXml).toContain('<w:autoHyphenation w:val="1"/>');
    expect(coreXml).toContain("<dc:title>季度总结报告</dc:title>");
    expect(coreXml).toContain("<dc:creator>王明</dc:creator>");
    expect(coreXml).toContain("<dc:subject>年度规划</dc:subject>");
    expect(coreXml).toContain("<cp:keywords>Word, 排版, 对比</cp:keywords>");
    expect(themeXml).toContain('val="D95D39"');
    expect(themeXml).toContain('typeface="黑体"');

    const reopened = openDocument(
      importDocx(bytes, { documentId: "reopened-meta", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    const meta = reopened.snapshot().metadata;
    expect(meta.title).toBe("季度总结报告");
    expect(meta.word?.pageBackground?.color.toUpperCase()).toBe("FFF8E7");
    expect(meta.word?.protection).toEqual({ edit: "trackedChanges", enforcement: true });
    expect(meta.word?.automaticHyphenation).toBe(true);
    expect(meta.word?.properties).toMatchObject({
      creator: "王明",
      subject: "年度规划",
      keywords: "Word, 排版, 对比",
      category: "内部文档",
    });
    expect(meta.word?.theme?.colors?.accent1?.toUpperCase()).toBe("D95D39");
    expect(meta.word?.theme?.majorEastAsia).toBe("黑体");
    expect(meta.sections![0]!.verticalAlign).toBe("center");
    expect(meta.sections![0]!.lineNumbers).toMatchObject({
      countBy: 2,
      start: 1,
      restart: "newPage",
    });
    bytes = exportDocx(reopened).bytes;
  }

  session.undo();
  expect(session.snapshot().metadata).toEqual(before.metadata);
});

test("ruby phonetic guide, drop cap, classic and w14 text effects, East Asian layout and fit text roundtrip strictly twice", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "word-parity-typo" }), { kind: "document" }),
  );
  const p1 = mainParagraphs(session)[0]!;

  applyOps(session, [
    {
      op: "replaceText",
      nodeId: p1,
      from: 0,
      to: 0,
      expectedText: "",
      text: "中文拼音指南与文字效果双行合一纵横2026调整宽度",
    },
  ]);

  applyOps(session, [
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
      from: 2,
      to: 6,
      expectedText: "拼音指南",
      rubyText: "pīn yīn zhǐ nán",
      align: "distributeSpace",
      fontSize: 6,
      raise: 11,
      languageId: "zh-CN",
    },
    {
      op: "formatWordText",
      nodeId: p1,
      from: 7,
      to: 11,
      expectedText: "文字效果",
      properties: {
        "--word-run-native-outline": "1",
        "--word-run-native-shadow": "1",
        "--word-text-glow": "6pt #4472c4",
        "--word-text-outline": "1pt #2f5496",
        "--word-text-shadow": "2pt 2pt 3pt #44546a",
        "--word-text-reflection": "1",
      },
    },
    {
      op: "formatWordText",
      nodeId: p1,
      from: 11,
      to: 15,
      expectedText: "双行合一",
      properties: { "--word-ea-combine": "round" },
    },
    {
      op: "formatWordText",
      nodeId: p1,
      from: 17,
      to: 21,
      expectedText: "2026",
      properties: { "--word-ea-vert": "compress" },
    },
    {
      op: "formatWordText",
      nodeId: p1,
      from: 21,
      to: 25,
      expectedText: "调整宽度",
      properties: { "--word-fit-text": "60pt" },
    },
  ]);

  // Base text stays pure UTF-16 without ruby text polluting paragraph offsets.
  expect(session.getNode(p1)!.text).toBe("中文拼音指南与文字效果双行合一纵横2026调整宽度");

  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const docXml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(docXml).toMatch(/<w:framePr\b[^>]*w:dropCap="drop"[^>]*w:lines="3"[^>]*w:hSpace="120"/);
    expect(docXml).toContain("<w:ruby>");
    expect(docXml).toContain("pīn yīn zhǐ nán");
    expect(docXml).toContain("<w:outline");
    expect(docXml).toContain("<w:shadow");
    expect(docXml).toContain("<w14:glow");
    expect(docXml).toContain("<w14:textOutline");
    expect(docXml).toContain("<w14:shadow");
    expect(docXml).toContain("<w14:reflection");
    expect(docXml).toMatch(/<w:eastAsianLayout\b[^>]*w:combine="1"[^>]*w:combineBrackets="round"/);
    expect(docXml).toMatch(/<w:eastAsianLayout\b[^>]*w:vert="1"[^>]*w:vertCompress="1"/);
    expect(docXml).toMatch(/<w:fitText\b[^>]*w:val="1200"/);

    const reopened = openDocument(
      importDocx(bytes, { documentId: "reopened-typo", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    const paras = mainParagraphs(reopened);
    // Native Word Drop Cap splits the initial character into its own w:framePr paragraph
    // followed by the continuation paragraph, preserving exact combined UTF-16 text.
    expect(paras.map((id) => reopened.getNode(id)!.text).join("")).toBe("中文拼音指南与文字效果双行合一纵横2026调整宽度");
    expect(reopened.getNode(paras[0]!)!.text).toBe("中");
    expect(reopened.snapshot().html).toContain('data-word-ruby="true"');
    expect(reopened.snapshot().html).toContain('data-word-ruby-text="pīn yīn zhǐ nán"');
    bytes = exportDocx(reopened).bytes;
  }

  // Test removeWordRuby + undo
  const rubyNodeId = session.getNode(p1)!.childNodeIds.find((id) => session.getNode(id)?.attributes["data-word-ruby"] === "true")!;
  applyOps(session, [{ op: "removeWordRuby", nodeId: rubyNodeId }]);
  expect(session.getNode(rubyNodeId)?.attributes["data-word-ruby"]).toBeUndefined();
  session.undo();
  expect(session.getNode(rubyNodeId)?.attributes["data-word-ruby"]).toBe("true");
});

test("table diagonal borders, splitWordTable, rowspan split rejection and undo roundtrip strictly", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "word-parity-table" }), { kind: "document" }),
  );
  const p1 = mainParagraphs(session)[0]!;
  applyOps(session, [
    {
      op: "insertWordTable",
      nodeId: p1,
      position: "after",
      rows: 3,
      columns: 2,
      cells: [
        ["项目", "金额"],
        ["甲", "120"],
        ["乙", "380"],
      ],
    },
  ]);

  const tableId = session.snapshot().metadata.word!.tables[0]!.nodeId;
  const tbodyId = session.getNode(tableId)!.childNodeIds[0]!;
  const firstRowId = session.getNode(tbodyId)!.childNodeIds[0]!;
  const topLeftCellId = session.getNode(firstRowId)!.childNodeIds[0]!;

  applyOps(session, [
    {
      op: "formatWordCell",
      nodeId: topLeftCellId,
      properties: {
        "--word-border-tl2br": "1pt solid #185abd",
        "--word-border-tr2bl": "0.5pt dashed #aa0000",
      },
    },
  ]);

  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const docXml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(docXml).toMatch(/<w:tl2br\b[^>]*w:color="185ABD"/i);
    expect(docXml).toMatch(/<w:tr2bl\b[^>]*w:color="AA0000"/i);
    const reopened = openDocument(
      importDocx(bytes, { documentId: "reopened-diag", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(reopened.snapshot().html).toContain("--word-border-tl2br");
    bytes = exportDocx(reopened).bytes;
  }

  // Split table at row 2 -> produces two independent tables with preserved column grids
  applyOps(session, [{ op: "splitWordTable", nodeId: tableId, atRow: 2 }]);
  expect(session.snapshot().metadata.word!.tables).toHaveLength(2);
  const splitBytes = exportDocx(session).bytes;
  const reopenedSplit = openDocument(
    importDocx(splitBytes, { documentId: "reopened-split", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  expect(reopenedSplit.snapshot().metadata.word!.tables).toHaveLength(2);

  session.undo();
  expect(session.snapshot().metadata.word!.tables).toHaveLength(1);

  // Merge vertically across rows 0..1 and confirm splitting across rowspan is rejected atomically
  applyOps(session, [
    {
      op: "editWordTable",
      nodeId: tableId,
      change: { kind: "mergeCells", row: 0, column: 0, rows: 2, columns: 1 },
    },
  ]);
  expect(() => applyOps(session, [{ op: "splitWordTable", nodeId: tableId, atRow: 1 }])).toThrow(
    /vertically merged/,
  );
});

test("wildcard and special-character findWordText and replacements work accurately", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "word-parity-find" }), { kind: "document" }),
  );
  const p1 = mainParagraphs(session)[0]!;
  applyOps(session, [
    {
      op: "replaceText",
      nodeId: p1,
      from: 0,
      to: 0,
      expectedText: "",
      text: "Order #102\tStatus: OK  Order #205\tStatus: DONE",
    },
  ]);

  const tabMatches = findWordText(session, "^tStatus:", { specialChars: true });
  expect(tabMatches).toHaveLength(2);

  const wildcardMatches = findWordText(session, "Order #[0-9][0-9][0-9]", { wildcards: true });
  expect(wildcardMatches.map((m) => m.expectedText)).toEqual(["Order #102", "Order #205"]);

  const ops = wordTextReplacementOperations(session, tabMatches, " -> ", { specialChars: true });
  applyOps(session, ops);
  expect(session.getNode(p1)!.text).toBe("Order #102 ->  OK  Order #205 ->  DONE");
});

test("compareWordDocuments generates validated native tracked revisions that export and strictly reopen", () => {
  const s1 = createDocumentSession(
    openDocument(createWordDocument({ documentId: "orig", title: "合同草案" }), { kind: "document" }),
  );
  const p1 = mainParagraphs(s1)[0]!;
  applyOps(s1, [
    { op: "replaceText", nodeId: p1, from: 0, to: 0, expectedText: "", text: "第一条 交付周期为三十天。" },
    { op: "insertFragment", nodeId: p1, position: "after", html: "<p>第二条 违约金为总额的百分之五。</p>" },
  ]);

  const s2 = createDocumentSession(
    openDocument(createWordDocument({ documentId: "rev", title: "合同修订版" }), { kind: "document" }),
  );
  const p2 = mainParagraphs(s2)[0]!;
  applyOps(s2, [
    { op: "replaceText", nodeId: p2, from: 0, to: 0, expectedText: "", text: "第一条 交付周期为十五天。" },
    { op: "insertFragment", nodeId: p2, position: "after", html: "<p>第二条 违约金为总额的百分之五。</p><p>第三条 双方友好协商。</p>" },
  ]);

  const comparedPkg = compareWordDocuments(s1, s2, { author: "法务审阅", title: "合同对比稿" });
  const compared = openDocument(comparedPkg, { kind: "document" });
  expect(compared.snapshot().metadata.title).toBe("合同对比稿");
  expect(compared.snapshot().metadata.word!.trackChanges).toBe(true);
  expect(compared.snapshot().metadata.word!.revisions.some((r) => r.kind === "insert")).toBe(true);
  expect(compared.snapshot().metadata.word!.revisions.some((r) => r.kind === "delete")).toBe(true);
  expect(compared.snapshot().metadata.word!.revisions.every((r) => r.author === "法务审阅")).toBe(true);

  const docxBytes = exportDocx(compared).bytes;
  const docXml = strFromU8(unzipSync(docxBytes)["word/document.xml"]!);
  expect(docXml).toContain("<w:ins ");
  expect(docXml).toContain("<w:del ");
  const reopened = openDocument(
    importDocx(docxBytes, { documentId: "reopened-compare", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  expect(reopened.snapshot().metadata.word!.revisions.length).toBeGreaterThan(0);
});

test("formatWordDatePicture formats Chinese and ISO DATE/TIME field switches accurately", () => {
  const sampleDate = new Date(2026, 10, 27, 9, 5, 8); // 2026-11-27 09:05:08
  expect(formatWordDatePicture(sampleDate, 'DATE \\@ "yyyy年M月d日"', false)).toBe("2026年11月27日");
  expect(formatWordDatePicture(sampleDate, 'DATE \\@ "yyyy-MM-dd"', false)).toBe("2026-11-27");
  expect(formatWordDatePicture(sampleDate, 'TIME \\@ "HH:mm:ss"', true)).toBe("09:05:08");
});
