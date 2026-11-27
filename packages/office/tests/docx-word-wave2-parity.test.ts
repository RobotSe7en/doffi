import { createHash } from "node:crypto";
import { expect, test } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  createDocumentSession,
  createWordDocument,
  formatWordBibliography,
  formatWordCitation,
  openDocument,
  WORD_ART_PAGE_BORDER_STYLES,
} from "@doffi/core";
import type {
  ContentDocumentModel,
  DocumentEditingSession,
  DocumentOperation,
  WordCitationSource,
} from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";

const TINY_PNG = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

function applyOps(
  session: DocumentEditingSession,
  operations: readonly DocumentOperation[],
  resources?: { path: string; mimeType: string; bytes: Uint8Array }[],
) {
  const snap = session.snapshot();
  return session.commit(
    session.prepare(
      {
        protocolVersion: 1,
        documentId: snap.manifest.documentId,
        baseRevision: snap.manifest.revision,
        operations: [...operations],
      },
      resources
        ? {
            assets: resources.map((item) => ({
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
}

function mainParagraphs(model: ContentDocumentModel): string[] {
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

test("bibliography sources, 5 citation styles and native customXml roundtrip strictly twice with undo", () => {
  const sources: WordCitationSource[] = [
    {
      tag: "Zha2025",
      sourceType: "JournalArticle",
      title: "大型文档排版系统的增量布局算法",
      authors: [
        { first: "伟", last: "张" },
        { first: "娜", last: "李" },
      ],
      year: "2025",
      journalName: "软件学报",
      volume: "36",
      issue: "4",
      pages: "1120-1135",
    },
    {
      tag: "Knu1984",
      sourceType: "Book",
      title: "The TeXbook",
      authors: [{ first: "Donald E.", last: "Knuth" }],
      year: "1984",
      publisher: "Addison-Wesley",
      city: "Reading, MA",
    },
  ];

  // Verify all 5 citation & bibliography styles directly
  expect(formatWordCitation(sources[0]!, "APA")).toBe("(张伟 & 李娜, 2025)");
  expect(formatWordCitation(sources[1]!, "APA")).toBe("(Knuth, 1984)");
  expect(formatWordCitation(sources[1]!, "MLA", { pages: "42" })).toBe("(Knuth 42)");
  expect(formatWordCitation(sources[0]!, "GB7714", { ordinal: 1 })).toBe("[1]");
  expect(formatWordCitation(sources[1]!, "Chicago", { pages: "12" })).toBe("(Knuth 1984, 12)");
  expect(formatWordCitation(sources[1]!, "IEEE", { ordinal: 2 })).toBe("[2]");

  const gbText = formatWordBibliography(sources, "GB7714").map((e) => e.text).join("\n");
  expect(gbText).toContain("[1]");
  expect(gbText).toContain("[J]");
  expect(gbText).toContain("[M]");
  expect(formatWordBibliography(sources, "IEEE").map((e) => e.text).join("\n")).toContain("[1]");
  expect(formatWordBibliography(sources, "APA").map((e) => e.text).join("\n")).toContain("The TeXbook");
  expect(formatWordBibliography(sources, "MLA").map((e) => e.text).join("\n")).toContain("Addison-Wesley");
  expect(formatWordBibliography(sources, "Chicago").map((e) => e.text).join("\n")).toContain("Reading, MA");

  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "wave2-bib" }), { kind: "document" }),
  );
  const articleId = rootId(session);
  const p1 = mainParagraphs(session)[0]!;
  const before = session.snapshot();

  applyOps(session, [
    {
      op: "setWordOptions",
      nodeId: articleId,
      bibliography: {
        style: "GB7714",
        sources,
      },
    },
    {
      op: "replaceText",
      nodeId: p1,
      from: 0,
      to: 0,
      expectedText: "",
      text: "参见文献 ",
    },
    {
      op: "insertWordField",
      nodeId: p1,
      offset: 5,
      instruction: " CITATION Zha2025 \\l 2052 ",
      cachedText: "[1]",
    },
    {
      op: "insertFragment",
      nodeId: p1,
      position: "after",
      html: "<p></p>",
    },
  ]);

  const p2 = mainParagraphs(session)[1]!;
  applyOps(session, [
    {
      op: "insertWordField",
      nodeId: p2,
      offset: 0,
      instruction: " BIBLIOGRAPHY ",
      cachedText: "参考文献",
    },
  ]);

  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const parts = unzipSync(bytes);
    const customXmlKey = Object.keys(parts).find(
      (k) => k.startsWith("customXml/item") && k.endsWith(".xml") && !k.includes("Props"),
    );
    expect(customXmlKey).toBeDefined();
    const bibXml = strFromU8(parts[customXmlKey!]!);
    expect(bibXml).toContain("http://schemas.openxmlformats.org/officeDocument/2006/bibliography");
    expect(bibXml).toContain("<b:Tag>Zha2025</b:Tag>");
    expect(bibXml).toContain("<b:Tag>Knu1984</b:Tag>");
    expect(bibXml).toContain("The TeXbook");
    const docRels = strFromU8(parts["word/_rels/document.xml.rels"]!);
    expect(docRels).toContain("customXml");

    const reopened = openDocument(
      importDocx(bytes, { documentId: "reopen-bib-" + pass, unsupportedContent: "reject" }),
      { kind: "document" },
    );
    const metaBib = reopened.snapshot().metadata.word?.bibliography;
    expect(metaBib?.style).toBe("GB7714");
    expect(metaBib?.sources).toHaveLength(2);
    expect(metaBib?.sources[0]?.tag).toBe("Zha2025");
    expect(metaBib?.sources[1]?.title).toBe("The TeXbook");
    bytes = exportDocx(reopened).bytes;
  }

  session.undo();
  session.undo();
  expect(session.snapshot().metadata).toEqual(before.metadata);
});

test("XE/INDEX, TA/TOA and advanced nested Word fields roundtrip strictly twice", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "wave2-fields", title: "高级域测试" }), {
      kind: "document",
    }),
  );
  const p1 = mainParagraphs(session)[0]!;
  applyOps(session, [
    {
      op: "replaceText",
      nodeId: p1,
      from: 0,
      to: 0,
      expectedText: "",
      text: "测试段落 ",
    },
    {
      op: "insertWordField",
      nodeId: p1,
      offset: 5,
      instruction: ' XE "排版引擎:分页算法" ',
      cachedText: "",
    },
    {
      op: "insertWordField",
      nodeId: p1,
      offset: 5,
      instruction: ' TA \\l "民法典第577条" \\s "民法典第577条" \\c 2 ',
      cachedText: "",
    },
    {
      op: "insertWordField",
      nodeId: p1,
      offset: 5,
      instruction: ' IF 2 > 1 "成立" "不成立" ',
      cachedText: "成立",
    },
    {
      op: "insertWordField",
      nodeId: p1,
      offset: 5,
      instruction: ' INDEX \\c "2" \\h "A" ',
      cachedText: "索引",
    },
    {
      op: "insertWordField",
      nodeId: p1,
      offset: 5,
      instruction: ' TOA \\h \\c "2" \\p ',
      cachedText: "引文目录",
    },
    {
      op: "insertWordField",
      nodeId: p1,
      offset: 5,
      instruction: " = 2026 \\* CHINESENUM3 ",
      cachedText: "二〇二六",
    },
  ]);

  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const docXml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(docXml).toContain("XE");
    expect(docXml).toContain("INDEX");
    expect(docXml).toContain("TA ");
    expect(docXml).toContain("TOA ");
    expect(docXml).toContain("CHINESENUM3");
    const reopened = openDocument(
      importDocx(bytes, { documentId: "reopen-fields-" + pass, unsupportedContent: "reject" }),
      { kind: "document" },
    );
    const instructions = reopened.snapshot().metadata.word!.fields.map((f) => f.instruction).join("\n");
    expect(instructions).toContain("CHINESENUM3");
    expect(instructions).toContain("INDEX");
    expect(instructions).toContain("TOA");
    expect(instructions).toContain("XE");
    expect(instructions).toContain("TA");
    bytes = exportDocx(reopened).bytes;
  }
});

test("picture bullets (drawing export + VML pict import) and numbering style links roundtrip strictly twice with undo", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "wave2-picbullet" }), { kind: "document" }),
  );
  const articleId = rootId(session);
  const p1 = mainParagraphs(session)[0]!;
  const before = session.snapshot();

  applyOps(
    session,
    [
      {
        op: "setWordNumbering",
        nodeId: articleId,
        numberingId: "10",
        name: "PicBulletScheme",
        multiLevelType: "hybridMultilevel",
        styleLink: "CustomListStyle",
        levels: [
          {
            level: 0,
            start: 1,
            format: "bullet",
            text: "•",
            suffix: "tab",
            cssClass: "wl-10-0",
            pictureBulletAssetPath: "assets/bullet-dot.png",
            pictureBulletWidthPt: 11,
            pictureBulletHeightPt: 11,
            pStyle: "Heading1",
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
        text: "图片项目符号段落",
      },
      {
        op: "setWordList",
        nodeId: p1,
        numberingId: "10",
        level: 0,
      },
    ],
    [{ path: "assets/bullet-dot.png", mimeType: "image/png", bytes: TINY_PNG }],
  );

  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const parts = unzipSync(bytes);
    const numXml = strFromU8(parts["word/numbering.xml"]!);
    const numRels = strFromU8(parts["word/_rels/numbering.xml.rels"]!);
    expect(numXml).toContain("<w:numPicBullet");
    expect(numXml).toContain("<w:drawing>");
    expect(numXml).toContain('<w:lvlPicBulletId w:val="0"/>');
    expect(numXml).toContain('<w:pStyle w:val="Heading1"/>');
    expect(numXml).toContain("<w:isLgl/>");
    expect(numXml).toContain('<w:name w:val="PicBulletScheme"/>');
    expect(numXml).toContain('<w:multiLevelType w:val="hybridMultilevel"/>');
    expect(numXml).toContain('<w:styleLink w:val="CustomListStyle"/>');
    expect(numRels).toContain("media/");

    const reopened = openDocument(
      importDocx(bytes, { documentId: "reopen-pic-" + pass, unsupportedContent: "reject" }),
      { kind: "document" },
    );
    const num = reopened.snapshot().metadata.word?.numbering.find(
      (n) => n.name === "PicBulletScheme",
    );
    expect(num?.name).toBe("PicBulletScheme");
    expect(num?.multiLevelType).toBe("hybridMultilevel");
    expect(num?.styleLink).toBe("CustomListStyle");
    expect(num?.levels[0]?.pictureBulletAssetPath).toMatch(/^assets\//);
    expect(num?.levels[0]?.pStyle).toBe("Heading1");
    expect(num?.levels[0]?.isLgl).toBe(true);
    bytes = exportDocx(reopened).bytes;
  }

  // Also test legacy VML <w:pict> inside <w:numPicBullet> imports cleanly and upgrades to valid <w:drawing> on export
  const parts = unzipSync(bytes);
  const relMatch = /Id="(rId\d+)"[^>]*relationships\/image/.exec(
    strFromU8(parts["word/_rels/numbering.xml.rels"]!),
  );
  expect(relMatch).toBeTruthy();
  const vmlNumXml = strFromU8(parts["word/numbering.xml"]!).replace(
    /<w:numPicBullet[\s\S]*?<\/w:numPicBullet>/,
    `<w:numPicBullet w:numPicBulletId="0"><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:10.5pt;height:10.5pt"><v:imagedata r:id="${relMatch![1]}"/></v:shape></w:pict></w:numPicBullet>`,
  );
  parts["word/numbering.xml"] = strToU8(vmlNumXml);
  const vmlDocx = zipSync(parts);
  const fromVml = openDocument(
    importDocx(vmlDocx, { documentId: "from-vml-pic", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  const vmlLevel = fromVml.snapshot().metadata.word?.numbering.find(
    (n) => n.name === "PicBulletScheme",
  )?.levels[0];
  expect(vmlLevel?.pictureBulletAssetPath).toMatch(/^assets\//);
  expect(vmlLevel?.pictureBulletWidthPt).toBeCloseTo(10.5, 1);

  session.undo();
  expect(session.snapshot().metadata).toEqual(before.metadata);
});

test("attached template, art page borders up to 31pt, WordArt gradient/3D and cover page roundtrip strictly twice", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "wave2-art-tpl" }), { kind: "document" }),
  );
  const articleId = rootId(session);
  const secId = session.snapshot().metadata.sections![0]!.rootNodeId;
  const p1 = mainParagraphs(session)[0]!;

  expect(WORD_ART_PAGE_BORDER_STYLES.length).toBeGreaterThanOrEqual(16);

  applyOps(session, [
    {
      op: "setWordOptions",
      nodeId: articleId,
      template: {
        name: "AnnualReport.dotx",
        target: "AnnualReport.dotx",
        linkStyles: true,
      },
      buildingBlocks: [
        {
          id: "bb-custom-1",
          name: "法律声明",
          gallery: "autoText",
          category: "法务",
          html: "<p>本报告版权归本公司所有。</p>",
        },
      ],
    },
    {
      op: "setWordSection",
      nodeId: secId,
      changes: {
        pageBorders: {
          offsetFrom: "page",
          display: "allPages",
          zOrder: "front",
          top: { style: "stars", width: 24, color: "2F5496", space: 24 },
          bottom: { style: "apples", width: 31, color: "C00000", space: 24 },
          left: { style: "triple", width: 6, color: "44546A", space: 24 },
          right: { style: "doubleWave", width: 6, color: "548235", space: 24 },
        },
      },
    },
    {
      op: "replaceText",
      nodeId: p1,
      from: 0,
      to: 0,
      expectedText: "",
      text: "渐变三维艺术字标题",
    },
    {
      op: "formatWordText",
      nodeId: p1,
      from: 0,
      to: 9,
      expectedText: "渐变三维艺术字标题",
      properties: {
        "font-size": "28pt",
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
      position: "before",
      html: '<p data-word-cover-page="true" style="background-color:#eef4fc;border-top:6pt solid #2f5496;font-size:28pt;font-weight:700;color:#1f3864">年度报告封面<br data-word-break="page"></p>',
    },
  ]);

  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const parts = unzipSync(bytes);
    const docXml = strFromU8(parts["word/document.xml"]!);
    const settingsXml = strFromU8(parts["word/settings.xml"]!);
    const appXml = strFromU8(parts["docProps/app.xml"]!);

    expect(settingsXml).toContain("<w:attachedTemplate ");
    expect(settingsXml).toContain('<w:linkStyles w:val="1"/>');
    expect(appXml).toContain("<Template>AnnualReport.dotx</Template>");
    expect(docXml).toMatch(/<w:top\b[^>]*w:val="stars"[^>]*w:sz="24"/);
    expect(docXml).toMatch(/<w:bottom\b[^>]*w:val="apples"[^>]*w:sz="31"/);
    expect(docXml).toContain('w:val="triple"');
    expect(docXml).toContain('w:val="doubleWave"');
    expect(docXml).toContain("<w14:textFill>");
    expect(docXml).toContain("<w14:gradFill>");
    expect(docXml).toContain("<w14:props3d>");
    expect(docXml).toContain("<w14:bevelT");
    expect(docXml).toContain("<w14:scene3d>");
    expect(docXml).toContain('<w:br w:type="page"/>');

    const reopened = openDocument(
      importDocx(bytes, { documentId: "reopen-art-" + pass, unsupportedContent: "reject" }),
      { kind: "document" },
    );
    const meta = reopened.snapshot().metadata;
    expect(meta.word?.template).toMatchObject({
      name: "AnnualReport.dotx",
      target: "AnnualReport.dotx",
      linkStyles: true,
    });
    expect(meta.sections![0]!.pageBorders?.top?.style).toBe("stars");
    expect(meta.sections![0]!.pageBorders?.bottom?.width).toBe(31);
    expect(reopened.snapshot().html).toContain("--word-text-fill-gradient");
    expect(reopened.snapshot().html).toContain("--word-text-bevel");
    expect(reopened.snapshot().html).toContain("--word-text-3d");
    bytes = exportDocx(reopened).bytes;
  }

  // Verify unsafe attachedTemplate URI is rejected in strict mode
  const unsafeParts = unzipSync(bytes);
  unsafeParts["word/_rels/settings.xml.rels"] = strToU8(
    strFromU8(unsafeParts["word/_rels/settings.xml.rels"]!).replace(
      'Target="AnnualReport.dotx"',
      'Target="javascript:alert(1)"',
    ),
  );
  expect(() =>
    importDocx(zipSync(unsafeParts), { documentId: "unsafe-tpl", unsupportedContent: "reject" }),
  ).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));

  // Remove cover page atomically and undo
  const coverId = mainParagraphs(session).find(
    (id) => session.getNode(id)?.attributes["data-word-cover-page"] === "true",
  )!;
  expect(coverId).toBeDefined();
  applyOps(session, [{ op: "removeNode", nodeId: coverId }]);
  expect(session.getNode(coverId)).toBeUndefined();
  session.undo();
  expect(session.getNode(coverId)?.text).toContain("年度报告封面");
});
