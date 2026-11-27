import { expect, test } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import {
  createDocumentSession,
  createWordDocument,
  openDocument,
} from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { body, mutateDocx, styles } from "./regression-fixtures.js";

function mainParagraphs(session: ReturnType<typeof createDocumentSession>): string[] {
  const main = session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!;
  const queue = [main.rootNodeId];
  const paragraphs: string[] = [];
  for (let i = 0; i < queue.length; i++) {
    const node = session.getNode(queue[i]!)!;
    if (/^p$|^h[1-6]$/.test(node.tagName)) paragraphs.push(node.nodeId);
    queue.push(...node.childNodeIds);
  }
  return paragraphs;
}

const W14_NS = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const withW14 = (xml: string): string =>
  xml.includes("xmlns:w14=") ? xml : xml.replace(/<(w:(?:document|styles))\b/, `<$1 ${W14_NS}`);

test("double strikethrough, extended underlines, hidden text, position, emphasis, character borders, and OpenType w14 features survive strict import, edits, undo and two native reopens", () => {
  const docx = mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      withW14(
        styles(
          '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>' +
            '<w:style w:type="character" w:styleId="OpenTypeStyle"><w:rPr>' +
            '<w14:ligatures w14:val="standardContextual"/>' +
            '<w14:numForm w14:val="oldStyle"/>' +
            '<w14:numSpacing w14:val="tabular"/>' +
            '<w14:stylisticSets><w14:styleSet w14:id="1" w14:val="on"/><w14:styleSet w14:id="4" w14:val="on"/></w14:stylisticSets>' +
            '<w14:cntxtAlts w14:val="1"/>' +
            "</w:rPr></w:style>",
        ),
      ),
    "word/document.xml": (source) =>
      body(
        source,
        "<w:p>" +
          '<w:r><w:rPr><w:dstrike/><w:u w:val="single" w:color="CC0000"/></w:rPr><w:t>DoubleStrikeUnderline</w:t></w:r>' +
          '<w:r><w:rPr><w:vanish/></w:rPr><w:t>SecretText</w:t></w:r>' +
          '<w:r><w:rPr><w:position w:val="12"/></w:rPr><w:t>Raised6pt</w:t></w:r>' +
          '<w:r><w:rPr><w:position w:val="-8"/></w:rPr><w:t>Lowered4pt</w:t></w:r>' +
          '<w:r><w:rPr><w:em w:val="dot"/></w:rPr><w:t>着重圆点</w:t></w:r>' +
          '<w:r><w:rPr><w:em w:val="comma"/></w:rPr><w:t>着重顿号</w:t></w:r>' +
          '<w:r><w:rPr><w:em w:val="circle"/></w:rPr><w:t>着重空心圆</w:t></w:r>' +
          '<w:r><w:rPr><w:em w:val="underDot"/></w:rPr><w:t>下圆点</w:t></w:r>' +
          '<w:r><w:rPr><w:bdr w:val="single" w:sz="6" w:space="0" w:color="224466"/></w:rPr><w:t>Bordered</w:t></w:r>' +
          '<w:r><w:rPr><w:u w:val="words"/></w:rPr><w:t>Words Only</w:t></w:r>' +
          '<w:r><w:rPr><w:u w:val="thick"/></w:rPr><w:t>ThickU</w:t></w:r>' +
          '<w:r><w:rPr><w:u w:val="dotDash"/></w:rPr><w:t>DotDashU</w:t></w:r>' +
          '<w:r><w:rPr><w:u w:val="wavyDouble"/></w:rPr><w:t>WavyDoubleU</w:t></w:r>' +
          '<w:r><w:rPr><w:rStyle w:val="OpenTypeStyle"/></w:rPr><w:t>fi 1234</w:t></w:r>' +
          "</w:p>",
      ),
  });

  const session = createDocumentSession(
    openDocument(importDocx(docx, { documentId: "word-char-features", unsupportedContent: "reject" }), {
      kind: "document",
    }),
  );
  const paraId = mainParagraphs(session)[0]!;

  // Edit text inside the paragraph and verify every feature survives export + two strict reopens.
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "word-char-features",
      baseRevision: 0,
      operations: [
        { op: "replaceText", nodeId: paraId, from: 0, to: 0, expectedText: "", text: "Prefix-" },
      ],
    }),
  );

  let exported = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const zip = unzipSync(exported);
    const docXml = strFromU8(zip["word/document.xml"]!);
    const stylesXml = strFromU8(zip["word/styles.xml"]!);

    expect(docXml).toContain('<w:dstrike w:val="1"/>');
    expect(docXml.toLowerCase()).toContain('w:u w:val="single" w:color="cc0000"');
    expect(docXml).toContain('<w:vanish w:val="1"/>');
    expect(docXml).toContain('<w:position w:val="12"/>');
    expect(docXml).toContain('<w:position w:val="-8"/>');
    expect(docXml).toContain('<w:em w:val="dot"/>');
    expect(docXml).toContain('<w:em w:val="comma"/>');
    expect(docXml).toContain('<w:em w:val="circle"/>');
    expect(docXml).toContain('<w:em w:val="underDot"/>');
    expect(docXml.toLowerCase()).toContain('<w:bdr w:val="single" w:sz="6" w:space="0" w:color="224466"/>');
    expect(docXml).toContain('<w:u w:val="words"/>');
    expect(docXml).toContain('<w:u w:val="thick"/>');
    expect(docXml).toContain('<w:u w:val="dotDash"/>');
    expect(docXml).toContain('<w:u w:val="wavyDouble"/>');
    expect(stylesXml).toContain('mc:Ignorable="w14"');
    expect(stylesXml).toContain('<w14:ligatures w14:val="standardContextual"/>');
    expect(stylesXml).toContain('<w14:numForm w14:val="oldStyle"/>');
    expect(stylesXml).toContain('<w14:numSpacing w14:val="tabular"/>');
    expect(stylesXml).toContain('<w14:styleSet w14:id="1" w14:val="on"/>');
    expect(stylesXml).toContain('<w14:styleSet w14:id="4" w14:val="on"/>');
    expect(stylesXml).toContain('<w14:cntxtAlts w14:val="1"/>');

    for (const rPr of docXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/g) ?? [])
      expect((rPr.match(/<w:dstrike\b/g) ?? []).length).toBeLessThanOrEqual(1);

    const reopened = openDocument(
      importDocx(exported, { documentId: "word-char-reopen-" + pass, unsupportedContent: "reject" }),
      { kind: "document" },
    );
    exported = exportDocx(reopened).bytes;
  }

  session.undo();
  expect(session.getNode(paraId)!.text).not.toContain("Prefix-");
  expect(session.getNode(paraId)!.text).toContain("DoubleStrikeUnderline");
});

test("invalid character position fails with InvalidDocument and unknown w14 element preserves natively in compatible mode", () => {
  const invalidPosition = mutateDocx("01-text-numbering.docx", {
    "word/document.xml": (source) =>
      body(source, '<w:p><w:r><w:rPr><w:position w:val="5000"/></w:rPr><w:t>Bad</w:t></w:r></w:p>'),
  });
  expect(() => importDocx(invalidPosition, { documentId: "bad-pos" })).toThrowError(
    expect.objectContaining({ code: "InvalidDocument" }),
  );

  const mixedW14 = mutateDocx("01-text-numbering.docx", {
    "word/document.xml": (source) =>
      withW14(
        body(
          source,
          '<w:p><w:r><w:rPr><w14:ligatures w14:val="all"/><w14:unknownEffect/></w:rPr><w:t>MixedW14</w:t></w:r></w:p>',
        ),
      ),
  });
  expect(() =>
    importDocx(mixedW14, { documentId: "strict-w14", unsupportedContent: "reject" }),
  ).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));

  const compat = openDocument(importDocx(mixedW14, { documentId: "compat-w14" }), {
    kind: "document",
  });
  const roundtripXml = strFromU8(unzipSync(exportDocx(compat).bytes)["word/document.xml"]!);
  expect(roundtripXml).toMatch(/<w14:ligatures\b[^>]*w14:val="all"/);
  expect(roundtripXml).toContain("w14:unknownEffect");
});

test("selection pane rename, hide/show, and z-order reorder roundtrip cleanly to DOCX and undo", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "selection-pane-ops" }), { kind: "document" }),
  );
  const paraId = mainParagraphs(session)[0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "selection-pane-ops",
      baseRevision: 0,
      operations: [
        {
          op: "insertWordDrawing",
          nodeId: paraId,
          offset: 0,
          geometry: { x: 20, y: 20, width: 80, height: 60, rotation: 0, flipH: false, flipV: false },
          layout: { mode: "floating", horizontal: "column", vertical: "paragraph", wrap: "none", zOrder: 10 },
          object: { kind: "shape", preset: "rect", text: "A" },
        },
        {
          op: "insertWordDrawing",
          nodeId: paraId,
          offset: 0,
          geometry: { x: 40, y: 40, width: 80, height: 60, rotation: 0, flipH: false, flipV: false },
          layout: { mode: "floating", horizontal: "column", vertical: "paragraph", wrap: "none", zOrder: 20 },
          object: { kind: "shape", preset: "ellipse", text: "B" },
        },
      ],
    }),
  );

  const [first, second] = session.snapshot().metadata.word!.drawings!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "selection-pane-ops",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        { op: "setWordDrawingName", nodeId: first!.nodeId, name: "主徽标图形" },
        {
          op: "setWordDrawingLayout",
          nodeId: first!.nodeId,
          layout: {
            ...(JSON.parse(JSON.stringify(first!.layout)) as import("@doffi/core").WordDrawingLayout),
            hidden: true,
            zOrder: 30,
          },
        },
      ],
    }),
  );

  const exported = exportDocx(session).bytes;
  const docXml = strFromU8(unzipSync(exported)["word/document.xml"]!);
  expect(docXml).toContain('name="主徽标图形"');
  expect(docXml).toContain('hidden="1"');
  expect(docXml).toContain('relativeHeight="30"');

  const reopened = openDocument(
    importDocx(exported, { documentId: "selection-pane-reopen", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  const reopenedFirst = reopened
    .snapshot()
    .metadata.word!.drawings!.find((item) => item.name === "主徽标图形")!;
  expect(reopenedFirst).toBeDefined();
  expect(reopenedFirst.layout.hidden).toBe(true);
  expect(reopenedFirst.layout.zOrder).toBe(30);

  session.undo();
  const undoneFirst = session
    .snapshot()
    .metadata.word!.drawings!.find((item) => item.nodeId === first!.nodeId)!;
  expect(undoneFirst.name).toBeUndefined();
  expect(undoneFirst.layout.hidden).toBeUndefined();
  expect(undoneFirst.layout.zOrder).toBe(10);
  expect(second!.layout.zOrder).toBe(20);
});

test("reorderWordNodes sorts paragraphs and table rows atomically, rejects merged rows, and text<->table roundtrips cleanly", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "sort-and-table" }), { kind: "document" }),
  );
  const sectionId = session.snapshot().metadata.sections![0]!.rootNodeId;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "sort-and-table",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: sectionId,
          html: "<p>香蕉</p><p>苹果</p><p>樱桃</p>",
        },
      ],
    }),
  );

  const [p1, p2, p3] = mainParagraphs(session);
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "sort-and-table",
      baseRevision: session.snapshot().manifest.revision,
      operations: [{ op: "reorderWordNodes", nodeId: sectionId, childNodeIds: [p2!, p1!, p3!] }],
    }),
  );
  expect(mainParagraphs(session).map((id) => session.getNode(id)!.text)).toEqual([
    "苹果",
    "香蕉",
    "樱桃",
  ]);
  session.undo();
  expect(mainParagraphs(session).map((id) => session.getNode(id)!.text)).toEqual([
    "香蕉",
    "苹果",
    "樱桃",
  ]);

  // Text to table via insertWordTable(cells) + removeNode, then sort table rows, then export & reopen.
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "sort-and-table",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        {
          op: "insertWordTable",
          nodeId: p1!,
          position: "before",
          rows: 3,
          columns: 2,
          cells: [
            ["名称", "分数"],
            ["乙", "90"],
            ["甲", "100"],
          ],
        },
        { op: "removeNode", nodeId: p1! },
        { op: "removeNode", nodeId: p2! },
        { op: "removeNode", nodeId: p3! },
      ],
    }),
  );

  const tableId = session.snapshot().metadata.word!.tables[0]!.nodeId;
  const tableNode = session.getNode(tableId)!;
  const rowContainer =
    session.getNode(tableNode.childNodeIds[0]!)?.tagName === "tbody"
      ? tableNode.childNodeIds[0]!
      : tableId;
  const rows = session.getNode(rowContainer)!.childNodeIds;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "sort-and-table",
      baseRevision: session.snapshot().manifest.revision,
      operations: [{ op: "reorderWordNodes", nodeId: rowContainer, childNodeIds: [rows[2]!, rows[1]!] }],
    }),
  );

  const exported = exportDocx(session).bytes;
  const reopened = openDocument(
    importDocx(exported, { documentId: "sort-table-reopen", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  const reopenedTable = reopened.getNode(reopened.snapshot().metadata.word!.tables[0]!.nodeId)!;
  expect(reopenedTable.text).toContain("名称分数甲100乙90");

  // Reject sorting vertically merged rows.
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "sort-and-table",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        {
          op: "editWordTable",
          nodeId: tableId,
          change: { kind: "mergeCells", row: 1, column: 0, rows: 2, columns: 1 },
        },
      ],
    }),
  );
  const mergedRows = session.getNode(rowContainer)!.childNodeIds;
  expect(() =>
    session.prepare({
      protocolVersion: 1,
      documentId: "sort-and-table",
      baseRevision: session.snapshot().manifest.revision,
      operations: [{ op: "reorderWordNodes", nodeId: rowContainer, childNodeIds: [mergedRows[2]!, mergedRows[1]!] }],
    }),
  ).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
});

test("resolved and reopened Word comment threads roundtrip cleanly with w15:done flags", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "comment-resolve-rt" }), { kind: "document" }),
  );
  const paraId = mainParagraphs(session)[0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "comment-resolve-rt",
      baseRevision: 0,
      operations: [
        { op: "replaceText", nodeId: paraId, from: 0, to: 0, expectedText: "", text: "审阅目标文本" },
      ],
    }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "comment-resolve-rt",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        {
          op: "addAnnotation",
          annotation: {
            annotationId: "ann-1",
            documentId: "comment-resolve-rt",
            baseRevision: session.snapshot().manifest.revision,
            status: "open",
            instruction: "请核对数值",
            anchor: { kind: "text", nodeId: paraId, from: 0, to: 6, quote: "审阅目标文本" },
            author: { name: "审阅人甲" },
            date: "2026-11-26T08:00:00Z",
          },
        },
      ],
    }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "comment-resolve-rt",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        {
          op: "addWordCommentReply",
          nodeId: paraId,
          annotationId: "ann-1",
          html: "<p>已核对无误</p>",
          author: { name: "编辑人乙" },
          date: "2026-11-26T08:05:00Z",
        },
        {
          op: "updateAnnotation",
          annotationId: "ann-1",
          changes: { status: "resolved" },
        },
      ],
    }),
  );

  const resolvedDocx = exportDocx(session).bytes;
  const resolvedExt = strFromU8(unzipSync(resolvedDocx)["word/commentsExtended.xml"]!);
  expect(resolvedExt).toContain('w15:done="1"');

  const reopenedResolved = openDocument(
    importDocx(resolvedDocx, { documentId: "comment-resolved-reopen", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  expect(reopenedResolved.snapshot().annotations[0]!.status).toBe("resolved");
  expect(reopenedResolved.snapshot().annotations[0]!.replies?.length).toBe(1);

  // Reopen the comment thread and verify w15:done="0".
  const reopenSession = createDocumentSession(reopenedResolved);
  reopenSession.commit(
    reopenSession.prepare({
      protocolVersion: 1,
      documentId: "comment-resolved-reopen",
      baseRevision: 0,
      operations: [
        {
          op: "updateAnnotation",
          annotationId: reopenSession.getAnnotations()[0]!.annotationId,
          changes: { status: "open" },
        },
      ],
    }),
  );
  const reopenedDocx = exportDocx(reopenSession).bytes;
  const reopenedExt = strFromU8(unzipSync(reopenedDocx)["word/commentsExtended.xml"]!);
  expect(reopenedExt).toContain('w15:done="0"');
  const finalModel = openDocument(
    importDocx(reopenedDocx, { documentId: "comment-open-reopen", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  expect(finalModel.snapshot().annotations[0]!.status).toBe("open");
});
