import { expect, test } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  createAssetDescriptor,
  createDocumentSession,
  createWordDocument,
  openDocument,
} from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocxWithReport } from "@doffi/office";
function source() {
  return createDocumentSession(
    openDocument(createWordDocument({ documentId: "controls" }), { kind: "document" }),
  );
}
test("editable native rich text, checkbox, date and list controls reopen twice with semantic values and cached text", () => {
  const session = source(),
    operations: DocumentOperation[] = [
      {
        op: "insertWordControl",
        nodeId: "paragraph-1",
        control: {
          kind: "richText",
          level: "block",
          alias: "内容",
          tag: "body",
          lock: "sdtLocked",
        },
        html: '<p><span style="font-weight:bold">中文</span> test</p>',
      },
      {
        op: "insertWordControl",
        nodeId: "paragraph-1",
        control: {
          kind: "checkBox",
          level: "inline",
          checked: true,
          checkedSymbol: 0x2611,
          uncheckedSymbol: 0x2610,
        },
      },
      {
        op: "insertWordControl",
        nodeId: "paragraph-1",
        control: {
          kind: "date",
          level: "inline",
          fullDate: "2028-02-29",
          dateFormat: "yyyy年M月d日",
          locale: "zh-CN",
        },
      },
      {
        op: "insertWordControl",
        nodeId: "paragraph-1",
        control: {
          kind: "dropDown",
          level: "inline",
          items: [
            { displayText: "甲", value: "a" },
            { displayText: "乙", value: "b" },
          ],
          value: "b",
        },
      },
    ];
  session.commit(
    session.prepare({ protocolVersion: 1, documentId: "controls", baseRevision: 0, operations }),
  );
  let bytes = exportDocx(session).bytes;
  for (let i = 0; i < 2; i++) {
    const imported = importDocxWithReport(bytes, {
      documentId: "reopen",
      unsupportedContent: "reject",
    });
    expect(imported.report.preserved).toHaveLength(0);
    const model = openDocument(imported.files, { kind: "document" }),
      controls = model.snapshot().metadata.word!.controls!;
    expect(controls).toHaveLength(4);
    expect(controls.find((control) => control.kind === "checkBox")).toMatchObject({
      checked: true,
      checkedSymbol: 0x2611,
    });
    expect(controls.find((control) => control.kind === "date")).toMatchObject({
      fullDate: "2028-02-29T00:00:00Z",
    });
    expect(model.getNode(controls.find((control) => control.kind === "date")!.nodeId)!.text).toBe(
      "2028年2月29日",
    );
    expect(controls.find((control) => control.kind === "dropDown")?.value).toBe("b");
    bytes = exportDocx(model).bytes;
  }
});
test("unmodeled control properties and bound controls preserve their exact native data without swallowing a containing table", () => {
  const session = source();
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "controls",
      baseRevision: 0,
      operations: [
        {
          op: "insertWordControl",
          nodeId: "paragraph-1",
          control: { kind: "text", level: "inline" },
          text: "可编辑",
        },
      ],
    }),
  );
  const files = unzipSync(exportDocx(session).bytes);
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace("<w:sdtPr>", '<w:sdtPr><w:label w:val="77"/>'),
  );
  const imported = importDocxWithReport(zipSync(files), { documentId: "preserved" }),
    model = openDocument(imported.files, { kind: "document" });
  expect(model.snapshot().metadata.word!.controls).toHaveLength(1);
  expect(imported.report.preserved).toHaveLength(1);
  const xml = strFromU8(unzipSync(exportDocx(model).bytes)["word/document.xml"]!);
  expect(xml).toMatch(/:label\b[^>]*:val="77"/);
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace(
      "<w:sdtPr>",
      '<w:sdtPr><w:dataBinding w:storeItemID="{00000000-0000-0000-0000-000000000000}" w:xpath="/root/value"/>',
    ),
  );
  const tableSource = strFromU8(files["word/document.xml"]!),
    paragraph = tableSource.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/)![0];
  files["word/document.xml"] = strToU8(
    tableSource.replace(
      paragraph,
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr/>' +
        paragraph +
        "</w:tc></w:tr></w:tbl>",
    ),
  );
  const bound = importDocxWithReport(zipSync(files), { documentId: "bound" });
  expect(
    openDocument(bound.files, { kind: "document" }).snapshot().metadata.word!.tables,
  ).toHaveLength(1);
  expect(
    openDocument(bound.files, { kind: "document" }).snapshot().metadata.word!.controls ?? [],
  ).toHaveLength(0);
  expect(
    strFromU8(
      unzipSync(exportDocx(openDocument(bound.files, { kind: "document" })).bytes)[
        "word/document.xml"
      ]!,
    ),
  ).toContain("dataBinding");
});
test("picture replacement stages authorized bytes, removes obsolete drawing semantics and survives native reopen and undo", () => {
  const session = source();
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "controls",
      baseRevision: 0,
      operations: [
        {
          op: "insertWordControl",
          nodeId: "paragraph-1",
          control: { kind: "picture", level: "block" },
        },
      ],
    }),
  );
  const control = session.snapshot().metadata.word!.controls![0]!,
    bytes = Uint8Array.from(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aZs8AAAAASUVORK5CYII=",
        "base64",
      ),
    ),
    descriptor = createAssetDescriptor("assets/picture.png", "image/png", bytes);
  for (let i = 0; i < 2; i++)
    session.commit(
      session.prepare(
        {
          protocolVersion: 1,
          documentId: "controls",
          baseRevision: session.snapshot().manifest.revision,
          operations: [
            {
              op: "setWordControlPicture",
              nodeId: control.nodeId,
              assetPath: descriptor.path,
              alt: "替换 " + i,
            },
          ],
        },
        { assets: [{ descriptor, bytes }] },
      ),
    );
  expect(session.snapshot().metadata.word!.drawings).toHaveLength(1);
  const model = openDocument(
    importDocxWithReport(exportDocx(session).bytes, {
      documentId: "picture",
      unsupportedContent: "reject",
    }).files,
    { kind: "document" },
  );
  expect(model.snapshot().metadata.word!.controls![0]!.kind).toBe("picture");
  expect(model.snapshot().metadata.word!.drawings![0]!.kind).toBe("image");
  session.undo();
  expect(session.snapshot().metadata.word!.drawings).toHaveLength(1);
  expect(session.snapshot().html).toContain("替换 0");
});
test("old form field properties survive while cached results and their surrounding run remain editable", () => {
  const files = unzipSync(exportDocx(source()).bytes);
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace(
      "<w:body>",
      '<w:body><w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="legacy"/><w:enabled/><w:textInput><w:default w:val="初始"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>初始</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    ),
  );
  const imported = importDocxWithReport(zipSync(files), { documentId: "legacy" }),
    session = createDocumentSession(openDocument(imported.files, { kind: "document" }));
  expect(session.snapshot().metadata.word!.fields).toHaveLength(1);
  const resultId = session
    .snapshot()
    .metadata.word!.fields[0]!.resultNodeIds!.find((id) => session.getNode(id)?.text === "初始")!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "legacy",
      baseRevision: 0,
      operations: [
        { op: "replaceText", nodeId: resultId, from: 0, to: 2, expectedText: "初始", text: "修改" },
      ],
    }),
  );
  const xml = strFromU8(unzipSync(exportDocx(session).bytes)["word/document.xml"]!);
  expect(xml).toMatch(/:ffData\b/);
  expect(xml).toContain("FORMTEXT");
  expect(xml).toContain("修改");
  expect(
    openDocument(importDocxWithReport(exportDocx(session).bytes, { documentId: "again" }).files, {
      kind: "document",
    }).snapshot().metadata.word!.fields,
  ).toHaveLength(1);
});
