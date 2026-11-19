import { expect, test } from "vitest";
import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  createWordDocument,
  createDocumentSession,
  openDocument,
  parseWordMath,
} from "@doffi/core";
import { exportDocx, importDocxWithReport } from "@doffi/office";
test("multiple native math groups stay in one editable paragraph with inherited font roles and character formatting", () => {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "math-groups" }), { kind: "document" }),
    ),
    a = parseWordMath("x_1+漢\u{e0100}"),
    b = parseWordMath("\\frac{y}{2}"),
    nodes = [
      ...a.nodes,
      ...b.nodes.map((node) => ({
        ...node,
        mathId: "b-" + node.mathId,
        ...("children" in node ? { children: node.children.map((id) => "b-" + id) } : {}),
      })),
      {
        mathId: "group",
        kind: "equationGroup" as const,
        children: [a.rootMathId, "b-" + b.rootMathId],
        style: {
          fonts: { ascii: "Cambria Math", hAnsi: "Cambria Math", eastAsia: "等线" },
          fontSize: 18,
          color: "#224466" as const,
          bold: true,
          language: "zh-CN",
        },
      },
    ];
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "math-groups",
      baseRevision: 0,
      operations: [
        {
          op: "insertWordEquation",
          nodeId: "paragraph-1",
          offset: 0,
          display: "block",
          alignment: "left",
          math: { rootMathId: "group", nodes },
        },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let index = 0; index < 2; index++) {
    const imported = importDocxWithReport(bytes, {
      documentId: "math-groups-reopen",
      unsupportedContent: "reject",
    });
    expect(imported.report.preserved).toHaveLength(0);
    const model = openDocument(imported.files, { kind: "document" }),
      equation = model.snapshot().metadata.word!.equations![0]!;
    expect(model.snapshot().metadata.word!.equations).toHaveLength(1);
    expect(equation.nodes.find((node) => node.mathId === equation.rootMathId)).toMatchObject({
      kind: "equationGroup",
      children: expect.arrayContaining([expect.any(String), expect.any(String)]),
    });
    expect(
      equation.nodes
        .filter((node) => node.kind === "text")
        .every(
          (node) =>
            node.style?.fontSize === 18 &&
            node.style?.fonts?.eastAsia === "等线" &&
            node.style?.bold === true,
        ),
    ).toBe(true);
    expect(
      strFromU8(unzipSync(bytes)["word/document.xml"]!).match(/<m:oMath[ >]/g) ?? [],
    ).toHaveLength(2);
    bytes = exportDocx(model).bytes;
  }
  const before = session.snapshot(),
    equation = before.metadata.word!.equations![0]!;
  expect(() =>
    session.prepare({
      protocolVersion: 1,
      documentId: "math-groups",
      baseRevision: before.manifest.revision,
      operations: [
        {
          op: "setWordEquation",
          nodeId: equation.nodeId,
          math: {
            rootMathId: equation.rootMathId,
            nodes: JSON.parse(JSON.stringify(equation.nodes)),
          },
          display: "inline",
        },
      ],
    }),
  ).toThrow();
  expect(session.snapshot()).toBe(before);
});
test("typed fractions, indexed radicals and scripts reopen twice as editable OMML without cached equation XML", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "equations" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "equations",
      baseRevision: 0,
      operations: [
        {
          op: "insertWordEquation",
          nodeId: "paragraph-1",
          offset: 0,
          math: parseWordMath("\\frac{x_1^2+1}{\\sqrt[3]{y}}+\\alpha"),
          display: "block",
          alignment: "left",
          fontFamily: "Cambria Math",
        },
      ],
    }),
  );
  let bytes = exportDocx(session).bytes;
  for (let i = 0; i < 2; i++) {
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(xml).toContain("m:oMath");
    expect(xml).toContain("m:f>");
    expect(xml).toContain("m:sSubSup>");
    expect(xml).toContain('m:val="left"');
    const imported = importDocxWithReport(bytes, {
      documentId: "math",
      unsupportedContent: "reject",
    });
    expect(imported.report.preserved).toHaveLength(0);
    const model = openDocument(imported.files, { kind: "document" }),
      equations = model.snapshot().metadata.word!.equations!;
    expect(equations).toHaveLength(1);
    expect(equations[0]!.nodes.some((node) => node.kind === "fraction")).toBe(true);
    expect(
      equations[0]!.nodes.some((node) => node.kind === "radical" && node.children.length === 2),
    ).toBe(true);
    bytes = exportDocx(model).bytes;
  }
});
test("common native math run font roles, sizes, colours and text styles retain their editable source values", () => {
  const files = unzipSync(
    new Uint8Array(
      readFileSync(new URL("../../../examples/docx/01-text-numbering.docx", import.meta.url)),
    ),
  );
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace(
      "<w:body>",
      '<w:body><w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:sty m:val="bi"/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:eastAsia="等线"/><w:color w:val="AA2244"/><w:sz w:val="30"/></w:rPr><m:t>漢字x</m:t></m:r></m:oMath></w:p>',
    ),
  );
  const imported = importDocxWithReport(zipSync(files), {
    documentId: "styled",
    unsupportedContent: "reject",
  });
  expect(imported.report.preserved).toHaveLength(0);
  const model = openDocument(imported.files, { kind: "document" }),
    token = model
      .snapshot()
      .metadata.word!.equations![0]!.nodes.find((node) => node.kind === "text")!;
  expect(token.style).toMatchObject({
    fonts: { ascii: "Cambria Math", eastAsia: "等线" },
    color: "#AA2244",
    fontSize: 15,
  });
  expect(token.kind === "text" && token.variant).toBe("boldItalic");
  const xml = strFromU8(unzipSync(exportDocx(model).bytes)["word/document.xml"]!);
  expect(xml).toContain("等线");
  expect(xml).toContain("AA2244");
});
test("unsupported math layout remains an intact native equation beside editable surrounding text", () => {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "unsupported-math" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "unsupported-math",
      baseRevision: 0,
      operations: [
        { op: "insertWordEquation", nodeId: "paragraph-1", offset: 0, math: parseWordMath("x") },
      ],
    }),
  );
  const files = unzipSync(exportDocx(session).bytes);
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace("<m:rPr>", '<m:rPr><m:brk m:alnAt="1"/>'),
  );
  const imported = importDocxWithReport(zipSync(files), { documentId: "opaque" });
  expect(imported.report.preserved.length).toBeGreaterThan(0);
  expect(
    strFromU8(
      unzipSync(exportDocx(openDocument(imported.files, { kind: "document" })).bytes)[
        "word/document.xml"
      ]!,
    ),
  ).toMatch(/<[^>]*:brk\b/);
});
