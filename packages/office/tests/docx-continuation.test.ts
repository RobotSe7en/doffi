import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { importDocx, exportDocx } from "@doffi/office";
import { strFromU8, unzipSync } from "fflate";
import { mutateDocx, styles } from "./regression-fixtures.js";

function fixture() {
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () =>
      styles(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Empty"><w:basedOn w:val="Normal"/></w:style><w:style w:type="character" w:styleId="Italic"><w:rPr><w:i/></w:rPr></w:style><w:style w:type="character" w:styleId="Other"/>',
      ),
    "word/document.xml": (source) =>
      source.replace(
        /(<w:tc><w:tcPr><w:vMerge\/><\/w:tcPr>)[\s\S]*?<\/w:tc>/,
        '$1<w:p><w:pPr><w:pStyle w:val="Empty"/><w:pBdr><w:bottom w:val="single" w:sz="8" w:color="224466"/></w:pBdr><w:spacing w:before="60" w:after="80" w:line="300" w:lineRule="auto"/><w:tabs><w:tab w:val="clear" w:pos="720"/><w:tab w:val="decimal" w:pos="1440" w:leader="dot"/></w:tabs><w:numPr><w:numId w:val="0"/></w:numPr><w:outlineLvl w:val="9"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="Italic"/><w:noProof/><w:color w:val="224466"/></w:rPr></w:r></w:p><w:p/></w:tc>',
      ),
  });
}
function current() {
  const session = createDocumentSession(
    openDocument(
      importDocx(fixture(), { documentId: "continuation", unsupportedContent: "reject" }),
      { kind: "document" },
    ),
  );
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "continuation",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  return { session, apply };
}
test("empty merged continuation paragraph and run formats, tab clears and style associations survive two strict native reopens", () => {
  const { session } = current();
  expect(session.snapshot().metadata.word!.nativeContent?.fragments ?? []).toEqual([]);
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const native = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(native).toMatch(/<w:tab[^>]*w:val="clear"[^>]*w:pos="720"/);
    expect(native).toContain('w:rStyle w:val="Italic"');
    const model = openDocument(
        importDocx(bytes, { documentId: "continued", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
      entry = model.snapshot().metadata.word!.tables[0]!.continuationFormats![0]!,
      first = entry.emptyParagraphs![0]!;
    expect(entry.emptyParagraphs).toHaveLength(2);
    expect(first).toMatchObject({
      styleId: "Empty",
      outlineLevel: 9,
      numberingDisabled: true,
      tabs: [
        { position: 36, alignment: "clear", leader: "none" },
        { position: 72, alignment: "decimal", leader: "dot" },
      ],
      runs: [{ styleId: "Italic" }],
    });
    expect(
      model
        .snapshot()
        .styles.map((sheet) => sheet.css)
        .join(""),
    ).toContain("margin-top:3pt");
    expect(native).toMatch(/<w:bottom[^>]*w:color="224466"/);
    expect(native).toContain('w:after="80"');
    expect(native).toContain('w:sz w:val="28"');
    expect(native).toContain('w:noProof w:val="1"');
    bytes = exportDocx(model).bytes;
  }
});
test("editing merged text and moving its continuation rows and columns keep empty formats and one-step history", () => {
  const { session, apply } = current(),
    table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    nodes = [table];
  for (let index = 0; index < nodes.length; index++)
    nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
  const paragraph = nodes.find(
    (id) => session.getNode(id)!.tagName === "p" && session.getNode(id)!.text.includes("纵向合并"),
  )!;
  const before = session.snapshot();
  apply([
    { op: "replaceText", nodeId: paragraph, from: 0, to: 0, expectedText: "", text: "编辑" },
    { op: "editWordTable", nodeId: table, change: { kind: "insertRows", at: 0, count: 1 } },
    { op: "editWordTable", nodeId: table, change: { kind: "insertColumns", at: 0, count: 1 } },
  ]);
  const entry = session.snapshot().metadata.word!.tables[0]!.continuationFormats![0]!,
    original = before.metadata.word!.tables[0]!.continuationFormats![0]!;
  expect(entry).toMatchObject({ row: original.row + 1, column: original.column + 1 });
  expect(entry.emptyParagraphs![0]!.styleId).toBe("Empty");
  const reopened = openDocument(
    importDocx(exportDocx(session).bytes, { documentId: "edited", unsupportedContent: "reject" }),
    { kind: "document" },
  );
  expect(reopened.snapshot().html).toContain("编辑纵向合并");
  expect(
    reopened.snapshot().metadata.word!.tables[0]!.continuationFormats![0]!.emptyParagraphs![0]!
      .styleId,
  ).toBe("Empty");
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
});
test("deleting named styles migrates empty continuation associations and undo restores them", () => {
  const { session, apply } = current(),
    root = session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
  apply([
    { op: "deleteWordStyle", nodeId: root, styleId: "Empty", replacementStyleId: "Normal" },
    { op: "deleteWordStyle", nodeId: root, styleId: "Italic", replacementStyleId: "Other" },
  ]);
  const paragraph =
    session.snapshot().metadata.word!.tables[0]!.continuationFormats![0]!.emptyParagraphs![0]!;
  expect(paragraph.styleId).toBe("Normal");
  expect(paragraph.runs[0]!.styleId).toBe("Other");
  expect(() =>
    importDocx(exportDocx(session).bytes, {
      documentId: "styles-reopened",
      unsupportedContent: "reject",
    }),
  ).not.toThrow();
  session.undo();
  expect(
    session.snapshot().metadata.word!.tables[0]!.continuationFormats![0]!.emptyParagraphs![0]!
      .styleId,
  ).toBe("Empty");
});
test("merging another region retains independent continuation formats outside that region", () => {
  const { session, apply } = current(),
    table = session.snapshot().metadata.word!.tables[0]!,
    before = session.snapshot(),
    entry = table.continuationFormats![0]!;
  apply([
    {
      op: "editWordTable",
      nodeId: table.nodeId,
      change: {
        kind: "mergeCells",
        row: entry.row - 1,
        column: entry.column + 1,
        rows: 2,
        columns: 2,
      },
    },
  ]);
  expect(session.snapshot().metadata.word!.tables[0]!.continuationFormats).toEqual(
    table.continuationFormats,
  );
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const model = openDocument(
        importDocx(bytes, { documentId: "local-merge", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
      retained = model
        .snapshot()
        .metadata.word!.tables[0]!.continuationFormats!.find(
          (item) => item.row === entry.row && item.column === entry.column,
        )!;
    expect(retained.emptyParagraphs![0]!.styleId).toBe("Empty");
    expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).toContain('w:color="224466"');
    bytes = exportDocx(model).bytes;
  }
  session.undo();
  expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
  expect(session.snapshot().html).toBe(before.html);
});
test("unmerging turns continuation paragraphs, mark fonts, empty runs and tab declarations into editable cell source", () => {
  const { session, apply } = current(),
    table = session.snapshot().metadata.word!.tables[0]!,
    before = session.snapshot(),
    entry = table.continuationFormats![0]!,
    text = session.getNode(table.nodeId)!.text;
  apply([
    {
      op: "editWordTable",
      nodeId: table.nodeId,
      change: { kind: "splitCell", row: entry.row - 1, column: entry.column },
    },
  ]);
  expect(session.getNode(table.nodeId)!.text).toBe(text);
  expect(session.snapshot().html).toContain('data-word-numbering-off="true"');
  expect(session.snapshot().html).toContain('data-word-style="Empty"');
  expect(
    session
      .snapshot()
      .metadata.word!.tabs.some((item) => item.stops.some((tab) => tab.alignment === "clear")),
  ).toBe(true);
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(xml).toContain('w:pStyle w:val="Empty"');
    expect(xml).toContain('w:rStyle w:val="Italic"');
    expect(xml).toContain('w:sz w:val="28"');
    expect(xml).toContain('w:noProof w:val="1"');
    expect(xml).toContain('<w:numId w:val="0"');
    const model = openDocument(
      importDocx(bytes, { documentId: "unmerged", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(model.getNode(model.snapshot().metadata.word!.tables[0]!.nodeId)!.text).toBe(text);
    bytes = exportDocx(model).bytes;
  }
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  expect(session.snapshot().metadata.word).toEqual(before.metadata.word);
});
test("refined splits retain original continuation boundaries while an unrelated split preserves their formats", () => {
  const { session, apply } = current(),
    table = session.snapshot().metadata.word!.tables[0]!,
    entry = table.continuationFormats![0]!;
  apply([
    { op: "editWordTable", nodeId: table.nodeId, change: { kind: "splitCell", row: 0, column: 0 } },
  ]);
  expect(session.snapshot().metadata.word!.tables[0]!.continuationFormats).toEqual(
    table.continuationFormats,
  );
  session.undo();
  apply([
    {
      op: "editWordTable",
      nodeId: table.nodeId,
      change: { kind: "splitCell", row: entry.row - 1, column: entry.column, rows: 3, columns: 2 },
    },
  ]);
  expect(
    session
      .snapshot()
      .metadata.word!.tables[0]!.continuationFormats!.filter(
        (item) => item.emptyParagraphs?.[0]?.styleId === "Empty",
      ),
  ).toHaveLength(2);
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const model = openDocument(
      importDocx(bytes, { documentId: "refined", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(
      model
        .snapshot()
        .metadata.word!.tables[0]!.continuationFormats!.filter(
          (item) => item.emptyParagraphs?.[0]?.styleId === "Empty",
        ),
    ).toHaveLength(2);
    bytes = exportDocx(model).bytes;
  }
});
test("splitting many formatted continuation paragraphs respects the source node budget and rolls back atomically", () => {
  const bytes = mutateDocx("02-table-merges.docx", {
      "word/document.xml": (source) =>
        source.replace(
          /(<w:tc><w:tcPr><w:vMerge\/><\/w:tcPr>)[\s\S]*?<\/w:tc>/,
          "$1" + "<w:p/>".repeat(40) + "</w:tc>",
        ),
    }),
    session = createDocumentSession(
      openDocument(
        importDocx(bytes, { documentId: "split-budget", unsupportedContent: "reject" }),
        { kind: "document", limits: { maxNodes: 800 } },
      ),
      { limits: { maxNodes: 800 } },
    ),
    before = session.snapshot(),
    table = before.metadata.word!.tables[0]!,
    entry = table.continuationFormats![0]!;
  expect(() =>
    session.prepare({
      protocolVersion: 1,
      documentId: "split-budget",
      baseRevision: 0,
      operations: [
        {
          op: "editWordTable",
          nodeId: table.nodeId,
          change: {
            kind: "splitCell",
            row: entry.row - 1,
            column: entry.column,
            rows: 2,
            columns: 100,
          },
        },
      ],
    }),
  ).toThrowError(
    expect.objectContaining({
      code: "ResourceLimitExceeded",
      message: expect.stringContaining("Split continuation"),
    }),
  );
  expect(session.snapshot()).toEqual(before);
  expect(session.history().undoSteps).toBe(0);
});
