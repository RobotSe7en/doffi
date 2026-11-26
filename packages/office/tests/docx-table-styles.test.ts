import { expect, test } from "vitest";
import { strFromU8, unzipSync, strToU8, zipSync } from "fflate";
import {
  openDocument,
  createDocumentSession,
  createWordDocument,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
  WordStyleSchema,
  getWordTextDefaults,
} from "@doffi/core";
import type { DocumentNode, SourceStyleNode, DocumentOperation } from "@doffi/core";
import { importDocx, exportDocx } from "@doffi/office";
import { tableStylesFixture, rowStylesFixture } from "./table-style-fixture.js";
import { tableDefaultCases, tableDefaultFixture, rowHeightFixture } from "./table-default-fixture.js";
test("text defaults follow live table conditions and paragraph styles independently of character overrides through native reopens", () => {
  const session = createDocumentSession(
      openDocument(
        importDocx(tableStylesFixture(), {
          documentId: "table-font-defaults",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    paragraph = descendants(session, table).find(
      (node) => node.tagName === "p" && node.text === "R0C1",
    )!,
    run = descendants(session, paragraph.nodeId).find(
      (node) => node.attributes["data-word-style"] === "Toggle",
    )!,
    defaults = getWordTextDefaults(session, run.nodeId);
  expect(defaults.color).toBe("#cc0000");
  expect(defaults["text-decoration"]).toContain("underline");
  expect(defaults["font-weight"] === "bold" || Number(defaults["font-weight"]) >= 600).toBe(true);
  expect(css(session, run.nodeId)["font-weight"]).toMatch(/normal|400/);
  const root = session.getNode(
    session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
  )!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "table-font-defaults",
      baseRevision: 0,
      operations: [
        {
          op: "setWordStyle",
          nodeId: root,
          style: { kind: "paragraph", styleId: "Local", cssClass: "local" },
          runProperties: { "font-size": "18pt", "font-style": "italic" },
        },
        { op: "setWordParagraph", nodeId: paragraph.nodeId, styleId: "Local" },
        { op: "formatWordTable", nodeId: table, styleId: "Blue" },
      ],
    }),
  );
  expect(getWordTextDefaults(session, paragraph.nodeId)).toMatchObject({
    color: "#0000cc",
    "font-size": "18pt",
    "font-style": "italic",
  });
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const reopened = createDocumentSession(
        openDocument(
          importDocx(bytes, { documentId: "table-font-reopen", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
      ),
      node = descendants(reopened, reopened.snapshot().metadata.word!.tables[0]!.nodeId).find(
        (node) => node.tagName === "p" && node.text === "R0C1",
      )!;
    expect(getWordTextDefaults(reopened, node.nodeId)).toMatchObject({
      color: "#0000cc",
      "font-size": "18pt",
      "font-style": "italic",
    });
    bytes = exportDocx(reopened).bytes;
  }
  session.undo();
  expect(getWordTextDefaults(session, paragraph.nodeId).color).toBe("#cc0000");
});
function css(session: ReturnType<typeof createDocumentSession>, id: string) {
  const resolver = new SourceStyleResolver(
      session.snapshot().styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    project = (id: string): SourceStyleNode => {
      const node = session.getNode(id)!;
      return {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
    };
  return resolver.get(project(id));
}
function descendants(
  session: ReturnType<typeof createDocumentSession>,
  id: string,
): DocumentNode[] {
  const node = session.getNode(id)!;
  return [node, ...node.childNodeIds.flatMap((child) => descendants(session, child))];
}
test("shares repeated table rule declarations while keeping each table independently editable", () => {
  const archive = unzipSync(tableStylesFixture());
  archive["word/document.xml"] = strToU8(
    strFromU8(archive["word/document.xml"]!).replace(/<w:tbl>[\s\S]*?<\/w:tbl>/, (table) => table + table),
  );
  const open = (bytes: Uint8Array) => createDocumentSession(openDocument(importDocx(bytes, {
    documentId: "grouped-table-css",
    unsupportedContent: "reject",
  }), { kind: "document" }));
  const session = open(zipSync(archive));
  const tables = session.snapshot().metadata.word!.tables;
  const colors = (current: typeof session) => current.snapshot().metadata.word!.tables.map(({ nodeId }) => {
    const firstCell = descendants(current, nodeId).find((node) => node.tagName === "td")!;
    return css(current, firstCell.nodeId)["background-color"];
  });
  const selectors = tables.map((table) => {
    const context = table.styleContexts!.find(({ kind }) => kind === "cell")!;
    return ("." + context.cssClass).repeat(5);
  });
  const ruleSelectors = session.snapshot().styles[0]!.css.match(/[^{}]+(?=\{)/g) ?? [];
  expect(ruleSelectors.some((list) => selectors.every((selector) => list.split(",").includes(selector))))
    .toBe(true);
  expect(colors(session)).toEqual(["#aa1111", "#aa1111"]);
  session.commit(session.prepare({
    protocolVersion: 1,
    documentId: "grouped-table-css",
    baseRevision: 0,
    operations: [{ op: "formatWordTable", nodeId: tables[0]!.nodeId, styleId: "Blue" }],
  }));
  expect(colors(session)).toEqual(["#1111aa", "#aa1111"]);
  const reopened = open(exportDocx(session).bytes);
  expect(colors(reopened)).toEqual(["#1111aa", "#aa1111"]);
  reopened.dispose();
  session.dispose();
});
test("switching table styles recomposes conditional cells and run toggles, retains direct formats, and remains style-driven after two native reopens", () => {
  const session = createDocumentSession(
    openDocument(
      importDocx(tableStylesFixture(), { documentId: "table-style", unsupportedContent: "reject" }),
      { kind: "document" },
    ),
  );
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "table-style",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  const table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    get = () => descendants(session, table),
    cells = () => get().filter((node) => node.tagName === "td"),
    runs = () => get().filter((node) => node.attributes["data-word-run"]);
  expect(css(session, cells()[0]!.nodeId)["background-color"]).toBe("#aa1111");
  expect(css(session, runs()[0]!.nodeId)["font-weight"]).toBe("bold");
  expect(css(session, runs()[1]!.nodeId)["font-weight"]).toBe("normal");
  apply([
    { op: "formatWordCell", nodeId: cells()[2]!.nodeId, properties: { "padding-left": "17pt" } },
    { op: "formatWordTable", nodeId: table, styleId: "Blue" },
  ]);
  expect(css(session, cells()[0]!.nodeId)["background-color"]).toBe("#1111aa");
  expect(css(session, cells()[2]!.nodeId)["background-color"]).toBe("#0000cc");
  expect(css(session, cells()[2]!.nodeId)["padding-left"]).toBe("17pt");
  expect(css(session, cells()[3]!.nodeId)["background-color"]).toBe("#00ff00");
  expect(css(session, runs()[1]!.nodeId)["font-weight"]).toBe("normal");
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(xml.match(/<w:shd/g)).toHaveLength(1);
    expect(xml).not.toContain('w:fill="1111AA"');
    const reopened = createDocumentSession(
        openDocument(importDocx(bytes, { documentId: "reopen", unsupportedContent: "reject" }), {
          kind: "document",
        }),
      ),
      all = descendants(reopened, reopened.snapshot().metadata.word!.tables[0]!.nodeId),
      first = all.find((node) => node.tagName === "td")!;
    expect(css(reopened, first.nodeId)["background-color"]).toBe("#1111aa");
    expect(
      all
        .filter((node) => node.attributes["data-word-run"])
        .map((node) => css(reopened, node.nodeId)["font-weight"]),
    ).toEqual(["bold", "normal", "normal", "normal"]);
    bytes = exportDocx(reopened).bytes;
    reopened.dispose();
  }
  session.undo();
  expect(css(session, cells()[0]!.nodeId)["background-color"]).toBe("#aa1111");
  session.redo();
  const style = WordStyleSchema.parse(
    session.snapshot().metadata.word!.styles.find((style) => style.styleId === "Blue")!,
  );
  apply([
    {
      op: "setWordStyle",
      nodeId: session.getNode(
        session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!
          .rootNodeId,
      )!.parentNodeId!,
      style,
      cellProperties: { "background-color": "#dddddd" },
    },
  ]);
  expect(css(session, cells()[2]!.nodeId)["background-color"]).toBe("#dddddd");
  expect(css(session, cells()[0]!.nodeId)["background-color"]).toBe("#1111aa");
});
test("table-style histories retain original conditional options, reopen and reject without freezing cell paints", () => {
  const session = createDocumentSession(
      openDocument(
        importDocx(tableStylesFixture(), {
          documentId: "tracked-table",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    root = session.getNode(
      session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
    )!.parentNodeId!,
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "tracked-table",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    { op: "setWordReview", nodeId: root, trackChanges: true, author: "甲" },
    { op: "formatWordTable", nodeId: table, styleId: "Blue", lookMask: 1536, rowBandSize: 2 },
  ]);
  const first = descendants(session, table).find((node) => node.tagName === "td")!;
  expect(css(session, first.nodeId)["background-color"]).toBe("#0000cc");
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const reopened = createDocumentSession(
        openDocument(
          importDocx(bytes, { documentId: "tracked-reopen", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
      ),
      snapshot = reopened.snapshot(),
      revision = snapshot.metadata.word!.revisions.find((item) => item.property === "table")!;
    expect(revision.before?.attributes?.["data-word-style"]).toBe("Red");
    expect(revision.before?.tableStyleOptions?.lookMask).toBe(1568);
    bytes = exportDocx(reopened).bytes;
    reopened.commit(
      reopened.prepare({
        protocolVersion: 1,
        documentId: "tracked-reopen",
        baseRevision: 0,
        operations: [
          {
            op: "reviewWordRevision",
            nodeId: revision.nodeId,
            revisionId: revision.revisionId!,
            decision: "reject",
          },
        ],
      }),
    );
    expect(
      css(
        reopened,
        descendants(reopened, revision.nodeId).find((node) => node.tagName === "td")!.nodeId,
      )["background-color"],
    ).toBe("#aa1111");
    reopened.dispose();
  }
  apply([{ op: "reviewWordRevisions", nodeId: root, decision: "reject" }]);
  expect(css(session, first.nodeId)["background-color"]).toBe("#aa1111");
  session.undo();
  expect(css(session, first.nodeId)["background-color"]).toBe("#0000cc");
});
test("table structural edits recompute first-row conditional styles and direct paragraph fonts continue to inherit into runs", () => {
  const session = createDocumentSession(
      openDocument(
        importDocx(tableStylesFixture(), {
          documentId: "structure-style",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "structure-style",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([{ op: "editWordTable", nodeId: table, change: { kind: "insertRows", at: 0, count: 1 } }]);
  const nodes = descendants(session, table),
    cells = nodes.filter((node) => node.tagName === "td");
  expect(css(session, cells[0]!.nodeId)["background-color"]).toBe("#aa1111");
  expect(css(session, cells[2]!.nodeId)["background-color"]).toBe("#cc0000");
  const paragraph = nodes.find((node) => node.tagName === "p" && node.text === "R0C0")!,
    runNode = descendants(session, paragraph.nodeId).find(
      (node) => node.attributes["data-word-run"],
    )!;
  apply([{ op: "setStyle", nodeId: paragraph.nodeId, property: "font-size", value: "22pt" }]);
  expect(css(session, runNode.nodeId)["font-size"]).toBe("22pt");
  apply([{ op: "editWordTable", nodeId: table, change: { kind: "deleteRows", at: 0, count: 1 } }]);
  expect(
    css(session, descendants(session, table).find((node) => node.tagName === "td")!.nodeId)[
      "background-color"
    ],
  ).toBe("#aa1111");
});
test("explicit cell formatting equal to a table style is retained when reopened and another style is applied", () => {
  const session = createDocumentSession(
      openDocument(
        importDocx(tableStylesFixture(), {
          documentId: "equal-direct",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    cell = descendants(session, table).filter((node) => node.tagName === "td")[2]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "equal-direct",
      baseRevision: 0,
      operations: [
        {
          op: "formatWordCell",
          nodeId: cell.nodeId,
          properties: { "background-color": "#cc0000" },
        },
      ],
    }),
  );
  const reopened = createDocumentSession(
      openDocument(
        importDocx(exportDocx(session).bytes, {
          documentId: "equal-reopen",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    id = reopened.snapshot().metadata.word!.tables[0]!.nodeId;
  reopened.commit(
    reopened.prepare({
      protocolVersion: 1,
      documentId: "equal-reopen",
      baseRevision: 0,
      operations: [{ op: "formatWordTable", nodeId: id, styleId: "Blue" }],
    }),
  );
  expect(
    css(reopened, descendants(reopened, id).filter((node) => node.tagName === "td")[2]!.nodeId)[
      "background-color"
    ],
  ).toBe("#cc0000");
});

test.each([undefined, 0, 1])(
  "Word table band size %s keeps missing/zero unbanded and preserves explicit native zero",
  (size) => {
    const files = unzipSync(tableStylesFixture());
    files["word/styles.xml"] = strToU8(
      strFromU8(files["word/styles.xml"]!).replace(
        '<w:style w:type="table" w:styleId="Red">',
        '<w:style w:type="table" w:styleId="Red"><w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:fill="123456"/></w:tcPr></w:tblStylePr>',
      ),
    );
    files["word/document.xml"] = strToU8(
      strFromU8(files["word/document.xml"]!)
        .replace(
          "<w:tblPr>",
          "<w:tblPr>" + (size !== undefined ? '<w:tblStyleRowBandSize w:val="' + size + '"/>' : ""),
        )
        .replace(/<w:tblLook[^>]*\/>/, '<w:tblLook w:val="0400"/>'),
    );
    let bytes: Uint8Array = zipSync(files);
    for (let pass = 0; pass < 2; pass++) {
      const session = createDocumentSession(
          openDocument(importDocx(bytes, { documentId: "bands", unsupportedContent: "reject" }), {
            kind: "document",
          }),
        ),
        table = session.snapshot().metadata.word!.tables[0]!.nodeId,
        cell = descendants(session, table).find((node) => node.tagName === "td")!;
      expect(css(session, cell.nodeId)["background-color"]).toBe(
        size === 1 ? "#123456" : "#cc0000",
      );
      expect(session.snapshot().metadata.word!.tables[0]!.rowBandSize).toBe(size);
      bytes = exportDocx(session).bytes;
      session.dispose();
    }
  },
);

test("conditional row headers and cantSplit remain style-driven; column conditions apply cell margins to whole rows", () => {
  const session = createDocumentSession(
      openDocument(
        importDocx(rowStylesFixture(), { documentId: "row-styles", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    ),
    table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    rows = descendants(session, table).filter((node) => node.tagName === "tr"),
    cells = descendants(session, table).filter((node) => node.tagName === "td");
  expect(css(session, rows[0]!.nodeId)).toMatchObject({
    "break-inside": "avoid",
    "--word-row-repeat-header": "1",
  });
  expect(css(session, rows[1]!.nodeId)["break-inside"]).toBe("auto");
  expect(css(session, cells[1]!.nodeId)).toMatchObject({
    "padding-left": "13pt",
    "vertical-align": "bottom",
  });
  const output = unzipSync(exportDocx(session).bytes),
    document = strFromU8(output["word/document.xml"]!),
    stylesXml = strFromU8(output["word/styles.xml"]!);
  expect(document).not.toMatch(/<w:(?:cantSplit|tblHeader)/);
  expect(stylesXml).toContain('<w:tblHeader w:val="1"');
  const root = session.getNode(
    session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
  )!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "row-styles",
      baseRevision: 0,
      operations: [
        { op: "setWordReview", nodeId: root, trackChanges: true, author: "甲" },
        { op: "formatWordTable", nodeId: table, repeatHeaderRows: 0, allowRowSplit: true },
      ],
    }),
  );
  const reopened = createDocumentSession(
      openDocument(
        importDocx(exportDocx(session).bytes, {
          documentId: "row-reopen",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    reopenedTable = reopened.snapshot().metadata.word!.tables[0]!.nodeId;
  expect(
    descendants(reopened, reopenedTable).find((node) => node.tagName === "tr")!.attributes[
      "data-word-repeat-header"
    ],
  ).toBe("false");
  reopened.commit(
    reopened.prepare({
      protocolVersion: 1,
      documentId: "row-reopen",
      baseRevision: 0,
      operations: [
        {
          op: "reviewWordRevisions",
          nodeId: reopened.getNode(
            reopened.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!
              .rootNodeId,
          )!.parentNodeId!,
          decision: "reject",
        },
      ],
    }),
  );
  const first = descendants(reopened, reopenedTable).find((node) => node.tagName === "tr")!;
  expect(first.attributes["data-word-repeat-header"]).toBeUndefined();
  expect(css(reopened, first.nodeId)["break-inside"]).toBe("avoid");
});
test("Word table style bold properties reset state and ignored wholeTable definitions roundtrip without affecting layout", () => {
  const files = unzipSync(tableStylesFixture());
  files["word/styles.xml"] = strToU8(
    strFromU8(files["word/styles.xml"]!)
      .replace(
        '<w:style w:type="table" w:styleId="Red">',
        '<w:style w:type="table" w:styleId="Base"><w:rPr><w:b/></w:rPr></w:style><w:style w:type="table" w:styleId="Red"><w:basedOn w:val="Base"/><w:tblStylePr w:type="wholeTable"><w:rPr><w:sz w:val="72"/></w:rPr></w:tblStylePr>',
      )
      .replace(
        '<w:rPr><w:color w:val="CC0000"/></w:rPr>',
        '<w:rPr><w:b/><w:color w:val="CC0000"/></w:rPr>',
      ),
  );
  let bytes: Uint8Array = zipSync(files);
  for (let pass = 0; pass < 2; pass++) {
    const session = createDocumentSession(
        openDocument(
          importDocx(bytes, { documentId: "word-table-reset", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
      ),
      id = session.snapshot().metadata.word!.tables[0]!.nodeId,
      runs = descendants(session, id).filter((node) => node.attributes["data-word-run"]);
    expect(css(session, runs[0]!.nodeId)["font-weight"]).toBe("bold");
    expect(css(session, runs[2]!.nodeId)["font-weight"]).toBe("bold");
    expect(css(session, runs[0]!.nodeId)["font-size"]).not.toBe("36pt");
    bytes = exportDocx(session).bytes;
    expect(strFromU8(unzipSync(bytes)["word/styles.xml"]!)).toContain('w:type="wholeTable"');
    session.dispose();
  }
});

test("clearing a table style exports an explicit neutral style instead of reapplying the document default", () => {
  const files = unzipSync(tableStylesFixture());
  files["word/styles.xml"] = strToU8(
    strFromU8(files["word/styles.xml"]!).replace(
      'w:styleId="Red"',
      'w:styleId="Red" w:default="1"',
    ),
  );
  const session = createDocumentSession(
      openDocument(
        importDocx(zipSync(files), { documentId: "clear-style", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    ),
    id = session.snapshot().metadata.word!.tables[0]!.nodeId;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "clear-style",
      baseRevision: 0,
      operations: [{ op: "formatWordTable", nodeId: id, styleId: null }],
    }),
  );
  const count = session.snapshot().metadata.word!.styles.length;
  let bytes = exportDocx(session).bytes;
  expect(session.snapshot().metadata.word!.styles).toHaveLength(count);
  for (let pass = 0; pass < 2; pass++) {
    const reopened = createDocumentSession(
        openDocument(
          importDocx(bytes, { documentId: "plain-reopen", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
      ),
      table = reopened.snapshot().metadata.word!.tables[0]!.nodeId;
    expect(
      css(reopened, descendants(reopened, table).find((node) => node.tagName === "td")!.nodeId)[
        "background-color"
      ],
    ).toBe("transparent");
    expect(reopened.getNode(table)!.attributes["data-word-style"]).not.toBe("Red");
    bytes = exportDocx(reopened).bytes;
    reopened.dispose();
  }
});

test("condition definitions edit each component, recompute after structure changes and remove or undo as one transaction", () => {
  const session = createDocumentSession(
      openDocument(
        importDocx(tableStylesFixture(), {
          documentId: "condition-edit",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    root = session.getNode(
      session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
    )!.parentNodeId!,
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "condition-edit",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      ),
    definition = () =>
      WordStyleSchema.parse(
        session.snapshot().metadata.word!.styles.find((style) => style.styleId === "Red")!,
      );
  apply([
    {
      op: "setWordStyle",
      nodeId: root,
      style: definition(),
      tableCondition: {
        region: "firstRow",
        cellProperties: { "background-color": "#663399" },
        runProperties: { "font-style": "italic" },
        rowProperties: { "--word-row-repeat-header": "1", "break-inside": "avoid" },
      },
    },
  ]);
  const cells = () => descendants(session, table).filter((node) => node.tagName === "td"),
    firstRun = () => descendants(session, table).find((node) => node.attributes["data-word-run"])!;
  expect(css(session, cells()[0]!.nodeId)["background-color"]).toBe("#663399");
  expect(css(session, firstRun().nodeId)["font-style"]).toBe("italic");
  expect(
    css(session, descendants(session, table).find((node) => node.tagName === "tr")!.nodeId)[
      "--word-row-repeat-header"
    ],
  ).toBe("1");
  apply([{ op: "editWordTable", nodeId: table, change: { kind: "insertRows", at: 0, count: 1 } }]);
  expect(css(session, cells()[2]!.nodeId)["background-color"]).toBe("#cc0000");
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const reopened = createDocumentSession(
        openDocument(
          importDocx(bytes, { documentId: "condition-reopen", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
      ),
      id = reopened.snapshot().metadata.word!.tables[0]!.nodeId;
    expect(
      css(reopened, descendants(reopened, id).find((node) => node.tagName === "td")!.nodeId)[
        "background-color"
      ],
    ).toBe("#663399");
    bytes = exportDocx(reopened).bytes;
    reopened.dispose();
  }
  apply([
    {
      op: "setWordStyle",
      nodeId: root,
      style: definition(),
      tableCondition: { region: "firstRow", clear: true },
    },
  ]);
  expect(css(session, cells()[0]!.nodeId)["background-color"]).toBe("#cc0000");
  session.undo();
  expect(css(session, cells()[0]!.nodeId)["background-color"]).toBe("#663399");
});

test("preserved native conditional properties follow their region after reordering and are collected when that region is cleared", () => {
  const files = unzipSync(rowStylesFixture());
  files["word/styles.xml"] = strToU8(
    strFromU8(files["word/styles.xml"]!).replace(
      '<w:tblStylePr w:type="firstRow"><w:trPr>',
      '<w:tblStylePr w:type="firstRow"><w:trPr><w:hidden/>',
    ),
  );
  const session = createDocumentSession(
      openDocument(importDocx(zipSync(files), { documentId: "native-condition" }), {
        kind: "document",
      }),
    ),
    root = session.getNode(
      session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
    )!.parentNodeId!,
    definition = WordStyleSchema.parse(
      session.snapshot().metadata.word!.styles.find((style) => style.styleId === "Red")!,
    );
  expect(
    session
      .snapshot()
      .metadata.word!.nativeContent!.fragments.filter((fragment) => fragment.styleId === "Red")
      .map((fragment) => fragment.styleRegion),
  ).toEqual(["firstRow"]);
  definition.conditions!.reverse();
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "native-condition",
      baseRevision: 0,
      operations: [{ op: "setWordStyle", nodeId: root, style: definition }],
    }),
  );
  const red = (bytes: Uint8Array) =>
      strFromU8(unzipSync(bytes)["word/styles.xml"]!).match(
        /<w:style[^>]*w:styleId="Red"[^>]*>([\s\S]*?)<\/w:style>/,
      )![1]!,
    output = red(exportDocx(session).bytes);
  expect(
    output.match(/<w:tblStylePr[^>]*w:type="firstRow"[^>]*>([\s\S]*?)<\/w:tblStylePr>/)![1],
  ).toContain("hidden");
  expect(
    output.match(/<w:tblStylePr[^>]*w:type="firstCol"[^>]*>([\s\S]*?)<\/w:tblStylePr>/)![1],
  ).not.toContain("hidden");
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "native-condition",
      baseRevision: 1,
      operations: [
        {
          op: "setWordStyle",
          nodeId: root,
          style: definition,
          tableCondition: { region: "firstRow", clear: true },
        },
      ],
    }),
  );
  expect(red(exportDocx(session).bytes)).not.toContain("hidden");
  expect(
    session
      .snapshot()
      .metadata.word!.nativeContent?.fragments.some(
        (fragment) => fragment.styleId === "Red" && fragment.styleRegion === "firstRow",
      ) ?? false,
  ).toBe(false);
  session.undo();
  expect(red(exportDocx(session).bytes)).toContain("hidden");
});

test("new tables use isolated named defaults and switching to plain removes grid defaults while retaining direct cell formats", () => {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "default-grid" }), { kind: "document" }),
    ),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "default-grid",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      ),
    definitions = session.snapshot().metadata.word!.styles;
  apply([{ op: "insertWordTable", nodeId: "paragraph-1", position: "after", rows: 2, columns: 2 }]);
  const table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    cells = descendants(session, table).filter((node) => node.tagName === "td");
  expect(session.getNode(table)!.attributes["data-word-style"]).toBe("TableGrid");
  expect(session.snapshot().metadata.word!.styles).toEqual(definitions);
  expect(css(session, table)["border-top-style"]).toBe("solid");
  expect(css(session, cells[0]!.nodeId)["padding-left"]).toBe("4pt");
  expect(cells.every((cell) => !cell.attributes.style?.includes("border"))).toBe(true);
  apply([
    { op: "formatWordCell", nodeId: cells[0]!.nodeId, properties: { "padding-left": "17pt" } },
    { op: "formatWordTable", nodeId: table, styleId: "TableNormal" },
  ]);
  expect(css(session, table)["border-top-style"] ?? "none").toBe("none");
  expect(css(session, cells[0]!.nodeId)["padding-left"]).toBe("17pt");
  expect(css(session, cells[1]!.nodeId)["padding-left"]).toBe("5.4pt");
  let bytes = exportDocx(session).bytes;
  expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).not.toContain("tblBorders");
  for (let pass = 0; pass < 2; pass++) {
    const reopened = createDocumentSession(
        openDocument(
          importDocx(bytes, { documentId: "grid-reopen", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
      ),
      id = reopened.snapshot().metadata.word!.tables[0]!.nodeId;
    expect(css(reopened, id)["border-top-style"] ?? "none").toBe("none");
    expect(
      css(reopened, descendants(reopened, id).find((node) => node.tagName === "td")!.nodeId)[
        "padding-left"
      ],
    ).toBe("17pt");
    bytes = exportDocx(reopened).bytes;
    reopened.dispose();
  }
  session.undo();
  expect(css(session, table)["border-top-style"]).toBe("solid");
});

// These native defaults were also queried in Word with independent synthetic
// sources. IDs, default bindings, built-in names and custom flags vary separately.
test.each(tableDefaultCases)("empty native table style $kind resolves its own defaults through editing and two reopens", ({ kind, padding }) => {
  const open = (bytes: Uint8Array) => createDocumentSession(openDocument(importDocx(bytes, {
    documentId: "native-table-defaults", unsupportedContent: "reject",
  }), { kind: "document" }));
  for (const override of [false, true]) {
    let session = open(tableDefaultFixture(kind, { override }));
    for (let pass = 0; pass < 3; pass++) {
      const table = session.snapshot().metadata.word!.tables[0]!.nodeId,
        cells = descendants(session, table).filter(node => node.tagName === "td");
      expect(css(session, table)["margin-left"]).toBe("0pt");
      expect(css(session, cells[0]!.nodeId)["padding-left"]).toBe((override ? 7 : padding) + "pt");
      expect(css(session, cells[0]!.nodeId)["padding-right"]).toBe((override ? 9 : padding) + "pt");
      expect(css(session, cells[1]!.nodeId)["padding-right"]).toBe(padding + "pt");
      const before = session.snapshot().html;
      session.commit(session.prepare({ protocolVersion: 1, documentId: "native-table-defaults",
        baseRevision: session.snapshot().manifest.revision,
        operations: [{ op: "formatWordTable", nodeId: table, properties: { "background-color": "#112233" } }],
      }));
      expect(css(session, cells[1]!.nodeId)["padding-right"]).toBe(padding + "pt");
      session.undo(); expect(session.snapshot().html).toBe(before);
      const bytes = exportDocx(session).bytes; session.dispose();
      session = open(bytes);
    }
    session.dispose();
  }
});

test("native direct inside borders survive row height recomposition, undo and editable export", () => {
  let session = createDocumentSession(openDocument(importDocx(rowHeightFixture(), {
    documentId: "native-row-strokes", unsupportedContent: "reject",
  }), { kind: "document" }));
  for (let pass = 0; pass < 3; pass++) {
    const table = session.snapshot().metadata.word!.tables[0]!.nodeId,
      cell = descendants(session, table).filter(node => node.tagName === "td")[3]!;
    expect(css(session, cell.nodeId)["border-top-width"]).toBe("1pt");
    expect(css(session, cell.nodeId)["border-left-width"]).toBe("1pt");
    const before = session.snapshot().html;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "native-row-strokes",
      baseRevision: session.snapshot().manifest.revision,
      operations: [{ op: "editWordTable", nodeId: table, change: { kind: "resize", rowIndex: 1, rowHeight: 70 } }],
    }));
    expect(css(session, cell.nodeId)["border-top-width"]).toBe("1pt");
    expect(css(session, cell.nodeId)["border-left-width"]).toBe("1pt");
    session.undo(); expect(session.snapshot().html).toBe(before);
    const bytes = exportDocx(session).bytes; session.dispose();
    session = createDocumentSession(openDocument(importDocx(bytes, { documentId: "native-row-strokes", unsupportedContent: "reject" }), { kind: "document" }));
  }
  session.dispose();
});
