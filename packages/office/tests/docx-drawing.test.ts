import { expect, test } from "vitest";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { unzipSync, strFromU8, zipSync, strToU8 } from "fflate";
import { createWordDocument, openDocument, createDocumentSession } from "@doffi/core";
import type { DocumentOperation, WordDrawing } from "@doffi/core";
import { exportDocx, importDocx, importDocxWithReport } from "@doffi/office";
const geometry = { x: 24, y: 15, width: 120, height: 75, rotation: 25, flipH: true, flipV: false },
  layout = {
    mode: "floating" as const,
    horizontal: "margin" as const,
    vertical: "paragraph" as const,
    wrap: "square" as const,
    distances: { top: 2, right: 4, bottom: 6, left: 8 },
    zOrder: 9,
  };
test("authorized image watermarks cover first and alternating headers, reopen and release an unused source image on removal", () => {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "picture-watermark" }), { kind: "document" }),
    ),
    bytes = new Uint8Array(
      readFileSync(new URL("../../../examples/xlsx/reference/13-date-1904-1.png", import.meta.url)),
    ),
    path = "assets/watermark.png";
  session.commit(
    session.prepare(
      {
        protocolVersion: 1,
        documentId: "picture-watermark",
        baseRevision: 0,
        operations: [
          { op: "setWordOptions", nodeId: "document-root", evenAndOddHeaders: true },
          { op: "setWordSection", nodeId: "section-root-1", changes: { differentFirstPage: true } },
          {
            op: "setWordWatermark",
            nodeId: "section-root-1",
            watermark: {
              kind: "image",
              assetPath: path,
              alt: "水印",
              width: 200,
              height: 120,
              rotation: 0,
            },
          },
        ],
      },
      {
        assets: [
          {
            descriptor: {
              path,
              mimeType: "image/png",
              sha256: createHash("sha256").update(bytes).digest("hex"),
            },
            bytes,
          },
        ],
      },
    ),
  );
  expect(
    session.snapshot().metadata.word!.drawings!.filter((drawing) => drawing.watermark),
  ).toHaveLength(3);
  const reopened = openDocument(
    importDocx(exportDocx(session).bytes, {
      documentId: "picture-watermark-native",
      unsupportedContent: "reject",
    }),
    { kind: "document" },
  );
  expect(
    reopened
      .snapshot()
      .metadata.word!.drawings!.filter((drawing) => drawing.watermark && drawing.kind === "image"),
  ).toHaveLength(3);
  expect(reopened.snapshot().metadata.sections![0]!.differentFirstPage).toBe(true);
  expect(reopened.snapshot().metadata.word!.evenAndOddHeaders).toBe(true);
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "picture-watermark",
      baseRevision: 1,
      operations: [{ op: "setWordWatermark", nodeId: "section-root-1", watermark: null }],
    }),
  );
  expect(session.serialize().has(path)).toBe(false);
  session.undo();
  expect(session.serialize().get(path)).toEqual(bytes);
});
function fixture() {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "drawings" }), { kind: "document" }),
  );
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "drawings",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  apply([
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 0,
      geometry,
      layout,
      object: { kind: "shape", preset: "roundRect", text: "图形文字😀" },
      properties: {
        "--office-fill-kind": "solid",
        "--office-fill-color": "#22aa66",
        "--office-line-kind": "solid",
        "--office-line-color": "#224466",
        "--office-line-width": "2pt",
      },
    },
  ]);
  const first = session.snapshot().metadata.word!.drawings![0]!;
  apply([
    {
      op: "setWordDrawingText",
      nodeId: first.nodeId,
      text: {
        insets: { top: 2, right: 4, bottom: 6, left: 8 },
        vertical: "middle",
        wrap: false,
        direction: "vertical",
        rotation: 10,
      },
    },
  ]);
  apply([
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 0,
      geometry: { ...geometry, x: 220, rotation: 0, flipH: false },
      layout,
      object: { kind: "textBox", text: "独立文本框" },
    },
  ]);
  return { session, apply, first };
}
test("editable floating shapes and Word text boxes preserve geometry, paint, text and layout through two native reopens", () => {
  const { session } = fixture();
  let model = session;
  for (let pass = 0; pass < 2; pass++) {
    const exported = exportDocx(model),
      files = unzipSync(exported.bytes),
      xml = strFromU8(files["word/document.xml"]!);
    expect(xml).toContain("wps:txbx");
    expect(xml).toContain("wp:anchor");
    expect(xml).toContain("22aa66");
    expect(xml).not.toContain("transient:native");
    expect(xml).not.toContain("<wps:wsp><wps:cNvPr");
    const imported = importDocxWithReport(exported.bytes, {
      documentId: "drawings",
      unsupportedContent: "reject",
    });
    expect(imported.report.preserved).toHaveLength(0);
    model = createDocumentSession(openDocument(imported.files, { kind: "document" }));
    const drawings = model.snapshot().metadata.word!.drawings!;
    expect(drawings).toHaveLength(2);
    const shape = drawings.find((item) => item.kind === "shape")!;
    expect(shape.geometry).toEqual(geometry);
    expect(shape.layout).toMatchObject(layout);
    expect(shape.kind === "shape" && shape.text).toMatchObject({
      insets: { top: 2, right: 4, bottom: 6, left: 8 },
      vertical: "middle",
      wrap: false,
      direction: "vertical",
      rotation: 10,
    });
    const story = shape.kind === "shape" ? shape.textStoryId! : "";
    expect(
      model.getNode(
        model.snapshot().metadata.word!.stories.find((item) => item.storyId === story)!.rootNodeId,
      )!.text,
    ).toBe("图形文字😀");
  }
  const story = model.snapshot().metadata.word!.stories.find((item) => item.kind === "textBox")!,
    paragraph = model.getNode(story.rootNodeId)!.childNodeIds[0]!;
  model.commit(
    model.prepare({
      protocolVersion: 1,
      documentId: "drawings",
      baseRevision: 0,
      operations: [
        { op: "replaceText", nodeId: paragraph, from: 0, to: 0, expectedText: "", text: "edited" },
      ],
    }),
  );
  expect(strFromU8(unzipSync(exportDocx(model).bytes)["word/document.xml"]!)).toContain("edited");
});
test("group coordinates and bound connectors reopen as editable DrawingML objects", () => {
  const { session, apply } = fixture(),
    objects = session.snapshot().metadata.word!.drawings!;
  apply([
    { op: "groupWordDrawings", nodeId: "paragraph-1", nodeIds: objects.map((item) => item.nodeId) },
  ]);
  const group = session.snapshot().metadata.word!.drawings!.find((item) => item.kind === "group")!;
  apply([
    {
      op: "setWordDrawingGeometry",
      nodeId: group.nodeId,
      geometry: { ...group.geometry, rotation: 30, flipH: true },
    },
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 0,
      geometry,
      layout,
      object: {
        kind: "connector",
        connector: {
          routing: "elbow",
          startNodeId: objects[0]!.nodeId,
          endNodeId: objects[1]!.nodeId,
          startSite: 3,
          endSite: 1,
        },
      },
    },
  ]);
  let bytes = exportDocx(session).bytes;
  for (let pass = 0; pass < 2; pass++) {
    const reopened = openDocument(
        importDocx(bytes, { documentId: "group", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
      word = reopened.snapshot().metadata.word!;
    expect(word.drawings).toHaveLength(4);
    expect(word.drawings!.find((item) => item.kind === "group")!.geometry.rotation).toBe(30);
    const connector = word.drawings!.find((item) => item.kind === "connector") as Extract<
      WordDrawing,
      { kind: "connector" }
    >;
    expect(connector.connector.startNodeId).toBeDefined();
    expect(connector.connector.endNodeId).toBeDefined();
    bytes = exportDocx(reopened).bytes;
  }
});
test("new authorized image bytes export with editable crop, rotation and native relationships", () => {
  const { session } = fixture(),
    bytes = new Uint8Array(
      readFileSync(new URL("../../../examples/xlsx/reference/13-date-1904-1.png", import.meta.url)),
    ),
    path = "assets/upload.png";
  session.commit(
    session.prepare(
      {
        protocolVersion: 1,
        documentId: "drawings",
        baseRevision: session.snapshot().manifest.revision,
        operations: [
          {
            op: "insertWordDrawing",
            nodeId: "paragraph-1",
            offset: 0,
            geometry,
            layout: { mode: "inline" },
            object: {
              kind: "image",
              assetPath: path,
              alt: "图像",
              crop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.1 },
            },
          },
        ],
      },
      {
        assets: [
          {
            descriptor: {
              path,
              mimeType: "image/png",
              sha256: createHash("sha256").update(bytes).digest("hex"),
            },
            bytes,
          },
        ],
      },
    ),
  );
  const output = exportDocx(session).bytes,
    reopened = openDocument(
      importDocx(output, { documentId: "image", unsupportedContent: "reject" }),
      { kind: "document" },
    ),
    image = reopened.snapshot().metadata.word!.drawings!.find((item) => item.kind === "image")!;
  expect(image.geometry.rotation).toBe(25);
  expect(image.kind === "image" && image.crop).toEqual({
    left: 0.1,
    top: 0.2,
    right: 0.3,
    bottom: 0.1,
  });
  expect(
    reopened
      .snapshot()
      .manifest.assets.some(
        (item) => item.sha256 === createHash("sha256").update(bytes).digest("hex"),
      ),
  ).toBe(true);
});
test("unsupported drawing geometry preserves the entire native graphic and leaves neighbouring text editable", () => {
  const { session } = fixture(),
    files = unzipSync(exportDocx(session).bytes);
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replace('prst="roundRect"', 'prst="cloud"'),
  );
  const imported = importDocxWithReport(zipSync(files), { documentId: "opaque" });
  expect(imported.report.preserved.length).toBeGreaterThan(0);
  const output = exportDocx(openDocument(imported.files, { kind: "document" }));
  expect(strFromU8(unzipSync(output.bytes)["word/document.xml"]!)).toContain('prst="cloud"');
});
