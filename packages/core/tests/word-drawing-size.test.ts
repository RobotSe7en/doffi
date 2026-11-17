import { expect, test } from "vitest";
import {
  createWordDocument,
  createDocumentSession,
  openDocument,
  wordDrawingDimensions,
  wordDrawingTransforms,
} from "@doffi/core";
import type { DocumentOperation, WordDrawing, PageSetup, DocumentSection } from "@doffi/core";
const geometry = { x: 10, y: 12, width: 160, height: 100, rotation: 0, flipH: false, flipV: false };
const setup: PageSetup = {
  unit: "pt",
  width: 420,
  height: 500,
  margins: { left: 36, right: 48, top: 24, bottom: 60 },
};
const drawing: WordDrawing = {
  kind: "shape",
  nodeId: "shape",
  preset: "rect",
  geometry,
  layout: { mode: "floating" },
};
test("all native relative reference regions use physical page dimensions, independently per axis", () => {
  const widths = {
      page: 420,
      margin: 336,
      leftMargin: 36,
      rightMargin: 48,
      insideMargin: 36,
      outsideMargin: 48,
    },
    heights = {
      page: 500,
      margin: 416,
      topMargin: 24,
      bottomMargin: 60,
      insideMargin: 500,
      outsideMargin: 500,
    };
  for (const [relativeFrom, value] of Object.entries(widths)) {
    const measured = wordDrawingDimensions(
      {
        ...drawing,
        layout: {
          mode: "floating",
          relativeWidth: {
            relativeFrom: relativeFrom as NonNullable<
              WordDrawing["layout"]["relativeWidth"]
            >["relativeFrom"],
            percent: 50,
          },
        },
      },
      { index: 0, setup },
    );
    expect(measured.width).toBe(value / 2);
    expect(measured.height).toBe(100);
  }
  for (const [relativeFrom, value] of Object.entries(heights)) {
    const measured = wordDrawingDimensions(
      {
        ...drawing,
        layout: {
          mode: "floating",
          relativeHeight: {
            relativeFrom: relativeFrom as NonNullable<
              WordDrawing["layout"]["relativeHeight"]
            >["relativeFrom"],
            percent: 125,
          },
        },
      },
      { index: 0, setup },
    );
    expect(measured.height).toBe(value * 1.25);
    expect(measured.width).toBe(160);
  }
  expect(
    wordDrawingDimensions(
      {
        ...drawing,
        layout: {
          mode: "floating",
          relativeWidth: { relativeFrom: "insideMargin", percent: 100 },
          relativeHeight: { relativeFrom: "outsideMargin", percent: 50 },
        },
      },
      { index: 1, setup },
      {
        sectionId: "section",
        rootNodeId: "section-root",
        gutterPosition: "top",
      } as DocumentSection,
    ),
  ).toMatchObject({ width: 48, height: 30 });
  expect(
    wordDrawingDimensions(
      {
        ...drawing,
        layout: { mode: "floating", relativeWidth: { relativeFrom: "page", percent: 0 } },
      },
      { index: 0, setup },
    ),
  ).toEqual(geometry);
});
test("relative sizing, copying, fixed conversion and invalid/stale batches share one history boundary", () => {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "size-transactions", pageSetup: setup }), {
        kind: "document",
      }),
    ),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "size-transactions",
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
      layout: { mode: "floating" },
      object: { kind: "shape", preset: "rect", text: "尺寸文字😀" },
    },
  ]);
  const first = session.snapshot().metadata.word!.drawings![0]!,
    before = session.snapshot();
  const layout = {
    ...first.layout,
    wrapPolygon: first.layout.wrapPolygon?.map((point) => ({ ...point })),
    relativeWidth: { relativeFrom: "page" as const, percent: 50 },
    relativeHeight: { relativeFrom: "margin" as const, percent: 25 },
  };
  const stale = session.prepare({
    protocolVersion: 1,
    documentId: "size-transactions",
    baseRevision: 1,
    operations: [{ op: "setWordDrawingLayout", nodeId: first.nodeId, layout }],
  });
  apply([{ op: "setWordDrawingLayout", nodeId: first.nodeId, layout }]);
  expect(() => session.commit(stale)).toThrowError(
    expect.objectContaining({ code: "RevisionConflict" }),
  );
  expect(session.snapshot().metadata.word!.drawings![0]!.geometry).toEqual(geometry);
  const current = session.snapshot();
  for (const percent of [-1, Infinity, 2147483.648]) {
    expect(() =>
      apply([
        {
          op: "setWordDrawingGeometry",
          nodeId: first.nodeId,
          geometry: { ...geometry, width: 999 },
        },
        {
          op: "setWordDrawingLayout",
          nodeId: first.nodeId,
          layout: { ...layout, relativeWidth: { relativeFrom: "page", percent } },
        },
      ]),
    ).toThrow();
    expect(session.snapshot()).toEqual(current);
  }
  expect(() =>
    apply([
      { op: "setWordDrawingLayout", nodeId: first.nodeId, layout: { ...layout, mode: "inline" } },
    ]),
  ).toThrow();
  session.undo();
  expect(session.snapshot().metadata).toEqual(before.metadata);
  session.redo();
  apply([{ op: "copyWordDrawing", nodeId: first.nodeId }]);
  expect(session.snapshot().metadata.word!.drawings!.map((d) => d.layout)).toEqual([
    layout,
    layout,
  ]);
  apply([{ op: "setWordDrawingLayout", nodeId: first.nodeId, layout: { mode: "floating" } }]);
  expect(session.snapshot().metadata.word!.drawings![0]!.layout.relativeWidth).toBeUndefined();
  session.undo();
  expect(openDocument(session.serialize(), { kind: "document" }).snapshot()).toEqual(
    session.snapshot(),
  );
});
test("bound connector projection follows relative target dimensions without changing source", () => {
  const objects: WordDrawing[] = [
    drawing,
    { ...drawing, nodeId: "target", geometry: { ...geometry, x: 300 } },
    {
      kind: "connector",
      nodeId: "connector",
      geometry: { ...geometry, x: 170, y: 62, width: 130, height: 0 },
      layout: { mode: "floating" },
      connector: {
        routing: "straight",
        startNodeId: "shape",
        endNodeId: "target",
        startSite: 3,
        endSite: 1,
      },
    },
  ];
  const source = {
    getNode: (id: string) => ({
      parentNodeId: id === "p" ? null : "p",
      tagName: id === "p" ? "p" : "span",
    }),
  };
  const word = openDocument(createWordDocument({ documentId: "connector-sizes" }), {
    kind: "document",
  }).snapshot().metadata.word!;
  const transforms = wordDrawingTransforms({ ...word, drawings: objects }, source);
  const base = transforms.connectorGeometry("connector");
  const projected = wordDrawingTransforms(
    {
      ...word,
      drawings: objects.map((d) =>
        d.nodeId === "shape" ? { ...d, geometry: { ...d.geometry, width: 210 } } : d,
      ),
    },
    source,
  ).connectorGeometry("connector");
  expect(projected.x - base.x).toBeCloseTo(50);
  expect(projected.width).toBeCloseTo(base.width - 50);
  expect(objects[0]!.geometry.width).toBe(160);
});
test("grouping relative drawings requires explicit measured extents and remains an atomic batch", () => {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "size-groups", pageSetup: setup }), {
        kind: "document",
      }),
    ),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "size-groups",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply(
    [0, 200].map((x) => ({
      op: "insertWordDrawing" as const,
      nodeId: "paragraph-1",
      offset: 0,
      geometry: { ...geometry, x },
      layout: {
        mode: "floating" as const,
        relativeWidth: { relativeFrom: "page" as const, percent: 25 },
      },
      object: { kind: "shape" as const, preset: "rect" as const },
    })),
  );
  const before = session.snapshot(),
    ids = before.metadata.word!.drawings!.map((d) => d.nodeId);
  expect(() => apply([{ op: "groupWordDrawings", nodeId: "paragraph-1", nodeIds: ids }])).toThrow(
    /fixed extents/,
  );
  expect(session.snapshot()).toEqual(before);
  apply([
    ...before.metadata.word!.drawings!.flatMap((d) => [
      {
        op: "setWordDrawingGeometry" as const,
        nodeId: d.nodeId,
        geometry: wordDrawingDimensions(d, { index: 0, setup }),
      },
      {
        op: "setWordDrawingLayout" as const,
        nodeId: d.nodeId,
        layout: { mode: "floating" as const },
      },
    ]),
    { op: "groupWordDrawings", nodeId: "paragraph-1", nodeIds: ids },
  ]);
  const group = session.snapshot().metadata.word!.drawings!.find((d) => d.kind === "group")!;
  expect(group.geometry.width).toBe(305);
  session.undo();
  expect(session.snapshot().metadata).toEqual(before.metadata);
});
for (const kind of ["shape", "textBox"] as const)
  test("table insertion in a relative text story uses its own text frame: " + kind, () => {
    const session = createDocumentSession(
        openDocument(createWordDocument({ documentId: "size-story", pageSetup: setup }), {
          kind: "document",
        }),
      ),
      apply = (operations: DocumentOperation[]) =>
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: "size-story",
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
        layout: { mode: "floating", relativeWidth: { relativeFrom: "page", percent: 50 } },
        object:
          kind === "shape"
            ? { kind: "shape", preset: "rect", text: "对象文字" }
            : { kind: "textBox", text: "对象文字" },
      },
    ]);
    const d = session.snapshot().metadata.word!.drawings![0]!,
      story = session
        .snapshot()
        .metadata.word!.stories.find((s) => "textStoryId" in d && s.storyId === d.textStoryId)!,
      p = session.getNode(story.rootNodeId)!.childNodeIds[0]!;
    apply([{ op: "insertWordTable", nodeId: p, position: "after", rows: 1, columns: 2 }]);
    expect(
      session.snapshot().metadata.word!.tables[0]!.grid.reduce((a, b) => a + b, 0),
    ).toBeCloseTo(195.6);
  });
