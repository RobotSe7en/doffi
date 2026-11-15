import { expect, test } from "vitest";
import { createWordDocument, openDocument, createDocumentSession } from "@doffi/core";
import type { DocumentOperation, WordDrawing } from "@doffi/core";
const geometry = { x: 30, y: 20, width: 100, height: 70, rotation: 0, flipH: false, flipV: false };
function create() {
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
  return { session, apply, drawings: () => session.snapshot().metadata.word!.drawings! };
}
const insert = (text: string, x = 30): DocumentOperation => ({
  op: "insertWordDrawing",
  nodeId: "paragraph-1",
  offset: 1,
  object: { kind: "shape", preset: "roundRect", text },
  geometry: { ...geometry, x },
  layout: { mode: "floating", wrap: "none" },
});
test("insert drawings at precise UTF-16 positions and retain independent text ownership and atomic undo", () => {
  const { session, apply, drawings } = create();
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "前😀后" },
  ]);
  const before = session.snapshot();
  apply([insert("图形文字")]);
  expect(session.getNode("paragraph-1")!.text).toBe("前\ufffc😀后");
  const object = drawings()[0]!;
  expect("textStoryId" in object && session.getNode(object.textStoryId!)!.text).toBe("图形文字");
  expect(openDocument(session.serialize(), { kind: "document" }).snapshot()).toEqual(
    session.snapshot(),
  );
  session.undo();
  expect(session.snapshot().html).toBe(before.html);
  session.redo();
  expect(drawings()).toHaveLength(1);
  expect(() =>
    apply([
      {
        op: "setWordDrawingGeometry",
        nodeId: object.nodeId,
        geometry: { ...geometry, width: 150 },
      },
      { op: "setStyle", nodeId: object.nodeId, property: "width", value: "5pt" },
    ]),
  ).toThrow();
  expect(drawings()[0]!.geometry.width).toBe(100);
});
test("group, move, copy and ungroup retain source IDs, owned stories and bound connectors", () => {
  const { session, apply, drawings } = create();
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
    insert("A"),
    insert("B", 250),
  ]);
  const [a, b] = drawings();
  apply([
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 0,
      object: {
        kind: "connector",
        connector: {
          routing: "straight",
          startNodeId: a!.nodeId,
          endNodeId: b!.nodeId,
          startSite: 3,
          endSite: 1,
        },
      },
      geometry,
      layout: { mode: "floating", wrap: "none" },
    },
  ]);
  const connector = drawings().find((item) => item.kind === "connector")!;
  expect(connector.geometry.width).toBeGreaterThan(100);
  apply([{ op: "groupWordDrawings", nodeId: "paragraph-1", nodeIds: [a!.nodeId, b!.nodeId] }]);
  const group = drawings().find((item) => item.kind === "group")!;
  const start = drawings().find((item) => item.nodeId === connector.nodeId)!.geometry.x;
  apply([
    {
      op: "setWordDrawingGeometry",
      nodeId: group.nodeId,
      geometry: { ...group.geometry, x: group.geometry.x + 25 },
    },
  ]);
  expect(drawings().find((item) => item.nodeId === connector.nodeId)!.geometry.x).toBeCloseTo(
    start + 25,
  );
  apply([{ op: "copyWordDrawing", nodeId: group.nodeId }]);
  expect(drawings().filter((item) => item.kind === "group")).toHaveLength(2);
  expect(
    session.snapshot().metadata.word!.stories.filter((story) => story.kind === "textBox"),
  ).toHaveLength(4);
  apply([{ op: "ungroupWordDrawing", nodeId: group.nodeId }]);
  expect(drawings().find((item) => item.nodeId === a!.nodeId)!.geometry.x).toBeCloseTo(55);
  apply([{ op: "removeNode", nodeId: "paragraph-1" }]);
  expect(drawings()).toHaveLength(0);
  expect(
    session.snapshot().metadata.word!.stories.filter((story) => story.kind === "textBox"),
  ).toHaveLength(0);
  session.undo();
  expect(drawings().length).toBeGreaterThan(0);
});
test("deleting a bound target detaches the endpoint without discarding its last native position", () => {
  const { apply, drawings } = create();
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
    insert("A"),
  ]);
  const a = drawings()[0]!;
  apply([
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 0,
      geometry,
      layout: { mode: "floating", wrap: "none" },
      object: { kind: "connector", connector: { routing: "elbow", startNodeId: a.nodeId } },
    },
  ]);
  const connector = drawings()[1] as Extract<WordDrawing, { kind: "connector" }>,
    last = connector.geometry;
  apply([{ op: "removeWordDrawing", nodeId: a.nodeId }]);
  const next = drawings()[0] as typeof connector;
  expect(next.connector.startNodeId).toBeUndefined();
  expect(next.geometry).toEqual(last);
});
test("ungrouping a rotated child under nonuniform scaling retains an exact native coordinate carrier", () => {
  const { session, apply, drawings } = create();
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
    insert("A"),
    insert("B", 250),
  ]);
  const ids = drawings().map((item) => item.nodeId);
  apply([{ op: "groupWordDrawings", nodeId: "paragraph-1", nodeIds: ids }]);
  const group = drawings().find((item) => item.kind === "group")!;
  apply([
    {
      op: "setWordDrawingGeometry",
      nodeId: ids[0]!,
      geometry: { ...drawings().find((item) => item.nodeId === ids[0])!.geometry, rotation: 35 },
    },
    {
      op: "setWordDrawingGeometry",
      nodeId: group.nodeId,
      geometry: { ...group.geometry, width: group.geometry.width * 2, rotation: 23 },
    },
  ]);
  const before = drawings().find((item) => item.nodeId === ids[0])!.geometry;
  apply([{ op: "ungroupWordDrawing", nodeId: group.nodeId }]);
  expect(drawings().find((item) => item.nodeId === ids[0])!.geometry).toEqual(before);
  const carrier = drawings().find((item) => item.kind === "group")!;
  expect(session.getNode(carrier.nodeId)!.childNodeIds).toEqual([ids[0]]);
  expect(carrier.geometry.width).toBe(group.geometry.width * 2);
  expect(openDocument(session.serialize(), { kind: "document" }).snapshot()).toEqual(
    session.snapshot(),
  );
});
