import { readFileSync } from "node:fs";
import { zipSync, unzipSync, strFromU8, strToU8 } from "fflate";
import {
  createDocumentSession,
  createWordDocument,
  createAssetDescriptor,
  openDocument,
} from "@doffi/core";
import type { DocumentOperation, WordContentSemantics, DocumentEditingSession } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";

export function wordObjectContentFixture(native = false): {
  session: DocumentEditingSession;
  paragraphId: string;
} {
  const documentId = "object-clipboard-source";
  let session = createDocumentSession(
    openDocument(createWordDocument({ documentId }), { kind: "document" }),
  );
  const apply = (
    operations: DocumentOperation[],
    options?: Parameters<typeof session.prepare>[1],
  ) =>
    session.commit(
      session.prepare(
        {
          protocolVersion: 1,
          documentId,
          baseRevision: session.snapshot().manifest.revision,
          operations,
        },
        options,
      ),
    );
  const geometry = { x: 10, y: 15, width: 90, height: 60, rotation: 15, flipH: true, flipV: false },
    layout = {
      mode: "floating" as const,
      horizontal: "margin" as const,
      vertical: "paragraph" as const,
      wrap: "square" as const,
      distances: { top: 1, right: 2, bottom: 3, left: 4 },
      zOrder: 7,
    };
  apply([
    {
      op: "replaceText",
      nodeId: "paragraph-1",
      from: 0,
      to: 0,
      expectedText: "",
      text: "LEFTRIGHT",
    },
    {
      op: "insertFragment",
      nodeId: "section-root-1",
      position: "end",
      html: "<p>PRIVATE-BODY</p>",
    },
  ]);
  const privateP = session.getNode("section-root-1")!.childNodeIds.at(-1)!;
  const bytes = new Uint8Array(
      readFileSync(new URL("../../../examples/xlsx/reference/13-date-1904-1.png", import.meta.url)),
    ),
    path = "assets/selected-picture.png";
  apply(
    [
      {
        op: "insertWordDrawing",
        nodeId: "paragraph-1",
        offset: 4,
        geometry,
        layout,
        object: { kind: "shape", preset: "roundRect", text: "OWNED😀" },
        properties: {
          "--office-fill-kind": "solid",
          "--office-fill-color": "#22aa66",
          "--office-line-width": "2pt",
        },
      },
      {
        op: "insertWordDrawing",
        nodeId: "paragraph-1",
        offset: 4,
        geometry: { ...geometry, x: 180, rotation: 0 },
        layout,
        object: { kind: "textBox", text: "BOX-BODY" },
      },
      {
        op: "insertWordDrawing",
        nodeId: "paragraph-1",
        offset: 4,
        geometry: { ...geometry, x: 290 },
        layout,
        object: {
          kind: "image",
          assetPath: path,
          alt: "Selected image",
          crop: { left: 0.2, top: 0.1, right: 0.1, bottom: 0.2 },
        },
      },
    ],
    { assets: [{ descriptor: createAssetDescriptor(path, "image/png", bytes), bytes }] },
  );
  const [a, b, image] = session.snapshot().metadata.word!.drawings!;
  apply([
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 4,
      geometry,
      layout,
      object: {
        kind: "connector",
        connector: {
          routing: "elbow",
          startNodeId: a!.nodeId,
          endNodeId: b!.nodeId,
          startSite: 3,
          endSite: 1,
        },
      },
    },
  ]);
  const connector = session
    .snapshot()
    .metadata.word!.drawings!.find((d) => d.kind === "connector")!;
  apply([{ op: "groupWordDrawings", nodeId: "paragraph-1", nodeIds: [a!.nodeId, b!.nodeId] }]);
  const inner = session.snapshot().metadata.word!.drawings!.find((d) => d.kind === "group")!;
  apply([
    {
      op: "groupWordDrawings",
      nodeId: "paragraph-1",
      nodeIds: [inner.nodeId, image!.nodeId, connector.nodeId],
    },
  ]);
  const outer = session
    .snapshot()
    .metadata.word!.drawings!.filter((d) => d.kind === "group")
    .at(-1)!;
  apply([
    {
      op: "setWordDrawingGeometry",
      nodeId: outer.nodeId,
      geometry: {
        ...outer.geometry,
        width: outer.geometry.width * 1.2,
        height: outer.geometry.height * 0.8,
        rotation: 21,
      },
    },
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 4,
      geometry: { ...geometry, x: 0, y: 0, width: 240, height: 150, rotation: 0, flipH: false },
      layout: { mode: "inline" },
      object: {
        kind: "chart",
        chart: {
          kind: "combo",
          title: "CURRENT-CHART",
          legend: "bottom",
          showLabels: true,
          categories: ["甲", "乙", "丙"],
          series: [
            { seriesId: "a", name: "A", values: [1, 2, 3], kind: "column" },
            { seriesId: "b", name: "B", values: [3, 2, 1], kind: "line", axis: "secondary" },
          ],
        },
      },
    },
    {
      op: "insertWordDrawing",
      nodeId: privateP,
      offset: 0,
      geometry,
      layout,
      object: { kind: "textBox", text: "PRIVATE-TEXT-STORY" },
    },
  ]);
  const owner = session.snapshot().metadata.word!.drawings!.find((d) => d.nodeId === a!.nodeId)!;
  if (!("textStoryId" in owner) || !owner.textStoryId) throw new Error("Missing owned story");
  const story = session
      .snapshot()
      .metadata.word!.stories.find((s) => s.storyId === owner.textStoryId)!,
    paragraph = session.getNode(story.rootNodeId)!.childNodeIds[0]!;
  apply([
    { op: "insertWordField", nodeId: paragraph, offset: 0, instruction: " PAGE ", cachedText: "7" },
    {
      op: "insertWordDrawing",
      nodeId: paragraph,
      offset: 0,
      geometry: { ...geometry, x: 0, y: 0, width: 20, height: 15, rotation: 0 },
      layout: { mode: "inline" },
      object: { kind: "image", assetPath: path, alt: "Owned image" },
    },
  ]);
  if (native) {
    const files = unzipSync(exportDocx(session).bytes),
      bookPath = Object.keys(files).find((p) => p.endsWith(".xlsx"))!,
      book = unzipSync(files[bookPath]!);
    const sheet = Object.keys(book).find((p) => /^xl\/worksheets\/.*xml$/.test(p))!;
    book[sheet] = strToU8(
      strFromU8(book[sheet]!).replace(
        "</sheetData>",
        '<row r="99"><c r="A99" t="inlineStr"><is><t>PRIVATE-WORKBOOK-CELL</t></is></c></row></sheetData>',
      ),
    );
    files[bookPath] = zipSync(book);
    session = createDocumentSession(
      openDocument(importDocx(zipSync(files), { documentId, unsupportedContent: "reject" }), {
        kind: "document",
      }),
    );
    const chart = session.snapshot().metadata.word!.drawings!.find((d) => d.kind === "chart")!;
    if (chart.kind !== "chart") throw new Error("Missing chart");
    apply([
      {
        op: "setWordChart",
        nodeId: chart.nodeId,
        changes: {
          series: chart.chart.series.map((s, i) => ({
            ...JSON.parse(JSON.stringify(s)),
            values: i ? [4, 5, 6] : [-7, 23, 8],
          })),
        },
      },
    ]);
  }
  const paragraphId = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
    .map((m) => session.getNode(m[1]!)!)
    .find(
      (n) => n.tagName === "p" && n.text.startsWith("LEFT") && n.text.endsWith("RIGHT"),
    )!.nodeId;
  return { session, paragraphId };
}

/** Test transport assembled from actual selected source ownership, with no
 * receiver lookup or whole-document metadata fallback. */
export function wordObjectContentGraph(
  source: DocumentEditingSession,
  paragraphId: string,
): { html: string; semantics: WordContentSemantics } {
  const word = source.snapshot().metadata.word!,
    ids = new Set<string>(),
    storyIds = new Set<string>(),
    pending = [paragraphId];
  while (pending.length) {
    const id = pending.pop()!;
    if (ids.has(id)) continue;
    ids.add(id);
    pending.push(...source.getNode(id)!.childNodeIds);
    const drawing = word.drawings?.find((d) => d.nodeId === id);
    if (
      drawing &&
      "textStoryId" in drawing &&
      drawing.textStoryId &&
      !storyIds.has(drawing.textStoryId)
    ) {
      storyIds.add(drawing.textStoryId);
      pending.push(word.stories.find((s) => s.storyId === drawing.textStoryId)!.rootNodeId);
    }
  }
  return {
    html: source
      .getNode(paragraphId)!
      .childNodeIds.filter(
        (id) => source.getNode(id)!.attributes["data-word-paragraph-mark"] !== "true",
      )
      .map((id) => source.getNode(id)!.html)
      .join(""),
    semantics: JSON.parse(
      JSON.stringify({
        drawings: word.drawings?.filter((d) => ids.has(d.nodeId)),
        fields: word.fields.filter((f) => ids.has(f.nodeId)),
        stories: word.stories
          .filter((s) => storyIds.has(s.storyId))
          .map((s) => ({ ...s, html: source.getNode(s.rootNodeId)!.html })),
      }),
    ) as WordContentSemantics,
  };
}
