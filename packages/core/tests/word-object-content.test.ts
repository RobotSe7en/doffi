import { expect, test } from "vitest";
import {
  createWordDocument,
  createDocumentSession,
  openDocument,
  wordDrawingTransforms,
} from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import {
  wordObjectContentFixture,
  wordObjectContentGraph,
} from "../../office/tests/word-object-content-fixture.js";

for (const native of [false, true])
  for (const tracked of [false, true])
    test(
      "semantic object graph allocates independent drawings, connections and owned stories " +
        native +
        " " +
        tracked,
      () => {
        const source = wordObjectContentFixture(native),
          graph = wordObjectContentGraph(source.session, source.paragraphId),
          documentId = "object-target";
        expect(JSON.stringify(graph)).not.toMatch(/PRIVATE-BODY|PRIVATE-TEXT-STORY/);
        const session = createDocumentSession(
          openDocument(createWordDocument({ documentId }), { kind: "document" }),
        );
        const imagePaths = new Set(
          [
            ...(graph.html + (graph.semantics.stories ?? []).map((s) => s.html).join("")).matchAll(
              /src="([^"]+)"/g,
            ),
          ].map((m) =>
            decodeURIComponent(
              new URL(
                m[1]!,
                "https://doffi.invalid/" + source.session.snapshot().manifest.entrypoint,
              ).pathname.slice(1),
            ),
          ),
        );
        const assets = source.session
          .snapshot()
          .manifest.assets.filter(
            (a) =>
              imagePaths.has(a.path) ||
              graph.semantics.drawings?.some(
                (d) =>
                  d.kind === "chart" && d.chart.native?.parts.some((p) => p.assetPath === a.path),
              ),
          )
          .map((descriptor) => ({
            descriptor,
            bytes: source.session.serialize().get(descriptor.path)!,
          }));
        const apply = (operations: DocumentOperation[]) =>
          session.commit(
            session.prepare(
              {
                protocolVersion: 1,
                documentId,
                baseRevision: session.snapshot().manifest.revision,
                operations,
              },
              operations.some((op) => op.op === "pasteWordInlineContent") ? { assets } : undefined,
            ),
          );
        apply([
          {
            op: "replaceText",
            nodeId: "paragraph-1",
            from: 0,
            to: 0,
            expectedText: "",
            text: "AB",
          },
          {
            op: "setWordReview",
            nodeId: "document-root",
            trackChanges: tracked,
            author: "Receiver",
          },
        ]);
        const before = session.snapshot();
        const operation = {
          op: "pasteWordInlineContent" as const,
          nodeId: "paragraph-1",
          from: 1,
          to: 1,
          expectedText: "",
          ...graph,
        };
        apply([operation]);
        const once = session.snapshot(),
          copied = once.metadata.word!.drawings!;
        expect(copied).toHaveLength(graph.semantics.drawings!.length);
        expect(new Set(copied.map((d) => d.nodeId)).size).toBe(copied.length);
        expect(
          copied.some((d) => graph.semantics.drawings!.some((s) => s.nodeId === d.nodeId)),
        ).toBe(false);
        const originalConnector = graph.semantics.drawings!.find((d) => d.kind === "connector")!,
          connector = copied.find((d) => d.kind === "connector")!;
        if (connector.kind === "connector" && originalConnector.kind === "connector") {
          expect(connector.connector.startNodeId).not.toBe(originalConnector.connector.startNodeId);
          expect(copied.some((d) => d.nodeId === connector.connector.startNodeId)).toBe(true);
        }
        for (const d of copied)
          if ("textStoryId" in d && d.textStoryId)
            expect(
              session.getNode(
                once.metadata.word!.stories.find((s) => s.storyId === d.textStoryId)!.rootNodeId,
              )!.text,
            ).toMatch(/OWNED|BOX-BODY/);
        const cropped = copied.find((d) => d.kind === "image" && d.crop);
        expect(cropped?.kind === "image" ? cropped.crop : undefined).toEqual({
          left: 0.2,
          top: 0.1,
          right: 0.1,
          bottom: 0.2,
        });
        const sourceTransforms = wordDrawingTransforms(
            source.session.snapshot().metadata.word!,
            source.session,
          ),
          targetTransforms = wordDrawingTransforms(once.metadata.word!, session);
        for (let i = 0; i < copied.length; i++)
          expect(targetTransforms.world(copied[i]!.nodeId)).toEqual(
            sourceTransforms.world(graph.semantics.drawings![i]!.nodeId),
          );
        apply([operation]);
        expect(session.snapshot().metadata.word!.drawings).toHaveLength(copied.length * 2);
        expect(
          session.snapshot().metadata.word!.stories.filter((s) => s.kind === "textBox"),
        ).toHaveLength(graph.semantics.stories!.length * 2);
        session.undo();
        expect(session.snapshot().html).toBe(once.html);
        let bytes = exportDocx(session).bytes;
        for (let pass = 0; pass < 2; pass++) {
          const reopened = openDocument(
              importDocx(bytes, { documentId, unsupportedContent: "reject" }),
              { kind: "document" },
            ),
            word = reopened.snapshot().metadata.word!;
          expect(word.drawings).toHaveLength(copied.length);
          const chart = word.drawings!.find((d) => d.kind === "chart")!;
          if (chart.kind === "chart")
            expect(chart.chart.series[0]!.values).toEqual(native ? [-7, 23, 8] : [1, 2, 3]);
          expect(word.fields.some((f) => f.instruction.trim() === "PAGE")).toBe(true);
          bytes = exportDocx(reopened).bytes;
        }
        if (tracked) {
          apply([{ op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" }]);
          expect(session.snapshot().metadata.word!.drawings).toHaveLength(0);
          expect(
            session.snapshot().metadata.word!.stories.filter((s) => s.kind === "textBox"),
          ).toHaveLength(0);
          expect(session.getNode("main-story")!.text).toBe("AB");
          expect(session.snapshot().manifest.assets).toEqual(before.manifest.assets);
          session.undo();
        }
        session.undo();
        expect(session.snapshot().html).toBe(before.html);
      },
    );

test("a connector alone retains fallback geometry and never binds to receiver identities", () => {
  const source = wordObjectContentFixture(),
    drawing = source.session
      .snapshot()
      .metadata.word!.drawings!.find((d) => d.kind === "connector")!;
  const documentId = "connector-copy",
    session = createDocumentSession(
      openDocument(createWordDocument({ documentId }), { kind: "document" }),
    );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId,
      baseRevision: 0,
      operations: [
        {
          op: "pasteWordInlineContent",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          html: source.session.getNode(drawing.nodeId)!.html,
          semantics: { drawings: [JSON.parse(JSON.stringify(drawing))] },
        },
      ],
    }),
  );
  const result = session.snapshot().metadata.word!.drawings![0]!;
  if (result.kind !== "connector") throw new Error("Missing connector");
  expect(result.connector).toEqual({ routing: "elbow" });
  expect(result.geometry).toEqual(drawing.geometry);
});

test("missing owned stories and foreign graph references roll back the complete paste", () => {
  const source = wordObjectContentFixture(),
    graph = wordObjectContentGraph(source.session, source.paragraphId),
    documentId = "bad-object-target",
    session = createDocumentSession(
      openDocument(createWordDocument({ documentId }), { kind: "document" }),
    );
  const before = session.snapshot();
  delete graph.semantics.stories;
  expect(() =>
    session.prepare({
      protocolVersion: 1,
      documentId,
      baseRevision: 0,
      operations: [
        {
          op: "pasteWordInlineContent",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          ...graph,
        },
      ],
    }),
  ).toThrow();
  expect(session.snapshot()).toEqual(before);
});

for (const kind of [
  "column",
  "line",
  "pie",
  "bar",
  "area",
  "doughnut",
  "scatter",
  "combo",
] as const)
  test("clipboard chart remains independently editable through native reopens: " + kind, () => {
    const documentId = "chart-copy-" + kind;
    let source = createDocumentSession(
      openDocument(createWordDocument({ documentId }), { kind: "document" }),
    );
    const apply = (operations: DocumentOperation[]) =>
      source.commit(
        source.prepare({
          protocolVersion: 1,
          documentId,
          baseRevision: source.snapshot().manifest.revision,
          operations,
        }),
      );
    apply([
      {
        op: "insertWordDrawing",
        nodeId: "paragraph-1",
        offset: 0,
        geometry: { x: 0, y: 0, width: 240, height: 160, rotation: 0, flipH: false, flipV: false },
        layout: { mode: "inline" },
        object: {
          kind: "chart",
          chart: {
            kind,
            legend: "bottom",
            showLabels: true,
            categories: ["A", "B", "C"],
            series: [
              {
                seriesId: "series",
                name: "CURRENT",
                values: [1, 2, 3],
                ...(kind === "scatter" ? { xValues: [2, 5, 9] } : {}),
                ...(kind === "combo" ? { kind: "column" as const } : {}),
              },
              ...(kind === "combo"
                ? [
                    {
                      seriesId: "secondary",
                      name: "B",
                      values: [2, 3, 1],
                      kind: "line" as const,
                      axis: "secondary" as const,
                    },
                  ]
                : []),
            ],
          },
        },
      },
    ]);
    source = createDocumentSession(
      openDocument(
        importDocx(exportDocx(source).bytes, { documentId, unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    );
    let chart = source.snapshot().metadata.word!.drawings![0]!;
    if (chart.kind !== "chart") throw new Error("Missing chart");
    apply([
      {
        op: "setWordChart",
        nodeId: chart.nodeId,
        changes: {
          series: chart.chart.series.map((series, index) => ({
            ...JSON.parse(JSON.stringify(series)),
            values: index ? [2, 3, 1] : [9, 10, -3],
          })),
        },
      },
    ]);
    chart = source.snapshot().metadata.word!.drawings![0]!;
    if (chart.kind !== "chart") throw new Error("Missing chart");
    const target = createDocumentSession(
      openDocument(createWordDocument({ documentId: "chart-target" }), { kind: "document" }),
    );
    const assets = source
      .snapshot()
      .manifest.assets.filter(
        (a) =>
          chart.kind === "chart" && chart.chart.native!.parts.some((p) => p.assetPath === a.path),
      )
      .map((descriptor) => ({ descriptor, bytes: source.serialize().get(descriptor.path)! }));
    target.commit(
      target.prepare(
        {
          protocolVersion: 1,
          documentId: "chart-target",
          baseRevision: 0,
          operations: [
            {
              op: "pasteWordInlineContent",
              nodeId: "paragraph-1",
              from: 0,
              to: 0,
              expectedText: "",
              html: source.getNode(chart.nodeId)!.html,
              semantics: { drawings: [JSON.parse(JSON.stringify(chart))] },
            },
          ],
        },
        { assets },
      ),
    );
    let bytes = exportDocx(target).bytes;
    for (let pass = 0; pass < 2; pass++) {
      const reopened = openDocument(
          importDocx(bytes, { documentId: "chart-target", unsupportedContent: "reject" }),
          { kind: "document" },
        ),
        copy = reopened.snapshot().metadata.word!.drawings![0]!;
      if (copy.kind !== "chart") throw new Error("Missing copied chart");
      expect(copy.chart.kind).toBe(kind);
      expect(copy.chart.series[0]!.values).toEqual([9, 10, -3]);

      bytes = exportDocx(reopened).bytes;
    }
  });

for (const decision of ["accept", "reject"] as const)
  test(
    "selected complete drawing cut couples all source objects to one native move: " + decision,
    () => {
      const { session, paragraphId } = wordObjectContentFixture(true),
        graph = wordObjectContentGraph(session, paragraphId),
        word = session.snapshot().metadata.word!;
      const inParagraph = (id: string) => {
        for (
          let current: string | null = id;
          current;
          current = session.getNode(current)!.parentNodeId
        )
          if (current === paragraphId) return true;
        return false;
      };
      const drawingIds = new Set(word.drawings!.map((d) => d.nodeId));
      const selected = word
        .drawings!.filter(
          (d) =>
            inParagraph(d.nodeId) &&
            (() => {
              for (
                let parent = session.getNode(d.nodeId)!.parentNodeId;
                parent && parent !== paragraphId;
                parent = session.getNode(parent)!.parentNodeId
              )
                if (drawingIds.has(parent)) return false;
              return true;
            })(),
        )
        .sort(
          (a, b) =>
            session.snapshot().html.indexOf('data-node-id="' + a.nodeId + '"') -
            session.snapshot().html.indexOf('data-node-id="' + b.nodeId + '"'),
        );
      const html = selected.map((d) => session.getNode(d.nodeId)!.html).join("");
      const apply = (operations: DocumentOperation[]) =>
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: session.snapshot().manifest.documentId,
            baseRevision: session.snapshot().manifest.revision,
            operations,
          }),
        );
      let root = session.getNode(paragraphId)!;
      while (root.parentNodeId) root = session.getNode(root.parentNodeId)!;
      apply([{ op: "setWordReview", nodeId: root.nodeId, trackChanges: true, author: "Receiver" }]);
      const before = session.snapshot(),
        cutId = "object-set-cut";
      apply([
        {
          op: "cutWordDrawings",
          nodeId: selected[0]!.nodeId,
          nodeIds: selected.map((d) => d.nodeId),
          cutId,
        },
      ]);
      apply([
        {
          op: "pasteWordInlineContent",
          nodeId: paragraphId,
          from: 0,
          to: 0,
          expectedText: "",
          html,
          semantics: graph.semantics,
          cutId,
        },
      ]);
      expect(session.snapshot().metadata.word!.moves).toHaveLength(2);
      const reopened = createDocumentSession(
        openDocument(
          importDocx(exportDocx(session).bytes, {
            documentId: "objects-moved",
            unsupportedContent: "reject",
          }),
          { kind: "document" },
        ),
      );
      const move = reopened.snapshot().metadata.word!.revisions.find((r) => r.kind === "moveTo")!;
      reopened.commit(
        reopened.prepare({
          protocolVersion: 1,
          documentId: "objects-moved",
          baseRevision: reopened.snapshot().manifest.revision,
          operations: [
            {
              op: "reviewWordRevision",
              nodeId: move.nodeId,
              revisionId: move.revisionId ?? move.nodeId,
              decision,
            },
          ],
        }),
      );
      expect(reopened.snapshot().metadata.word!.drawings).toHaveLength(
        before.metadata.word!.drawings!.length,
      );
      expect(reopened.snapshot().metadata.word!.revisions).toHaveLength(0);
      session.undo();
      session.undo();
      expect(session.snapshot().html).toBe(before.html);
    },
  );
