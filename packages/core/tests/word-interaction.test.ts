import { expect, test } from "vitest";
import {
  createDocumentSession,
  createWordDocument,
  openDocument,
  parseWordMath,
} from "@doffi/core";
import type { DocumentOperation, EditProposal, WordChart } from "@doffi/core";
function setup() {
  const session = createDocumentSession(
    openDocument(createWordDocument({ documentId: "scoped-word" }), { kind: "document" }),
  );
  const apply = (operations: DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "scoped-word",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  const chart: WordChart = {
    kind: "column",
    legend: "none",
    showLabels: false,
    categories: ["一"],
    series: [{ seriesId: "s", name: "收入", values: [1] }],
  };
  apply([
    {
      op: "replaceText",
      nodeId: "paragraph-1",
      from: 0,
      to: 0,
      expectedText: "",
      text: "选中正文",
    },
    { op: "insertFragment", nodeId: "paragraph-1", position: "after", html: "<p>私有段落</p>" },
    {
      op: "insertWordEquation",
      nodeId: "paragraph-1",
      offset: 0,
      display: "inline",
      math: parseWordMath("x+1"),
    },
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 0,
      geometry: { x: 0, y: 0, width: 300, height: 200, rotation: 0, flipH: false, flipV: false },
      layout: { mode: "inline" },
      object: { kind: "chart", chart },
    },
    {
      op: "insertWordDrawing",
      nodeId: "paragraph-1",
      offset: 0,
      geometry: { x: 0, y: 0, width: 100, height: 40, rotation: 0, flipH: false, flipV: false },
      layout: { mode: "inline" },
      object: { kind: "textBox", text: "附属文字" },
    },
  ]);
  const privateNode = session.getNode("section-root-1")!.childNodeIds.at(-1)!;
  apply([
    {
      op: "insertWordEquation",
      nodeId: privateNode,
      offset: 0,
      display: "inline",
      math: parseWordMath("SECRET+42"),
    },
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "a",
        documentId: "scoped-word",
        baseRevision: 1,
        status: "open",
        anchor: { kind: "node", nodeId: "paragraph-1" },
        instruction: "修改图表",
      },
    },
  ]);
  apply([
    {
      op: "addWordCommentReply",
      nodeId: "paragraph-1",
      annotationId: "a",
      html: "<p>批注回复</p>",
    },
  ]);
  return { session, apply, privateNode };
}
test("Word request context includes owned stories and current objects while excluding unrelated semantic bodies", () => {
  const { session, privateNode } = setup(),
    request = session.buildInteractionRequest({ annotationIds: ["a"] });
  const serialized = JSON.stringify(request);
  expect(serialized).toContain("附属文字");
  expect(serialized).toContain("批注回复");
  expect(serialized).not.toContain("SECRET");
  expect(request.context.blocks.some((block) => block.nodeId === privateNode)).toBe(false);
  expect(request.context.metadata.word!.equations).toHaveLength(1);
  expect(request.context.metadata.word!.drawings).toHaveLength(2);
  const drawing = request.context.metadata.word!.drawings!.find(
    (drawing) => drawing.kind === "chart",
  )!;
  const proposal: EditProposal = {
    protocolVersion: 1,
    requestId: request.requestId,
    proposalId: "p",
    documentId: "scoped-word",
    baseRevision: session.snapshot().manifest.revision,
    annotationIds: ["a"],
    summary: "调整数值",
    operations: [
      {
        op: "setWordChart",
        nodeId: drawing.nodeId,
        changes: { series: [{ seriesId: "s", name: "收入", values: [8] }] },
      },
    ],
  };
  const candidate = session.prepareProposal(request, proposal);
  expect(
    candidate.changes.some(
      (change) => change.nodeId === drawing.nodeId && change.kind === "modified",
    ),
  ).toBe(true);
  expect(candidate.preview.annotations[0]!.status).toBe("resolved");
  expect(session.snapshot().annotations[0]!.status).toBe("open");
  session.commit(candidate);
  session.undo();
  expect(
    session.snapshot().metadata.word!.drawings!.find((drawing) => drawing.kind === "chart")!.chart
      .series[0]!.values,
  ).toEqual([1]);
});
test("Word proposals recheck copied ownership and reject modified context or unrelated object targets", () => {
  const { session, privateNode } = setup(),
    request = session.buildInteractionRequest({ annotationIds: ["a"] }),
    object = request.context.metadata.word!.drawings!.find(
      (drawing) => drawing.kind === "textBox",
    )!,
    proposal: EditProposal = {
      protocolVersion: 1,
      requestId: request.requestId,
      proposalId: "copy",
      documentId: "scoped-word",
      baseRevision: session.snapshot().manifest.revision,
      annotationIds: ["a"],
      summary: "复制",
      operations: [{ op: "copyWordDrawing", nodeId: object.nodeId }],
    };
  expect(() => session.prepareProposal(request, proposal)).not.toThrow();
  const changed = JSON.parse(JSON.stringify(request)) as typeof request;
  changed.context.metadata.word!.equations!.push(
    JSON.parse(JSON.stringify(session.snapshot().metadata.word!.equations!.at(-1))),
  );
  expect(() => session.prepareProposal(changed, proposal)).toThrow();
  expect(() =>
    session.prepareProposal(request, {
      ...proposal,
      operations: [
        { op: "replaceText", nodeId: privateNode, from: 0, to: 0, expectedText: "", text: "越界" },
      ],
    }),
  ).toThrow();
});
