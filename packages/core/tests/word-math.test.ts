import { expect, test } from "vitest";
import {
  parseWordMath,
  validateWordMath,
  createWordDocument,
  createDocumentSession,
  openDocument,
} from "@doffi/core";
import type { DocumentOperation } from "@doffi/core";
test("linear math preserves Unicode variation sequences and builds typed fractions, radicals and combined scripts", () => {
  const tree = parseWordMath("\\frac{x_1^2+漢\u{e0100}}{\\sqrt[3]{y}}+\\alpha");
  expect(tree.nodes.some((node) => node.kind === "fraction")).toBe(true);
  expect(tree.nodes.some((node) => node.kind === "script" && node.mode === "subsup")).toBe(true);
  expect(tree.nodes.some((node) => node.kind === "radical" && node.children.length === 2)).toBe(
    true,
  );
  expect(tree.nodes.some((node) => node.kind === "text" && node.text === "漢\u{e0100}")).toBe(true);
  validateWordMath(tree);
  expect(() => parseWordMath("\\frac{x}{")).toThrow();
  expect(() => parseWordMath("\\unknown{x}")).toThrow();
});
test("equation transactions retain atomic text and invalidate measured regions when only math metadata changes", () => {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "math" }), { kind: "document" }),
    ),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "math",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    { op: "insertWordEquation", nodeId: "paragraph-1", offset: 0, math: parseWordMath("x^2") },
  ]);
  const equation = session.snapshot().metadata.word!.equations![0]!;
  expect(session.getNode("paragraph-1")!.text).toBe("\ufffc");
  apply([
    {
      op: "addAnnotation",
      annotation: {
        annotationId: "region",
        documentId: "math",
        baseRevision: 1,
        status: "open",
        instruction: "公式",
        anchor: {
          kind: "region",
          surfaceId: "main-story",
          rect: { x: 0, y: 0, width: 30, height: 30 },
          nodeIds: [equation.nodeId],
          layoutKey: "measured",
        },
      },
    },
  ]);
  const before = session.snapshot();
  apply([
    { op: "setWordEquation", nodeId: equation.nodeId, math: parseWordMath("\\frac{x+1}{2}") },
  ]);
  expect(session.snapshot().annotations[0]!.status).toBe("needs_reanchor");
  expect(session.snapshot().html).toBe(before.html);
  session.undo();
  expect(session.snapshot().metadata.word!.equations![0]!.nodes).toEqual(
    before.metadata.word!.equations![0]!.nodes,
  );
  expect(session.snapshot().annotations[0]!.status).toBe("open");
  apply([{ op: "removeWordEquation", nodeId: equation.nodeId }]);
  expect(session.snapshot().metadata.word!.equations).toHaveLength(0);
});
test("math trees reject cycles, orphan nodes and invalid matrix dimensions", () => {
  expect(() =>
    validateWordMath({ rootMathId: "a", nodes: [{ mathId: "a", kind: "row", children: ["a"] }] }),
  ).toThrow();
  expect(() =>
    validateWordMath({
      rootMathId: "a",
      nodes: [
        { mathId: "a", kind: "text", text: "x" },
        { mathId: "b", kind: "text", text: "y" },
      ],
    }),
  ).toThrow();
  expect(() =>
    validateWordMath({
      rootMathId: "a",
      nodes: [
        { mathId: "a", kind: "matrix", children: ["r", "s"] },
        { mathId: "r", kind: "matrixRow", children: ["x"] },
        { mathId: "s", kind: "matrixRow", children: ["y", "z"] },
        { mathId: "x", kind: "text", text: "x" },
        { mathId: "y", kind: "text", text: "y" },
        { mathId: "z", kind: "text", text: "z" },
      ],
    }),
  ).toThrow();
});

test("linear matrix and n-ary templates retain complete argument trees without orphan nodes", () => {
  const tree = parseWordMath("\\matrix{a&b\\\\c&d}+\\sum_{i=1}^{n}{i}");
  validateWordMath(tree);
  expect(tree.nodes.filter((node) => node.kind === "matrixRow")).toHaveLength(2);
  const sum = tree.nodes.find((node) => node.kind === "nary");
  expect(sum?.kind === "nary" && sum.hideLower).toBe(false);
  expect(sum?.kind === "nary" && sum.hideUpper).toBe(false);
  expect(() => parseWordMath("\\matrix{a&b\\\\c}")).toThrow();
});
