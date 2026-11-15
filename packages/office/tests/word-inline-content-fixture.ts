import {
  createWordDocument,
  createDocumentSession,
  openDocument,
  parseWordMath,
} from "@doffi/core";
import type { WordInlineContentSemantics, DocumentOperation } from "@doffi/core";
export function wordInlineContentFixture(empty = false): {
  html: string;
  semantics: WordInlineContentSemantics;
} {
  const session = createDocumentSession(
      openDocument(createWordDocument({ documentId: "portable-inline-source" }), {
        kind: "document",
      }),
    ),
    operations: DocumentOperation[] = [];
  if (!empty)
    operations.push(
      {
        op: "insertWordEquation",
        nodeId: "paragraph-1",
        offset: 0,
        math: parseWordMath("\\frac{x+1}{2}"),
      },
      {
        op: "insertWordControl",
        nodeId: "paragraph-1",
        control: { kind: "checkBox", level: "inline", checked: true },
      },
    );
  operations.push({
    op: "insertWordField",
    nodeId: "paragraph-1",
    offset: 0,
    instruction: " PAGE ",
    cachedText: empty ? "" : "7",
  });
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "portable-inline-source",
      baseRevision: 0,
      operations,
    }),
  );
  if (!empty) {
    const offset = session.getNode("paragraph-1")!.text.length;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "portable-inline-source",
        baseRevision: 1,
        operations: [
          {
            op: "insertFragment",
            nodeId: "paragraph-1",
            position: "end",
            html: "<span>LINK</span>",
          },
          {
            op: "insertWordLink",
            nodeId: "paragraph-1",
            from: offset,
            to: offset + 4,
            expectedText: "LINK",
            target: "https://example.com/source",
          },
        ],
      }),
    );
  }
  const word = session.snapshot().metadata.word!;
  return {
    html: session
      .getNode("paragraph-1")!
      .childNodeIds.map((id) => session.getNode(id)!.html)
      .join(""),
    semantics: JSON.parse(
      JSON.stringify({
        fields: word.fields,
        equations: word.equations,
        controls: word.controls,
        hyperlinks: word.hyperlinks,
      }),
    ) as WordInlineContentSemantics,
  };
}
