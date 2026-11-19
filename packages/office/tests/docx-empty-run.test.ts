import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { importDocx, exportDocx } from "@doffi/office";
import { strFromU8, unzipSync } from "fflate";
import { mutateDocx, body, styles } from "./regression-fixtures.js";

for (const opaque of [false, true])
  test(
    "empty native runs retain their own font, passive flags and " +
      (opaque ? "opaque effects" : "format history") +
      " across editing and two reopens",
    () => {
      const bytes = mutateDocx("01-text-numbering.docx", {
          "word/styles.xml": () =>
            styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
          "word/document.xml": (source) =>
            body(
              source,
              '<w:p><w:r><w:rPr><w:sz w:val="64"/><w:noProof/>' +
                (opaque
                  ? "<w:emboss/>"
                  : '<w:rPrChange w:id="7" w:author="A"><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrChange>') +
                "</w:rPr></w:r><w:r><w:t>正文</w:t></w:r></w:p>",
            ),
        }),
        session = createDocumentSession(
          openDocument(
            importDocx(bytes, {
              documentId: "empty-run",
              ...(opaque ? {} : { unsupportedContent: "reject" as const }),
            }),
            { kind: "document" },
          ),
        ),
        nodes = [session.snapshot().metadata.word!.stories[0]!.rootNodeId];
      for (let index = 0; index < nodes.length; index++)
        nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
      const paragraph = nodes.find((id) => session.getNode(id)!.tagName === "p")!;
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "empty-run",
          baseRevision: 0,
          operations: [
            {
              op: "replaceText",
              nodeId: paragraph,
              from: 0,
              to: 0,
              expectedText: "",
              text: "编辑",
            },
          ],
        }),
      );
      let output = exportDocx(session).bytes;
      for (let pass = 0; pass < 2; pass++) {
        const xml = strFromU8(unzipSync(output)["word/document.xml"]!),
          runs = xml.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) ?? [];
        expect(
          runs.some((run) => run.includes('w:sz w:val="64"') && !/<w:(?:t|br|tab)\b/.test(run)),
        ).toBe(true);
        expect(xml).toContain("noProof");
        expect(xml).toContain(opaque ? "emboss" : "rPrChange");
        const reopened = openDocument(
          importDocx(output, {
            documentId: "empty-reopen",
            ...(opaque ? {} : { unsupportedContent: "reject" as const }),
          }),
          { kind: "document" },
        );
        expect(
          reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text,
        ).toBe("编辑正文");
        if (!opaque) expect(reopened.snapshot().metadata.word!.revisions[0]!.property).toBe("run");
        output = exportDocx(reopened).bytes;
      }
      session.undo();
      expect(session.getNode(paragraph)!.text).toBe("正文");
    },
  );
