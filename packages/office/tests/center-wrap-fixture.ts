import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import { exportDocx } from "@doffi/office";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { floatingTableFixture } from "./floating-table-fixture.js";
export const centerWrapText =
  "中英协作 ABCdef 𠀀\u{e0100} 😀 文字沿对象两侧排列，保留同一段落和准确选区。 ";
export function centerWrapFixture(kind: "drawing" | "table" | "image", repeat = 80): Uint8Array {
  if (kind === "table") {
    const parts = unzipSync(
      floatingTableFixture(
        '<w:tblpPr w:horzAnchor="text" w:vertAnchor="text" w:tblpXSpec="center" w:tblpY="1" w:leftFromText="120" w:rightFromText="120" w:bottomFromText="120"/>',
        1,
      ),
    );
    parts["word/document.xml"] = strToU8(
      strFromU8(parts["word/document.xml"]!)
        .replace("<w:tr>", '<w:tr><w:trPr><w:trHeight w:val="1400" w:hRule="exact"/></w:trPr>')
        .replace("AFTER text wraps beside the table. ", centerWrapText.repeat(repeat)),
    );
    return zipSync(parts);
  }
  const media = unzipSync(
    new Uint8Array(readFileSync(new URL("../../../examples/docx/03-images.docx", import.meta.url))),
  );
  const path = Object.keys(media).find(
      (path) => path.startsWith("word/media/") && path.endsWith(".png"),
    )!,
    image = media[path]!;
  const session = createDocumentSession(
    openDocument(
      createWordDocument({
        documentId: "center-wrap",
        pageSetup: {
          unit: "pt",
          width: 420,
          height: 500,
          margins: { top: 36, right: 36, bottom: 36, left: 36 },
        },
      }),
      { kind: "document" },
    ),
  );
  session.commit(
    session.prepare(
      {
        protocolVersion: 1,
        documentId: "center-wrap",
        baseRevision: 0,
        operations: [
          {
            op: "replaceText",
            nodeId: "paragraph-1",
            from: 0,
            to: 0,
            expectedText: "",
            text: centerWrapText.repeat(repeat),
          },
          { op: "setWordParagraph", nodeId: "paragraph-1", properties: { "line-height": "16pt" } },
          {
            op: "insertWordDrawing",
            nodeId: "paragraph-1",
            offset: 0,
            geometry: {
              x: 0,
              y: 12,
              width: 100,
              height: 72,
              rotation: 0,
              flipH: false,
              flipV: false,
            },
            layout: {
              mode: "floating",
              horizontal: "column",
              vertical: "paragraph",
              alignH: "center",
              wrap: "square",
              wrapSide: "bothSides",
              distances: { left: 6, right: 6, top: 3, bottom: 5 },
            },
            object:
              kind === "image"
                ? { kind: "image", assetPath: "assets/center.png", alt: "居中图片" }
                : { kind: "shape", preset: "rect" },
            properties:
              kind === "image"
                ? {}
                : {
                    "--office-fill-kind": "solid",
                    "--office-fill-color": "#facc15",
                    "--office-line-kind": "solid",
                    "--office-line-color": "#224466",
                    "--office-line-width": "1pt",
                  },
          },
        ],
      },
      kind === "image"
        ? {
            assets: [
              {
                descriptor: {
                  path: "assets/center.png",
                  mimeType: "image/png",
                  sha256: createHash("sha256").update(image).digest("hex"),
                },
                bytes: image,
              },
            ],
          }
        : undefined,
    ),
  );
  return exportDocx(session).bytes;
}
