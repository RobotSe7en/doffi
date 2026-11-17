import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { drawingWrapFixture } from "./drawing-wrap-fixture.js";
export const WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
/** Inject independently authored native dimensions into a ZIP/XML source. */
export function drawingSizeFixture(
  options: {
    widthFrom?: string;
    heightFrom?: string;
    width?: string;
    height?: string;
    wrap?: "none" | "square" | "through";
    inCell?: boolean;
    count?: number;
  } = {},
): Uint8Array {
  const parts = unzipSync(
    drawingWrapFixture({
      wrap: options.wrap ?? "none",
      inCell: options.inCell ?? false,
      count: options.count ?? 1,
    }),
  );
  const width = options.width ?? "50000",
    height = options.height ?? "25000";
  const rel = (tag: string, value: string, from: string, amount: string) =>
    `<wp14:${tag} xmlns:wp14="${WP14}" relativeFrom="${from}"><wp14:${value}>${amount}</wp14:${value}></wp14:${tag}>`;
  parts["word/document.xml"] = strToU8(
    strFromU8(parts["word/document.xml"]!)
      .replaceAll(
        "<w:drawing ",
        '<w:drawing xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="' +
          WP14 +
          '" mc:Ignorable="wp14" ',
      )
      .replaceAll(
        "</wp:anchor>",
        (width === "abs" ? "" : rel("sizeRelH", "pctWidth", options.widthFrom ?? "page", width)) +
          (height === "abs"
            ? ""
            : rel("sizeRelV", "pctHeight", options.heightFrom ?? "margin", height)) +
          "</wp:anchor>",
      ),
  );
  return zipSync(parts);
}

export function mirroredDrawingSizeFixture(from: string): Uint8Array {
  const parts = unzipSync(drawingSizeFixture({ widthFrom: from, width: "100000", height: "abs" }));
  let raw = strFromU8(parts["word/document.xml"]!)
    .replace(/<w:t xml:space="preserve">[\s\S]*?<\/w:t>/g, "<w:t>镜像页相对尺寸</w:t>")
    .replace('w:right="720"', 'w:right="1440"');
  const p = raw.match(/<w:p>[\s\S]*?<\/w:p>/)![0];
  raw = raw.replace(
    p,
    p +
      p
        .replace("<w:pPr>", "<w:pPr><w:pageBreakBefore/>")
        .replace('wp:docPr id="1"', 'wp:docPr id="2"'),
  );
  parts["word/document.xml"] = strToU8(raw);
  parts["word/settings.xml"] = strToU8(
    strFromU8(parts["word/settings.xml"]!).replace(
      "</w:settings>",
      "<w:mirrorMargins/></w:settings>",
    ),
  );
  return zipSync(parts);
}
