import type { Css } from "../export/css.js";
import { attributes } from "../export/xml.js";
import { unsupported, rgba } from "./values.js";
import type { NativeSurface, PresentationExportContext } from "./context.js";
import type { SourceElement } from "../export/source.js";
import {
  colorXml,
  fillXml,
  lineXml,
  effectXml,
  drawingAmount as amount,
} from "../export/drawing-paint.js";
export { colorXml, fillXml, lineXml, effectXml } from "../export/drawing-paint.js";
function refPrefix(css: Css, prefix: string): string | undefined {
  const result =
    css[prefix + "-source"] === "reference" ? css[prefix + "-reference-prefix"] : undefined;
  if (
    result !== undefined &&
    ![
      "--office-fill-reference",
      "--office-line-reference",
      "--office-background-reference",
      "--office-effect-reference",
    ].includes(result)
  )
    unsupported("A native style color reference is invalid.");
  return result;
}
/** Keep the two native color stages rather than flattening theme references:
 * A lives in fmtScheme and B is the object's style-reference color. */
export function objectAppearance(
  context: PresentationExportContext,
  node: SourceElement,
  surface: NativeSurface,
  shapeKind: "geometry" | "text" | "connector" | "group",
  preset?: string,
): {
  properties: string;
  style: string;
} {
  let css = context.styles.get(node, surface),
    raw = context.styles.raw(node, surface);
  const own = context.styles.own(node, surface),
    refs: {
      kind: "fill" | "line" | "effect";
      index: number;
      color: string;
    }[] = [],
    has = (prefix: string): boolean =>
      Object.keys(own).some((name) => name.startsWith(prefix + "-")),
    placeholder = context.objects.get(node.attrs["data-node-id"]!)?.placeholder;
  if (
    (shapeKind === "geometry" || shapeKind === "text") &&
    preset !== "line" &&
    !css["--office-fill-kind"] &&
    !css["--office-fill-color"] &&
    css["background-color"] &&
    rgba(css["background-color"]!, css.color).alpha > 0
  ) {
    css = {
      ...css,
      "--office-fill-kind": "solid",
      "--office-fill-color": css["background-color"]!,
      "--office-fill-source": "direct",
    };
    raw = { ...raw, "--office-fill-color": raw["background-color"] ?? css["background-color"]! };
  }
  let properties = "";
  const fillReference =
    (css["--office-fill-kind"] || css["--office-fill-color"]
      ? refPrefix(css, "--office-fill")
      : undefined) ??
    Array.from({ length: Math.min(1024, amount(css, "--office-fill-stop-count", 0)) }, (_, i) =>
      refPrefix(css, "--office-fill-stop-" + i),
    ).find(Boolean);
  if (!placeholder || has("--office-fill")) {
    if (fillReference) {
      const style = fillXml(css, "--office-fill", raw, true),
        index = context.format(surface.themeId, "fill", style);
      refs.push({ kind: "fill", index, color: colorXml(css, fillReference, raw) });
    } else
      properties += fillXml(
        css,
        "--office-fill",
        raw,
        false,
        shapeKind === "geometry" && preset !== "line" ? "#4472c4" : "none",
      );
  }
  const lineReference = (
    css["--office-line-kind"] || css["--office-line-color"]
      ? refPrefix(css, "--office-line")
      : undefined) ??
    Array.from({ length: Math.min(1024, amount(css, "--office-line-stop-count", 0)) }, (_, index) =>
      refPrefix(css, "--office-line-stop-" + index),
    ).find(Boolean);
  if (!placeholder || has("--office-line")) {
    if (lineReference) {
      const index = context.format(
        surface.themeId,
        "line",
        lineXml(css, "--office-line", raw, true),
      );
      refs.push({ kind: "line", index, color: colorXml(css, lineReference, raw) });
    } else
      properties += lineXml(
        css,
        "--office-line",
        raw,
        false,
        shapeKind === "connector" || preset === "line" ? "#000000" : "none",
      );
  }
  const effectReference = refPrefix(css, "--office-shadow");
  if (!placeholder || has("--office-shadow")) {
    if (effectReference) {
      const index = context.format(
        surface.themeId,
        "effect",
        "<a:effectStyle>" + effectXml(css, raw, true) + "</a:effectStyle>",
      );
      refs.push({ kind: "effect", index, color: colorXml(css, effectReference, raw) });
    } else properties += effectXml(css, raw);
  }
  const defaultWhite =
    shapeKind === "geometry" &&
    preset !== "line" &&
    !["linear", "none"].includes(css["--office-fill-kind"] ?? "") &&
    !css["--office-text-color"];
  if (!refs.length && !defaultWhite) return { properties, style: "" };
  let style = "<p:style>";
  for (const kind of ["line", "fill", "effect"] as const) {
    const ref = refs.find((ref) => ref.kind === kind),
      name = kind === "line" ? "lnRef" : kind === "fill" ? "fillRef" : "effectRef";
    style +=
      "<a:" +
      name +
      attributes({ idx: ref?.index ?? 0 }) +
      ">" +
      (ref?.color ?? '<a:srgbClr val="000000"/>') +
      "</a:" +
      name +
      ">";
  }
  const textCss = css["--office-text-color"]
      ? css
      : { ...css, "--office-text-color": css.color ?? "#000000" },
    textRaw = raw["--office-text-color"]
      ? raw
      : { ...raw, "--office-text-color": raw.color ?? css.color ?? "#000000" };
  style +=
    '<a:fontRef idx="none">' +
    colorXml(textCss, "--office-text", textRaw) +
    "</a:fontRef></p:style>";
  return { properties, style };
}
