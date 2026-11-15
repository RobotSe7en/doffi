import type { Css } from "./css.js";
import { empty } from "./xml.js";
import { length, nativeInteger } from "./values.js";
import { DocumentError } from "@doffi/core";
export function writeWordTablePosition(css: Css): string {
  const invalid = (): never => {
    throw new DocumentError("UnsupportedFeature", "The floating table position is invalid.");
  };
  const floating = css["--word-table-floating"];
  if (floating !== undefined && !["0", "1"].includes(floating)) invalid();
  let result = "";
  if (floating === "1") {
    const attrs: Record<string, string | number> = {};
    for (const [name, key, values] of [
      ["horzAnchor", "h-anchor", ["text", "margin", "page"]],
      ["vertAnchor", "v-anchor", ["text", "margin", "page"]],
      ["tblpXSpec", "x-align", ["left", "center", "right", "inside", "outside"]],
      ["tblpYSpec", "y-align", ["inline", "top", "center", "bottom", "inside", "outside"]],
    ] as const) {
      const raw = css["--word-table-position-" + key];
      if (raw !== undefined && !(key.endsWith("align") && raw === "none")) {
        if (!(values as readonly string[]).includes(raw)) invalid();
        attrs["w:" + name] = raw;
      }
    }
    for (const [name, key] of [
      ["tblpX", "x"],
      ["tblpY", "y"],
      ["leftFromText", "left-distance"],
      ["rightFromText", "right-distance"],
      ["topFromText", "top-distance"],
      ["bottomFromText", "bottom-distance"],
    ] as const) {
      const raw = css["--word-table-position-" + key];
      if (raw !== undefined) {
        const amount = length(raw, { font: 12, percent: 468 });
        if (key.endsWith("distance") && amount < 0) invalid();
        attrs["w:" + name] = nativeInteger(amount, 20);
      }
    }
    result = empty("w:tblpPr", attrs);
  }
  if (floating === "0") result = empty("w:tblpPr", {});
  const overlap = css["--word-table-overlap"];
  if (overlap !== undefined) {
    if (!["never", "overlap"].includes(overlap)) invalid();
    result += empty("w:tblOverlap", { "w:val": overlap });
  }
  return result;
}
