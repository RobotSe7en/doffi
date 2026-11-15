import { W, attr, child, check, checkAttributes, fail } from "./xml.js";
import type { XmlNode } from "./xml.js";
import { measurePoints } from "./measure.js";
const prefix = "--word-table-position-";
/** Preserve positioning separately from HTML layout; zero/one twip sentinels
 * keep their native meaning rather than being normalized to CSS floats. */
export function readWordTablePosition(properties: XmlNode): Record<string, string> {
  const result: Record<string, string> = {},
    position = child(properties, "tblpPr"),
    overlap = child(properties, "tblOverlap");
  if (position) {
    check(position, []);
    checkAttributes(
      position,
      [
        "leftFromText",
        "rightFromText",
        "topFromText",
        "bottomFromText",
        "vertAnchor",
        "horzAnchor",
        "tblpXSpec",
        "tblpX",
        "tblpYSpec",
        "tblpY",
      ],
      W,
    );
    result["--word-table-floating"] = position.attrs.size ? "1" : "0";
    for (const [name, key, values] of [
      ["horzAnchor", "h-anchor", ["text", "margin", "page"]],
      ["vertAnchor", "v-anchor", ["text", "margin", "page"]],
      ["tblpXSpec", "x-align", ["left", "center", "right", "inside", "outside"]],
      ["tblpYSpec", "y-align", ["inline", "top", "center", "bottom", "inside", "outside"]],
    ] as const) {
      const value = attr(position, name);
      if (value !== undefined) {
        if (!(values as readonly string[]).includes(value))
          fail(position, "A floating table anchor or alignment is invalid.");
        result[prefix + key] = value;
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
      const raw = attr(position, name);
      if (raw !== undefined) {
        const amount = measurePoints(raw, position, 20);
        if (key.endsWith("distance") && amount < 0)
          fail(position, "A floating table text distance cannot be negative.");
        result[prefix + key] = amount + "pt";
      }
    }
  }
  if (overlap) {
    check(overlap, []);
    checkAttributes(overlap, ["val"], W);
    const value = attr(overlap, "val");
    if (!value || !["never", "overlap"].includes(value))
      fail(overlap, "A table overlap policy is invalid.");
    result["--word-table-overlap"] = value;
  }
  return result;
}
