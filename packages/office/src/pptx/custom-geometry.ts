import { PresentationCustomGeometrySchema } from "@doffi/core";
import type { PresentationCustomGeometry } from "@doffi/core";
import { A, check, child, elements, fail, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { bool, integer, value } from "./context.js";

export function customGeometry(source: XmlNode): PresentationCustomGeometry {
  check(source, ["avLst", "gdLst", "ahLst", "cxnLst", "rect", "pathLst"], A);
  const adjustments = child(source, "avLst", A), handles = child(source, "ahLst", A), guides = child(source, "gdLst", A),
    sites = child(source, "cxnLst", A), rect = child(source, "rect", A), paths = child(source, "pathLst", A);
  check(adjustments, [], A);
  check(handles, [], A);
  check(guides, ["gd"], A);
  check(sites, ["cxn"], A);
  check(rect, [], A);
  check(paths, ["path"], A);
  if (!paths) fail(source, "Custom geometry requires a path list.");
  const position = (node: XmlNode | undefined) => {
    if (!node) fail(source, "A custom geometry point is missing.");
    check(node, [], A);
    return { x: value(node, "x"), y: value(node, "y") };
  };
  const parsed = PresentationCustomGeometrySchema.safeParse({
    guides: elements(guides, "gd", A).map(guide => {
      check(guide, [], A);
      const [operator, ...args] = (value(guide, "fmla") ?? "").trim().split(/\s+/);
      if (!["val", "*/", "+-"].includes(operator!)) unsupported(guide, "This custom geometry guide formula is outside the supported profile.");
      return { name: value(guide, "name"), operator, args };
    }),
    connections: elements(sites, "cxn", A).map(site => {
      check(site, ["pos"], A);
      return { angle: value(site, "ang"), ...position(child(site, "pos", A)) };
    }),
    ...(rect ? { textRect: { left: value(rect, "l"), top: value(rect, "t"), right: value(rect, "r"), bottom: value(rect, "b") } } : {}),
    paths: elements(paths, "path", A).map(path => {
      check(path, ["moveTo", "lnTo", "close"], A);
      const fill = value(path, "fill") ?? "norm";
      if (fill !== "norm" && fill !== "none" || !bool(path, "extrusionOk", true)) unsupported(path, "This custom path paint is outside the supported profile.");
      return {
        ...(value(path, "w") === undefined ? {} : { width: integer(path, "w") }),
        ...(value(path, "h") === undefined ? {} : { height: integer(path, "h") }),
        fill: fill === "norm", stroke: bool(path, "stroke", true),
        commands: elements(path, undefined, A).map(command => {
          if (command.name === "close") { check(command, [], A); return { kind: "close" }; }
          check(command, ["pt"], A);
          return { kind: command.name === "moveTo" ? "move" : "line", ...position(child(command, "pt", A)) };
        }),
      };
    }),
  });
  if (!parsed.success) fail(source, "A custom geometry is malformed or outside its path budget.");
  return parsed.data;
}
