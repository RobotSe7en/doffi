import type { DeepReadonly, WorkbookDrawingAnchor, WorkbookDrawingMarker } from "@doffi/core";
import { empty } from "../export/xml.js";
import { bool, element, textElement } from "./xml.js";

export const emu = (points: number): number => Math.round(points * 12700);
export function drawingMarker(tag: string, marker: WorkbookDrawingMarker): string {
  return element("xdr:" + tag, textElement("xdr:col", String(marker.column)) + textElement("xdr:colOff", String(emu(marker.offsetXPt))) +
    textElement("xdr:row", String(marker.row)) + textElement("xdr:rowOff", String(emu(marker.offsetYPt))));
}
export function drawingAnchor(anchor: DeepReadonly<WorkbookDrawingAnchor>, body: string, locked: boolean, printable: boolean): string {
  const client = empty("xdr:clientData", { fLocksWithSheet: bool(locked), fPrintsWithSheet: bool(printable) });
  const extent = (width: number, height: number): string => empty("xdr:ext", { cx: emu(width), cy: emu(height) });
  switch (anchor.kind) {
    case "absolute":
      return element("xdr:absoluteAnchor", empty("xdr:pos", { x: emu(anchor.xPt), y: emu(anchor.yPt) }) +
        extent(anchor.widthPt, anchor.heightPt) + body + client);
    case "one-cell":
      return element("xdr:oneCellAnchor", drawingMarker("from", anchor.from) + extent(anchor.widthPt, anchor.heightPt) + body + client);
    case "two-cell":
      return element("xdr:twoCellAnchor", drawingMarker("from", anchor.from) + drawingMarker("to", anchor.to) + body + client,
        { editAs: anchor.moveWithCells ? anchor.sizeWithCells ? "twoCell" : "oneCell" : "absolute" });
  }
}
