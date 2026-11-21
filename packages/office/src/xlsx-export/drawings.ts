import { DocumentError, WorkbookAxisDimensions } from "@doffi/core";
import type { DeepReadonly, DocumentReadLimits, WorkbookCalculatedCell, WorkbookDrawingAnchor, WorkbookDrawingMarker, WorkbookModel, WorkbookObject, WorkbookSheet } from "@doffi/core";
import type { OutputOpc } from "../export/opc.js";
import { children, parseSource } from "../export/source.js";
import type { SourceElement } from "../export/source.js";
import { declaration, empty } from "../export/xml.js";
import { A, C, R } from "../ooxml/xml.js";
import { XDR } from "../xlsx/drawing-anchor.js";
import { drawingAnchor, emu } from "./anchors.js";
import { WorkbookChartsWriter } from "./charts.js";
import { shapeAppearance, WorkbookDrawingStyles } from "./drawing-style.js";
import type { XlsxExportLoss } from "./options.js";
import type { WorkbookStylesWriter } from "./styles.js";
import { bool, element } from "./xml.js";

type Obj = DeepReadonly<WorkbookObject>;
type Sheet = DeepReadonly<WorkbookSheet>;
type Box = { x: number; y: number; w: number; h: number };
const imageTypes = new Map([["image/png", "png"], ["image/jpeg", "jpeg"], ["image/webp", "webp"]]);

function layout(sheet: Sheet): (anchor: DeepReadonly<WorkbookDrawingAnchor>) => Box {
  // Native anchors remain authoritative. xfrm is a regenerated display cache.
  const width = (characters: number): number => Math.floor(characters * 7) * .75;
  const columnWidth = width(sheet.defaultColumnWidth) + (sheet.defaultColumnWidthUnit === "baseCharacters" ? 3.75 : 0);
  const columns = new WorkbookAxisDimensions(16384, columnWidth, sheet.columnDimensions.map(column => ({
    ...column, sizePt: column.hidden ? 0 : column.width === undefined ? columnWidth : width(column.width),
  })));
  const rows = new WorkbookAxisDimensions(1048576, sheet.defaultRowsHidden ? 0 : sheet.defaultRowHeightPt, sheet.rowDimensions.map(row => ({
    ...row, sizePt: (row.hidden ?? sheet.defaultRowsHidden) ? 0 : row.heightPt ?? sheet.defaultRowHeightPt,
  })));
  const marker = (at: WorkbookDrawingMarker): { x: number; y: number } =>
    ({ x: columns.position(at.column) + at.offsetXPt, y: rows.position(at.row) + at.offsetYPt });
  return anchor => {
    if (anchor.kind === "absolute") return { x: anchor.xPt, y: anchor.yPt, w: anchor.widthPt, h: anchor.heightPt };
    const start = marker(anchor.from);
    if (anchor.kind === "one-cell") return { ...start, w: anchor.widthPt, h: anchor.heightPt };
    const end = marker(anchor.to);
    return { ...start, w: Math.max(0, end.x - start.x), h: Math.max(0, end.y - start.y) };
  };
}
function transform(node: SourceElement, box: Box, frame: boolean): string {
  // The frame xfrm shares the spreadsheet anchor's geometry, with no rotation
  // attributes in the native frame schema.
  if (frame) return element("xdr:xfrm", empty("a:off", { x: emu(box.x), y: emu(box.y) }) + empty("a:ext", { cx: emu(box.w), cy: emu(box.h) }));
  const value = (name: string): number | undefined => node.attrs[name] === undefined ? undefined : Number(node.attrs[name] === "true");
  const rotation = node.attrs["data-office-rotation"];
  return element("a:xfrm", empty("a:off", { x: emu(box.x), y: emu(box.y) }) + empty("a:ext", { cx: emu(box.w), cy: emu(box.h) }), {
    rot: rotation === undefined ? undefined : Math.round(Number(rotation) * 60000),
    flipH: value("data-office-flip-h"), flipV: value("data-office-flip-v"),
  });
}
function geometry(node: SourceElement, preset: string): string {
  const adjustments = Object.entries(node.attrs).flatMap(([key, value]) => key.startsWith("data-office-adjust-") ?
    [empty("a:gd", { name: key.slice("data-office-adjust-".length), fmla: "val " + value })] : []);
  return element("a:prstGeom", element("a:avLst", adjustments.join("")), { prst: preset });
}

export class WorkbookDrawingsWriter {
  private readonly resources = new Map<string, string>();
  private readonly assets: ReadonlyMap<string, ReturnType<WorkbookModel["snapshot"]>["manifest"]["assets"][number]>;
  private readonly tree: ReturnType<typeof parseSource> | undefined;
  private readonly styles: WorkbookDrawingStyles;
  private readonly charts: WorkbookChartsWriter | undefined;
  private readonly chartDefinitions: ReadonlyMap<string, NonNullable<ReturnType<WorkbookModel["snapshot"]>["metadata"]["workbook"]["charts"]>[number]>;
  private readonly objects = new Map<string, Obj[]>();

  constructor(private readonly model: WorkbookModel, private readonly opc: OutputOpc, styles: WorkbookStylesWriter,
    limits: DocumentReadLimits, calculated: ReadonlyMap<string, WorkbookCalculatedCell>, private readonly losses: XlsxExportLoss[], maxPoints: number) {
    const snapshot = model.snapshot();
    this.assets = new Map(snapshot.manifest.assets.map(asset => [asset.path, asset]));
    this.chartDefinitions = new Map(snapshot.metadata.workbook.charts?.map(chart => [chart.chartId, chart]));
    this.tree = snapshot.content && snapshot.metadata.workbook.objects?.length ? parseSource(snapshot.content.html, limits, "workbook-content") : undefined;
    this.styles = new WorkbookDrawingStyles(styles, opc, snapshot);
    this.charts = this.tree ? new WorkbookChartsWriter(model, opc, this.styles, this.tree, calculated, losses, maxPoints) : undefined;
    for (const object of snapshot.metadata.workbook.objects ?? []) {
      const group = this.objects.get(object.sheetId) ?? [];
      group.push(object);
      this.objects.set(object.sheetId, group);
    }
  }
  private image(node: SourceElement, part: string): string {
    const image = children(node, "img")[0]!, path = image.attrs.src!, asset = this.assets.get(path), extension = asset && imageTypes.get(asset.mimeType);
    if (!asset || !extension) throw new DocumentError("UnsupportedFeature", "The picture has no supported embedded XLSX asset.");
    let native = this.resources.get(path);
    if (!native) {
      const bytes = this.model.getAsset(path, this.opc.parts.limits.maxFileBytes);
      if (!bytes) throw new DocumentError("InvalidDocument", "The declared workbook picture bytes are missing.");
      native = "image" + (this.resources.size + 1) + "." + extension;
      this.opc.binary("xl/media/" + native, asset.mimeType, bytes);
      this.resources.set(path, native);
    }
    const crop = Object.fromEntries([["l", "left"], ["t", "top"], ["r", "right"], ["b", "bottom"]].flatMap(([native, side]) => {
      const value = node.attrs["data-office-crop-" + side];
      return value === undefined ? [] : [[native!, Math.round(Number(value) * 100000)]];
    }));
    const rotates = node.attrs["data-office-image-rotate-with-shape"];
    return element("xdr:blipFill", empty("a:blip", { "r:embed": this.opc.relate(part, "image", "../media/" + native) }) +
      (Object.keys(crop).length ? empty("a:srcRect", crop) : "") + "<a:stretch><a:fillRect/></a:stretch>",
    { rotWithShape: rotates === undefined ? undefined : bool(rotates !== "false") });
  }
  private object(object: Obj, node: SourceElement, id: number, geometryBox: Box, part: string): string {
    const frame = object.kind === "chart", picture = object.kind === "image";
    const kind = frame ? "GraphicFrame" : picture ? "Pic" : "Sp";
    const attrs = { id, name: object.name ?? object.kind + " " + id, descr: node.attrs["data-office-alt"], title: node.attrs.title,
      hidden: node.attrs["data-office-hidden"] === undefined ? undefined : bool(node.attrs["data-office-hidden"] === "true") };
    const lockName = frame ? "graphicFrameLocks" : picture ? "picLocks" : "spLocks";
    const locks = object.locks ? empty("a:" + lockName, Object.fromEntries(Object.entries(object.locks).map(([key, value]) => [key, bool(value)]))) : "";
    const props = picture ? { preferRelativeResize: bool(object.preferRelativeResize) } : frame ? {} : { txBox: object.kind === "text" ? 1 : undefined };
    const nonvisual = element("xdr:nv" + kind + "Pr", element("xdr:cNvPr", this.styles.hyperlink(node, part), attrs) +
      element("xdr:cNv" + kind + "Pr", locks, props));
    const xfrm = transform(node, geometryBox, frame);
    if (frame) return element("xdr:graphicFrame", nonvisual + xfrm + element("a:graphic", element("a:graphicData",
      this.charts!.write(this.chartDefinitions.get(object.chartId)!, node, object.sheetId, part), { uri: C })));
    const style = this.styles.style(node);
    const shape = element("xdr:spPr", xfrm + geometry(node, object.kind === "geometry" ? object.preset : "rect") + shapeAppearance(style));
    if (picture) return element("xdr:pic", nonvisual + this.image(node, part) + shape);
    const text = object.text || children(node, "p").length ? this.styles.richBody("xdr:txBody", node, part, object.text) : "";
    return element("xdr:sp", nonvisual + shape + text);
  }

  sheet(sheet: Sheet, sheetPart: string, index: number): string {
    const objects = this.objects.get(sheet.sheetId);
    if (!objects?.length) return "";
    const path = "drawing" + index + ".xml", name = "xl/drawings/" + path;
    const drawing = this.opc.part(name, "application/vnd.openxmlformats-officedocument.drawing+xml");
    drawing.add(declaration + '<xdr:wsDr xmlns:xdr="' + XDR + '" xmlns:a="' + A + '" xmlns:c="' + C + '" xmlns:r="' + R + '">');
    const box = layout(sheet), used = new Set<number>();
    let next = 1;
    for (const object of objects) {
      const candidate = Number(object.nativeId), id = Number.isInteger(candidate) && candidate > 0 && candidate <= 0xffff_ffff && !used.has(candidate)
        ? candidate : next++;
      used.add(id);
      while (used.has(next)) next++;
      const node = this.tree!.byId.get(object.nodeId)!;
      if (object.anchor.kind === "two-cell" && !object.anchor.moveWithCells && object.anchor.sizeWithCells)
        this.losses.push({ kind: "metadata", sheetId: sheet.sheetId, nodeId: object.nodeId, message: "The native drawing anchor cannot independently size with cells while its position stays fixed." });
      if (object.kind === "chart" && (Number(node.attrs["data-office-rotation"] ?? 0) ||
        node.attrs["data-office-flip-h"] === "true" || node.attrs["data-office-flip-v"] === "true"))
        this.losses.push({ kind: "format", sheetId: sheet.sheetId, nodeId: object.nodeId, message: "The XLSX chart frame cannot carry rotation or flips." });
      const native = this.object(object, node, id, box(object.anchor), name);
      drawing.add(drawingAnchor(object.anchor, native, object.locked, object.printable));
    }
    drawing.add("</xdr:wsDr>");
    drawing.finish();
    return empty("drawing", { "r:id": this.opc.relate(sheetPart, "drawing", "../drawings/" + path) });
  }
}
