import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";
import type { Annotation, AnnotationReply, DeepReadonly, WorkbookComment, WorkbookCommentAuthor, WorkbookCommentMention, WorkbookCoordinate, WorkbookModel, WorkbookNoteBox } from "@doffi/core";
import type { OutputOpc } from "../export/opc.js";
import { declaration, empty, escapeXml } from "../export/xml.js";
import { XDR } from "../xlsx/drawing-anchor.js";
import { COMMENT_REVISION } from "../xlsx/comment-attributes.js";
import { THREADED_COMMENTS } from "../xlsx/threaded-comments.js";
import { encode } from "../xlsx/values.js";
import { VML, VML_EXCEL, VML_OFFICE } from "../xlsx/vml-notes.js";
import { drawingMarker } from "./anchors.js";
import type { XlsxExportReport } from "./options.js";
import type { WorkbookStylesWriter } from "./styles.js";
import { richCellText } from "./text.js";
import { address, bool, element, rangeRef, sheetContentType, sheetNamespaces, stringText, textElement, workbookPart, xstring } from "./xml.js";

type Classic = DeepReadonly<Extract<WorkbookComment, { format: "classic" }>>;
type Modern = DeepReadonly<Extract<WorkbookComment, { format: "modern" }>>;
type NoteBox = DeepReadonly<WorkbookNoteBox>;
type Message = DeepReadonly<Annotation | AnnotationReply>;
type MutableReport = { [K in "exportedAnnotationIds" | "unexportedAnnotations" | "losses"]: XlsxExportReport[K][number][] };
type Entry = { annotation: DeepReadonly<Annotation>; native?: DeepReadonly<WorkbookComment>; cell?: WorkbookCoordinate; suffix: string };
const relationPrefix = "http://schemas.microsoft.com/office/2017/10/relationships/";
const guidPattern = /^\{[\da-f]{8}-(?:[\da-f]{4}-){3}[\da-f]{12}\}$/i;
const datePattern = /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)?$/;

class CommentIds {
  private readonly used = new Set<string>();
  constructor(private readonly namespace: string) {}
  one(native: string | undefined, identity: string): string {
    if (native && guidPattern.test(native) && !this.used.has(native.toUpperCase())) {
      const normalized = native.toUpperCase();
      this.used.add(normalized);
      return normalized;
    }
    for (let serial = 0;; serial++) {
      const hex = bytesToHex(sha256(encode(JSON.stringify([this.namespace, identity, serial])))).slice(0, 32).toUpperCase();
      const id = "{" + hex.slice(0, 8) + "-" + hex.slice(8, 12) + "-4" + hex.slice(13, 16) + "-8" +
        hex.slice(17, 20) + "-" + hex.slice(20) + "}";
      if (this.used.has(id)) continue;
      this.used.add(id);
      return id;
    }
  }
}

function authorName(name: string): string {
  // A legacy author token matching the modern compatibility convention is
  // interpreted as a placeholder by Office and the reader.
  return name.replace(/tc=(?=\{[\da-f-]+\})/gi, "tc =");
}
function commentProperties(box: NoteBox, styles: WorkbookStylesWriter): string {
  const anchor = box.anchor;
  if (anchor.kind !== "two-cell") return "";
  const css = box.cssClass ? styles.css(box.cssClass) : {};
  return element("commentPr", element("anchor", drawingMarker("from", anchor.from) + drawingMarker("to", anchor.to),
    { moveWithCells: bool(anchor.moveWithCells), sizeWithCells: bool(anchor.sizeWithCells) }), {
    locked: bool(box.locked), defaultSize: bool(box.defaultSize), print: bool(box.printable), disabled: bool(box.disabled),
    autoFill: bool(box.autoFill), autoLine: bool(box.autoLine), lockText: bool(box.lockText), justLastX: bool(box.justifyLastLine),
    autoScale: bool(box.autoScale), altText: box.altText === undefined ? undefined : xstring(box.altText),
    textHAlign: css["--office-note-horizontal"], textVAlign: css["--office-note-vertical"],
  });
}

function vmlNote(box: NoteBox, cell: WorkbookCoordinate, shapeId: number, styles: WorkbookStylesWriter): string {
  const anchor = box.anchor;
  if (anchor.kind !== "two-cell") return "";
  const css = box.cssClass ? styles.css(box.cssClass) : {};
  const markers = [anchor.from, anchor.to].flatMap(marker =>
    [marker.column, Math.round(marker.offsetXPt / .75), marker.row, Math.round(marker.offsetYPt / .75)]);
  const enabled = (name: string, value: boolean): string => textElement("x:" + name, value ? "True" : "False");
  const client = textElement("x:Anchor", markers.join(", ")) + textElement("x:Row", String(cell.row)) + textElement("x:Column", String(cell.column)) +
    enabled("MoveWithCells", anchor.moveWithCells) + enabled("SizeWithCells", anchor.sizeWithCells) + enabled("Visible", box.visible) +
    enabled("Locked", box.locked) + enabled("PrintObject", box.printable) + enabled("LockText", box.lockText) + enabled("DefaultSize", box.defaultSize) +
    enabled("Disabled", box.disabled) + enabled("AutoFill", box.autoFill) + enabled("AutoLine", box.autoLine) + enabled("AutoScale", box.autoScale) +
    enabled("JustLastX", box.justifyLastLine) + (css["--office-note-horizontal"] ? textElement("x:TextHAlign", css["--office-note-horizontal"]) : "") +
    (css["--office-note-vertical"] ? textElement("x:TextVAlign", css["--office-note-vertical"]) : "");
  const fill = css["--office-note-fill"], stroke = css["--office-note-stroke"], shadow = css["--office-note-shadow"];
  const fillXml = empty("v:fill", { on: fill === "none" ? "f" : "t", color: fill && fill !== "none" ? fill : undefined, opacity: css["--office-note-fill-opacity"] });
  const strokeAttrs = Object.fromEntries(["joinstyle", "endcap", "dashstyle"].map(key => [key, css["--office-note-stroke-" + key]]));
  for (const end of ["start", "end"]) for (const key of ["arrow", "arrowwidth", "arrowlength"])
    strokeAttrs[end + key] = css["--office-note-" + end + "-" + key];
  const strokeXml = empty("v:stroke", { ...strokeAttrs, on: stroke === "none" ? "f" : "t", color: stroke && stroke !== "none" ? stroke : undefined,
    weight: css["--office-note-stroke-width"], opacity: css["--office-note-stroke-opacity"] });
  const shadowXml = shadow && shadow !== "none" ? empty("v:shadow", { on: "t", color: shadow,
    offset: (css["--office-note-shadow-x"] ?? "2pt") + "," + (css["--office-note-shadow-y"] ?? "2pt"),
    opacity: css["--office-note-shadow-opacity"], obscured: css["--office-note-shadow-obscured"] === "true" ? "t" : "f" }) : "";
  const inset = ["left", "top", "right", "bottom"].map((side, index) => css["padding-" + side] ?? (index % 2 ? "3.6pt" : "7.2pt")).join(",");
  const textbox = element("v:textbox", "<div/>", { inset, style: "mso-direction-alt:auto" });
  const style = "position:absolute;visibility:" + (box.visible ? "visible" : "hidden") + ";mso-wrap-style:" +
    (css["--office-note-wrap-style"] ?? "none") + (css["--office-note-z-index"] ? ";z-index:" + css["--office-note-z-index"] : "");
  return element("v:shape", fillXml + strokeXml + shadowXml + textbox + element("x:ClientData", client, { ObjectType: "Note" }), {
    id: "_x0000_s" + shapeId, type: "#_x0000_t202", style, "o:insetmode": css["--office-note-inset-mode"],
    alt: box.altText === undefined ? undefined : xstring(box.altText),
  });
}

export class WorkbookCommentsWriter {
  private readonly entries = new Map<string, Entry[]>();
  private readonly authors: Map<string, DeepReadonly<WorkbookCommentAuthor>>;
  private readonly personIds = new Map<string, string>();
  private readonly people: string[] = [];
  private readonly ids: CommentIds;

  constructor(private readonly model: WorkbookModel, private readonly opc: OutputOpc, private readonly styles: WorkbookStylesWriter, private readonly report: MutableReport) {
    const snapshot = model.snapshot();
    this.ids = new CommentIds(snapshot.manifest.documentId);
    const { workbook } = snapshot.metadata;
    this.authors = new Map(workbook.commentAuthors?.map(author => [author.authorId, author]));
    const comments = new Map(workbook.comments?.map(comment => [comment.annotationId, comment]));
    const sheets = new Set(workbook.sheets.map(sheet => sheet.sheetId));
    const used = new Set<string>();
    for (const annotation of snapshot.annotations) {
      const { annotationId, anchor } = annotation, native = comments.get(annotationId);
      if (anchor.kind !== "cells" || !sheets.has(anchor.sheetId)) {
        report.unexportedAnnotations.push({ annotationId, reason: "unsupported_anchor" }); continue;
      }
      const cell = anchor.ranges[0]?.from;
      if (!cell && native?.format !== "modern") {
        report.unexportedAnnotations.push({ annotationId, reason: "missing_coordinate" }); continue;
      }
      const key = cell && JSON.stringify([anchor.sheetId, cell.row, cell.column]);
      if (key && used.has(key)) {
        report.unexportedAnnotations.push({ annotationId, reason: "representative_collision" }); continue;
      }
      if (key) used.add(key);
      let suffix = "";
      if (anchor.ranges.length > 1 || anchor.ranges[0] && rangeRef(anchor.ranges[0]) !== address(cell!)) {
        suffix = "\n\n[doffi ranges: " + anchor.ranges.map(range => rangeRef(range)).join(", ") + "]";
        report.losses.push({ annotationId, kind: "anchor", message: "The native comment is attached to the first cell of the first source range." });
      }
      if (annotation.status === "needs_reanchor")
        report.losses.push({ annotationId, kind: "anchor", message: "The source annotation requires reanchoring." });
      report.exportedAnnotationIds.push(annotationId);
      const entries = this.entries.get(anchor.sheetId) ?? [];
      entries.push({ annotation, ...(native ? { native } : {}), ...(cell ? { cell } : {}), suffix });
      this.entries.set(anchor.sheetId, entries);
    }
  }

  private body(message: Message): { plain: string; rich?: string } {
    const node = message.bodyNodeId && this.model.getNode(message.bodyNodeId);
    if (!node) return { plain: message.instruction ?? "" };
    const paragraphs = node.childNodeIds.map(id => this.model.getNode(id)!);
    const plain = paragraphs.map(paragraph => paragraph.text).join("\n");
    if (!paragraphs.length) return { plain: node.text };
    const rich = paragraphs.map(paragraph => paragraph.html.replace(/^<[^>]+>|<\/[^>]+>$/g, "")).join("\n");
    return { plain, rich };
  }
  private text(entry: Entry, message: Message, suffix: string): string {
    const { plain } = this.body(message);
    if (plain.length + suffix.length <= 32767) return plain + suffix;
    if (plain.length <= 32767) {
      this.report.losses.push({ annotationId: entry.annotation.annotationId, kind: "anchor", message: "The range description exceeded the native comment length." });
      return plain;
    }
    // Preserve complete Unicode characters while obeying the cell-comment limit.
    this.report.losses.push({ annotationId: entry.annotation.annotationId, kind: "format", message: "The comment text exceeded the native UTF-16 length limit." });
    return plain.slice(0, 32767).replace(/[\ud800-\udbff]$/, "");
  }
  private person(authorId: string | undefined, name: string): string {
    const author = authorId && this.authors.get(authorId), key = author ? author.authorId : JSON.stringify(name);
    const previous = this.personIds.get(key);
    if (previous) return previous;
    const id = this.ids.one(author ? author.nativeId : undefined, "person:" + key);
    this.personIds.set(key, id);
    this.people.push(empty("person", { id, displayName: xstring(author ? author.name : name),
      userId: author && author.userId !== undefined ? xstring(author.userId) : undefined,
      providerId: author && author.providerId !== undefined ? xstring(author.providerId) : undefined }));
    return id;
  }
  private mentions(list: readonly DeepReadonly<WorkbookCommentMention>[] | undefined): string {
    return list?.length ? element("mentions", list.map(mention => empty("mention", {
      mentionId: this.ids.one(mention.mentionId, "mention:" + mention.mentionId),
      mentionpersonId: this.person(mention.authorId, ""), startIndex: mention.startIndex, length: mention.length,
    })).join("")) : "";
  }
  private modern(entry: Entry, native: Modern | undefined, root: string): string {
    const { annotation, cell } = entry;
    const message = (source: Message, id: string, ref: WorkbookCoordinate | undefined, authorId: string | undefined,
      mentions: readonly DeepReadonly<WorkbookCommentMention>[] | undefined, parentId?: string, suffix = ""): string => {
      const date = source.date && datePattern.test(source.date) ? source.date : undefined;
      if (source.date && !date) this.report.losses.push({ annotationId: annotation.annotationId, kind: "metadata", message: "The comment timestamp has no native dateTime representation." });
      if (/<span[\s>]/i.test(this.body(source).rich ?? ""))
        this.report.losses.push({ annotationId: annotation.annotationId, kind: "format", message: "Modern XLSX comments carry plain text." });
      const state = parentId ? source.status : native?.recoveryStatus ?? source.status;
      if (state === "closed") this.report.losses.push({ annotationId: annotation.annotationId, kind: "status", message: "A closed source comment is saved as a resolved native comment." });
      return element("threadedComment", textElement("text", xstring(this.text(entry, source, suffix))) + this.mentions(mentions), {
        id, ref: ref ? address(ref) : undefined, personId: this.person(authorId, source.author?.name ?? ""),
        parentId, dT: date, done: bool(state === "resolved" || state === "closed"),
      });
    };
    const replies = annotation.replies ?? [], bindings = new Map(native?.replies.map(reply => [reply.replyId, reply]));
    const ids = new Map(replies.map(reply => [reply.replyId, this.ids.one(reply.sourceId, "reply:" + annotation.annotationId + ":" + reply.replyId)]));
    return message(annotation, root, cell, native?.authorId, native?.mentions, undefined, entry.suffix) + replies.map(reply => {
      const binding = bindings.get(reply.replyId);
      const nativeCell = native?.recoveryReasons?.includes("mismatched-reply-location") ? binding?.nativeCell ?? cell : cell;
      return message(reply, ids.get(reply.replyId)!, nativeCell, binding?.authorId, binding?.mentions,
        (binding?.parentReplyId ? ids.get(binding.parentReplyId) : undefined) ?? root);
    }).join("");
  }

  sheet(sheetId: string, sheetPart: string, index: number): string {
    const entries = this.entries.get(sheetId);
    if (!entries?.length) return "";
    const authors: string[] = [], authorIds = new Map<string, number>();
    const classic: string[] = [], threaded: string[] = [], shapes: string[] = [];
    const shapeIds = new Set<number>();
    const author = (name: string): number => {
      const known = authorIds.get(name); if (known !== undefined) return known;
      const id = authors.length;
      authors.push(textElement("author", xstring(name)));
      authorIds.set(name, id);
      return id;
    };
    for (const entry of entries) {
      const { annotation, native, cell } = entry, id = annotation.annotationId;
      if (native?.format === "modern" || (!native && annotation.replies?.length)) {
        const modern = native?.format === "modern" ? native : undefined;
        const root = this.ids.one(annotation.sourceId, "thread:" + id);
        threaded.push(this.modern(entry, modern, root));
        if (cell && !modern?.recoveryReasons?.includes("missing-placeholder")) classic.push(element("comment",
          element("text", stringText(this.text(entry, annotation, entry.suffix))), {
            ref: address(cell), authorId: author("tc=" + root),
            "xr:uid": modern?.compatibilityUids?.length === 1 ? modern.compatibilityUids[0] : undefined,
          }));
        continue;
      }
      const note: Classic | undefined = native?.format === "classic" ? native : undefined;
      const position = cell!, { plain, rich } = this.body(annotation), text = this.text(entry, annotation, entry.suffix);
      const body = rich && text === plain + entry.suffix ? richCellText("<td>" + rich + escapeXml(entry.suffix) + "</td>", text, this.styles) : stringText(text);
      let shapeId = note?.nativeShapeId;
      if (note?.box && note.box.anchor.kind !== "two-cell")
        this.report.losses.push({ annotationId: id, kind: "anchor", message: "The classic note box requires two cell markers in XLSX." });
      if (note?.box && note.boxSource !== "commentPr") {
        if (!shapeId || shapeIds.has(shapeId)) {
          shapeId = 1025 + shapes.length;
          while (shapeIds.has(shapeId)) shapeId++;
        }
        const shape = vmlNote(note.box, position, shapeId, this.styles);
        if (shape) { shapes.push(shape); shapeIds.add(shapeId); }
      }
      const commentPr = note?.commentPrBox ?? (note?.boxSource === "commentPr" ? note.box : undefined);
      let creator = authorName(annotation.author?.name ?? "");
      if (note?.compatibilityThreadId) creator = "tc=" + note.compatibilityThreadId;
      classic.push(element("comment", element("text", body) + (commentPr ? commentProperties(commentPr, this.styles) : ""), {
        ref: address(position), authorId: author(creator), shapeId, guid: annotation.sourceId && guidPattern.test(annotation.sourceId)
          ? this.ids.one(annotation.sourceId, "classic:" + id) : undefined, "xr:uid": note?.nativeUid,
      }));
      if (annotation.status === "resolved" || annotation.status === "closed") this.report.losses.push({ annotationId: id, kind: "status", message: "A classic note has no resolved or closed state." });
      if (annotation.replies?.length) this.report.losses.push({ annotationId: id, kind: "thread", message: "A classic note cannot contain replies." });
    }
    if (classic.length) {
      const part = this.opc.part("xl/comments" + index + ".xml", sheetContentType("comments"));
      part.add(declaration + "<comments" + sheetNamespaces + ' xmlns:xdr="' + XDR + '" xmlns:xr="' + COMMENT_REVISION + '">' +
        element("authors", authors.join("")) + "<commentList>");
      for (const note of classic) part.add(note);
      part.add("</commentList></comments>");
      part.finish();
      this.opc.relate(sheetPart, "comments", "../comments" + index + ".xml");
    }
    if (threaded.length) {
      const path = "threadedComments/threadedComment" + index + ".xml", part = this.opc.part("xl/" + path, "application/vnd.ms-excel.threadedcomments+xml");
      part.add(declaration + '<ThreadedComments xmlns="' + THREADED_COMMENTS + '">');
      for (const thread of threaded) part.add(thread);
      part.add("</ThreadedComments>");
      part.finish();
      this.opc.relate(sheetPart, relationPrefix + "threadedComment", "../" + path);
    }
    if (!shapes.length) return "";
    const path = "drawings/vmlDrawing" + index + ".vml", part = this.opc.part("xl/" + path, "application/vnd.openxmlformats-officedocument.vmlDrawing");
    part.add(declaration + '<xml xmlns:v="' + VML + '" xmlns:o="' + VML_OFFICE + '" xmlns:x="' + VML_EXCEL + '">' +
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path o:connecttype="rect"/></v:shapetype>');
    for (const shape of shapes) part.add(shape);
    part.add("</xml>");
    part.finish();
    return empty("legacyDrawing", { "r:id": this.opc.relate(sheetPart, "vmlDrawing", "../" + path) });
  }

  finish(): void {
    if (!this.people.length) return;
    const part = this.opc.part("xl/persons/person.xml", "application/vnd.ms-excel.person+xml");
    part.add(declaration + '<personList xmlns="' + THREADED_COMMENTS + '">');
    for (const person of this.people) part.add(person);
    part.add("</personList>");
    part.finish();
    this.opc.relate(workbookPart, relationPrefix + "person", "persons/person.xml");
  }
}
