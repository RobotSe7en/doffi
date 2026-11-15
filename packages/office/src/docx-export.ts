import type { NativeCommentBinding } from "./docx-native-comments.js";
import { writeWordEmptyContinuation } from "./docx-table-continuation.js";
import { writeWordRowGapWidths } from "./docx-row-gaps.js";
import { wordCompatibilityNames } from "./ooxml/word-compatibility.js";
import { writeWordConditionalFormat } from "./docx-conditional-format.js";
import { ExportWordRevisions } from "./docx-revisions-export.js";
import { writeWordPageBorders } from "./docx-page-borders.js";
import { W15, CID, CEX } from "./docx-comments.js";
import { WordChartWriter } from "./docx-chart-export.js";
import { writeWordControlProperties } from "./docx-control.js";
import { writeNoteOptions } from "./docx-note-options.js";
import { writeWordMath } from "./docx-math.js";
import { DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, openDocument } from "@doffi/core";
import type {
  DeepReadonly,
  ContentDocumentModel,
  DocumentPackage,
  DocumentReadLimits,
  DocumentSection,
  ContentDocumentSnapshot,
  PageSetup,
  WordMetadata,
  WordDocumentMetadata,
  WordDrawing,
} from "@doffi/core";
import type { DocumentZipOptions } from "./document-zip.js";
import { writeOpcZip } from "./document-zip.js";
import { xml } from "./ooxml/xml.js";
import { ExportAnnotations } from "./export/annotations.js";
import type { ExportComment, UnexportedAnnotation } from "./export/annotations.js";
import { ExportCss } from "./export/css.js";
import type { Css } from "./export/css.js";
import { contentType, OutputOpc } from "./export/opc.js";
import {
  rowExceptionProperties,
  rowProperties,
  cellProperties,
  difference,
  paragraphKeys,
  paragraphProperties,
  pick,
  points,
  runKeys,
  runProperties,
  tableProperties,
  unsupported,
} from "./export/properties.js";
import { children, contains, element, parseSource } from "./export/source.js";
import type { SourceElement, SourceNode, SourceTree } from "./export/source.js";
import type { HtmlNumbering } from "./export/styles.js";
import { ExportStyles } from "./export/styles.js";
import { color, fontName, nativeInteger } from "./export/values.js";
import {
  attributes,
  declaration,
  empty,
  escapeXml,
  namespaces,
  orderedProperties,
  value,
  XmlParts,
} from "./export/xml.js";

import { writeWordDrawing } from "./docx-drawing-export.js";
import { nativeMarker, nativeObjectMarker, restoreDocxNative } from "./docx-native-export.js";

export interface DocxExportOptions extends DocumentZipOptions {}
export interface DocxExportReport {
  readonly documentId: string;
  readonly revision: number;
  readonly exportedAnnotationIds: readonly string[];
  readonly preservedNativeIds?: readonly string[];
  readonly unexportedAnnotations: readonly UnexportedAnnotation[];
}
export interface DocxExportResult {
  readonly bytes: Uint8Array;
  readonly report: DocxExportReport;
}
interface Buffer {
  add(value: string): void;
  finish(): void;
}
type Field = DeepReadonly<WordMetadata["fields"][number]>;
interface FieldState {
  definition: Field;
  offset: number;
  result: boolean;
}
interface InlineRevision {
  tag: string;
  revision: DeepReadonly<WordMetadata["revisions"][number]>;
}
const paragraphTags = new Set(["p", "h1", "h2", "h3", "h4", "h5", "h6"]);
const blockTags = new Set([
  ...paragraphTags,
  "table",
  "div",
  "section",
  "article",
  "ul",
  "ol",
  "li",
]);
const tableRows = (table: SourceElement): SourceElement[] =>
  children(table).flatMap((node) =>
    node.tag === "tr"
      ? [node]
      : ["thead", "tbody", "tfoot"].includes(node.tag)
        ? children(node, "tr")
        : [],
  );
const borderKeys = (css: Css): string[] =>
  Object.keys(css).filter((key) => key.startsWith("border-"));
const date = (raw: string | undefined): string | undefined => {
  if (
    raw !== undefined &&
    (!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)$/.test(raw) ||
      !Number.isFinite(Date.parse(raw)))
  )
    unsupported("A comment or revision date cannot be represented as an XML dateTime.");
  return raw;
};
class Exporter {
  private readonly files: DocumentPackage;
  private readonly snapshot: ContentDocumentSnapshot & {
    readonly metadata: DeepReadonly<WordDocumentMetadata>;
  };
  private readonly tree: SourceTree;
  private readonly css: ExportCss;
  private readonly styles: ExportStyles;
  private readonly annotations: ExportAnnotations;
  private readonly opc: OutputOpc;
  private readonly packageBudget = { bytes: 0, nodes: 0, files: 0 };
  private readonly charts: WordChartWriter;
  private readonly fields: ReadonlyMap<string, Field>;
  private readonly fieldEnds: ReadonlyMap<string, Field>;
  private readonly links: ReadonlyMap<string, DeepReadonly<WordMetadata["hyperlinks"][number]>>;
  private readonly linkEnds: ReadonlySet<string>;
  private readonly revisionIds = new Map<DeepReadonly<WordMetadata["revisions"][number]>, number>();
  private readonly paragraphRevisions = new Map<
    string,
    DeepReadonly<WordMetadata["revisions"][number]>[]
  >();
  private readonly review: ExportWordRevisions;
  private readonly revisions: ReadonlyMap<string, DeepReadonly<WordMetadata["revisions"][number]>>;
  private readonly bookmarkStarts: ReadonlyMap<string, { id: number; name: string }>;
  private readonly bookmarkEnds: ReadonlyMap<string, number>;
  private readonly auxiliary = new Map<
    string,
    { path: string; kind: "header" | "footer"; relation: string }
  >();
  private readonly media = new Map<string, string>();
  private readonly mainStories: SourceElement[];
  private part = "word/document.xml";
  private output: Buffer;
  private fieldStack: FieldState[] = [];
  private deleted = false;
  private inlineRevisions: InlineRevision[] = [];
  private splitRevisionId = 0;
  private width = 468;
  private readonly moveStarts = new Map<
    string,
    { move: DeepReadonly<NonNullable<WordMetadata["moves"]>[number]>; id: number }
  >();
  private readonly moveEnds = new Map<
    string,
    { move: DeepReadonly<NonNullable<WordMetadata["moves"]>[number]>; id: number }
  >();
  private readonly noteIds = new Map<string, number>();
  private drawingId = 0;
  private controlId = 0;
  private readonly drawingIds = new Map<string, number>();
  private readonly drawings: ReadonlyMap<string, DeepReadonly<WordDrawing>>;
  private mappedCells = 0;
  private tableWork = 0;
  private readonly flowStyles = new WeakMap<SourceElement, Css>();
  private readonly visibility = new WeakMap<SourceNode, boolean>();
  private readonly emittedText = new WeakSet<SourceNode>();
  constructor(
    document: ContentDocumentModel,
    private readonly limits: DocumentReadLimits,
  ) {
    const snapshot = document.snapshot();
    if (snapshot.metadata.presentation !== undefined)
      unsupported("PPT sources cannot be exported as DOCX.");
    this.files = document.serialize();
    this.snapshot = { ...snapshot, metadata: snapshot.metadata };
    this.tree = parseSource(this.snapshot.html, limits);
    const word = this.snapshot.metadata.word;
    for (const item of word?.conditionalFormats ?? [])
      this.conditionalFormats.set(item.nodeId, item.format);
    for (const table of word?.tables ?? [])
      for (const context of table.styleContexts ?? [])
        this.tableContexts.set(context.nodeId, context);
    for (const [index, revision] of (word?.revisions ?? []).entries()) {
      this.revisionIds.set(revision, index);
      if (revision.kind.startsWith("paragraph")) {
        const entries = this.paragraphRevisions.get(revision.nodeId) ?? [];
        entries.push(revision);
        this.paragraphRevisions.set(revision.nodeId, entries);
      }
    }
    for (const [ordinal, move] of (word?.moves ?? []).entries()) {
      const item = { move, id: word!.revisions.length + ordinal };
      this.moveStarts.set(move.startNodeId, item);
      this.moveEnds.set(move.endNodeId, item);
    }
    this.drawings = new Map((word?.drawings ?? []).map((item) => [item.nodeId, item]));
    for (const item of word?.drawings ?? []) this.drawingIds.set(item.nodeId, ++this.drawingId);
    this.directCss = new ExportCss([], limits);
    this.css = new ExportCss(
      this.snapshot.styles.map((style) => style.css),
      limits,
      new Set(word?.revisions.map((item) => item.nodeId)),
    );
    for (const node of this.tree.elements)
      if (contains(this.tree.root, node)) this.assertFormatting(node, this.css.get(node));
    const main = word?.stories
      .filter((story) => story.kind === "main")
      .map((story) => this.tree.byId.get(story.rootNodeId)!) ?? [this.tree.root];
    this.styles = new ExportStyles(
      this.css,
      word,
      this.snapshot.metadata.locale,
      limits,
      this.css.get(main[0]!),
    );
    this.mainStories = main;
    this.review = new ExportWordRevisions(word, this.css, this.styles);
    const stories =
      word?.stories
        .filter((story) => story.kind !== "comment")
        .map((story) => this.tree.byId.get(story.rootNodeId)!) ?? main;
    this.opc = new OutputOpc(new XmlParts(limits, this.packageBudget));
    this.charts = new WordChartWriter(this.opc, this.files, limits, this.packageBudget);
    this.output = this.opc.part(
      this.part,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    );
    this.fields = new Map(word?.fields.map((field) => [field.nodeId, field]) ?? []);
    this.fieldEnds = new Map(
      word?.fields.filter((field) => field.endNodeId).map((field) => [field.endNodeId!, field]) ??
        [],
    );
    this.links = new Map(word?.hyperlinks.map((link) => [link.nodeId, link]) ?? []);
    this.linkEnds = new Set(
      word?.hyperlinks.flatMap((link) => (link.endNodeId ? [link.endNodeId] : [])) ?? [],
    );
    this.revisions = new Map(
      word?.revisions
        .filter((revision) => ["insert", "delete", "moveFrom", "moveTo"].includes(revision.kind))
        .map((revision) => [revision.nodeId, revision]) ?? [],
    );
    this.bookmarkStarts = new Map(
      word?.bookmarks.map((bookmark, index) => [
        bookmark.startNodeId,
        { id: index, name: bookmark.name },
      ]) ?? [],
    );
    this.bookmarkEnds = new Map(
      word?.bookmarks.map((bookmark, index) => [bookmark.endNodeId, index]) ?? [],
    );
    for (const story of word?.stories ?? [])
      if (story.kind === "header" || story.kind === "footer") {
        const path = "word/" + story.kind + (this.auxiliary.size + 1) + ".xml";
        this.auxiliary.set(story.storyId, {
          path,
          kind: story.kind,
          relation: this.opc.relate(this.part, story.kind, path.slice(5)),
        });
      }
    const mark = (node: SourceNode): void => {
      if (!element(node)) this.emittedText.add(node);
      else node.children.forEach(mark);
    };
    for (const parent of this.tree.elements)
      if (
        paragraphTags.has(parent.tag) ||
        ["article", "section", "div", "li", "td", "th"].includes(parent.tag)
      ) {
        for (const node of this.flow(parent.children, parent))
          if (!element(node) || this.flowStyles.has(node) || !blockTags.has(node.tag)) mark(node);
      }
    this.annotations = new ExportAnnotations(
      this.tree,
      this.snapshot.annotations,
      stories,
      limits.maxNodes * 128,
      (node) => this.emittedText.has(node),
      word,
    );
  }
  private add(value: string): void {
    this.output.add(value);
  }
  private openInlineRevision(frame: InlineRevision, id: number): void {
    this.add(
      "<w:" +
        frame.tag +
        attributes({
          "w:id": id,
          "w:author": frame.revision.author ?? "",
          "w:date": date(frame.revision.date),
        }) +
        ">",
    );
  }
  /** Hyperlinks are paragraph content, not children of CT_RunTrackChange.
   * Split the current revision around their boundaries with fresh native IDs. */
  private outsideInlineRevisions(markup: string): void {
    for (const frame of [...this.inlineRevisions].reverse()) this.add("</w:" + frame.tag + ">");
    this.add(markup);
    for (const frame of this.inlineRevisions)
      this.openInlineRevision(
        frame,
        this.revisionIds.size + this.moveStarts.size + this.splitRevisionId++,
      );
  }
  private assertFormatting(node: SourceElement, css: Css): void {
    if (this.drawings.has(node.attrs["data-node-id"] ?? "")) return;
    if (css["word-spacing"] && !["normal", "0", "0pt", "0px"].includes(css["word-spacing"]))
      unsupported("Native DOCX cannot preserve independent word spacing in this profile.");
    if (css["font-weight"] && !["normal", "bold", "400", "700"].includes(css["font-weight"]))
      unsupported("Native DOCX supports regular and bold font weights in this profile.");
    if (["pre", "pre-line", "break-spaces"].includes(css["white-space"] ?? ""))
      unsupported("This whitespace layout has no supported native DOCX representation.");
    if (css["list-style-position"] === "inside")
      unsupported("Inside HTML list markers have no supported native DOCX representation.");
    if (css["white-space"] === "nowrap") {
      let cell: SourceElement | undefined = node;
      while (cell && !["td", "th"].includes(cell.tag)) cell = cell.parent;
      if (!cell || this.css.get(cell)["white-space"] !== "nowrap")
        unsupported(
          "Nonwrapping text outside a nonwrapping table cell has no supported native DOCX representation.",
        );
    }
    if (["td", "th"].includes(node.tag) && css.height && !["auto", "initial"].includes(css.height))
      unsupported("Cell height must be applied to its row for native DOCX export.");
    if (!paragraphTags.has(node.tag) && !["table", "td", "th"].includes(node.tag))
      for (const side of ["top", "right", "bottom", "left"]) {
        const style = css["border-" + side + "-style"],
          width = css["border-" + side + "-width"];
        if (
          style &&
          !["none", "hidden"].includes(style) &&
          width !== "0" &&
          width !== "0pt" &&
          width !== "0px"
        )
          unsupported("Borders on this HTML element have no supported native DOCX representation.");
      }
    if (
      css["text-decoration-line"]?.includes("line-through") &&
      !["solid", "double"].includes(css["text-decoration-style"] ?? "solid")
    )
      unsupported("This strike decoration has no supported native DOCX representation.");
    for (const key of ["min-width", "max-width", "min-height", "max-height"])
      if (css[key] && !["none", "auto", "0", "0px", "0pt", "initial"].includes(css[key]))
        unsupported(
          "CSS minimum and maximum dimensions have no supported native DOCX representation.",
        );
    if (!["table", "td", "th", "tr", "img"].includes(node.tag)) {
      for (const key of ["width", "height"])
        if (css[key] && !["auto", "initial"].includes(css[key]))
          unsupported(
            "Fixed dimensions on this HTML element have no supported native DOCX representation.",
          );
      for (const side of ["top", "right", "bottom", "left"])
        if (css["padding-" + side] && points(css["padding-" + side]!, css, this.width) !== 0)
          unsupported("Padding on this HTML element has no supported native DOCX representation.");
    }
    if (
      node.tag === "table" &&
      css["border-collapse"] === "separate" &&
      points(
        (css["--word-table-cell-spacing"] ?? css["border-spacing"] ?? "0").trim().split(/\s+/)[0]!,
        css,
      ) === 0
    )
      unsupported("Separate zero-spacing CSS borders have no native Word equivalent.");
    if (node.tag === "tr" && css["background-color"] !== "transparent")
      unsupported("Row background color must be applied to cells for native DOCX export.");
    if (
      css["white-space"] === "normal" &&
      (paragraphTags.has(node.tag) ||
        (["td", "th", "li"].includes(node.tag) &&
          !children(node).some(
            (child) =>
              paragraphTags.has(child.tag) || ["table", "ul", "ol", "div"].includes(child.tag),
          )))
    ) {
      const lines: string[][] = [[]];
      let normal = true;
      const visit = (child: SourceNode): void => {
        if (!element(child)) {
          if (this.css.get(child.parent)["white-space"] !== "normal") normal = false;
          lines.at(-1)!.push(child.text);
        } else if (child.tag === "br") lines.push([]);
        else if (child.tag === "img") lines.at(-1)!.push("\ufffc");
        else child.children.forEach(visit);
      };
      node.children.forEach(visit);
      if (
        normal &&
        lines.some((parts) => {
          const line = parts.join("");
          return line !== line.replace(/[\t\n\r\f ]+/g, " ").replace(/^ | $/g, "");
        })
      )
        unsupported(
          "Normalize collapsible HTML whitespace or use pre-wrap before native DOCX export.",
        );
    }
  }
  private setup(section: DeepReadonly<DocumentSection>): DeepReadonly<PageSetup> {
    return section.pageSetup ?? this.snapshot.metadata.pageSetup;
  }
  private sectionTextWidth(section: DeepReadonly<DocumentSection>, columns = true): number {
    const setup = this.setup(section),
      available =
        setup.width -
        setup.margins.left -
        setup.margins.right -
        (section.gutterPosition !== "top" ? (section.gutter ?? 0) : 0);
    return columns && section.columns
      ? section.columns.widths
        ? Math.min(...section.columns.widths)
        : (available - section.columns.gap * (section.columns.count - 1)) / section.columns.count
      : available;
  }
  private sectionProperties(section: DeepReadonly<DocumentSection>, review = true): string {
    const setup = this.setup(section);
    let out = "<w:sectPr" + (review ? this.nativeMark(section.rootNodeId) : "") + ">";
    for (const [key, kind] of [
      ["headers", "header"],
      ["footers", "footer"],
    ] as const)
      for (const [type, id] of Object.entries(section[key] ?? {})) {
        const story = id ? this.auxiliary.get(id) : undefined;
        if (!story || story.kind !== kind)
          unsupported("A section auxiliary story cannot be exported.");
        out += empty("w:" + kind + "Reference", { "w:type": type, "r:id": story!.relation });
      }
    out +=
      writeNoteOptions("footnote", section.footnoteOptions) +
      writeNoteOptions("endnote", section.endnoteOptions);
    out += value("type", section.breakType ?? "nextPage");
    out += empty("w:pgSz", {
      "w:w": nativeInteger(setup.width, 20, 1),
      "w:h": nativeInteger(setup.height, 20, 1),
      ...(setup.width > setup.height ? { "w:orient": "landscape" } : {}),
    });
    out += empty("w:pgMar", {
      "w:top": nativeInteger(setup.margins.top, 20, 0),
      "w:right": nativeInteger(setup.margins.right, 20, 0),
      "w:bottom": nativeInteger(setup.margins.bottom, 20, 0),
      "w:left": nativeInteger(setup.margins.left, 20, 0),
      "w:header": nativeInteger(section.headerDistance ?? 36, 20, 0),
      "w:footer": nativeInteger(section.footerDistance ?? 36, 20, 0),
      "w:gutter": nativeInteger(section.gutter ?? 0, 20, 0),
    });
    out += writeWordPageBorders(section.pageBorders);
    if (section.pageNumberStart !== undefined || section.pageNumberFormat)
      out += empty("w:pgNumType", {
        "w:start": section.pageNumberStart,
        "w:fmt": section.pageNumberFormat,
      });
    if (section.columns) {
      const columns = section.columns,
        attrs = {
          "w:num": columns.count,
          "w:space": nativeInteger(columns.gap, 20, 0),
          "w:equalWidth": columns.widths ? 0 : 1,
          "w:sep": columns.separator ? 1 : 0,
        };
      out += columns.widths
        ? "<w:cols" +
          attributes(attrs) +
          ">" +
          columns.widths
            .map((width, index) =>
              empty("w:col", {
                "w:w": nativeInteger(width, 20, 1),
                "w:space": nativeInteger(columns.spaces?.[index] ?? columns.gap, 20, 0),
              }),
            )
            .join("") +
          "</w:cols>"
        : empty("w:cols", attrs);
    }
    if (section.noEndnote !== undefined) out += value("noEndnote", section.noEndnote ? 1 : 0);
    out += value("titlePg", section.differentFirstPage ? "1" : "0");
    if (section.gutterPosition === "right") out += value("rtlGutter", 1);
    if (section.grid)
      out += empty("w:docGrid", {
        "w:type": section.grid.type,
        "w:linePitch":
          section.grid.linePitch === undefined
            ? undefined
            : nativeInteger(section.grid.linePitch, 20, 1),
        "w:charSpace": section.grid.characterSpacing,
      });
    if (review)
      out += this.review.section(this.tree.byId.get(section.rootNodeId)!, (previous) =>
        this.sectionProperties(
          { sectionId: section.sectionId, rootNodeId: section.rootNodeId, ...previous },
          false,
        ),
      );
    out += "</w:sectPr>";
    return out;
  }
  private runDifference(css: Css, expected: Css): Record<string, string> {
    const properties = { ...difference(css, expected, runKeys) };
    for (const key of [
      "font-family",
      "--word-font-ascii",
      "--word-font-hAnsi",
      "--word-font-eastAsia",
      "--word-font-cs",
    ])
      if (properties[key] && expected[key] && fontName(properties[key]) === fontName(expected[key]))
        delete properties[key];
    if (properties.color && expected.color && color(properties.color) === color(expected.color))
      delete properties.color;
    const weight = (raw: string | undefined): number =>
      raw === "bold" ? 700 : raw === "normal" ? 400 : Number(raw);
    if (
      properties["font-weight"] &&
      weight(properties["font-weight"]) === weight(expected["font-weight"])
    )
      delete properties["font-weight"];
    if (
      properties["text-decoration-color"] &&
      expected["text-decoration-color"]?.toLowerCase() === "currentcolor" &&
      color(properties["text-decoration-color"], css.color) === color(css.color!)
    )
      delete properties["text-decoration-color"];
    return properties;
  }
  private nativeRun(
    css: Css,
    expected: Css,
    character?: string,
    role?: string,
    language?: string,
    changes = "",
  ): string {
    const properties = this.runDifference(css, expected);
    // A script span selects its source role. Preserve the other three source
    // families even when a host changes just one script's font-family.
    if (role) {
      if (
        fontName(css["font-family"]!) !==
        fontName(expected["--word-font-" + role] ?? expected["font-family"]!)
      )
        properties["font-family"] = css["font-family"]!;
      else delete properties["font-family"];
    }
    const r =
      (character ? value("rStyle", this.styles.ids.get(character) ?? character) : "") +
      runProperties(
        {
          ...properties,
          ...(css["--word-language-val"] === undefined && language
            ? { "--word-language-val": JSON.stringify(language) }
            : {}),
        },
        expected,
        false,
        role,
      );
    return r || changes ? "<w:rPr>" + r + changes + "</w:rPr>" : "";
  }
  private nativeMark(nodeId: string | undefined): string {
    return this.snapshot.metadata.word?.nativeContent ? nativeMarker(nodeId) : "";
  }
  private nativeRunOwner(node: SourceNode): string | undefined {
    for (let current: SourceNode | undefined = node; current; current = current.parent)
      if (element(current)) {
        const id = current.attrs["data-node-id"];
        if (
          id &&
          this.snapshot.metadata.word?.nativeContent?.fragments.some(
            (f) => f.nodeId === id && f.ownerName === "r",
          )
        )
          return id;
      }
    return undefined;
  }
  private nativeRunReset(node: SourceNode): string {
    if (!this.snapshot.metadata.word?.nativeContent) return "";
    for (let current: SourceNode | undefined = node; current; current = current.parent)
      if (
        element(current) &&
        (current.attrs["data-word-reset-text"] === "true" ||
          current.attrs["data-word-current-format"] === "true")
      )
        return ' dn:reset-run="1"';
    return "";
  }
  private emitText(
    node: SourceNode,
    text: string,
    css: Css,
    expected: Css,
    character?: string,
    role?: string,
    language?: string,
  ): void {
    const boundaries = [...new Set([0, text.length, ...this.annotations.offsets(node)])].sort(
      (a, b) => a - b,
    );
    for (let index = 0; index < boundaries.length; index++) {
      const start = boundaries[index]!;
      this.add(this.annotations.at(node, start));
      const end = boundaries[index + 1];
      if (end === undefined || end <= start) continue;
      const textName = this.deleted ? "delText" : "t",
        raw = text.slice(start, end);
      const pieces = raw.split(/([\t\n])/);
      this.add(
        "<w:r" +
          this.nativeMark(this.nativeRunOwner(node)) +
          this.nativeRunReset(node) +
          ">" +
          this.nativeRun(
            css,
            expected,
            character,
            role ?? (/^[\u0000-\u007f]*$/.test(raw) ? "ascii" : undefined),
            language,
            this.review.run(node, expected),
          ),
      );
      for (const piece of pieces)
        if (piece)
          this.add(
            piece === "\t"
              ? "<w:tab/>"
              : piece === "\n"
                ? "<w:br/>"
                : "<w:" +
                  textName +
                  ' xml:space="preserve">' +
                  escapeXml(piece) +
                  "</w:" +
                  textName +
                  ">",
          );
      this.add("</w:r>");
    }
  }
  private fieldControl(type: "begin" | "separate" | "end", field?: Field): void {
    const source = field ? this.tree.byId.get(field.nodeId) : undefined;
    this.add(
      "<w:r" +
        this.nativeMark(source ? this.nativeRunOwner(source) : undefined) +
        (source ? this.nativeRunReset(source) : "") +
        ">" +
        empty("w:fldChar", {
          "w:fldCharType": type,
          ...(field
            ? {
                "w:fldLock": field.locked === undefined ? undefined : field.locked ? 1 : 0,
                "w:dirty": field.dirty === undefined ? undefined : field.dirty ? 1 : 0,
              }
            : {}),
        }) +
        "</w:r>",
    );
  }
  private instruction(state: FieldState, until: number): void {
    if (until < state.offset || until > state.definition.instruction.length)
      unsupported("A nested field instruction offset cannot be exported.");
    if (until > state.offset)
      this.add(
        '<w:r><w:instrText xml:space="preserve">' +
          escapeXml(state.definition.instruction.slice(state.offset, until)) +
          "</w:instrText></w:r>",
      );
    state.offset = until;
  }
  private moveMarker(nodeId: string): string | undefined {
    const start = this.moveStarts.get(nodeId),
      end = this.moveEnds.get(nodeId),
      item = start ?? end;
    if (!item) return;
    const name =
      "move" + (item.move.kind === "from" ? "From" : "To") + "Range" + (start ? "Start" : "End");
    return (
      "<w:" +
      name +
      this.nativeMark(nodeId) +
      attributes({
        "w:id": item.id,
        ...(start
          ? {
              "w:name": item.move.moveId,
              "w:author": item.move.author,
              "w:date": date(item.move.date),
              "w:colFirst": item.move.columnFirst,
              "w:colLast": item.move.columnLast,
              "w:displacedByCustomXml": item.move.startDisplaced,
            }
          : { "w:displacedByCustomXml": item.move.endDisplaced }),
      }) +
      "/>"
    );
  }
  private paragraphRevision(nodeId: string): string {
    const revisions = [...(this.paragraphRevisions.get(nodeId) ?? [])],
      order = ["paragraphInsert", "paragraphDelete", "paragraphMoveFrom", "paragraphMoveTo"];
    return revisions
      .sort((a, b) => order.indexOf(a.kind) - order.indexOf(b.kind))
      .map((revision) =>
        empty(
          "w:" +
            (
              {
                paragraphInsert: "ins",
                paragraphDelete: "del",
                paragraphMoveFrom: "moveFrom",
                paragraphMoveTo: "moveTo",
              } as Record<string, string>
            )[revision.kind],
          {
            "w:id": this.revisionIds.get(revision)!,
            "w:author": revision.author ?? "",
            "w:date": date(revision.date),
          },
        ),
      )
      .join("");
  }
  private inline(
    node: SourceNode,
    paragraphStyle: string | undefined,
    character?: string,
    role?: string,
    language?: string,
  ): void {
    if (!element(node)) {
      const css = this.css.get(node.parent);
      const textCss = ["td", "th"].includes(node.parent.tag)
        ? { ...css, "vertical-align": "baseline", "background-color": "transparent" }
        : css;
      this.emitText(
        node,
        node.text,
        textCss,
        this.styleContext(node, this.styles.runContext(paragraphStyle, character)),
        character,
        role,
        language,
      );
      return;
    }
    const control = this.snapshot.metadata.word?.controls?.find(
      (control) => control.nodeId === node.attrs["data-node-id"],
    );
    if (control) {
      this.add(this.annotations.at(node, 0));
      this.add(
        "<w:sdt" +
          this.nativeMark(control.nodeId) +
          ">" +
          writeWordControlProperties(
            control,
            ++this.controlId,
            this.nativeRun(
              this.css.get(node),
              this.styleContext(node, this.styles.runContext(paragraphStyle, character)),
              character,
              role,
              language,
            ),
          ) +
          "<w:sdtContent>",
      );
      for (const child of node.children)
        this.inline(child, paragraphStyle, character, role, language);
      this.add("</w:sdtContent></w:sdt>");
      this.add(this.annotations.at(node, node.to - node.from));
      return;
    }
    if (this.drawings.has(node.attrs["data-node-id"] ?? "")) {
      this.add(this.annotations.at(node, 0));
      this.add(
        writeWordDrawing(
          {
            word: this.snapshot.metadata.word!,
            tree: this.tree,
            css: (node) => this.css.get(node),
            raw: (node) => this.css.specified(node),
            id: (id) => this.drawingIds.get(id) ?? unsupported("A drawing identity is missing."),
            image: (node) => this.imageRelation(node),
            chart: (chart) => {
              const node = this.tree.byId.get(chart.nodeId)!;
              return this.charts.write(
                chart.chart,
                this.part,
                this.css.get(node),
                this.css.specified(node),
              );
            },
            text: (id, width) => this.captureStory(id, width),
            marker: (node) => {
              let container = node.parent;
              while (container && !container.attrs["data-word-run"]) container = container.parent;
              return (
                this.nativeMark(node.attrs["data-node-id"]!) +
                (this.snapshot.metadata.word?.nativeContent && container
                  ? ' dn:container="' + escapeXml(container.attrs["data-node-id"]!) + '"'
                  : "")
              );
            },
          },
          this.drawings.get(node.attrs["data-node-id"]!)!,
        ),
      );
      this.add(this.annotations.at(node, node.to - node.from));
      return;
    }
    if (node.attrs["data-word-native-object"]) {
      this.add(this.annotations.at(node, 0));
      let container = node.parent;
      while (container && !container.attrs["data-word-run"]) container = container.parent;
      this.add(
        nativeObjectMarker(
          node.attrs["data-word-native-object"],
          false,
          node.attrs["data-node-id"],
          container?.attrs["data-node-id"],
        ),
      );
      this.add(this.annotations.at(node, node.to - node.from));
      return;
    }
    const id = node.attrs["data-node-id"] ?? "",
      css = this.css.get(node);
    const move = this.moveMarker(id);
    if (move !== undefined) {
      this.add(move);
      return;
    }
    if (node.attrs["data-word-equation"]) {
      const equation = this.snapshot.metadata.word?.equations?.find(
        (equation) => equation.nodeId === id,
      );
      if (!equation) unsupported("The equation definition is missing.");
      this.add(this.annotations.at(node, 0));
      this.add(writeWordMath(equation));
      this.add(this.annotations.at(node, node.to - node.from));
      return;
    }
    if (node.attrs["data-word-note-reference"]) {
      const reference = this.snapshot.metadata.word?.noteReferences?.find(
          (reference) => reference.nodeId === id,
        ),
        note = this.snapshot.metadata.word?.notes?.find(
          (note) => note.noteId === reference?.noteId,
        );
      if (!reference || !note) unsupported("A note reference is missing its source definition.");
      this.add(this.annotations.at(node, 0));
      this.add(
        "<w:r" +
          this.nativeMark(id) +
          ">" +
          this.nativeRun(
            css,
            this.styleContext(node, this.styles.runContext(paragraphStyle, character)),
            character,
            role,
            language,
          ) +
          empty("w:" + note.kind + "Reference", {
            "w:id": this.noteIds.get(note.noteId),
            "w:customMarkFollows": reference.customMark ? 1 : undefined,
          }) +
          (reference.customMark ? "<w:t>" + escapeXml(reference.customMark) + "</w:t>" : "") +
          "</w:r>",
      );
      this.add(this.annotations.at(node, node.to - node.from));
      return;
    }
    if (node.attrs["data-word-note-number"]) {
      const note = this.snapshot.metadata.word?.notes?.find((note) => {
        const story = this.snapshot.metadata.word!.stories.find(
          (story) => story.storyId === note.storyId,
        );
        return !!story && contains(this.tree.byId.get(story.rootNodeId)!, node);
      });
      if (!note) unsupported("A note number requires its note story.");
      const writeNumber = (): void =>
        this.add(
          "<w:r>" +
            this.nativeRun(
              css,
              this.styleContext(node, this.styles.runContext(paragraphStyle, character)),
              character,
              role,
              language,
            ) +
            empty("w:" + note.kind + "Ref") +
            "</w:r>",
        );
      // Older source writers could place inserted content inside the counter
      // span. Emit its actual children in order instead of dropping their fields.
      let written = false;
      for (const child of node.children) {
        if (element(child)) {
          this.inline(child, paragraphStyle, character, role, language);
          continue;
        }
        const at = written ? -1 : child.text.indexOf("\ufffc");
        if (at < 0) {
          this.inline(child, paragraphStyle, character, role, language);
          continue;
        }
        if (at)
          this.inline(
            { ...child, text: child.text.slice(0, at), to: child.from + at },
            paragraphStyle,
            character,
            role,
            language,
          );
        writeNumber();
        written = true;
        if (at + 1 < child.text.length)
          this.inline(
            { ...child, text: child.text.slice(at + 1), from: child.from + at + 1 },
            paragraphStyle,
            character,
            role,
            language,
          );
      }
      if (!written) writeNumber();
      return;
    }
    if (node.attrs["data-word-note-separator"]) {
      this.add("<w:r>" + empty("w:" + node.attrs["data-word-note-separator"]) + "</w:r>");
      return;
    }
    if (node.attrs["data-word-reset-text"] === "true") character = undefined;
    const requestedStyle =
      node.attrs["data-word-style"] ?? node.attrs["data-word-ignored-character-style"];
    if (requestedStyle) character = requestedStyle;
    role =
      node.attrs["data-word-script"] === "eastAsia"
        ? "eastAsia"
        : node.attrs["data-word-script"] === "cs"
          ? "cs"
          : node.attrs["data-word-script"] === "hAnsi"
            ? "hAnsi"
            : node.attrs["data-word-script"] === "ascii"
              ? "ascii"
              : role;
    language = node.attrs.lang ?? language;
    if (node.attrs["data-word-paragraph-mark"]) return;
    const field = this.fields.get(id),
      end = this.fieldEnds.get(id),
      parent = this.fieldStack.at(-1);
    if (parent && !parent.result) {
      const nested =
        field &&
        parent.definition.instructionFields?.find((item) => item.fieldNodeId === field.nodeId);
      if (nested) this.instruction(parent, nested.offset);
      else if (
        parent.definition.resultNodeIds?.includes(id) ||
        node.attrs["data-word-field-result"] === parent.definition.nodeId
      ) {
        this.instruction(parent, parent.definition.instruction.length);
        this.fieldControl("separate");
        parent.result = true;
      }
    }
    this.add(this.annotations.at(node, 0));
    const bookmark = this.bookmarkStarts.get(id);
    if (bookmark)
      this.add(empty("w:bookmarkStart", { "w:id": bookmark.id, "w:name": bookmark.name }));
    const bookmarkEnd = this.bookmarkEnds.get(id);
    if (bookmarkEnd !== undefined) this.add(empty("w:bookmarkEnd", { "w:id": bookmarkEnd }));
    if (end) {
      const current = this.fieldStack.pop();
      if (current?.definition !== end)
        unsupported("Unbalanced complex field controls cannot be exported.");
      if (!current.result) {
        this.instruction(current, end.instruction.length);
        this.fieldControl("separate");
      }
      this.fieldControl("end");
      return;
    }
    if (field) {
      if (!field.endNodeId) {
        if (this.inlineRevisions.length) {
          this.fieldControl("begin", field);
          const state: FieldState = { definition: field, offset: 0, result: true };
          this.instruction(state, field.instruction.length);
          this.fieldControl("separate");
          this.fieldStack.push(state);
          for (const child of node.children)
            this.inline(child, paragraphStyle, character, role, language);
          this.fieldStack.pop();
          this.fieldControl("end");
          return;
        }
        this.add(
          "<w:fldSimple" +
            attributes({
              "w:instr": field.instruction,
              "w:fldLock": field.locked === undefined ? undefined : field.locked ? 1 : 0,
              "w:dirty": field.dirty === undefined ? undefined : field.dirty ? 1 : 0,
            }) +
            ">",
        );
        for (const child of node.children)
          this.inline(child, paragraphStyle, character, role, language);
        this.add("</w:fldSimple>");
        return;
      }
      this.fieldControl("begin", field);
      const state = { definition: field, offset: 0, result: false };
      this.fieldStack.push(state);
      this.instruction(state, field.instructionFields?.[0]?.offset ?? field.instruction.length);
      return;
    }
    const link = this.links.get(id),
      revision = this.revisions.get(id);
    if (this.linkEnds.has(id)) {
      this.outsideInlineRevisions("</w:hyperlink>");
      return;
    }
    if (link)
      this.outsideInlineRevisions(
        "<w:hyperlink" +
          attributes({
            "w:anchor": link.bookmark,
            "r:id": link.target
              ? this.opc.relate(this.part, "hyperlink", link.target, true)
              : undefined,
          }) +
          ">",
      );
    const previousDeleted = this.deleted;
    const revisionTag = revision
      ? (
          { insert: "ins", delete: "del", moveFrom: "moveFrom", moveTo: "moveTo" } as Record<
            string,
            string
          >
        )[revision.kind]
      : undefined;
    if (revision) {
      const frame = { tag: revisionTag!, revision };
      this.openInlineRevision(frame, this.revisionIds.get(revision)!);
      this.inlineRevisions.push(frame);
      this.deleted = ["delete", "moveFrom"].includes(revision.kind);
    }
    if (node.tag === "br") {
      this.add(
        "<w:r" +
          this.nativeMark(this.nativeRunOwner(node)) +
          this.nativeRunReset(node) +
          ">" +
          this.nativeRun(
            css,
            this.styleContext(node, this.styles.runContext(paragraphStyle, character)),
            character,
            role,
            language,
          ) +
          empty(
            "w:br",
            ["page", "column"].includes(node.attrs["data-word-break"] ?? "")
              ? { "w:type": node.attrs["data-word-break"] }
              : {},
          ) +
          "</w:r>",
      );
      this.add(this.annotations.at(node, 1));
    } else if (node.tag === "img") {
      this.image(node, css, paragraphStyle, character, role, language);
      this.add(this.annotations.at(node, 1));
    } else if (node.attrs["data-word-tab"])
      this.emitText(
        node,
        "\t",
        css,
        this.styleContext(node, this.styles.runContext(paragraphStyle, character)),
        character,
        role,
        language,
      );
    else if (node.attrs["data-word-run"] === "true" && !node.children.length) {
      const expected = this.styleContext(node, this.styles.runContext(paragraphStyle, character));
      this.add(
        "<w:r" +
          this.nativeMark(this.nativeRunOwner(node)) +
          this.nativeRunReset(node) +
          ">" +
          this.nativeRun(
            css,
            expected,
            character,
            role,
            language,
            this.review.run(node, expected),
          ) +
          "</w:r>",
      );
    } else
      for (const child of node.children)
        this.inline(child, paragraphStyle, character, role, language);
    if (revision) {
      this.add("</w:" + revisionTag + ">");
      this.inlineRevisions.pop();
      this.deleted = previousDeleted;
    }
    if (link && !link.endNodeId) this.outsideInlineRevisions("</w:hyperlink>");
  }
  private imageRelation(node: SourceElement): string {
    const parts = this.snapshot.manifest.entrypoint.split("/").slice(0, -1);
    // Source was validated against this entrypoint by core; decode exactly one
    // URL layer, then look up only the already supplied resource bytes.
    for (const token of node.attrs.src!.split("/")) {
      const segment = decodeURIComponent(token);
      if (segment === "..") parts.pop();
      else if (segment !== ".") parts.push(segment);
    }
    const src = parts.join("/"),
      descriptor = this.snapshot.manifest.assets.find((asset) => asset.path === src);
    if (!descriptor)
      throw new DocumentError("AssetUnavailable", "An exported image resource is missing.");
    let path = this.media.get(src);
    if (!path) {
      const ext: Record<string, string> = {
        "image/png": "png",
        "image/jpeg": "jpeg",
        "image/webp": "webp",
      };
      if (!ext[descriptor.mimeType])
        unsupported("An image container is outside the native DOCX export profile.");
      path = "word/media/image" + (this.media.size + 1) + "." + ext[descriptor.mimeType];
      this.media.set(src, path);
      this.opc.binary(path, descriptor.mimeType, this.files.get(src)!);
    }
    return this.opc.relate(this.part, "image", path.slice(5));
  }
  private captureStory(storyId: string, width: number): string {
    const story = this.snapshot.metadata.word!.stories.find((story) => story.storyId === storyId)!;
    const output = this.output,
      chunks: string[] = [];
    let length = 0;
    this.output = {
      add: (value) => {
        length += value.length;
        if (length > this.limits.maxFileBytes)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Drawing text exceeds the native part budget.",
          );
        chunks.push(value);
      },
      finish: () => {},
    };
    try {
      this.withStory(this.part, () => {
        this.width = width;
        this.blocks(this.tree.byId.get(story.rootNodeId)!.children);
      });
      return chunks.join("");
    } finally {
      this.output = output;
    }
  }
  private image(
    node: SourceElement,
    css: Css,
    paragraphStyle?: string,
    character?: string,
    role?: string,
    language?: string,
  ): void {
    const width = css.width ? points(css.width, css, this.width) : Number(node.attrs.width) * 0.75,
      height = css.height ? points(css.height, css, this.width) : Number(node.attrs.height) * 0.75;
    if (!(width > 0) || !(height > 0))
      unsupported("An inline image requires finite native dimensions.");
    const cx = nativeInteger(width, 12700, 1),
      cy = nativeInteger(height, 12700, 1),
      id = ++this.drawingId;
    const rel = this.imageRelation(node);
    this.add(
      "<w:r>" +
        this.nativeRun(
          css,
          this.styleContext(node, this.styles.runContext(paragraphStyle, character)),
          character,
          role,
          language,
        ) +
        '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
        empty("wp:extent", { cx, cy }) +
        empty("wp:docPr", { id, name: "Image " + id, descr: node.attrs.alt ?? "" }) +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
        "<pic:nvPicPr>" +
        empty("pic:cNvPr", { id: 0, name: "Image " + id, descr: node.attrs.alt ?? "" }) +
        "<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>" +
        empty("a:blip", { "r:embed": rel }) +
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/>' +
        empty("a:ext", { cx, cy }) +
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>',
    );
  }
  private readonly conditionalFormats = new Map<
    string,
    import("@doffi/core").DeepReadonly<import("@doffi/core").WordConditionalFormatValue>
  >();
  private conditionalFormat(node: SourceElement): string {
    return writeWordConditionalFormat(
      this.conditionalFormats.get(node.attrs["data-node-id"] ?? ""),
    );
  }
  private readonly directCss: ExportCss;
  private readonly tableContexts = new Map<string, { cssClass: string; kind: string }>();
  private tablePropertyDifference(
    node: SourceElement,
    css: Css,
    kind: "table" | "row" | "cell",
  ): Css {
    const result = { ...difference(css, this.styleContext(node, {}, kind), Object.keys(css)) };
    if (node.attrs.style) {
      const direct = this.directCss.specified({
        tag: "span",
        attrs: { style: node.attrs.style },
        parent: undefined,
      });
      Object.assign(result, pick(css, Object.keys(direct)));
    }
    if (
      kind === "table" &&
      Object.keys(result).some(
        (key) => key === "--word-table-floating" || key.startsWith("--word-table-position-"),
      )
    )
      for (const [key, value] of Object.entries(css))
        if (key === "--word-table-floating" || key.startsWith("--word-table-position-"))
          result[key] = value;
    return result;
  }
  private styleContext(
    node: SourceNode,
    fallback: Css,
    kind?: "table" | "row" | "cell" | "paragraph",
  ): Css {
    const contexts: { name: string; node: SourceElement }[] = [];
    let reset = false;
    for (let current: SourceNode | undefined = node; current; current = current.parent) {
      if (!element(current)) continue;
      const context = this.tableContexts.get(current.attrs["data-node-id"] ?? "");
      if (
        context &&
        (!reset || context.kind !== "run") &&
        ((!kind && ["run", "paragraph"].includes(context.kind)) || context.kind === kind)
      )
        contexts.unshift({ name: context.cssClass, node: current });
      if (current.attrs["data-word-reset-text"] === "true") reset = true;
      if (kind || current.tag === "table") break;
    }
    let result = fallback;
    for (const context of contexts) {
      const parent = context.node.parent ? this.css.get(context.node.parent) : {},
        variables = Object.fromEntries(
          Object.entries(parent).filter(([key]) => key.startsWith("--word-border-inside-")),
        );
      result = this.css.classStyle(context.name, { ...result, ...variables });
    }
    return result;
  }
  private paragraph(
    node: SourceElement,
    section?: string,
    paragraphId?: string,
    list?: { id: string; level: number },
  ): void {
    const css = this.flowStyles.get(node) ?? this.css.get(node),
      id = node.attrs["data-node-id"] ?? "",
      style = node.attrs["data-word-style"];
    this.add("<w:p" + this.nativeMark(id) + attributes({ "w14:paraId": paragraphId }) + ">");
    let props =
      this.conditionalFormat(node) +
      (style ? value("pStyle", this.styles.ids.get(style) ?? style) : "");
    const numId = node.attrs["data-word-numbering"]
      ? this.styles.numberingIds.get(node.attrs["data-word-numbering"])
      : list?.id;
    if (node.attrs["data-word-numbering-off"] === "true")
      props += "<w:numPr>" + value("numId", 0) + "</w:numPr>";
    else if (numId)
      props +=
        "<w:numPr>" +
        value("ilvl", node.attrs["data-word-level"] ?? list?.level ?? 0) +
        value("numId", numId) +
        "</w:numPr>";
    const expectedParagraph = this.styleContext(
        node,
        this.styles.paragraphContext(style),
        "paragraph",
      ),
      paragraphCss = { ...css };
    for (const side of ["top", "right", "bottom", "left"])
      if (
        !paragraphCss["border-" + side + "-style"] &&
        expectedParagraph["border-" + side + "-style"] &&
        expectedParagraph["border-" + side + "-style"] !== "none"
      ) {
        paragraphCss["border-" + side + "-style"] = "none";
        paragraphCss["border-" + side + "-width"] = "0pt";
      }
    props += paragraphProperties(
      {
        ...difference(paragraphCss, expectedParagraph, [
          ...paragraphKeys,
          ...borderKeys(paragraphCss),
        ]),
        "font-size": css["font-size"]!,
      },
      this.width,
    );
    const tabs = this.snapshot.metadata.word?.tabs.find((tab) => tab.paragraphNodeId === id);
    if (tabs?.stops.length)
      props +=
        "<w:tabs>" +
        tabs.stops
          .map((tab) =>
            empty("w:tab", {
              "w:val": tab.alignment,
              "w:leader": tab.leader,
              "w:pos": nativeInteger(tab.position),
            }),
          )
          .join("") +
        "</w:tabs>";
    if (node.attrs["data-word-outline-level"] !== undefined)
      props += value("outlineLvl", node.attrs["data-word-outline-level"]);
    else if (/^h[1-6]$/.test(node.tag)) props += value("outlineLvl", Number(node.tag[1]) - 1);
    const mark = children(node).find((child) => child.attrs["data-word-paragraph-mark"]);
    const markStyle = mark?.attrs["data-word-style"],
      markContext = this.styleContext(node, this.styles.runContext(style, markStyle), "paragraph");
    const markProperties = mark
      ? this.runDifference(this.css.get(mark), markContext)
      : pick(
          this.runDifference(css, markContext),
          runKeys.filter(
            (key) =>
              key !== "background-color" &&
              !key.startsWith("text-decoration") &&
              ![
                "vertical-align",
                "letter-spacing",
                "--word-run-shading",
                "--word-kerning-threshold",
              ].includes(key),
          ),
        );
    const markRevisions = this.paragraphRevision(id),
      markHistory = mark ? this.review.run(mark, markContext) : "";
    if (mark || Object.keys(markProperties).length || markRevisions || markHistory)
      props +=
        "<w:rPr" +
        (mark ? this.nativeMark(mark.attrs["data-node-id"]) + this.nativeRunReset(mark) : "") +
        ">" +
        markRevisions +
        (markStyle ? value("rStyle", this.styles.ids.get(markStyle) ?? markStyle) : "") +
        runProperties(markProperties, markContext) +
        markHistory +
        "</w:rPr>";
    props += section ?? "";
    props += this.review.paragraph(node, this.width);
    this.add("<w:pPr>" + orderedProperties("pPr", props) + "</w:pPr>");
    this.add(this.annotations.at(node, 0));
    for (const child of node.children) this.inline(child, style);
    this.add("</w:p>");
  }
  /** Group anonymous inline flow before writing paragraphs. Whitespace-only
   * nodes between blocks are formatting only under normal white-space; spaces
   * within an inline run and all pre-wrap characters retain their source leaf. */
  private flow(nodes: readonly SourceNode[], container = nodes[0]?.parent): SourceNode[] {
    const result: SourceNode[] = [];
    let pending: SourceNode[] = [];
    const flush = (): void => {
      if (!pending.length) return;
      const visible = pending.some((node) => this.visible(node));
      if (!visible) {
        for (const node of pending)
          if (
            element(node) &&
            (this.bookmarkStarts.has(node.attrs["data-node-id"] ?? "") ||
              this.bookmarkEnds.has(node.attrs["data-node-id"] ?? "") ||
              this.moveStarts.has(node.attrs["data-node-id"] ?? "") ||
              this.moveEnds.has(node.attrs["data-node-id"] ?? ""))
          )
            result.push(node);
      } else {
        const parent = container ?? pending[0]!.parent!;
        const paragraph: SourceElement = {
          tag: "p",
          attrs: {},
          parent,
          children: pending,
          from: pending[0]!.from,
          to: pending.at(-1)!.to,
        };
        const css = this.css.get(parent);
        this.flowStyles.set(
          paragraph,
          ["td", "th"].includes(parent.tag)
            ? {
                ...this.css.emptyStyle(),
                ...pick(css, [
                  ...runKeys,
                  "text-align",
                  "text-indent",
                  "line-height",
                  "widows",
                  "orphans",
                  ...Object.keys(css).filter((key) => key.startsWith("--")),
                ]),
                "background-color": "transparent",
                "vertical-align": "baseline",
              }
            : css,
        );
        result.push(paragraph);
      }
      pending = [];
    };
    for (const node of nodes) {
      if (element(node) && blockTags.has(node.tag)) {
        flush();
        result.push(node);
      } else pending.push(node);
    }
    flush();
    return result;
  }
  private visible(node: SourceNode): boolean {
    const known = this.visibility.get(node);
    if (known !== undefined) return known;
    const result = element(node)
      ? ["img", "br"].includes(node.tag) || node.children.some((child) => this.visible(child))
      : /[^\t\n\r\f ]/.test(node.text) ||
        (node.text.length > 0 && this.css.get(node.parent)["white-space"] !== "normal");
    this.visibility.set(node, result);
    return result;
  }
  private blocks(
    nodes: readonly SourceNode[],
    finalSection?: string,
    paragraphId?: string,
    hasFollowing = false,
  ): void {
    const significant = this.flow(nodes);
    if (!significant.length) {
      if (hasFollowing && !finalSection && !paragraphId) return;
      this.add(
        "<w:p" +
          attributes({ "w14:paraId": paragraphId }) +
          "><w:pPr>" +
          (finalSection ?? "") +
          "</w:pPr></w:p>",
      );
      return;
    }
    for (let index = 0; index < significant.length; index++) {
      const node = significant[index]!,
        last = index === significant.length - 1;
      if (element(node) && node.attrs["data-word-native-object"]) {
        this.add(this.annotations.at(node, 0, true));
        this.add(
          nativeObjectMarker(
            node.attrs["data-word-native-object"],
            true,
            node.attrs["data-node-id"],
          ),
        );
        this.add(this.annotations.at(node, node.to - node.from, true));
        if (last && finalSection) this.add("<w:p><w:pPr>" + finalSection + "</w:pPr></w:p>");
      } else if (
        element(node) &&
        node.from === node.to &&
        (this.bookmarkStarts.has(node.attrs["data-node-id"] ?? "") ||
          this.bookmarkEnds.has(node.attrs["data-node-id"] ?? "") ||
          this.moveStarts.has(node.attrs["data-node-id"] ?? "") ||
          this.moveEnds.has(node.attrs["data-node-id"] ?? ""))
      ) {
        this.inline(node, undefined);
        if (last && (finalSection || paragraphId))
          this.add(
            "<w:p" +
              attributes({ "w14:paraId": paragraphId }) +
              "><w:pPr>" +
              (finalSection ?? "") +
              "</w:pPr></w:p>",
          );
      } else if (
        element(node) &&
        node.attrs["data-word-content-control"] &&
        node.attrs["data-word-content-control"] !== "toc"
      ) {
        const control = this.snapshot.metadata.word?.controls?.find(
          (control) => control.nodeId === node.attrs["data-node-id"],
        );
        if (!control) unsupported("The content control definition is missing.");
        this.add(
          "<w:sdt" +
            this.nativeMark(control.nodeId) +
            ">" +
            writeWordControlProperties(
              control,
              ++this.controlId,
              this.nativeRun(this.css.get(node), this.styles.runContext(undefined)),
            ) +
            "<w:sdtContent>",
        );
        this.blocks(node.children, last ? finalSection : undefined, last ? paragraphId : undefined);
        this.add("</w:sdtContent></w:sdt>");
      } else if (element(node) && paragraphTags.has(node.tag))
        this.paragraph(node, last ? finalSection : undefined, last ? paragraphId : undefined);
      else if (element(node) && node.tag === "table") {
        this.table(node);
        if (last && (!hasFollowing || finalSection || paragraphId))
          this.add(
            "<w:p" +
              attributes({ "w14:paraId": paragraphId }) +
              "><w:pPr>" +
              (finalSection ?? "") +
              "</w:pPr></w:p>",
          );
      } else if (
        element(node) &&
        ["div", "section", "article", "li", "ul", "ol"].includes(node.tag)
      ) {
        if (node.tag === "ul" || node.tag === "ol")
          this.list(node, last ? finalSection : undefined, last ? paragraphId : undefined);
        else {
          const toc = node.attrs["data-word-content-control"] === "toc";
          if (toc)
            this.add(
              "<w:sdt" +
                this.nativeMark(node.attrs["data-node-id"]) +
                "><w:sdtPr>" +
                ["id", "alias", "tag"]
                  .map((key) => value(key, node.attrs["data-word-control-" + key]))
                  .join("") +
                "<w:docPartObj>" +
                value("docPartGallery", "Table of Contents") +
                "<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>",
            );
          this.blocks(
            node.children,
            last ? finalSection : undefined,
            last ? paragraphId : undefined,
            !last || hasFollowing,
          );
          if (toc) this.add("</w:sdtContent></w:sdt>");
        }
      } else {
        this.add(
          "<w:p" +
            attributes({ "w14:paraId": last ? paragraphId : undefined }) +
            "><w:pPr>" +
            (last ? (finalSection ?? "") : "") +
            "</w:pPr>",
        );
        this.inline(node, undefined);
        this.add("</w:p>");
      }
    }
  }
  private table(node: SourceElement): void {
    const css = this.css.get(node),
      info = this.snapshot.metadata.word?.tables.find(
        (table) => table.nodeId === node.attrs["data-node-id"],
      ),
      rows = tableRows(node);
    const style = node.attrs["data-word-style"];
    interface Cell {
      node: SourceElement;
      column: number;
      span: number;
      rows: number;
      continuation: boolean;
    }
    const matrix: Cell[][] = [];
    let active = new Map<number, Cell>(),
      columns = 0;
    for (const row of rows) {
      const rowCells: Cell[] = [];
      let column = 0,
        current = 0,
        remaining = active.size;
      const cells = children(row).filter((cell) => ["td", "th"].includes(cell.tag)),
        next = new Map<number, Cell>();
      while (current < cells.length || remaining > 0) {
        // The expansion is bounded before constructing a continuation object.
        if (++this.mappedCells > this.limits.maxNodes)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "DOCX table expansion exceeds its cell budget.",
          );
        const above = active.get(column);
        let cell: Cell;
        if (above) {
          remaining--;
          cell = { ...above, rows: above.rows - 1, continuation: true };
        } else {
          const source = cells[current++];
          if (!source) unsupported("A table row leaves an unsupported hole in its native grid.");
          cell = {
            node: source!,
            column,
            span: Number(source!.attrs.colspan ?? 1),
            rows: Number(source!.attrs.rowspan ?? 1),
            continuation: false,
          };
        }
        if (
          cell.span < 1 ||
          !Number.isSafeInteger(cell.span) ||
          !Number.isSafeInteger(cell.rows) ||
          cell.rows < 1 ||
          column + cell.span > this.limits.maxNodes
        )
          unsupported("A table span exceeds the native grid profile.");
        this.tableWork += cell.span;
        if (this.tableWork > this.limits.maxNodes * 128)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "DOCX table mapping exceeds its work budget.",
          );
        for (let offset = 1; offset < cell.span; offset++)
          if (active.has(column + offset)) unsupported("A table cell overlaps a vertical merge.");
        rowCells.push(cell);
        if (cell.rows > 1) next.set(column, cell);
        column += cell.span;
      }
      active = next;
      columns = Math.max(columns, column);
      matrix.push(rowCells);
    }
    if (
      active.size ||
      !columns ||
      matrix.some((row) => row.reduce((total, cell) => total + cell.span, 0) !== columns)
    )
      unsupported("A table has an incomplete or nonrectangular grid.");
    const tableWidth =
      css.width && css.width !== "auto" ? points(css.width, css, this.width) : this.width;
    const grid = info?.grid ?? Array.from({ length: columns }, () => tableWidth / columns);
    if (grid.length !== columns)
      unsupported("The current table structure no longer matches its native column grid.");
    this.add(
      "<w:tbl" +
        this.nativeMark(node.attrs["data-node-id"]) +
        "><w:tblPr>" +
        orderedProperties(
          "tblPr",
          value("tblStyle", this.styles.tableStyleId(style)) +
            (info?.rowBandSize !== undefined
              ? value("tblStyleRowBandSize", info.rowBandSize)
              : "") +
            (info?.columnBandSize !== undefined
              ? value("tblStyleColBandSize", info.columnBandSize)
              : "") +
            tableProperties(this.tablePropertyDifference(node, css, "table")) +
            (info?.lookMask !== undefined
              ? empty("w:tblLook", { "w:val": info.lookMask.toString(16).padStart(4, "0") })
              : "") +
            this.review.table(node),
        ) +
        "</w:tblPr><w:tblGrid>",
    );
    for (const width of grid) this.add(empty("w:gridCol", { "w:w": nativeInteger(width, 20, 1) }));
    this.add(this.review.grid(node) + "</w:tblGrid>");
    matrix.forEach((cells, index) => {
      const row = rows[index]!,
        rowCss = { ...this.tablePropertyDifference(row, this.css.get(row), "row") };
      const directHeader = row.attrs["data-word-repeat-header"];
      if (directHeader !== undefined)
        rowCss["--word-row-repeat-header"] = directHeader === "true" ? "1" : "0";
      else if (row.parent?.tag === "thead") rowCss["--word-row-repeat-header"] = "1";
      const before = cells[0]?.node.attrs["data-word-grid-gap"] === "before" ? cells[0]!.span : 0,
        after = cells.at(-1)?.node.attrs["data-word-grid-gap"] === "after" ? cells.at(-1)!.span : 0;
      for (const cell of cells)
        if (
          cell.node.attrs["data-word-grid-gap"] &&
          (cell.node.attrs["data-word-grid-gap"] === "before"
            ? cell.column !== 0
            : cell.column + cell.span !== grid.length)
        )
          unsupported("A native row omission must occupy its leading or trailing grid edge.");
      const exception = rowExceptionProperties(rowCss);
      this.add(
        "<w:tr" +
          this.nativeMark(row.attrs["data-node-id"]) +
          ">" +
          (exception ? "<w:tblPrEx>" + exception + "</w:tblPrEx>" : "") +
          "<w:trPr>" +
          this.conditionalFormat(row) +
          (before ? value("gridBefore", before) : "") +
          (after ? value("gridAfter", after) : "") +
          writeWordRowGapWidths(row.attrs) +
          rowProperties(rowCss) +
          this.review.rowMark(row) +
          this.review.row(row) +
          "</w:trPr>",
      );
      for (const cell of cells) {
        if (cell.node.attrs["data-word-grid-gap"]) continue;
        const continuation = cell.continuation
          ? info?.continuationFormats?.find(
              (item) => item.row === index && item.column === cell.column,
            )
          : undefined;
        const cellCss = this.css.get(cell.node),
          continuationClass = continuation?.cellCssClass,
          computed = continuationClass
            ? this.css.classStyle(continuationClass, this.css.get(row))
            : undefined,
          ownCellCss = continuationClass
            ? Object.fromEntries(
                this.css.classKeys(continuationClass).map((key) => [key, computed![key]!]),
              )
            : this.tablePropertyDifference(cell.node, cellCss, "cell");
        const direction = continuationClass
          ? continuation?.textDirection
          : (cell.node.attrs["data-word-text-direction"] ?? cellCss["--word-cell-text-direction"]);
        this.add(
          "<w:tc" +
            this.nativeMark(cell.node.attrs["data-node-id"]) +
            "><w:tcPr>" +
            orderedProperties(
              "tcPr",
              (cell.continuation
                ? writeWordConditionalFormat(continuation?.format)
                : this.conditionalFormat(cell.node)) +
                (cell.span > 1 ? value("gridSpan", cell.span) : "") +
                (cell.continuation || cell.rows > 1
                  ? value("vMerge", cell.continuation ? "continue" : "restart")
                  : "") +
                cellProperties(ownCellCss) +
                value("textDirection", direction) +
                (cell.continuation
                  ? ""
                  : this.review.cellMark(cell.node) + this.review.cell(cell.node)),
            ) +
            "</w:tcPr>",
        );
        if (cell.continuation)
          this.add(
            continuation?.emptyParagraphs
              ?.map((paragraph) =>
                writeWordEmptyContinuation(
                  paragraph,
                  this.css,
                  cell.node,
                  this.styles.ids,
                  this.styles.numberingIds,
                ),
              )
              .join("") ??
              continuation?.paragraphs
                ?.map((format) =>
                  format
                    ? "<w:p><w:pPr>" + writeWordConditionalFormat(format) + "</w:pPr></w:p>"
                    : "<w:p/>",
                )
                .join("") ??
              "<w:p/>",
          );
        else {
          const width = this.width;
          this.width = grid
            .slice(cell.column, cell.column + cell.span)
            .reduce((total, value) => total + value, 0);
          const cellWidth = this.width;
          for (const property of [
            "padding-left",
            "padding-right",
            "border-left-width",
            "border-right-width",
          ])
            if (cellCss[property]) this.width -= points(cellCss[property]!, cellCss, cellWidth);
          this.width = Math.max(1, this.width);
          this.blocks(cell.node.children);
          this.width = width;
        }
        this.add("</w:tc>");
      }
      this.add("</w:tr>");
    });
    this.add("</w:tbl>");
  }
  private list(node: SourceElement, section?: string, paragraphId?: string, depth = 0): void {
    if (depth > 8) unsupported("Native DOCX supports up to nine numbering levels.");
    const items = children(node, "li"),
      ordered = node.tag === "ol";
    if (node.attrs.reversed !== undefined)
      unsupported("Reversed HTML numbering is outside the native DOCX profile.");
    const formatMap: Record<string, string> = {
      decimal: "decimal",
      "lower-alpha": "lowerLetter",
      "upper-alpha": "upperLetter",
      "lower-roman": "lowerRoman",
      "upper-roman": "upperRoman",
      disc: "bullet",
      circle: "bullet",
      square: "bullet",
      none: "none",
    };
    const raw = this.css.get(node)["list-style-type"] ?? (ordered ? "decimal" : "disc"),
      format =
        formatMap[raw] ??
        unsupported("This list marker has no supported native DOCX representation.");
    const create = (start: number): string => {
      if (!Number.isSafeInteger(start) || start < 0 || start > 2147483647)
        unsupported("A list counter is outside the native numbering profile.");
      const id = String(this.styles.numberingIds.size + this.htmlLists.length + 1);
      this.htmlLists.push({
        id,
        start,
        level: depth,
        format,
        text:
          format === "bullet"
            ? raw === "square"
              ? "▪"
              : raw === "circle"
                ? "○"
                : "•"
            : format === "none"
              ? ""
              : "%" + (depth + 1) + ".",
      });
      return id;
    };
    let id = create(Number(node.attrs.start ?? 1));
    items.forEach((item, index) => {
      if (item.attrs.value !== undefined) id = create(Number(item.attrs.value));
      let first = true;
      const content = this.flow(item.children, item);
      let pending: SourceNode[] = [];
      const emit = (nodes: SourceNode[], last: boolean): void => {
        if (!nodes.length) return;
        const paragraph =
          nodes.length === 1 && element(nodes[0]!) && paragraphTags.has(nodes[0]!.tag)
            ? (nodes[0] as SourceElement)
            : { ...item, tag: "p", children: nodes };
        if (!this.flowStyles.has(paragraph) && paragraph.tag === "p" && paragraph !== nodes[0])
          this.flowStyles.set(paragraph, this.css.get(item));
        this.paragraph(
          paragraph,
          last && index === items.length - 1 ? section : undefined,
          last && index === items.length - 1 ? paragraphId : undefined,
          first ? { id, level: depth } : undefined,
        );
        first = false;
      };
      content.forEach((child, position) => {
        const block =
          element(child) &&
          (paragraphTags.has(child.tag) || ["ol", "ul", "table"].includes(child.tag));
        if (block) {
          emit(pending, false);
          pending = [];
          if (element(child) && ["ol", "ul"].includes(child.tag))
            this.list(
              child,
              position === content.length - 1 && index === items.length - 1 ? section : undefined,
              position === content.length - 1 && index === items.length - 1
                ? paragraphId
                : undefined,
              depth + 1,
            );
          else if (element(child) && child.tag === "table") this.table(child);
          else emit([child], position === content.length - 1);
        } else pending.push(child);
      });
      emit(pending, true);
      if (first && !content.length)
        this.paragraph(
          { ...item, tag: "p" },
          index === items.length - 1 ? section : undefined,
          index === items.length - 1 ? paragraphId : undefined,
          { id, level: depth },
        );
    });
  }
  private readonly htmlLists: HtmlNumbering[] = [];
  private withStory(part: string, body: () => void): void {
    const previous = {
      part: this.part,
      output: this.output,
      fields: this.fieldStack,
      deleted: this.deleted,
      width: this.width,
      inlineRevisions: this.inlineRevisions,
    };
    this.part = part;
    this.fieldStack = [];
    this.deleted = false;
    this.inlineRevisions = [];
    try {
      body();
      if (this.fieldStack.length) unsupported("A story has an unclosed field.");
    } finally {
      this.part = previous.part;
      this.output = previous.output;
      this.fieldStack = previous.fields;
      this.deleted = previous.deleted;
      this.width = previous.width;
      this.inlineRevisions = previous.inlineRevisions;
    }
  }
  write(): {
    files: ReadonlyMap<string, Uint8Array>;
    report: DocxExportReport;
    nativeComments: readonly NativeCommentBinding[];
  } {
    for (const kind of ["footnote", "endnote"] as const) {
      let next = 1;
      for (const note of this.snapshot.metadata.word?.notes ?? [])
        if (note.kind === kind)
          this.noteIds.set(
            note.noteId,
            (note.type ?? "normal") === "normal"
              ? next++
              : note.type === "separator"
                ? -1
                : note.type === "continuationSeparator"
                  ? 0
                  : next++,
          );
    }
    const sections = this.snapshot.metadata.sections
      .map((section) => ({
        section,
        root: this.tree.byId.get(section.rootNodeId)!,
        nodes: [] as SourceNode[],
      }))
      .sort(
        (left, right) =>
          left.root.from - right.root.from ||
          this.tree.elements.indexOf(left.root) - this.tree.elements.indexOf(right.root),
      );
    if (
      sections.some((left) =>
        sections.some((right) => left !== right && contains(left.root, right.root)),
      )
    )
      unsupported("Nested native section roots cannot be exported.");
    let current = sections[0]!;
    const collect = (node: SourceNode): void => {
      if (!element(node)) {
        current.nodes.push(node);
        return;
      }
      if (node.attrs["data-office-story"] && node.attrs["data-office-story"] !== "main") return;
      const section = sections.find((section) => section.root === node);
      if (section) {
        current = section;
        current.nodes.push(node);
        return;
      }
      if (
        !["article", "section", "div"].includes(node.tag) ||
        !sections.some((item) => contains(node, item.root))
      ) {
        current.nodes.push(node);
        return;
      }
      for (const child of node.children) collect(child);
    };
    this.mainStories.forEach(collect);
    this.add(
      declaration +
        "<w:document" +
        namespaces +
        this.nativeMark(this.tree.root.attrs["data-node-id"]) +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>',
    );
    sections.forEach(({ section, nodes }, index) => {
      this.width = this.sectionTextWidth(section);
      this.blocks(
        nodes,
        index === sections.length - 1 ? undefined : this.sectionProperties(section),
      );
    });
    this.add(this.sectionProperties(sections.at(-1)!.section) + "</w:body></w:document>");
    this.output.finish();
    if (this.fieldStack.length) unsupported("The main story has an unclosed field.");
    for (const story of this.snapshot.metadata.word?.stories ?? []) {
      const info = this.auxiliary.get(story.storyId);
      if (!info) continue;
      this.withStory(info.path, () => {
        const section =
          this.snapshot.metadata.sections.find((section) =>
            [
              ...Object.values(section.headers ?? {}),
              ...Object.values(section.footers ?? {}),
            ].includes(story.storyId),
          ) ?? this.snapshot.metadata.sections[0]!;
        this.width = this.sectionTextWidth(section, false);
        this.output = this.opc.part(info.path, contentType(info.kind));
        const tag = info.kind === "header" ? "hdr" : "ftr";
        this.add(declaration + "<w:" + tag + namespaces + this.nativeMark(story.rootNodeId) + ">");
        this.blocks(this.tree.byId.get(story.rootNodeId)!.children);
        this.add("</w:" + tag + ">");
        this.output.finish();
      });
    }
    this.writeNotes();
    this.writeComments();
    this.styles.write(this.opc);
    this.styles.writeNumbering(this.opc, this.htmlLists);
    this.writeSettings();
    const core = this.opc.part(
      "docProps/core.xml",
      "application/vnd.openxmlformats-package.core-properties+xml",
    );
    core.add(
      declaration +
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>' +
        escapeXml(this.snapshot.metadata.title ?? "") +
        "</dc:title></cp:coreProperties>",
    );
    core.finish();
    this.opc.relate("", "officeDocument", "word/document.xml");
    this.opc.relate(
      "",
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
      "docProps/core.xml",
    );
    this.opc.finish();
    const budget = { nodes: 0 };
    for (const [path, bytes] of this.opc.parts.files)
      if (path.endsWith(".xml") || path.endsWith(".rels")) xml(bytes, path, this.limits, budget);
    return {
      nativeComments: this.annotations.nativeBindings,
      files: this.opc.parts.files,
      report: Object.freeze({
        documentId: this.snapshot.manifest.documentId,
        revision: this.snapshot.manifest.revision,
        exportedAnnotationIds: Object.freeze(
          this.annotations.comments.map((comment) => comment.annotation.annotationId),
        ),
        unexportedAnnotations: Object.freeze(
          this.annotations.omitted.map((item) => Object.freeze({ ...item })),
        ),
      }),
    };
  }
  private writeNotes(): void {
    for (const kind of ["footnote", "endnote"] as const) {
      const notes = this.snapshot.metadata.word?.notes?.filter((note) => note.kind === kind) ?? [];
      if (!notes.length) continue;
      const path = "word/" + kind + "s.xml";
      this.withStory(path, () => {
        this.output = this.opc.part(
          path,
          "application/vnd.openxmlformats-officedocument.wordprocessingml." + kind + "s+xml",
        );
        this.add(declaration + "<w:" + kind + "s" + namespaces + ">");
        for (const [type, id, tag] of [
          ["separator", -1, "separator"],
          ["continuationSeparator", 0, "continuationSeparator"],
        ] as const)
          if (!notes.some((note) => note.type === type))
            this.add(
              "<w:" +
                kind +
                attributes({ "w:id": id, "w:type": type }) +
                "><w:p><w:r>" +
                empty("w:" + tag) +
                "</w:r></w:p></w:" +
                kind +
                ">",
            );
        for (const note of notes) {
          const story = this.snapshot.metadata.word!.stories.find(
            (story) => story.storyId === note.storyId,
          )!;
          this.add(
            "<w:" +
              kind +
              this.nativeMark(story.rootNodeId) +
              attributes({
                "w:id": this.noteIds.get(note.noteId),
                "w:type": note.type && note.type !== "normal" ? note.type : undefined,
              }) +
              ">",
          );
          this.blocks(this.tree.byId.get(story.rootNodeId)!.children);
          this.add("</w:" + kind + ">");
        }
        this.add("</w:" + kind + "s>");
        this.output.finish();
      });
      this.opc.relate("word/document.xml", kind + "s", kind + "s.xml");
    }
  }
  private writeComments(): void {
    if (
      !this.annotations.comments.length &&
      !this.snapshot.metadata.word?.nativeContent?.fragments.some(
        (fragment) =>
          !fragment.nodeId &&
          ["commentsEx", "commentsIds", "commentsExtensible"].includes(fragment.ownerName),
      )
    ) {
      this.writePeople([]);
      return;
    }
    const messages = this.annotations.messages,
      path = "word/comments.xml";
    this.withStory(path, () => {
      this.output = this.opc.part(path, contentType("comments"));
      this.add(
        declaration +
          "<w:comments" +
          namespaces +
          ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">',
      );
      for (const message of messages) {
        const body = message.reply ?? message.annotation;
        this.add(
          "<w:comment" +
            this.nativeMark(body.bodyNodeId) +
            attributes({
              "w:id": message.id,
              "w:author": body.author?.name ?? "",
              "w:initials": body.author?.initials,
              "w:date": date(body.date),
            }) +
            ">",
        );
        if (body.bodyNodeId)
          this.blocks(
            this.tree.byId.get(body.bodyNodeId)!.children,
            undefined,
            message.paragraphId,
          );
        else
          this.add(
            '<w:p w14:paraId="' +
              message.paragraphId +
              '"><w:r><w:t xml:space="preserve">' +
              escapeXml(body.instruction!) +
              "</w:t></w:r></w:p>",
          );
        this.add("</w:comment>");
      }
      this.add("</w:comments>");
      this.output.finish();
    });
    this.opc.relate("word/document.xml", "comments", "comments.xml");
    const extension = (
      file: string,
      type: string,
      ns: string,
      prefix: string,
      root: string,
      items: string,
    ): void => {
      const part = this.opc.part("word/" + file, type);
      part.add(
        declaration +
          "<" +
          prefix +
          ":" +
          root +
          " xmlns:" +
          prefix +
          '="' +
          ns +
          '">' +
          items +
          "</" +
          prefix +
          ":" +
          root +
          ">",
      );
      part.finish();
    };
    extension(
      "commentsExtended.xml",
      "application/vnd.ms-word.commentsExtended+xml",
      W15,
      "w15",
      "commentsEx",
      messages
        .map((message) => {
          const body = message.reply ?? message.annotation;
          return (
            "<w15:commentEx" +
            this.nativeMark(body.bodyNodeId) +
            attributes({
              "w15:paraId": message.paragraphId,
              "w15:paraIdParent": message.parentParagraphId,
              "w15:done": ["resolved", "closed"].includes(body.status ?? "open") ? 1 : 0,
            }) +
            "/>"
          );
        })
        .join(""),
    );
    this.opc.relate(
      "word/document.xml",
      "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
      "commentsExtended.xml",
    );
    extension(
      "commentsIds.xml",
      "application/vnd.ms-word.commentsIds+xml",
      CID,
      "w16cid",
      "commentsIds",
      messages
        .map(
          (message) =>
            "<w16cid:commentId" +
            this.nativeMark((message.reply ?? message.annotation).bodyNodeId) +
            attributes({
              "w16cid:paraId": message.paragraphId,
              "w16cid:durableId": message.durableId,
            }) +
            "/>",
        )
        .join(""),
    );
    this.opc.relate(
      "word/document.xml",
      "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
      "commentsIds.xml",
    );
    extension(
      "commentsExtensible.xml",
      "application/vnd.ms-word.commentsExtensible+xml",
      CEX,
      "w16cex",
      "commentsExtensible",
      messages
        .map(
          (message) =>
            "<w16cex:commentExtensible" +
            this.nativeMark((message.reply ?? message.annotation).bodyNodeId) +
            attributes({
              "w16cex:durableId": message.durableId,
              "w16cex:dateUtc": (message.reply ?? message.annotation).date
                ? new Date(date((message.reply ?? message.annotation).date)!).toISOString()
                : undefined,
              "w16cex:intelligentPlaceholder":
                message.followUp === undefined ? undefined : message.followUp ? 1 : 0,
            }) +
            "/>",
        )
        .join(""),
    );
    this.opc.relate(
      "word/document.xml",
      "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
      "commentsExtensible.xml",
    );
    this.writePeople(messages);
  }
  private writePeople(messages: readonly ExportComment[]): void {
    const people = new Map(
      this.snapshot.metadata.word?.people?.map((person) => [person.author, person]) ?? [],
    );
    for (const message of messages) {
      const name = (message.reply ?? message.annotation).author?.name ?? "";
      if (!people.has(name)) people.set(name, { author: name });
    }
    const items = [...people.values()]
      .map((person) => {
        const body = messages.find(
            (message) => (message.reply ?? message.annotation).author?.name === person.author,
          ),
          nodeId = body
            ? (body.reply ?? body.annotation).bodyNodeId
            : this.tree.root.attrs["data-node-id"];
        return (
          "<w15:person" +
          this.nativeMark(nodeId) +
          attributes({ "w15:author": person.author }) +
          ">" +
          (person.presence
            ? empty("w15:presenceInfo", {
                "w15:providerId": person.presence.providerId,
                "w15:userId": person.presence.userId,
              })
            : "") +
          "</w15:person>"
        );
      })
      .join("");
    if (!items) return;
    const out = this.opc.part("word/people.xml", "application/vnd.ms-word.people+xml");
    out.add(declaration + '<w15:people xmlns:w15="' + W15 + '">' + items + "</w15:people>");
    out.finish();
    this.opc.relate(
      "word/document.xml",
      "http://schemas.microsoft.com/office/2011/relationships/people",
      "people.xml",
    );
  }

  private noteSettings(kind: "footnote" | "endnote"): string {
    const word = this.snapshot.metadata.word,
      notes = word?.notes?.filter((note) => note.kind === kind) ?? [],
      options = kind === "footnote" ? word?.footnoteOptions : word?.endnoteOptions;
    return writeNoteOptions(
      kind,
      notes.length
        ? { format: kind === "footnote" ? "decimal" : "lowerRoman", ...options }
        : options,
      notes.length
        ? [
            -1,
            0,
            ...notes
              .filter((note) => note.type === "continuationNotice")
              .map((note) => this.noteIds.get(note.noteId)!),
          ]
        : [],
    );
  }
  private writeSettings(): void {
    const out = this.opc.part("word/settings.xml", contentType("settings")),
      tabs = this.snapshot.metadata.word?.tabs ?? [],
      intervals = new Set(tabs.map((tab) => tab.defaultInterval));
    if (intervals.size > 1) unsupported("DOCX has one document-wide default tab interval.");
    const sections = this.snapshot.metadata.sections;
    if (
      sections.some((section) => section.gutterPosition === "top" && section.gutter) &&
      sections.some((section) => section.gutterPosition !== "top" && section.gutter)
    )
      unsupported(
        "DOCX gutterAtTop is document-wide; nonzero top and side gutters cannot be mixed.",
      );
    out.add(
      declaration +
        "<w:settings" +
        namespaces +
        ">" +
        orderedProperties(
          "settings",
          value(
            "bordersDoNotSurroundHeader",
            this.snapshot.metadata.word?.bordersDoNotSurroundHeader ? 1 : 0,
          ) +
            value(
              "bordersDoNotSurroundFooter",
              this.snapshot.metadata.word?.bordersDoNotSurroundFooter ? 1 : 0,
            ) +
            value("mirrorMargins", this.snapshot.metadata.word?.mirrorMargins ? 1 : 0) +
            value("evenAndOddHeaders", this.snapshot.metadata.word?.evenAndOddHeaders ? 1 : 0) +
            value("defaultTabStop", nativeInteger(tabs[0]?.defaultInterval ?? 36, 20, 1)) +
            value("autoHyphenation", this.snapshot.metadata.word?.automaticHyphenation ? 1 : 0) +
            value("trackRevisions", this.snapshot.metadata.word?.trackChanges ? 1 : 0) +
            value(
              "gutterAtTop",
              sections.some((section) => section.gutterPosition === "top" && section.gutter)
                ? 1
                : 0,
            ) +
            value("updateFields", 0) +
            this.noteSettings("footnote") +
            this.noteSettings("endnote") +
            (wordCompatibilityNames.some(
              (name) => this.snapshot.metadata.word?.[name] !== undefined,
            )
              ? "<w:compat>" +
                wordCompatibilityNames
                  .filter((name) => this.snapshot.metadata.word?.[name] !== undefined)
                  .map((name) => value(name, this.snapshot.metadata.word![name] ? 1 : 0))
                  .join("") +
                "</w:compat>"
              : ""),
        ) +
        "</w:settings>",
    );
    out.finish();
    this.opc.relate("word/document.xml", "settings", "settings.xml");
  }
}
/** Generate a native, editable DOCX exclusively from the current validated HTML
 * package. No original Office bytes, DOM, filesystem or network is consulted. */
export function exportDocx(
  document: ContentDocumentModel,
  options: DocxExportOptions = {},
): DocxExportResult {
  const model = openDocument(document.serialize(), { ...options, kind: "document" }),
    limits = { ...DEFAULT_DOCUMENT_READ_LIMITS, ...options.limits };
  if (model.snapshot().manifest.kind !== "document")
    throw new DocumentError("UnsupportedFormat", "DOCX export requires a Word document package.");
  const output = new Exporter(model, limits).write();
  return Object.freeze({
    bytes: writeOpcZip(
      restoreDocxNative(model, output.files, limits, output.nativeComments),
      options,
    ),
    report: Object.freeze({
      ...output.report,
      preservedNativeIds: Object.freeze(
        model
          .snapshot()
          .metadata.word?.nativeContent?.fragments.map((fragment) => fragment.nativeId) ?? [],
      ),
    }),
  });
}
