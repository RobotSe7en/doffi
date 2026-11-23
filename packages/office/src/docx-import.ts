import { readEditableWordGrid, wordSectionTableWidth } from "./docx-table-grid.js";
import { wordRowExceptionProperties, wordRowExceptionCellProperties } from "@doffi/core";
import { commentControl, wordScope } from "./docx-native-comments.js";
import { readWordEmptyContinuation } from "./docx-table-continuation.js";
import { bindWordGridInsertion } from "./docx-grid-revision.js";
import { readWordRowGapCount, readWordRowGapWidths } from "./docx-row-gaps.js";
import { measurePoints } from "./ooxml/measure.js";
import { wordCompatibilityKey, wordCompatibilityNames } from "./ooxml/word-compatibility.js";
import { readWordTypography } from "./ooxml/word-typography.js";
import { bindWordConditionalFormat, readWordConditionalFormat } from "./docx-conditional-format.js";
import {
  resolveWordTableStyle,
  wordTableRegions,
  wordTableRowRegions,
  wordTableRowCellProperties,
} from "@doffi/core";
import { WordMoveReader } from "./docx-moves.js";
import { readWordPageBorders } from "./docx-page-borders.js";
import { readWordFormatRevision } from "./docx-revisions-import.js";
import { readWordCommentExtensions, wordCommentAnnotations } from "./docx-comments.js";
import type { WordCommentState as CommentState } from "./docx-comments.js";
import { readWordControl } from "./docx-control.js";
import { readNoteOptions } from "./docx-note-options.js";
import { M, readWordMath } from "./docx-math.js";
import { importWordDrawing, WPS, WPG } from "./docx-drawing-import.js";
import { createAssetDescriptor, wordScriptRole, DocumentError, openDocument } from "@doffi/core";
import type {
  ContentDocumentManifest,
  ContentDocumentMetadata,
  DocumentPackage,
  DocumentSection,
  PageSetup,
  WordMetadata,
  WordTableEmptyParagraph,
} from "@doffi/core";
import type { DocumentZipOptions } from "./document-zip.js";
import { openOpc } from "./ooxml/package.js";
import type { OpcPackage } from "./ooxml/package.js";
import {
  DocxStyles,
  declarations,
  escapeHtml,
  inheritedRun,
  mergeTabs,
  tableLookMask,
} from "./ooxml/styles.js";
import type { Css } from "./ooxml/styles.js";
import {
  A,
  MC,
  PIC,
  R,
  W,
  WP,
  attr,
  check,
  child,
  elements,
  fail,
  number,
  on,
  text,
  unsupported,
  unsupportedType,
} from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";

import { DocxNativeImport } from "./docx-native-import.js";

export interface DocxImportOptions extends DocumentZipOptions {
  readonly unsupportedContent?: "preserve" | "reject";
  readonly documentId: string;
}
interface Html {
  tag: string;
  attrs: Record<string, string>;
  children: (Html | string)[];
}
interface FieldState {
  definition: WordMetadata["fields"][number];
  result: boolean;
}

/** Serialized HTML size without allocating an escaped string or UTF-8 buffer. */
function htmlBytes(value: string): number {
  let bytes = 0;
  for (const character of value) {
    const code = character.codePointAt(0)!;
    bytes +=
      character === "&"
        ? 5
        : character === '"'
          ? 6
          : character === "<" || character === ">"
            ? 4
            : code <= 0x7f
              ? 1
              : code <= 0x7ff
                ? 2
                : code <= 0xffff
                  ? 3
                  : 4;
  }
  return bytes;
}
const serialize = (node: Html): string => {
  const attributes = Object.entries(node.attrs)
    .map(([name, value]) => " " + name + '="' + escapeHtml(value) + '"')
    .join("");
  return (
    "<" +
    node.tag +
    attributes +
    ">" +
    (["img", "br"].includes(node.tag)
      ? ""
      : node.children
          .map((item) => (typeof item === "string" ? escapeHtml(item) : serialize(item)))
          .join("") +
        "</" +
        node.tag +
        ">")
  );
};
const emptyWord = (): WordMetadata => ({
  stories: [],
  styles: [],
  numbering: [],
  hyperlinks: [],
  bookmarks: [],
  fields: [],
  tabs: [],
  revisions: [],
  tables: [],
});

class Importer {
  private serial = 0;
  private generatedNodes = 0;
  private generatedBytes = 0;
  private readonly textLengths = new WeakMap<Html, number>();
  private currentRun: Html | undefined;
  private paragraphRunClasses: string[] = ["word-default-run"];
  private readonly files = new Map<string, Uint8Array>();
  private readonly word: WordMetadata = emptyWord();
  private readonly sections: DocumentSection[] = [];
  private readonly assetSources = new Map<string, string>();
  private readonly assets: ContentDocumentManifest["assets"] = [];
  private readonly styles: DocxStyles;
  private readonly root: Html;
  private readonly auxiliary: Html[] = [];
  private readonly comments = new Map<string, CommentState>();
  private activeComments = new Set<string>();
  private moves = new WordMoveReader(this.word);
  private commentRangeCount = 0;
  private commentRangeBytes = 0;
  private fieldResultCount = 0;
  private fieldResultBytes = 0;
  private fieldStack: FieldState[] = [];
  private bookmarkStarts = new Map<string, { name: string; nodeId: string }>();
  private readonly loadedStories = new Map<string, string>();
  private readonly settings: XmlNode | undefined;
  private readonly defaultTab: number;
  private currentPart: string;
  private currentStory = "";
  private pendingNoteCustomMark: string | undefined;
  private readonly nativeNotes = new Map<string, string>();
  constructor(
    private readonly opc: OpcPackage,
    private readonly main: string,
    private readonly documentId: string,
    private readonly native?: DocxNativeImport,
  ) {
    this.currentPart = main;
    this.styles = new DocxStyles(opc, main, !!native);
    this.word.evenAndOddHeaders = on(
      child(
        opc.related(main, "settings") ? opc.read(opc.related(main, "settings")!) : undefined,
        "evenAndOddHeaders",
      ),
    );
    this.word.defaultRunCssClass = "word-default-run";
    this.word.styles = this.styles.styles;
    this.word.numbering = this.styles.numbering;
    this.root = this.make("article", {
      "data-office-kind": "document",
      style: declarations(this.styles.drawingThemeProperties()),
    });
    const settingsPath = opc.related(main, "settings");
    this.settings = settingsPath ? opc.read(settingsPath) : undefined;
    if (this.settings && (this.settings.uri !== W || this.settings.name !== "settings"))
      fail(this.settings, "The settings part has an invalid root.");
    const footnoteOptions = readNoteOptions(child(this.settings, "footnotePr")),
      endnoteOptions = readNoteOptions(child(this.settings, "endnotePr"));
    if (footnoteOptions) this.word.footnoteOptions = footnoteOptions;
    if (endnoteOptions) this.word.endnoteOptions = endnoteOptions;
    this.defaultTab =
      number(
        attr(child(this.settings, "defaultTabStop"), "val"),
        this.settings ?? opc.read(main),
        720,
      ) / 20;
    this.word.trackChanges = on(child(this.settings, "trackRevisions"));
    this.word.mirrorMargins = on(child(this.settings, "mirrorMargins"));
    this.word.bordersDoNotSurroundHeader = on(child(this.settings, "bordersDoNotSurroundHeader"));
    this.word.bordersDoNotSurroundFooter = on(child(this.settings, "bordersDoNotSurroundFooter"));
    const compat = child(this.settings, "compat");
    const mode = elements(compat).find((option) => wordCompatibilityKey(option) === "compatibilityMode");
    const modeValue = attr(mode, "val")?.trim();
    if (modeValue && /^\d+$/.test(modeValue)) {
      const parsed = Number(modeValue);
      if (parsed >= 1 && parsed <= 2147483647) this.word.compatibilityMode = parsed;
    }
    for (const name of wordCompatibilityNames) {
      const flag = child(compat, name);
      if (flag) this.word[name] = on(flag);
    }
    this.word.automaticHyphenation = on(child(this.settings, "autoHyphenation"));
    const typography = readWordTypography(this.settings);
    if (typography) this.word.typography = typography;
    for (const name of ["documentProtection", "writeProtection", "mailMerge", "attachedTemplate"]) {
      if (child(this.settings, name))
        unsupported(
          child(this.settings, name)!,
          "Protected or externally dependent document settings are not supported.",
        );
    }
  }
  private tableAvailable = 468;
  private activeTable: WordMetadata["tables"][number] | undefined;
  private context(
    node: Html,
    kind: "table" | "row" | "cell" | "paragraph" | "run",
    css: Css,
    info = this.activeTable,
  ): void {
    if (!info) return;
    const cssClass = "wt-" + this.id(node);
    info.styleContexts ??= [];
    info.styleContexts.push({ nodeId: this.id(node), kind, cssClass });
    node.attrs.class = [node.attrs.class, cssClass].filter(Boolean).join(" ");
    this.styles.addIndependentRule(("." + cssClass).repeat(5), css);
  }
  private make(tag: string, attrs: Record<string, string> = {}): Html {
    const id = "n" + (this.serial + 1);
    let bytes =
      (["img", "br"].includes(tag) ? tag.length + 2 : tag.length * 2 + 5) +
      ' data-node-id=""'.length +
      id.length;
    for (const [name, value] of Object.entries(attrs)) bytes += name.length + 4 + htmlBytes(value);
    this.reserveHtml(1, bytes);
    this.serial++;
    return {
      tag,
      attrs: { "data-node-id": id, ...attrs },
      children: [],
    };
  }
  private reserveHtml(nodes: number, bytes: number): void {
    if (
      this.generatedNodes + nodes > this.opc.limits.maxNodes ||
      this.generatedBytes + bytes >
        Math.min(this.opc.limits.maxFileBytes, this.opc.limits.maxPackageBytes)
    )
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Generated HTML exceeds the configured node or byte budget.",
      );
    this.generatedNodes += nodes;
    this.generatedBytes += bytes;
  }
  private id(node: Html): string {
    return node.attrs["data-node-id"]!;
  }
  private offset(node: Html, value: string): number {
    const from = this.textLengths.get(node) ?? 0;
    this.textLengths.set(node, from + value.length);
    if (this.currentRun && this.currentRun !== node)
      this.textLengths.set(
        this.currentRun,
        (this.textLengths.get(this.currentRun) ?? 0) + value.length,
      );
    return from;
  }
  private addText(node: Html, value: string): void {
    if (!value) return;
    const from = this.offset(node, value);
    let content = node;
    const field = this.fieldStack.filter((item) => item.result).pop();
    if (field) {
      content = this.make("span", {
        "data-word-field-result": field.definition.nodeId,
      });
      node.children.push(content);
      this.fieldResult(content);
    }
    if (node.tag === "span" && !node.attrs["data-word-tab"]) {
      // Flush one script part at a time; reserve text and wrapper budgets before
      // creating substrings or HTML nodes. Never build an unbounded parts array.
      let script: "ascii" | "hAnsi" | "eastAsia" | "cs" | undefined;
      let start = 0,
        end = 0,
        bytes = 0;
      const append = (): void => {
        if (!script) return;
        this.reserveHtml(1, bytes);
        let target = content;
        if (script !== "ascii") {
          target = this.make("span", {
            style:
              "font-family:var(--word-font-" +
              script +
              ",var(--word-font-ascii,var(--word-font-hAnsi,inherit)))",
            "data-word-script": script,
          });
          content.children.push(target);
        }
        target.children.push(value.slice(start, end));
      };
      for (const character of value) {
        const nextScript = wordScriptRole(character, script);
        if (script !== nextScript) {
          append();
          start = end;
          bytes = 0;
          script = nextScript;
        }
        end += character.length;
        bytes += htmlBytes(character);
      }
      append();
    } else {
      this.reserveHtml(1, htmlBytes(value));
      content.children.push(value);
    }
    this.record(node, from, value);
  }
  private fieldResult(node: Html): void {
    // Ancestor results also own nested fields' cached content.
    for (const field of this.fieldStack)
      if (field.result) {
        const bytes = 16 + this.id(node).length * 6;
        if (
          this.fieldResultCount + 1 > this.opc.limits.maxNodes ||
          this.fieldResultBytes + bytes >
            Math.min(this.opc.limits.maxFileBytes, this.opc.limits.maxPackageBytes)
        )
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Generated field result references exceed the configured node or byte budget.",
          );
        this.fieldResultCount++;
        this.fieldResultBytes += bytes;
        (field.definition.resultNodeIds ??= []).push(this.id(node));
      }
  }
  private record(node: Html, from: number, value: string): void {
    for (const id of this.activeComments) {
      const comment = this.comments.get(id);
      if (!comment)
        throw new DocumentError("InvalidDocument", "A comment range refers to a missing comment.");
      const previous = comment.ranges[comment.ranges.length - 1];
      const extend = previous?.nodeId === this.id(node) && previous.to === from;
      // Six bytes per UTF-16 unit bounds JSON escaping and UTF-8. The fixed
      // allowance covers keys, offsets and pretty-printing before allocation.
      const bytes = value.length * 6 + (extend ? 0 : 256 + this.id(node).length * 6);
      if (
        this.commentRangeCount + (extend ? 0 : 1) > this.opc.limits.maxNodes ||
        this.commentRangeBytes + bytes >
          Math.min(this.opc.limits.maxFileBytes, this.opc.limits.maxPackageBytes)
      )
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Generated comment ranges exceed the configured node or byte budget.",
        );
      this.commentRangeBytes += bytes;
      if (extend) {
        previous.to += value.length;
        previous.quote += value;
      } else {
        this.commentRangeCount++;
        comment.ranges.push({
          nodeId: this.id(node),
          from,
          to: from + value.length,
          quote: value,
        });
      }
    }
  }
  private story(
    kind: "main" | "header" | "footer" | "comment" | "textBox" | "footnote" | "endnote",
    storyId: string,
  ): Html {
    const root = this.make("div", { "data-office-story": kind });
    this.word.stories.push({ storyId, kind, rootNodeId: this.id(root) });
    return root;
  }
  private compat(node: XmlNode): XmlNode[] {
    const result: XmlNode[] = [];
    for (const item of node.content) {
      if (typeof item === "string") {
        if (item.trim()) fail(node, "Unexpected OOXML text.");
        continue;
      }
      if (item.uri === MC && item.name === "AlternateContent") {
        const supported = new Set([W, R, A, WP, PIC, WPS, WPG]);
        const choices = elements(item, "Choice", MC);
        if (
          this.native &&
          choices.some((choice) =>
            (attr(choice, "Requires", "") ?? "")
              .split(/\s+/)
              .filter(Boolean)
              .some((prefix) => !supported.has(choice.namespaces[prefix] ?? "")),
          )
        )
          unsupported(
            item,
            "The native compatibility choices are preserved together with their cached fallback.",
          );
        const branch =
          elements(item, "Choice", MC).find((choice) =>
            (attr(choice, "Requires", "") ?? "")
              .split(/\s+/)
              .filter(Boolean)
              .every((prefix) => supported.has(choice.namespaces[prefix] ?? "")),
          ) ?? child(item, "Fallback", MC);
        if (!branch) unsupported(item, "No supported compatibility alternative is available.");
        result.push(...this.compat(branch));
      } else result.push(item);
    }
    return result;
  }
  import(): DocumentPackage {
    const source = this.opc.read(this.main);
    this.native?.bind(source, this.root.attrs);
    if (source.uri !== W || source.name !== "document")
      fail(source, "The main part is not a Word document.");
    check(source, ["background", "body"]);
    if (child(source, "background"))
      unsupported(child(source, "background")!, "Document backgrounds are not supported.");
    const fontTable = this.opc.related(this.main, "fontTable");
    if (fontTable) {
      const table = this.opc.read(fontTable);
      for (const font of elements(table, "font"))
        for (const item of elements(font))
          if (item.name.startsWith("embed"))
            unsupported(
              item,
              "Embedded fonts require an authorized font adapter and are not supported by this reader.",
            );
    }
    const mainStory = this.story("main", "main");
    this.currentStory = "main";
    this.readComments();
    this.readNotes();
    const body = child(source, "body");
    if (!body) fail(source, "A Word body is missing.");
    let section = this.make("section");
    mainStory.children.push(section);
    let settings: XmlNode | undefined;
    const addSection = (properties: XmlNode | undefined): void => {
      const item = this.section(properties, section);
      this.sections.push(item);
    };
    const bodyItems = this.compat(body),
      sectionProperties: (XmlNode | undefined)[] = [];
    let upcoming: XmlNode | undefined;
    for (let index = bodyItems.length - 1; index >= 0; index--) {
      const item = bodyItems[index]!;
      const properties =
        item.name === "sectPr"
          ? item
          : item.name === "p"
            ? child(child(item, "pPr"), "sectPr")
            : undefined;
      if (properties) upcoming = properties;
      sectionProperties[index] = upcoming;
    }
    for (const [index, item] of bodyItems.entries()) {
      this.tableAvailable = wordSectionTableWidth(
        sectionProperties[index],
        this.sections.at(-1)?.pageSetup,
      );
      if (item.uri === W && item.name === "sectPr") {
        if (settings) fail(item, "A body has repeated section properties.");
        settings = item;
        continue;
      }
      if (settings) fail(item, "Body content follows final section properties.");
      const blocks = this.blocks([item]);
      section.children.push(...blocks);
      const properties =
        item.uri === W && item.name === "p" ? child(child(item, "pPr"), "sectPr") : undefined;
      if (properties) {
        addSection(properties);
        section = this.make("section");
        mainStory.children.push(section);
      }
    }
    addSection(settings);
    if (
      this.fieldStack.length ||
      this.activeComments.size ||
      this.bookmarkStarts.size ||
      this.moves.pending
    )
      fail(source, "A field, comment range or bookmark is not closed.");
    this.resolveDrawingConnections();
    this.root.children.push(mainStory, ...this.auxiliary);
    const annotations = wordCommentAnnotations(this.comments, this.word, this.documentId);
    const metadata: ContentDocumentMetadata = {
      ...(this.styles.language ? { locale: this.styles.language } : {}),
      pageSetup: this.sections[0]!.pageSetup!,
      sections: this.sections,
      word: this.word,
    };
    const core = this.opc.related("", "metadata/core-properties");
    if (core) {
      const properties = this.opc.read(core);
      const title = elements(properties, "title", "http://purl.org/dc/elements/1.1/")[0];
      if (title) metadata.title = text(title);
      const language = elements(properties, "language", "http://purl.org/dc/elements/1.1/")[0];
      if (language && text(language)) metadata.locale = text(language);
    }
    const manifest: ContentDocumentManifest = {
      format: "doffi",
      formatVersion: 1,
      documentId: this.documentId,
      kind: "document",
      revision: 0,
      entrypoint: "content.html",
      styles: ["styles.css"],
      metadata: "metadata.json",
      annotations: "annotations.json",
      assets: this.assets,
    };
    this.native?.finish(this.files, manifest, this.word);
    const css =
      ".word-main{white-space:normal}.word-main p,.word-main h1,.word-main h2,.word-main h3,.word-main h4,.word-main h5,.word-main h6{white-space:pre-wrap}.word-main table{border-collapse:collapse}.word-main td{vertical-align:top}.word-main ins{text-decoration:underline;color:#006400}.word-main del{text-decoration:line-through;color:#a00000}\n" +
      this.styles.stylesheet();
    mainStory.attrs.class = "word-main";

    this.put("manifest.json", JSON.stringify(manifest, null, 2));
    this.put("metadata.json", JSON.stringify(metadata, null, 2));
    this.put("annotations.json", JSON.stringify(annotations, null, 2));
    this.put("styles.css", css);
    this.put(
      "content.html",
      "<!doctype html><html" +
        (metadata.locale ? ' lang="' + escapeHtml(metadata.locale) + '"' : "") +
        '><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body>' +
        serialize(this.root) +
        "</body></html>",
    );
    return openDocument(this.files, { limits: this.opc.limits }).serialize();
  }
  private put(path: string, value: string): void {
    const codecs = globalThis as unknown as {
      TextEncoder: new () => { encode(value: string): Uint8Array };
    };
    this.files.set(path, new codecs.TextEncoder().encode(value));
  }
  private blocks(
    items: readonly XmlNode[],
    inherited: { p: Css; r: Css; runClasses?: readonly string[] } = {
      p: {},
      r: {},
    },
  ): Html[] {
    const result: Html[] = [];
    for (const item of items) {
      if (item.uri === M && (item.name === "oMath" || item.name === "oMathPara")) {
        const paragraph = this.make("p", { "data-word-math-paragraph": "true" });
        if (item.name === "oMathPara") {
          this.mathParagraph(item, paragraph);
        } else this.math(item, paragraph, "block");
        result.push(paragraph);
        continue;
      }
      if (item.uri !== W) unsupportedType(item);
      if (this.native?.isPlaceholder(item)) result.push(this.nativeObject(item, "p"));
      else if (item.name === "p") result.push(this.paragraph(item, inherited));
      else if (item.name === "tbl") result.push(this.table(item));
      else if (item.name === "sdt") {
        const properties = child(item, "sdtPr"),
          object = child(properties, "docPartObj");
        if (attr(child(object, "docPartGallery"), "val") === "Table of Contents") {
          check(properties, ["docPartObj", "id", "alias", "tag"]);
          check(object, ["docPartGallery", "docPartUnique"]);
          const wrapper = this.make("div", { "data-word-content-control": "toc" });
          this.native?.bind(item, wrapper.attrs);
          for (const name of ["id", "alias", "tag"]) {
            const value = attr(child(properties, name), "val");
            if (value !== undefined) wrapper.attrs["data-word-control-" + name] = value;
          }
          const content = child(item, "sdtContent");
          if (!content) fail(item, "A content control has no content.");
          wrapper.children.push(...this.blocks(this.compat(content), inherited));
          result.push(wrapper);
        } else result.push(this.control(item, "block", inherited));
      } else if (
        item.name === "bookmarkStart" ||
        item.name === "bookmarkEnd" ||
        item.name === "commentRangeStart" ||
        item.name === "commentRangeEnd" ||
        /^move(?:From|To)Range(?:Start|End)$/.test(item.name)
      ) {
        const marker = this.make("div");
        this.inline(item, marker);
        if (marker.children.length) result.push(marker);
      } else unsupportedType(item);
    }
    return result;
  }
  private paragraph(
    source: XmlNode,
    inherited: { p: Css; r: Css; runClasses?: readonly string[] } = {
      p: {},
      r: {},
    },
  ): Html {
    const properties = child(source, "pPr");
    const styleId = attr(child(properties, "pStyle"), "val") ?? this.styles.defaultParagraph;
    const style = this.styles.get(styleId);
    const paragraphRun = this.styles.runStyle(styleId, {
      ...this.styles.defaults.r,
      ...inherited.r,
    });
    const contextualParagraph = Object.fromEntries(
      Object.entries(inherited.p).filter(([name]) => style.p[name] === undefined),
    );
    const css: Css = {
      ...(this.activeTable ? {} : contextualParagraph),
      ...this.styles.paragraph(properties),
    };
    const effectiveParagraph = {
      ...this.styles.defaults.p,
      ...inherited.p,
      ...style.p,
      ...css,
    };
    const outline = child(properties, "outlineLvl");
    const level = outline ? number(attr(outline, "val"), outline) : style.outline;
    const tag = level !== undefined && level < 6 ? "h" + (level + 1) : "p";
    const paragraph = this.make(tag, {
      ...(outline ? { "data-word-outline-level": String(level) } : {}),
      style: declarations(css),
      ...(styleId ? { "data-word-style": styleId, class: this.styles.styleClass(styleId) } : {}),
    });
    bindWordConditionalFormat(this.word, properties, this.id(paragraph));
    if (this.activeTable)
      this.context(paragraph, "paragraph", {
        ...contextualParagraph,
        ...style.p,
        ...inheritedRun(paragraphRun),
      });
    if (
      effectiveParagraph["line-height"] &&
      /^\d+(?:\.\d+)?$/.test(effectiveParagraph["line-height"])
    )
      paragraph.attrs["data-word-line-rule"] = "auto";
    const numbering = this.styles.paragraphNumbering(source, properties, styleId),
      numbered = numbering !== undefined && numbering.id !== "0",
      tabs = child(properties, "tabs"),
      stops = mergeTabs(
        mergeTabs(style.tabs ?? this.styles.defaults.tabs, numbering?.tabs),
        tabs,
      );
    if (numbered || stops.length || tabs || this.hasTab(source))
      this.word.tabs.push({
        paragraphNodeId: this.id(paragraph),
        defaultInterval: this.defaultTab,
        stops,
      });
    if (numbering?.id === "0") paragraph.attrs["data-word-numbering-off"] = "true";
    if (numbered) {
      paragraph.attrs["data-word-numbering"] = numbering.id;
      paragraph.attrs["data-word-level"] = String(numbering.level);
      paragraph.attrs.style = declarations({
        ...numbering.paragraph,
        ...css,
      });
    }
    for (const change of elements(child(properties, "rPr")).filter(
      (item) => ["ins", "del", "moveFrom", "moveTo"].includes(item.name) && item.uri === W,
    ))
      this.moves.mark(change, paragraph);
    const oldRunClasses = this.paragraphRunClasses;
    this.paragraphRunClasses = [
      "word-default-run",
      ...(inherited.runClasses ?? []),
      ...this.styles.runClass(styleId),
    ];
    for (const item of this.compat(source)) {
      if (item.uri === W && item.name === "pPr") continue;
      this.inline(item, paragraph, paragraphRun);
    }
    const mark = child(properties, "rPr");
    if (mark) {
      const character = attr(child(mark, "rStyle"), "val"),
        known = character && this.styles.hasStyle(character, "character") ? character : undefined;
      const inheritedMark = this.styles.runStyle(
          known ?? this.styles.defaultCharacter,
          paragraphRun,
        ),
        markCss = this.styles.run(mark, inheritedMark, false, true, true);
      if (known) {
        const standalone = this.styles.get(known).r;
        for (const name of ["font-weight", "font-style", "font-variant", "text-transform"])
          if (
            standalone[name] !== undefined &&
            inheritedMark[name] !== standalone[name] &&
            markCss[name] === undefined
          )
            markCss[name] = inheritedMark[name]!;
      }
      const marker = this.make("span", {
        "data-word-paragraph-mark": "true",
        ...(known ? { "data-word-style": known, class: this.styles.styleClass(known) } : {}),
        style: declarations(markCss),
      });
      paragraph.children.push(marker);
      readWordFormatRevision(this.word, mark, marker, "run", this.styles, {
        inheritedRun: paragraphRun,
      });
      this.native?.bind(mark, marker.attrs);
    }
    readWordFormatRevision(this.word, properties, paragraph, "paragraph", this.styles, {
      defaultTab: this.defaultTab,
    });
    this.paragraphRunClasses = oldRunClasses;
    this.native?.bind(source, paragraph.attrs);
    return paragraph;
  }
  private hasTab(source: XmlNode): boolean {
    return source.content.some(
      (item) =>
        typeof item !== "string" && ((item.uri === W && item.name === "tab") || this.hasTab(item)),
    );
  }
  private fieldFlags(source: XmlNode): Pick<WordMetadata["fields"][number], "locked" | "dirty"> {
    const result: Pick<WordMetadata["fields"][number], "locked" | "dirty"> = {};
    for (const [name, key] of [
      ["fldLock", "locked"],
      ["dirty", "dirty"],
    ] as const) {
      const value = attr(source, name);
      if (value !== undefined) {
        if (!["0", "1", "true", "false", "on", "off"].includes(value))
          fail(source, "A field flag is invalid.");
        result[key] = ["1", "true", "on"].includes(value);
      }
    }
    return result;
  }
  private nativeObject(source: XmlNode, tag: "p" | "span"): Html {
    const object = this.make(tag, {});
    this.native!.bind(source, object.attrs);
    object.children.push("\ufffc");
    const active = new Set(this.activeComments),
      covered = new Set(active),
      original = this.native!.originalPlaceholder(source)!;
    const visit = (node: XmlNode, parent = true): void => {
      const scope = wordScope(node, parent),
        kind = scope ? commentControl(node) : undefined;
      if (kind) {
        const id = attr(node, "id"),
          comment = id === undefined ? undefined : this.comments.get(id);
        if (!comment) fail(node, "A native comment control targets a missing comment.");
        if (kind === "commentRangeStart") {
          if (active.has(id!)) fail(node, "A native comment range start is repeated.");
          active.add(id!);
          covered.add(id!);
        } else if (kind === "commentRangeEnd") {
          if (!active.delete(id!)) fail(node, "A native comment range end has no start.");
        } else comment.point = this.id(object);
      }
      const children =
        node.uri === MC && node.name === "AlternateContent"
          ? [elements(node, "Choice", MC)[0] ?? child(node, "Fallback", MC)].filter(
              (item): item is XmlNode => !!item,
            )
          : node.content.filter((item): item is XmlNode => typeof item !== "string");
      for (const item of children) visit(item, scope);
    };
    visit(original);
    this.activeComments = covered;
    this.record(object, 0, "\ufffc");
    this.activeComments = active;
    return object;
  }
  private inline(source: XmlNode, parent: Html, paragraphStyle: Css = {}): void {
    if (this.native?.isPlaceholder(source)) {
      this.offset(parent, "\ufffc");
      const object = this.nativeObject(source, "span");
      parent.children.push(object);
      this.fieldResult(object);
      return;
    }
    if (this.native?.preservedMove(source))
      unsupported(
        source,
        "A move crossing a native object boundary remains preserved as native content.",
      );
    if (
      this.native &&
      source.uri === W &&
      /^(?:bookmark|perm|(?:moveFrom|moveTo|customXmlIns|customXmlDel|customXmlMoveFrom|customXmlMoveTo)Range)(?:Start|End)$/.test(
        source.name,
      )
    ) {
      const marker = this.make("span");
      if (this.native.rangeMarker(source, marker.attrs)) {
        parent.children.push(marker);
        return;
      }
    }
    if (source.uri === M && source.name === "oMath") {
      this.math(source, parent, "inline");
      return;
    }
    if (source.uri === M && source.name === "oMathPara") {
      this.mathParagraph(source, parent);
      return;
    }
    if (source.uri !== W) unsupportedType(source);
    switch (source.name) {
      case "sdt":
        parent.children.push(this.control(source, "inline", { p: {}, r: paragraphStyle }));
        return;
      case "r": {
        const properties = child(source, "rPr");
        const requestedStyleId = attr(child(properties, "rStyle"), "val");
        const styleId =
          requestedStyleId && this.styles.hasStyle(requestedStyleId, "character")
            ? requestedStyleId
            : undefined;
        const inherited = this.styles.runStyle(styleId ?? this.styles.defaultCharacter, {
          ...this.styles.defaults.r,
          ...paragraphStyle,
        });
        const css = this.styles.run(properties, inherited, false, false, true);
        // Character-style toggles are relative to paragraph formatting; other
        // inherited font/color values remain inherited CSS rather than copies.
        if (styleId && !this.activeTable) {
          const standalone = this.styles.get(styleId).r;
          for (const name of ["font-weight", "font-style", "font-variant", "text-transform"]) {
            if (
              standalone[name] !== undefined &&
              inherited[name] !== standalone[name] &&
              css[name] === undefined
            )
              css[name] = inherited[name]!;
          }
        }
        const classes = [
          ...this.paragraphRunClasses,
          ...(styleId ? [this.styles.styleClass(styleId)] : []),
        ];
        const run = this.make("span", {
          style: declarations(css),
          class: classes.join(" "),
          "data-word-run": "true",
          ...(styleId
            ? {
                "data-word-style": styleId,
              }
            : {}),
        });
        if (this.activeTable)
          this.context(run, "run", {
            ...Object.fromEntries(
              Object.entries(inherited).filter(
                ([key, value]) =>
                  value !== paragraphStyle[key] ||
                  key === "background-color" ||
                  key.startsWith("text-decoration"),
              ),
            ),
          });
        if (requestedStyleId && !styleId)
          run.attrs["data-word-ignored-character-style"] = requestedStyleId;
        readWordFormatRevision(this.word, properties, run, "run", this.styles, {
          inheritedRun: { ...this.styles.defaults.r, ...paragraphStyle },
          runClasses: this.paragraphRunClasses,
        });
        this.native?.bind(source, run.attrs);
        parent.children.push(run);
        const language = attr(child(properties, "lang"), "val");
        if (language) run.attrs.lang = language;
        const previousRun = this.currentRun;
        this.currentRun = run;
        const contentNames = [
          "t",
          "delText",
          "br",
          "cr",
          "tab",
          "noBreakHyphen",
          "softHyphen",
          "drawing",
          "fldChar",
          "instrText",
          "commentReference",
          "lastRenderedPageBreak",
          "annotationRef",
          "footnoteReference",
          "endnoteReference",
          "footnoteRef",
          "endnoteRef",
          "separator",
          "continuationSeparator",
        ];
        for (const item of this.compat(source)) {
          if (item.uri === W && item.name === "rPr") continue;
          if (
            !this.native?.isPlaceholder(item) &&
            (item.uri !== W || !contentNames.includes(item.name))
          )
            unsupportedType(item, "The run contains unsupported or invalid inline content.");
          this.inline(item, run, paragraphStyle);
        }
        this.currentRun = previousRun;
        break;
      }
      case "t":
      case "delText": {
        if (this.pendingNoteCustomMark) {
          const reference = this.word.noteReferences!.find(
            (item) => item.nodeId === this.pendingNoteCustomMark,
          )!;
          reference.customMark = text(source);
          this.pendingNoteCustomMark = undefined;
          break;
        }
        if (source.content.some((item) => typeof item !== "string"))
          fail(source, "A text node contains XML elements.");
        const active = this.fieldStack[this.fieldStack.length - 1];
        if (active && !active.result)
          fail(source, "A complex field contains result text before its separator.");
        this.addText(parent, text(source));
        break;
      }
      case "br":
      case "cr": {
        const type = attr(source, "type") ?? "textWrapping";
        if (!["textWrapping", "page", "column"].includes(type))
          unsupported(source, "The break type is not supported.");
        const mark = this.make(
          "br",
          type === "page" || type === "column" ? { "data-word-break": type } : {},
        );
        const offset = this.offset(parent, "\n");
        parent.children.push(mark);
        this.fieldResult(mark);
        this.record(parent, offset, "\n");
        break;
      }
      case "tab": {
        const tab = this.make("span", { "data-word-tab": "true" });
        parent.children.push(tab);
        this.fieldResult(tab);
        this.addText(tab, "\t");
        break;
      }
      case "noBreakHyphen":
        this.addText(parent, "\u2011");
        break;
      case "softHyphen":
        this.addText(parent, "\u00ad");
        break;
      case "drawing": {
        const image = this.image(source);
        const value = (node: Html): string =>
          node.tag === "img"
            ? "\ufffc"
            : node.children
                .map((child) => (typeof child === "string" ? child : value(child)))
                .join("");
        const content = value(image),
          offset = this.offset(parent, content);
        parent.children.push(image);
        this.fieldResult(image);
        this.record(parent, offset, content);
        break;
      }
      case "hyperlink": {
        const link = this.make("span", { "data-word-link": "true" });
        parent.children.push(link);
        const id = attr(source, "id", R);
        const bookmark = attr(source, "anchor");
        let target: string | undefined;
        if (id) {
          const relation = this.opc.relationships(this.currentPart).get(id);
          if (!relation || !relation.type.endsWith("/hyperlink") || !relation.external)
            fail(source, "A hyperlink relationship is invalid.");
          target = relation.target;
        }
        this.word.hyperlinks.push({
          nodeId: this.id(link),
          ...(target ? { target } : {}),
          ...(bookmark ? { bookmark } : {}),
        });
        for (const item of this.compat(source)) this.inline(item, link, paragraphStyle);
        break;
      }
      case "footnoteReference":
      case "endnoteReference": {
        const kind = source.name === "footnoteReference" ? "footnote" : "endnote",
          nativeId = attr(source, "id"),
          noteId = this.nativeNotes.get(kind + "\n" + nativeId);
        if (!noteId) fail(source, "A note reference targets a missing native note.");
        const offset = this.offset(parent, "\ufffc"),
          marker = this.make("span", { "data-word-note-reference": kind });
        marker.children.push("\ufffc");
        parent.children.push(marker);
        (this.word.noteReferences ??= []).push({ nodeId: this.id(marker), noteId });
        this.native?.bind(source, marker.attrs);
        this.fieldResult(marker);
        this.record(parent, offset, "\ufffc");
        if (["1", "true", "on"].includes(attr(source, "customMarkFollows") ?? "0"))
          this.pendingNoteCustomMark = this.id(marker);
        break;
      }
      case "footnoteRef":
      case "endnoteRef": {
        const offset = this.offset(parent, "\ufffc"),
          marker = this.make("span", { "data-word-note-number": "true" });
        marker.children.push("\ufffc");
        parent.children.push(marker);
        this.fieldResult(marker);
        this.record(parent, offset, "\ufffc");
        break;
      }
      case "separator":
      case "continuationSeparator": {
        parent.children.push(this.make("span", { "data-word-note-separator": source.name }));
        break;
      }
      case "bookmarkStart": {
        const id = attr(source, "id");
        const name = attr(source, "name");
        if (id === undefined || !name || this.bookmarkStarts.has(id))
          fail(source, "A bookmark start is invalid.");
        const marker = this.make("span", { "data-word-bookmark": "start" });
        parent.children.push(marker);
        this.bookmarkStarts.set(id, { name, nodeId: this.id(marker) });
        break;
      }
      case "bookmarkEnd": {
        const id = attr(source, "id");
        const start = id !== undefined ? this.bookmarkStarts.get(id) : undefined;
        if (!start) fail(source, "A bookmark end has no start.");
        const marker = this.make("span", { "data-word-bookmark": "end" });
        parent.children.push(marker);
        this.bookmarkStarts.delete(id!);
        this.word.bookmarks.push({
          bookmarkId: this.currentStory + "-" + id,
          name: start.name,
          startNodeId: start.nodeId,
          endNodeId: this.id(marker),
        });
        break;
      }
      case "fldSimple": {
        if (this.fieldStack.length >= this.opc.limits.maxDepth)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Field nesting exceeds the configured depth.",
          );
        const field = this.make("span", { "data-word-field": "simple" });
        parent.children.push(field);
        this.fieldResult(field);
        const definition: WordMetadata["fields"][number] = {
          nodeId: this.id(field),
          instruction: attr(source, "instr") ?? "",
          resultState: "cached",
          ...this.fieldFlags(source),
          resultNodeIds: [],
        };
        const active = this.fieldStack[this.fieldStack.length - 1];
        if (active && !active.result)
          (active.definition.instructionFields ??= []).push({
            offset: active.definition.instruction.length,
            fieldNodeId: definition.nodeId,
          });
        this.word.fields.push(definition);
        const state = { definition, result: true };
        this.fieldStack.push(state);
        for (const item of this.compat(source)) this.inline(item, field, paragraphStyle);
        if (this.fieldStack.pop() !== state)
          fail(source, "A simple field has unbalanced nested controls.");
        break;
      }
      case "fldChar": {
        check(source, []);
        const type = attr(source, "fldCharType");
        if (type === "begin") {
          if (this.fieldStack.length >= this.opc.limits.maxDepth)
            throw new DocumentError(
              "ResourceLimitExceeded",
              "Field nesting exceeds the configured depth.",
            );
          const marker = this.make("span", { "data-word-field": "begin" });
          parent.children.push(marker);
          const definition: WordMetadata["fields"][number] = {
            nodeId: this.id(marker),
            instruction: "",
            resultState: "cached",
            ...this.fieldFlags(source),
            resultNodeIds: [],
          };
          const active = this.fieldStack[this.fieldStack.length - 1];
          if (active && !active.result)
            (active.definition.instructionFields ??= []).push({
              offset: active.definition.instruction.length,
              fieldNodeId: definition.nodeId,
            });
          this.word.fields.push(definition);
          this.fieldStack.push({ definition, result: false });
        } else {
          const active = this.fieldStack[this.fieldStack.length - 1];
          if (!active) fail(source, "A field control has no matching begin.");
          if (type === "separate") {
            if (active.result) fail(source, "A field has repeated separators.");
            active.result = true;
          } else if (type === "end") {
            const marker = this.make("span", {
              "data-word-field-end": active.definition.nodeId,
            });
            parent.children.push(marker);
            active.definition.endNodeId = this.id(marker);
            this.fieldStack.pop();
          } else fail(source, "A field control type is invalid.");
        }
        break;
      }
      case "instrText": {
        const active = this.fieldStack[this.fieldStack.length - 1];
        if (!active || active.result)
          fail(source, "Field instructions occur outside a field instruction range.");
        active.definition.instruction += text(source);
        break;
      }
      case "commentRangeStart": {
        const id = attr(source, "id");
        if (id === undefined || !this.comments.has(id) || this.activeComments.has(id))
          fail(source, "A comment range start is invalid.");
        this.activeComments.add(id);
        break;
      }
      case "commentRangeEnd": {
        const id = attr(source, "id");
        if (id === undefined || !this.activeComments.delete(id))
          fail(source, "A comment range end is invalid.");
        break;
      }
      case "commentReference": {
        const id = attr(source, "id");
        const state = id !== undefined ? this.comments.get(id) : undefined;
        if (!state) fail(source, "A comment reference is invalid.");
        const marker = this.make("span");
        parent.children.push(marker);
        state.point = this.id(marker);
        break;
      }
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd": {
        const marker = this.moves.marker(source, (tag, attrs) => this.make(tag, attrs));
        if (marker) parent.children.push(marker);
        break;
      }
      case "ins":
      case "del":
      case "moveFrom":
      case "moveTo": {
        const sourceId = attr(source, "id");
        if (sourceId === undefined) fail(source, "A revision has no identity.");
        const revision = this.make(["ins", "moveTo"].includes(source.name) ? "ins" : "del", {
          "data-word-run-overlay": "true",
        });
        parent.children.push(revision);
        const author = attr(source, "author");
        const date = attr(source, "date");
        this.word.revisions.push({
          nodeId: this.id(revision),
          sourceId,
          ...(source.name.startsWith("move")
            ? this.moves.revision(source)
            : { kind: source.name === "ins" ? ("insert" as const) : ("delete" as const) }),
          ...(author !== undefined ? { author } : {}),
          ...(date !== undefined ? { date } : {}),
        });
        for (const item of this.compat(source)) this.inline(item, revision, paragraphStyle);
        break;
      }
      case "annotationRef":
        parent.children.push(this.make("span", { "data-word-comment-mark": "true" }));
        break;
      case "proofErr":
      case "lastRenderedPageBreak":
        break; // Proofing and stale renderer positions carry no source content.
      default:
        unsupportedType(source);
    }
  }
  private section(
    source: XmlNode | undefined,
    root: Html,
    previousProperties = false,
  ): DocumentSection {
    check(source, [
      "headerReference",
      "footerReference",
      "footnotePr",
      "endnotePr",
      "type",
      "pgSz",
      "pgMar",
      "pgBorders",
      "paperSrc",
      "lnNumType",
      "pgNumType",
      "cols",
      "formProt",
      "vAlign",
      "noEndnote",
      "titlePg",
      "textDirection",
      "bidi",
      "rtlGutter",
      "docGrid",
      "printerSettings",
      "sectPrChange",
    ]);
    for (const feature of ["lnNumType"]) {
      if (child(source, feature)) unsupported(child(source, feature)!);
    }
    if (
      (child(source, "textDirection") && attr(child(source, "textDirection"), "val") !== "lrTb") ||
      (child(source, "vAlign") && attr(child(source, "vAlign"), "val") !== "top") ||
      on(child(source, "bidi")) ||
      on(child(source, "formProt"))
    )
      unsupported(source!, "The section writing direction or protection is not supported.");
    const grid = child(source, "docGrid");
    const gridType = attr(grid, "type") ?? "default";
    if (!["default", "lines", "linesAndChars", "snapToChars"].includes(gridType))
      unsupported(grid!, "The document grid type is not supported.");
    const columns = child(source, "cols");
    const unequal = columns && ["0", "false", "off"].includes(attr(columns, "equalWidth") ?? ""),
      columnDefinitions = unequal ? elements(columns, "col") : [];
    if (columns) check(columns, ["col"]);
    if (unequal && !columnDefinitions.length)
      fail(columns!, "Unequal columns require explicit column widths.");
    const size = child(source, "pgSz");
    const margins = child(source, "pgMar");
    if (source && !previousProperties) this.native?.bind(source, root.attrs);
    const previous = this.sections[this.sections.length - 1];
    const fallback = previous?.pageSetup ?? {
      unit: "pt" as const,
      width: 612,
      height: 792,
      margins: { top: 72, right: 72, bottom: 72, left: 72 },
    };
    const setup: PageSetup = {
      unit: "pt",
      width: size ? number(attr(size, "w"), size) / 20 : fallback.width,
      height: size ? number(attr(size, "h"), size) / 20 : fallback.height,
      margins: { ...fallback.margins },
    };
    if (margins) {
      for (const side of ["top", "right", "bottom", "left"] as const) {
        const value = attr(margins, side);
        if (value !== undefined) setup.margins[side] = number(value, margins) / 20;
      }
    }
    const headers = { ...previous?.headers };
    const footers = { ...previous?.footers };
    for (const [name, kind, refs] of [
      ["headerReference", "header", headers],
      ["footerReference", "footer", footers],
    ] as const) {
      for (const ref of elements(source, name)) {
        const type = attr(ref, "type") ?? "default";
        if (!["default", "first", "even"].includes(type))
          fail(ref, "A header or footer reference has an invalid type.");
        const relation = this.opc.relationships(this.main).get(attr(ref, "id", R) ?? "");
        if (!relation || relation.external || !relation.type.endsWith("/" + kind))
          fail(ref, "A header or footer relationship is invalid.");
        refs[type as "default" | "first" | "even"] = this.readStory(relation.target, kind);
      }
    }
    const breakType = attr(child(source, "type"), "val") ?? "nextPage";
    if (!["continuous", "nextPage", "nextColumn", "evenPage", "oddPage"].includes(breakType))
      unsupported(source!, "A section break type is not supported.");
    const pageNumber = child(source, "pgNumType");
    const pageNumberStart = attr(pageNumber, "start");
    const pageNumberFormat = attr(pageNumber, "fmt");
    if (
      pageNumberFormat &&
      !["decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"].includes(
        pageNumberFormat,
      )
    )
      unsupported(pageNumber!, "The page numbering format is not supported.");
    const definition: DocumentSection = {
      ...(pageNumberStart !== undefined
        ? { pageNumberStart: number(pageNumberStart, pageNumber!) }
        : {}),
      ...(pageNumberFormat
        ? {
            pageNumberFormat: pageNumberFormat as
              | "decimal"
              | "upperRoman"
              | "lowerRoman"
              | "upperLetter"
              | "lowerLetter",
          }
        : {}),
      sectionId: "section-" + (this.sections.length + 1),
      rootNodeId: this.id(root),
      ...(child(source, "pgBorders")
        ? { pageBorders: readWordPageBorders(child(source, "pgBorders"))! }
        : {}),
      ...(columns
        ? {
            columns: {
              count: unequal ? columnDefinitions.length : number(attr(columns, "num"), columns, 1),
              gap: number(attr(columns, "space"), columns, 720) / 20,
              separator: ["1", "true", "on"].includes(attr(columns, "sep") ?? ""),
              ...(unequal
                ? {
                    widths: columnDefinitions.map(
                      (column) => number(attr(column, "w"), column) / 20,
                    ),
                    spaces: columnDefinitions.map(
                      (column) => number(attr(column, "space"), column, 0) / 20,
                    ),
                  }
                : {}),
            },
          }
        : {}),
      ...(grid
        ? {
            grid: {
              type: gridType as "default" | "lines" | "linesAndChars" | "snapToChars",
              ...(attr(grid, "linePitch") !== undefined
                ? { linePitch: number(attr(grid, "linePitch"), grid) / 20 }
                : {}),
              ...(attr(grid, "charSpace") !== undefined
                ? { characterSpacing: number(attr(grid, "charSpace"), grid) }
                : {}),
            },
          }
        : {}),
      ...(margins && attr(margins, "gutter") !== undefined
        ? { gutter: number(attr(margins, "gutter"), margins) / 20 }
        : {}),
      gutterPosition: on(child(this.settings, "gutterAtTop"))
        ? "top"
        : on(child(source, "rtlGutter"))
          ? "right"
          : "left",
      pageSetup: setup,
      breakType: breakType as DocumentSection["breakType"] & string,
      differentFirstPage: on(child(source, "titlePg")),
      ...(child(source, "footnotePr")
        ? {
            footnoteOptions: readNoteOptions(
              child(source, "footnotePr"),
            ) as DocumentSection["footnoteOptions"],
          }
        : {}),
      ...(child(source, "endnotePr")
        ? {
            endnoteOptions: readNoteOptions(
              child(source, "endnotePr"),
            ) as DocumentSection["endnoteOptions"],
          }
        : {}),
      ...(child(source, "noEndnote") ? { noEndnote: on(child(source, "noEndnote")) } : {}),
      headerDistance: margins ? number(attr(margins, "header"), margins, 720) / 20 : 36,
      footerDistance: margins ? number(attr(margins, "footer"), margins, 720) / 20 : 36,
      headers,
      footers,
    };
    if (definition.columns?.widths) {
      const columns = definition.columns,
        available =
          setup.width -
          setup.margins.left -
          setup.margins.right -
          (definition.gutterPosition === "top" ? 0 : (definition.gutter ?? 0));
      if (
        columns.widths!.reduce(
          (sum, width, index) =>
            sum +
            width +
            (index < columns.count - 1 ? (columns.spaces?.[index] ?? columns.gap) : 0),
          0,
        ) >
        available + 0.1
      )
        unsupported(
          child(source, "cols")!,
          "Overflowing native columns retain their original layout.",
        );
    }
    const change = child(source, "sectPrChange");
    if (change) {
      if (previousProperties) fail(change, "Section property histories cannot be nested.");
      check(change, ["sectPr"]);
      const original = child(change, "sectPr"),
        id = attr(change, "id");
      if (!original || id === undefined)
        fail(change, "A section revision requires its previous properties and identity.");
      const before = this.section(original, root, true),
        { sectionId: _sectionId, rootNodeId: _rootNodeId, ...properties } = before;
      this.word.revisions.push({
        nodeId: this.id(root),
        revisionId: this.id(root) + "-format-section",
        sourceId: id,
        kind: "format",
        property: "section",
        before: { section: properties },
        ...(attr(change, "author") !== undefined ? { author: attr(change, "author")! } : {}),
        ...(attr(change, "date") !== undefined ? { date: attr(change, "date")! } : {}),
      });
    }
    return definition;
  }
  private readStory(part: string, kind: "header" | "footer"): string {
    const cached = this.loadedStories.get(part);
    if (cached) return cached;
    const source = this.opc.read(part);
    if (source.uri !== W || source.name !== (kind === "header" ? "hdr" : "ftr"))
      fail(source, "An auxiliary story has the wrong root.");
    const storyId = kind + "-" + this.loadedStories.size;
    this.loadedStories.set(part, storyId);
    const root = this.story(kind, storyId);
    root.attrs.class = "word-main";
    this.native?.bind(source, root.attrs);
    root.children.push(...this.readStoryContent(part, storyId, source));
    this.auxiliary.push(root);
    return storyId;
  }
  private readStoryContent(part: string, storyId: string, source: XmlNode): Html[] {
    const {
      currentPart,
      currentStory,
      fieldStack,
      activeComments,
      bookmarkStarts,
      currentRun,
      paragraphRunClasses,
      moves,
    } = this;
    this.currentPart = part;
    this.currentStory = storyId;
    this.moves = new WordMoveReader(this.word);
    this.fieldStack = [];
    this.activeComments = new Set();
    this.bookmarkStarts = new Map();
    this.currentRun = undefined;
    this.paragraphRunClasses = ["word-default-run"];
    try {
      const content = this.blocks(this.compat(source));
      if (
        this.fieldStack.length ||
        this.activeComments.size ||
        this.bookmarkStarts.size ||
        this.moves.pending
      )
        fail(source, "An auxiliary story has an unclosed content range.");
      return content;
    } finally {
      this.currentPart = currentPart;
      this.currentStory = currentStory;
      this.fieldStack = fieldStack;
      this.activeComments = activeComments;
      this.bookmarkStarts = bookmarkStarts;
      this.currentRun = currentRun;
      this.paragraphRunClasses = paragraphRunClasses;
      this.moves = moves;
    }
  }
  private control(
    source: XmlNode,
    level: "inline" | "block",
    inherited: { p: Css; r: Css; runClasses?: readonly string[] },
  ): Html {
    const definition = readWordControl(source, level),
      properties = child(source, "sdtPr"),
      content = child(source, "sdtContent");
    check(source, ["sdtPr", "sdtEndPr", "sdtContent"]);
    if (!content) fail(source, "A content control has no content.");
    const css = this.styles.run(child(properties, "rPr"), inherited.r),
      wrapper = this.make(level === "inline" ? "span" : "div", {
        "data-word-content-control": definition.kind,
        style: declarations(css),
      });
    this.native?.bind(source, wrapper.attrs);
    (this.word.controls ??= []).push({ ...definition, nodeId: this.id(wrapper) });
    if (level === "block")
      wrapper.children.push(
        ...this.blocks(this.compat(content), { ...inherited, r: { ...inherited.r, ...css } }),
      );
    else
      for (const item of this.compat(content))
        this.inline(item, wrapper, { ...inherited.r, ...css });
    return wrapper;
  }
  private mathParagraph(source: XmlNode, parent: Html): void {
    try {
      check(source, ["oMathParaPr", "oMath"], M);
      const properties = child(source, "oMathParaPr", M);
      check(properties, ["jc"], M);
      const align =
          attr(child(properties, "jc", M), "val", M) ??
          attr(child(child(this.settings, "mathPr", M), "defJc", M), "val", M) ??
          "centerGroup",
        maths = elements(source, "oMath", M);
      if (!maths.length || !["left", "right", "center", "centerGroup"].includes(align))
        unsupported(source, "This math paragraph is preserved natively.");
      this.math(source, parent, "block", align as "left" | "right" | "center" | "centerGroup");
    } catch (error) {
      if (error instanceof DocumentError && error.code === "UnsupportedFeature")
        unsupported(source, "This math paragraph is preserved natively.");
      throw error;
    }
  }
  private math(
    source: XmlNode,
    parent: Html,
    display: "inline" | "block",
    alignment?: "left" | "right" | "center" | "centerGroup",
  ): void {
    const defaults = child(this.settings, "mathPr", M),
      math = readWordMath(source, {
        display,
        ...(attr(child(defaults, "intLim", M), "val", M)
          ? { integralLimits: attr(child(defaults, "intLim", M), "val", M)! }
          : {}),
        ...(attr(child(defaults, "naryLim", M), "val", M)
          ? { naryLimits: attr(child(defaults, "naryLim", M), "val", M)! }
          : {}),
      }),
      marker = this.make("span", { "data-word-equation": "true" }),
      offset = this.offset(parent, "\ufffc");
    marker.children.push("\ufffc");
    parent.children.push(marker);
    (this.word.equations ??= []).push({
      nodeId: this.id(marker),
      display,
      ...math,
      ...(alignment ? { alignment } : {}),
      ...(attr(child(defaults, "mathFont", M), "val", M)
        ? { fontFamily: attr(child(defaults, "mathFont", M), "val", M)! }
        : {}),
    });
    this.native?.bind(source, marker.attrs);
    this.fieldResult(marker);
    this.record(parent, offset, "\ufffc");
  }
  private readNotes(): void {
    for (const kind of ["footnote", "endnote"] as const) {
      const part = this.opc.related(this.main, kind + "s");
      if (!part) continue;
      const root = this.opc.read(part);
      if (root.uri !== W || root.name !== kind + "s")
        fail(root, "A note part has an invalid root.");
      check(root, [kind]);
      for (const source of elements(root, kind)) {
        const id = attr(source, "id");
        if (
          id === undefined ||
          !Number.isSafeInteger(Number(id)) ||
          this.nativeNotes.has(kind + "\n" + id)
        )
          fail(source, "A native note identity is invalid or repeated.");
        const type = attr(source, "type") ?? "normal";
        if (!["normal", "separator", "continuationSeparator", "continuationNotice"].includes(type))
          fail(source, "A native note type is invalid.");
        const storyId = kind + "-" + id,
          story = this.story(kind, storyId);
        story.attrs.class = "word-main";
        this.nativeNotes.set(kind + "\n" + id, storyId);
        (this.word.notes ??= []).push({
          noteId: storyId,
          storyId,
          kind,
          type: type as "normal" | "separator" | "continuationSeparator" | "continuationNotice",
        });
        this.native?.bind(source, story.attrs);
        story.children.push(...this.readStoryContent(part, storyId, source));
        this.auxiliary.push(story);
      }
    }
  }
  private readComments(): void {
    const part = this.opc.related(this.main, "comments");
    if (!part) {
      if (this.opc.related(this.main, "commentsExtended"))
        fail(this.opc.read(this.main), "Comment status requires a classic comments part.");
      readWordCommentExtensions(
        this.opc,
        this.main,
        this.comments,
        this.word,
        this.id(this.root),
        this.native,
      );
      return;
    }
    const root = this.opc.read(part);
    if (root.uri !== W || root.name !== "comments")
      fail(root, "The comments part has an invalid root.");
    check(root, ["comment"]);
    for (const source of elements(root, "comment")) {
      const id = attr(source, "id");
      if (id === undefined || this.comments.has(id)) fail(source, "A comment identity is invalid.");
      const story = this.story("comment", "comment-" + id);
      story.attrs.class = "word-main";
      story.children.push(...this.readStoryContent(part, "comment-" + id, source));
      this.auxiliary.push(story);
      this.comments.set(id, { source, bodyNodeId: this.id(story), ranges: [] });
      this.native?.bind(source, story.attrs);
    }
    readWordCommentExtensions(
      this.opc,
      this.main,
      this.comments,
      this.word,
      this.id(this.root),
      this.native,
    );
  }

  private readonly drawingNativeIds = new Map<string, string>();
  private readonly drawingAliases = new Map<string, string | null>();
  private readonly drawingConnections: {
    source: XmlNode;
    nodeId: string;
    part: string;
    start?: string;
    end?: string;
    local: ReadonlyMap<string, string>;
  }[] = [];
  private resolveDrawingConnections(): void {
    for (const connection of this.drawingConnections) {
      const definition = this.word.drawings!.find((item) => item.nodeId === connection.nodeId)!;
      if (definition.kind !== "connector") continue;
      for (const side of ["start", "end"] as const) {
        const nativeId = connection[side];
        if (nativeId === undefined) continue;
        const key = connection.part + "\n" + nativeId,
          candidates = [
            connection.local.get(nativeId),
            this.drawingNativeIds.get(key),
            this.drawingAliases.get(key),
          ];
        const target = candidates.find((id) => id && id !== definition.nodeId);
        if (!target)
          unsupported(
            connection.source,
            "The connector endpoint is preserved with its native drawing.",
          );
        definition.connector[side === "start" ? "startNodeId" : "endNodeId"] = target;
      }
    }
  }
  private image(source: XmlNode): Html {
    return importWordDrawing(source, {
      word: this.word,
      opc: this.opc,
      part: this.currentPart,
      main: this.main,
      make: (tag, attrs) => this.make(tag, attrs),
      asset: (part, mime) => {
        let path = this.assetSources.get(part);
        if (path) return path;
        const ext = /\.([a-zA-Z0-9]{1,8})$/.exec(part)?.[1] ?? "bin";
        path = "assets/chart-" + (this.assets.length + 1) + "." + ext;
        const bytes = this.opc.files.get(part);
        if (!bytes) fail(source, "A chart resource is missing.");
        this.files.set(path, new Uint8Array(bytes));
        this.assets.push(createAssetDescriptor(path, mime, bytes));
        this.assetSources.set(part, path);
        return path;
      },
      image: (part, source) => {
        let path = this.assetSources.get(part);
        if (path) return path;
        const mime = this.opc.contentType(part),
          extensions: Record<string, string> = {
            "image/png": "png",
            "image/jpeg": "jpg",
            "image/webp": "webp",
          };
        if (!mime || !extensions[mime])
          unsupported(source, "This image encoding is preserved as native content.");
        path = "assets/image-" + (this.assets.length + 1) + "." + extensions[mime];
        const bytes = this.opc.files.get(part);
        if (!bytes) fail(source, "An embedded image resource is missing.");
        this.files.set(path, new Uint8Array(bytes));
        this.assets.push(createAssetDescriptor(path, mime, bytes));
        this.assetSources.set(part, path);
        return path;
      },
      story: (source) => {
        const id = "text-box-" + this.word.stories.length,
          root = this.story("textBox", id);
        root.attrs.class = "word-main";
        root.children.push(...this.readStoryContent(this.currentPart, id, source));
        this.auxiliary.push(root);
        return id;
      },
      native: (source, node) => this.native?.bind(source, node.attrs),
      identity: (source, id, nativeId) => {
        if (nativeId === undefined) return;
        number(nativeId, source);
        this.native?.identity(source, id, nativeId);
        const key = this.currentPart + "\n" + nativeId;
        if (source.uri === WP) {
          const previous = this.drawingNativeIds.get(key);
          if (previous && previous !== id) fail(source, "A drawing frame identity is repeated.");
          this.drawingNativeIds.set(key, id);
        } else if (nativeId !== "0") {
          const previous = this.drawingAliases.get(key);
          this.drawingAliases.set(key, previous === undefined || previous === id ? id : null);
        }
      },
      connector: (source, nodeId, start, end, local) =>
        this.drawingConnections.push({
          source,
          nodeId,
          part: this.currentPart,
          ...(start !== undefined ? { start } : {}),
          ...(end !== undefined ? { end } : {}),
          local,
        }),
    });
  }
  private validateMergeContinuation(cell: XmlNode): WordTableEmptyParagraph[] {
    if (cell.attrs.size || child(child(cell, "tcPr"), "tcFitText"))
      unsupported(cell, "A vertical merge continuation retains independent native properties.");
    const cellChange = child(child(cell, "tcPr"), "tcPrChange");
    if (cellChange)
      unsupported(cellChange, "A vertical merge continuation has independent formatting history.");

    return readWordEmptyContinuation(
      cell,
      this.styles,
      (node) => this.compat(node),
      (css) => {
        if (!Object.keys(css).length) return undefined;
        const name = "word-empty-continuation-" + this.styles.css.length;
        this.styles.css.push("." + name + "{" + declarations(css) + "}");
        return name;
      },
      this.defaultTab,
    );
  }

  private table(source: XmlNode): Html {
    check(source, ["tblPr", "tblGrid", "tr"]);
    const properties = child(source, "tblPr");
    const styleId = attr(child(properties, "tblStyle"), "val") ?? this.styles.defaultTable;
    const style = this.styles.get(styleId);
    const tableCss: Css = {
      "margin-left": "-5.4pt",
      ...style.table,
      ...resolveWordTableStyle(this.word, this.styles.readComponent, styleId, []).table,
      ...this.styles.table(properties),
    };
    const table = this.make("table", {
      style: declarations(this.styles.table(properties)),
      ...(styleId ? { "data-word-style": styleId, class: this.styles.styleClass(styleId) } : {}),
    });
    const gridSource = child(source, "tblGrid");
    const grid = readEditableWordGrid(
      source,
      this.styles,
      this.tableAvailable,
      this.opc.limits.maxNodes,
    );
    this.word.tables.push({ nodeId: this.id(table), grid });
    let previous = new Map<number, { cell: Html; span: number; css: Css }>();
    const rows = elements(source, "tr");
    let rowIndex = 0;
    const look = child(properties, "tblLook");
    const rawLook = attr(look, "val");
    if (rawLook && !/^[0-9a-f]{1,4}$/i.test(rawLook)) fail(look!, "A table look mask is invalid.");
    const mask = tableLookMask(look) ?? this.styles.tableSetting(styleId, "tableLookMask") ?? 0;
    const tableInfo = this.word.tables[this.word.tables.length - 1]!;
    this.context(
      table,
      "table",
      {
        "margin-left": "-5.4pt",
        ...style.table,
        ...resolveWordTableStyle(this.word, this.styles.readComponent, styleId, []).table,
      },
      tableInfo,
    );
    const ownMask = tableLookMask(look);
    if (ownMask !== undefined) tableInfo.lookMask = ownMask;
    for (const [name, key] of [
      ["tblStyleRowBandSize", "rowBandSize"],
      ["tblStyleColBandSize", "columnBandSize"],
    ] as const) {
      const item = child(properties, name);
      if (item) tableInfo[key] = number(attr(item, "val"), item);
    }
    const rowBand = tableInfo.rowBandSize ?? this.styles.tableSetting(styleId, "rowBandSize") ?? 0;
    const columnBand =
      tableInfo.columnBandSize ?? this.styles.tableSetting(styleId, "columnBandSize") ?? 0;
    for (const row of rows) {
      check(row, ["tblPrEx", "trPr", "tc"]);
      const exceptions = child(row, "tblPrEx");
      const exceptionCss = this.styles.rowException(exceptions);
      const rowProperties = child(row, "trPr");
      const rowCss = this.styles.row(rowProperties);
      const htmlRow = this.make("tr", {
        ...readWordRowGapWidths(rowProperties),
        style: declarations(rowCss),
        ...(child(rowProperties, "tblHeader")
          ? { "data-word-repeat-header": on(child(rowProperties, "tblHeader")) ? "true" : "false" }
          : {}),
      });
      if (Object.keys(exceptionCss).length) {
        const name = "word-row-exception-" + this.styles.css.length;
        this.styles.css.push("." + name + "{" + declarations(exceptionCss) + "}");
        htmlRow.attrs.class = name;
      }
      const rowMask =
        wordRowExceptionProperties(exceptionCss)["look-mask"] === undefined
          ? mask
          : Number(wordRowExceptionProperties(exceptionCss)["look-mask"]);
      const rowRegions = wordTableRowRegions(
        rowIndex,
        rows.length,
        grid.length,
        rowMask,
        rowBand,
        columnBand,
      );
      const rowStyle = resolveWordTableStyle(
          this.word,
          this.styles.readComponent,
          styleId,
          rowRegions,
        ),
        rowCell = wordTableRowCellProperties(rowStyle.cell);
      bindWordConditionalFormat(this.word, rowProperties, this.id(htmlRow));
      this.context(htmlRow, "row", rowStyle.row, tableInfo);
      for (const [kind, name] of [
        ["rowInsert", "ins"],
        ["rowDelete", "del"],
      ] as const) {
        const marker = child(rowProperties, name);
        if (marker) {
          check(marker, []);
          const id = attr(marker, "id");
          if (id === undefined) fail(marker, "A row revision requires an identity.");
          this.word.revisions.push({
            nodeId: this.id(htmlRow),
            sourceId: id,
            revisionId: this.id(htmlRow) + "-" + kind,
            kind,
            ...(this.moves.current(kind === "rowDelete" ? "from" : "to")
              ? { moveId: this.moves.current(kind === "rowDelete" ? "from" : "to")! }
              : {}),
            ...(attr(marker, "author") !== undefined ? { author: attr(marker, "author")! } : {}),
            ...(attr(marker, "date") !== undefined ? { date: attr(marker, "date")! } : {}),
          });
        }
      }
      readWordFormatRevision(this.word, rowProperties, htmlRow, "row", this.styles);
      readWordFormatRevision(this.word, exceptions, htmlRow, "tableException", this.styles);
      this.native?.bind(row, htmlRow.attrs);
      table.children.push(htmlRow);
      const before = readWordRowGapCount(rowProperties, "before");
      readWordRowGapCount(rowProperties, "after");
      const gap = (side: "before" | "after", start: number, count: number): Html =>
        this.make("td", {
          "data-word-grid-gap": side,
          ...(count > 1 ? { colspan: String(count) } : {}),
          style: "width:" + grid.slice(start, start + count).reduce((a, b) => a + b, 0) + "pt",
        });
      if (before) htmlRow.children.push(gap("before", 0, before));
      let column = before;
      const next = new Map<number, { cell: Html; span: number; css: Css }>();
      for (const cell of elements(row, "tc")) {
        const props = child(cell, "tcPr");
        if (child(props, "cellIns") && child(props, "cellDel"))
          fail(props!, "A native cell can carry only one insertion, deletion or merge mark.");
        const span = number(attr(child(props, "gridSpan"), "val"), cell, 1);
        if (span < 1) fail(cell, "A table cell exceeds its grid.");
        const merge = child(props, "vMerge");
        if (merge && (child(props, "cellIns") || child(props, "cellDel")))
          unsupported(
            source,
            "A cell revision in a vertical merge is preserved with its physical native cells.",
          );
        const value = attr(merge, "val") ?? "continue";
        const continuation = merge && value === "continue";
        if (merge && !["continue", "restart"].includes(value))
          fail(merge, "A vertical merge type is invalid.");
        let cellCss: Css;
        try {
          cellCss = this.styles.cell(props);
        } catch (error) {
          if (continuation && error instanceof DocumentError && error.code === "UnsupportedFeature")
            unsupported(cell, "A vertical merge continuation has independent native properties.");
          throw error;
        }
        const textDirection = attr(child(props, "textDirection"), "val");
        if (textDirection) delete cellCss["--word-cell-text-direction"];
        if (continuation) {
          const above = previous.get(column);
          if (!above || above.span !== span)
            fail(cell, "A vertical merge has no matching cell above.");
          try {
            const emptyParagraphs = this.validateMergeContinuation(cell),
              canonical = (css: Readonly<Css>) =>
                JSON.stringify(Object.entries(css).sort(([a], [b]) => a.localeCompare(b))),
              independent =
                canonical(above.css) !== canonical(cellCss) ||
                textDirection !== above.cell.attrs["data-word-text-direction"];
            let cellCssClass: string | undefined;
            if (independent) {
              cellCssClass = "word-continuation-cell-" + this.styles.css.length;
              this.styles.css.push("." + cellCssClass + "{" + declarations(cellCss) + "}");
            }
            const format = readWordConditionalFormat(props),
              paragraphs = elements(cell, "p").map(
                (paragraph) => readWordConditionalFormat(child(paragraph, "pPr")) ?? null,
              ),
              keepParagraphs = paragraphs.length > 1 || paragraphs.some(Boolean);
            if (independent || format || keepParagraphs || emptyParagraphs.length) {
              tableInfo.continuationFormats ??= [];
              tableInfo.continuationFormats.push({
                row: rowIndex,
                column,
                ...(cellCssClass ? { cellCssClass } : {}),
                ...(textDirection
                  ? { textDirection: textDirection as "lrTb" | "tbRl" | "btLr" }
                  : {}),
                ...(format ? { format } : {}),
                ...(keepParagraphs ? { paragraphs } : {}),
                ...(emptyParagraphs.length ? { emptyParagraphs } : {}),
              });
            }
          } catch (error) {
            if (error instanceof DocumentError && error.code === "UnsupportedFeature")
              unsupported(cell, "A vertical merge continuation has independent native properties.");
            throw error;
          }
          above.cell.attrs.rowspan = String(Number(above.cell.attrs.rowspan ?? 1) + 1);
          next.set(column, above);
        } else {
          const regions = wordTableRegions(
            rowIndex,
            column,
            rows.length,
            grid.length,
            span,
            rowMask,
            rowBand,
            columnBand,
          );
          const conditional = resolveWordTableStyle(
            this.word,
            this.styles.readComponent,
            styleId,
            regions,
          );
          const inner: Css = {};
          if (tableCss["--word-border-inside-h"] && table.children.length > 1)
            inner["border-top"] = "var(--word-border-inside-h)";
          if (tableCss["--word-border-inside-v"] && column > 0)
            inner["border-left"] = "var(--word-border-inside-v)";
          const htmlCell = this.make(on(child(rowProperties, "tblHeader")) ? "th" : "td", {
            style: declarations(cellCss),
            ...(span > 1 ? { colspan: String(span) } : {}),
          });
          bindWordConditionalFormat(this.word, props, this.id(htmlCell));
          this.context(
            htmlCell,
            "cell",
            {
              "padding-left": "5.4pt",
              "padding-right": "5.4pt",
              ...inner,
              ...wordRowExceptionCellProperties(
                wordRowExceptionProperties(exceptionCss),
                column,
                span,
                grid.length,
              ),
              ...style.cell,
              ...conditional.cell,
              ...rowCell,
              ...this.styles.margins(child(properties, "tblCellMar")),
              ...Object.fromEntries(
                Object.entries(
                  wordRowExceptionCellProperties(
                    wordRowExceptionProperties(exceptionCss),
                    column,
                    span,
                    grid.length,
                  ),
                ).filter(([key]) => key.startsWith("padding-")),
              ),
            },
            tableInfo,
          );
          const previousTable = this.activeTable,
            previousAvailable = this.tableAvailable;
          this.tableAvailable = grid.slice(column, column + span).reduce((a, b) => a + b, 0);
          this.activeTable = tableInfo;
          htmlCell.children.push(
            ...this.blocks(
              this.compat(cell).filter((item) => !(item.uri === W && item.name === "tcPr")),
              {
                p: { ...style.p, ...conditional.p },
                r: { ...style.r, ...conditional.r },
                runClasses: this.styles.runClass(styleId),
              },
            ),
          );
          this.activeTable = previousTable;
          this.tableAvailable = previousAvailable;
          if (textDirection) htmlCell.attrs["data-word-text-direction"] = textDirection;
          readWordFormatRevision(this.word, props, htmlCell, "cell", this.styles);
          for (const [kind, name] of [
            ["cellInsert", "cellIns"],
            ["cellDelete", "cellDel"],
          ] as const) {
            const marker = child(props, name);
            if (marker) {
              check(marker, []);
              const id = attr(marker, "id");
              if (id === undefined) fail(marker, "A cell revision requires an identity.");
              this.word.revisions.push({
                nodeId: this.id(htmlCell),
                sourceId: id,
                revisionId: this.id(htmlCell) + "-" + kind,
                kind,
                ...(attr(marker, "author") !== undefined
                  ? { author: attr(marker, "author")! }
                  : {}),
                ...(attr(marker, "date") !== undefined ? { date: attr(marker, "date")! } : {}),
              });
            }
          }
          this.native?.bind(cell, htmlCell.attrs);
          htmlRow.children.push(htmlCell);
          if (merge) next.set(column, { cell: htmlCell, span, css: cellCss });
        }
        column += span;
      }
      if (column < grid.length) htmlRow.children.push(gap("after", column, grid.length - column));
      previous = next;
      rowIndex++;
    }
    const gridChange = child(gridSource, "tblGridChange");
    if (gridChange) {
      check(gridChange, ["tblGrid"]);
      const before = child(gridChange, "tblGrid"),
        id = attr(gridChange, "id");
      if (!before || id === undefined)
        fail(gridChange, "A table grid revision requires its original grid and identity.");
      check(before, ["gridCol"]);
      const previous = elements(before, "gridCol").map((column) =>
        measurePoints(attr(column, "w"), column, 20, 0),
      );
      if (!previous.length || previous.some((width) => width <= 0))
        unsupported(gridChange, "An inferred original grid remains native.");
      const revision: WordMetadata["revisions"][number] = {
        nodeId: this.id(table),
        revisionId: this.id(table) + "-format-grid",
        sourceId: id,
        kind: "format",
        property: "tableGrid",
        before: { grid: previous },
        ...(attr(gridChange, "author") !== undefined
          ? { author: attr(gridChange, "author")! }
          : {}),
        ...(attr(gridChange, "date") !== undefined ? { date: attr(gridChange, "date")! } : {}),
      };
      this.word.revisions.push(revision);
      if (previous.length !== grid.length)
        bindWordGridInsertion(this.word, table, revision, gridChange, this.opc.limits.maxNodes);
    }
    const currentSpans = new Map(
      table.children.flatMap((row) =>
        typeof row === "string"
          ? []
          : row.children.flatMap((cell) =>
              typeof cell === "string"
                ? []
                : [[cell.attrs["data-node-id"]!, Number(cell.attrs.colspan ?? 1)] as const],
            ),
      ),
    );
    for (const revision of this.word.revisions)
      if (
        currentSpans.has(revision.nodeId) &&
        revision.property === "cell" &&
        revision.before?.attributes &&
        "colspan" in revision.before.attributes &&
        !revision.groupId &&
        currentSpans.get(revision.nodeId) !== Number(revision.before.attributes.colspan ?? 1)
      )
        unsupported(
          source,
          "A horizontal merge history without a complete native grid inverse remains native.",
        );
    const currentRows = new Map(
      table.children.flatMap((row, index) =>
        typeof row === "string"
          ? []
          : [[row.attrs["data-node-id"]!, child(rows[index], "trPr")] as const],
      ),
    );
    for (const revision of this.word.revisions)
      if (
        currentRows.has(revision.nodeId) &&
        revision.property === "row" &&
        revision.before?.rowOmissions &&
        !revision.groupId
      ) {
        const properties = currentRows.get(revision.nodeId);
        if (
          (["before", "after"] as const).some(
            (side) =>
              revision.before!.rowOmissions![side] !== readWordRowGapCount(properties, side),
          )
        )
          unsupported(
            source,
            "A row omission history without a complete native grid inverse remains native.",
          );
      }
    readWordFormatRevision(this.word, properties, table, "table", this.styles);
    this.native?.bind(source, table.attrs);
    return table;
  }
}

/** Import authorized DOCX bytes to the one persistent HTML document source. */
export function importDocx(bytes: Uint8Array, options: DocxImportOptions): DocumentPackage {
  if (!options || typeof options.documentId !== "string" || !options.documentId)
    throw new DocumentError(
      "InvalidDocument",
      "DOCX import requires a document identity supplied by its host.",
    );
  const opc = openOpc(bytes, options);
  const main = opc.related("", "officeDocument");
  if (
    !main ||
    opc.contentType(main) !==
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
  )
    throw new DocumentError("UnsupportedFormat", "The input is not a supported DOCX document.");
  if (
    options.unsupportedContent !== undefined &&
    !["preserve", "reject"].includes(options.unsupportedContent)
  )
    throw new DocumentError("InvalidDocument", "Unknown DOCX unsupported-content policy.");
  try {
    if (options.unsupportedContent === "reject")
      return new Importer(opc, main, options.documentId).import();
    const native = new DocxNativeImport(opc, main),
      view = native.initialize();
    for (;;) {
      native.beginAttempt();
      try {
        return new Importer(view, main, options.documentId, native).import();
      } catch (error) {
        if (!native.recover(error)) throw error;
      }
    }
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    throw new DocumentError(
      "InvalidDocument",
      "The DOCX document could not be mapped to the supported document model.",
    );
  }
}

export interface DocxImportReport {
  readonly documentId: string;
  readonly preserved: readonly {
    readonly nativeId: string;
    readonly nodeId?: string;
    readonly kind: "inline" | "block" | "property";
    readonly label: string;
  }[];
}
export interface DocxImportResult {
  readonly files: DocumentPackage;
  readonly report: DocxImportReport;
}
/** Same importer, with a bounded capability report derived from the saved source. */
export function importDocxWithReport(
  bytes: Uint8Array,
  options: DocxImportOptions,
): DocxImportResult {
  const files = importDocx(bytes, options),
    model = openDocument(files, {
      kind: "document",
      ...(options.limits ? { limits: options.limits } : {}),
    });
  const preserved = (model.snapshot().metadata.word?.nativeContent?.fragments ?? []).map(
    ({ nativeId, nodeId, kind, label }) =>
      Object.freeze({ nativeId, ...(nodeId ? { nodeId } : {}), kind, label }),
  );
  return Object.freeze({
    files,
    report: Object.freeze({ documentId: options.documentId, preserved: Object.freeze(preserved) }),
  });
}
