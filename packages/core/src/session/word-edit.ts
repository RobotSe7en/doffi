import { pasteWordBlockContent } from "./word-block-content.js";
import { pasteWordInlineContent } from "./word-inline-content.js";
import { restoreWordSplitContinuations } from "./word-split-continuation.js";
import { formatWordParagraphMark } from "./word-paragraph-mark.js";
import {
  trackWordColumnDeletion,
  prepareWordColumnInsertion,
  wordTableCells,
} from "./word-cell-revisions.js";
import {
  wordResetTextProperties,
  wordRunCharacterStyle,
  wordRunCharacterStyleAt,
} from "./word-reset-text.js";
import { ensureWordMainStoryContainer, relocateWordSectionRoot } from "./word-story-container.js";
import { moveWordContinuationFormats } from "./word-conditional-format.js";
import { wordContainerWidth } from "../document/word-container.js";
import { length } from "../document/css-values.js";
import type { SourceStyleNode } from "../document/source-css.js";
import { trackWordRowDeletion, wordTableRows } from "./word-row-revisions.js";
import { applyWordWatermark } from "./word-watermark.js";
import { recordWordFormat, prepareWordRunFormat } from "./word-format-revisions.js";
import { editWordStyle } from "./word-styles.js";
import { applyWordRevision, isWordRevisionOperation } from "./word-revisions.js";
import { editSourceTable } from "./table-edit.js";
import { copyWordStorySemantics } from "./word-story-copy.js";
import { wordTextDefaults, clearCoveredWordRuns } from "../document/word-format.js";
import { applyWordDrawing, isWordDrawingOperation } from "./word-drawing-edit.js";
import { applyWordReference, isWordReferenceOperation } from "./word-reference-edit.js";
import { applyWordMath, isWordMathOperation } from "./word-math-edit.js";
import { applyWordControl, isWordControlOperation } from "./word-control-edit.js";
import { applyWordComment, isWordCommentOperation } from "./word-comments.js";
import type { Annotation } from "../contracts/annotations.js";
import { checkLimit } from "../document/validation.js";
import type { DocumentReadLimits } from "../runtime.js";
import { DocumentError } from "../errors.js";
import type { WordOperation } from "../contracts/word-operations.js";
import type { WordDocumentMetadata } from "../contracts/document.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import type { SourceStyleEdit } from "./source-styles.js";
function fail(message: string): never {
  throw new DocumentError("InvalidOperation", message);
}
export function isWordOperation(operation: { readonly op: string }): operation is WordOperation {
  return (
    isWordRevisionOperation(operation) ||
    isWordCommentOperation(operation) ||
    isWordDrawingOperation(operation) ||
    isWordReferenceOperation(operation) ||
    isWordMathOperation(operation) ||
    isWordControlOperation(operation) ||
    [
      "pasteWordBlockContent",
      "pasteWordInlineContent",
      "setWordWatermark",
      "resetWordText",
      "cloneWordStory",
      "setWordOptions",
      "setWordTabs",
      "insertWordField",
      "setWordSection",
      "splitWordSection",
      "insertWordStory",
      "setWordStoryContent",
      "insertWordTable",
      "deleteWordTable",
      "editWordTable",
      "formatWordTable",
      "formatWordCell",
      "formatWordText",
      "formatWordParagraphMark",
      "applyWordCharacterStyle",
      "setWordParagraph",
      "setWordStyle",
      "deleteWordStyle",
      "setWordNumbering",
      "setWordList",
    ].includes(operation.op)
  );
}
export function applyWordOperation(
  metadata: WordDocumentMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  styles: SourceStyleEdit,
  operation: WordOperation,
  limits: DocumentReadLimits,
  annotations: Annotation[] = [],
): void {
  if (!metadata.word) {
    const source = parse(),
      rootId = source.nodeIds().find((id) => source.peekNode(id)!.parentNodeId === null)!;
    metadata.word = {
      stories: [{ storyId: "main", kind: "main", rootNodeId: rootId }],
      styles: [],
      numbering: [],
      hyperlinks: [],
      bookmarks: [],
      fields: [],
      tabs: [],
      revisions: [],
      tables: [],
    };
    edit.attributes(rootId, { "data-office-story": "main" });
  }
  const word = metadata.word;
  const node = parse().peekNode(operation.nodeId);
  if (!node) fail("The Word source target is missing.");
  if (operation.op === "formatWordParagraphMark") {
    formatWordParagraphMark(word, edit, parse, operation.nodeId, operation.properties);
    return;
  }
  if (operation.op === "setWordWatermark") {
    applyWordWatermark(metadata, edit, parse, operation, (operation) =>
      applyWordOperation(metadata, edit, parse, styles, operation, limits, annotations),
    );
    return;
  }
  if (isWordRevisionOperation(operation)) {
    applyWordRevision(word, edit, parse, operation, limits, metadata.sections, styles);
    return;
  }
  if (isWordCommentOperation(operation)) {
    applyWordComment(metadata, annotations, edit, parse, operation);
    return;
  }
  if (isWordDrawingOperation(operation)) {
    applyWordDrawing(metadata, edit, parse, operation, limits);
    return;
  }
  if (isWordReferenceOperation(operation)) {
    if (
      operation.op === "setWordNoteOptions" &&
      node.parentNodeId === null &&
      Object.keys(operation.options).length
    ) {
      const key = operation.kind === "footnote" ? "footnoteOptions" : "endnoteOptions";
      for (const section of metadata.sections) {
        const options = { ...section[key], ...operation.options };
        if (JSON.stringify(options) !== JSON.stringify(section[key] ?? {}))
          applyWordOperation(
            metadata,
            edit,
            parse,
            styles,
            { op: "setWordSection", nodeId: section.rootNodeId, changes: { [key]: options } },
            limits,
            annotations,
          );
      }
    }
    applyWordReference(metadata, edit, parse, operation);
    return;
  }
  if (isWordMathOperation(operation)) {
    applyWordMath(word, edit, parse, operation);
    return;
  }
  if (isWordControlOperation(operation)) {
    applyWordControl(word, edit, parse, operation);
    return;
  }
  if (operation.op === "setWordStyle" || operation.op === "deleteWordStyle") {
    editWordStyle(word, edit, parse, styles, operation, limits);
    return;
  }
  if (operation.op === "resetWordText") {
    if (!/^p$|^h[1-6]$/.test(node.tagName))
      fail("Resetting text formatting requires a paragraph range.");
    if (operation.resetParagraph) {
      if (operation.from !== 0 || operation.to !== parse().getNode(operation.nodeId)!.text.length)
        fail("Paragraph formatting can only be reset for a complete paragraph.");
      const normal = word.styles.find(
        (style) => style.kind === "paragraph" && style.styleId === "Normal",
      );
      applyWordOperation(
        metadata,
        edit,
        parse,
        styles,
        {
          op: "setWordParagraph",
          nodeId: operation.nodeId,
          clearDirectFormatting: true,
          ...(normal ? { styleId: normal.styleId } : {}),
        },
        limits,
      );
    }
    const range = word.trackChanges
        ? prepareWordRunFormat(word, edit, parse, styles, operation)
        : { ...operation, preferred: new Set<string>() },
      source = parse(),
      defaults = wordTextDefaults(
        word,
        source,
        operation.nodeId,
        styles.resolver(),
        limits.maxNodes * 8,
      ),
      changed: string[] = [];
    let groupId: string | undefined;
    edit.formatSourceText(
      range.nodeId,
      range,
      wordResetTextProperties(defaults),
      (change) => {
        changed.push(change.nodeId);
        if (!word.trackChanges) return;
        const previous = word.revisions.find(
            (item) =>
              item.nodeId === change.nodeId && item.kind === "format" && item.property === "run",
          ),
          id = previous?.revisionId ?? change.nodeId + "-format";
        groupId ??= id;
        if (previous) {
          previous.author = word.trackingAuthor ?? "";
          previous.date = new Date().toISOString();
          previous.groupId = groupId;
          const before = previous.before!,
            attributes = (before.attributes ??= {});
          if (attributes["data-word-current-format"] === undefined)
            attributes["data-word-current-format"] = null;
          if (attributes["data-word-reset-text"] === undefined)
            attributes["data-word-reset-text"] =
              source.peekNode(change.nodeId)?.attributes["data-word-reset-text"] ?? null;
          if (before.characterStyleId === undefined)
            before.characterStyleId =
              wordRunCharacterStyle(word, source, change.nodeId, attributes) ?? null;
        } else {
          const current = source.peekNode(change.nodeId),
            character = current
              ? wordRunCharacterStyle(word, source, change.nodeId)
              : wordRunCharacterStyleAt(word, source, range.nodeId, change.from);
          word.revisions.push({
            nodeId: change.nodeId,
            sourceId: id,
            revisionId: id,
            groupId,
            kind: "format",
            property: "run",
            author: word.trackingAuthor ?? "",
            date: new Date().toISOString(),
            before: {
              inheritNativeProperties: true,
              characterStyleId: character ?? null,
              attributes: {
                style: change.previousStyle ?? null,
                "data-word-reset-text": current?.attributes["data-word-reset-text"] ?? null,
                "data-word-current-format": current?.attributes["data-word-current-format"] ?? null,
              },
            },
          });
        }
      },
      range.preferred,
      true,
      true,
    );
    for (const id of changed) {
      const current = parse().peekNode(id)!,
        character = word.styles.find(
          (style) =>
            style.styleId === current.attributes["data-word-style"] && style.kind === "character",
        );
      if (character) {
        const revision = word.revisions.find(
          (item) => item.nodeId === id && item.property === "run" && item.kind === "format",
        );
        if (word.trackChanges && revision) {
          const attributes = (revision.before!.attributes ??= {});
          for (const key of ["data-word-style", "class"] as const)
            if (attributes[key] === undefined)
              attributes[key] = source.peekNode(id)?.attributes[key] ?? null;
        }
        edit.attributes(id, {
          "data-word-style": undefined,
          class: (current.attributes.class ?? "")
            .split(/\s+/)
            .filter((name) => name && name !== character.cssClass && name !== character.runCssClass)
            .join(" "),
        });
      }
      edit.attributes(id, { "data-word-reset-text": "true", "data-word-current-format": "true" });
    }
    if (!word.trackChanges)
      clearCoveredWordRuns(word, parse(), edit, operation.nodeId, operation.from, operation.to);
    return;
  }
  if (operation.op === "setWordOptions") {
    if (node.parentNodeId !== null) fail("Shared Word settings require the document root.");
    if (operation.mirrorMargins !== undefined) {
      word.mirrorMargins = operation.mirrorMargins;
      if (word.nativeContent)
        word.nativeContent.fragments = word.nativeContent.fragments.filter(
          (fragment) =>
            !(
              fragment.ownerName === "settings" &&
              fragment.propertyName ===
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main mirrorMargins"
            ),
        );
    }
    if (operation.evenAndOddHeaders !== undefined)
      word.evenAndOddHeaders = operation.evenAndOddHeaders;
    for (const key of ["noColumnBalance", "adjustLineHeightInTable"] as const)
      if (operation[key] !== undefined) word[key] = operation[key];
    if (operation.automaticHyphenation !== undefined)
      word.automaticHyphenation = operation.automaticHyphenation;
    for (const key of ["bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter"] as const)
      if (operation[key] !== undefined) word[key] = operation[key];
    return;
  }
  if (operation.op === "setWordTabs") {
    recordWordFormat(word, parse(), operation.nodeId, "paragraph");
    if (!/^p$|^h[1-6]$/.test(node.tagName)) fail("Tab stops require a paragraph.");
    const definition = {
        paragraphNodeId: operation.nodeId,
        defaultInterval: operation.defaultInterval,
        stops: operation.stops,
      },
      index = word.tabs.findIndex((item) => item.paragraphNodeId === operation.nodeId);
    if (index < 0) word.tabs.push(definition);
    else word.tabs[index] = definition;
    return;
  }
  if (operation.op === "insertWordField") {
    if (!/^p$|^h[1-6]$/.test(node.tagName)) fail("A field insertion requires a paragraph.");
    const cached = operation.cachedText ?? "1",
      escape = (value: string): string =>
        value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
    const source = '<span data-word-field="simple">' + escape(cached) + "</span>",
      id = (
        operation.offset === undefined
          ? edit.insertSource(operation.nodeId, source)
          : edit.insertSourceAt(operation.nodeId, operation.offset, source)
      ).rootNodeIds[0]!;
    word.fields.push({
      nodeId: id,
      instruction: operation.instruction,
      resultState: "cached",
      resultNodeIds: [id],
    });
    return;
  }
  if (operation.op === "setWordSection") {
    const section = metadata.sections.find((section) => section.rootNodeId === operation.nodeId);
    if (!section) fail("Page changes require a section root.");
    if (word.trackChanges) {
      const previous = word.revisions.find(
        (item) =>
          item.nodeId === operation.nodeId && item.kind === "format" && item.property === "section",
      );
      if (previous) {
        previous.author = word.trackingAuthor ?? "";
        previous.date = new Date().toISOString();
      } else {
        const before = JSON.parse(JSON.stringify(section)) as Partial<typeof section>;
        delete before.sectionId;
        delete before.rootNodeId;
        const id = operation.nodeId + "-format-section";
        word.revisions.push({
          nodeId: operation.nodeId,
          sourceId: id,
          revisionId: id,
          kind: "format",
          property: "section",
          before: { section: before, inheritNativeProperties: true },
          author: word.trackingAuthor ?? "",
          date: new Date().toISOString(),
        });
      }
    }
    if (operation.changes.pageBorders && word.nativeContent) {
      const revision = word.revisions.find(
        (item) =>
          item.nodeId === operation.nodeId && item.property === "section" && item.kind === "format",
      );
      word.nativeContent.fragments = word.nativeContent.fragments.flatMap((fragment) => {
        if (
          fragment.nodeId !== section.rootNodeId ||
          fragment.revisionPrevious ||
          !(
            fragment.propertyName?.endsWith(" pgBorders") ||
            fragment.propertyPath.some((name) => name.endsWith(" pgBorders"))
          )
        )
          return [fragment];
        if (!word.trackChanges || !revision) return [];
        return [
          {
            ...fragment,
            propertyPath: [
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main sectPrChange",
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main sectPr",
              ...fragment.propertyPath,
            ],
            ...(fragment.propertyIndices
              ? { propertyIndices: [0, 0, ...fragment.propertyIndices] }
              : {}),
            revisionPrevious: true,
            revisionSourceId: revision.sourceId,
          },
        ];
      });
    }
    Object.assign(section, operation.changes);
    return;
  }
  if (operation.op === "splitWordSection") {
    if (!/^p$|^h[1-6]$/.test(node.tagName)) fail("A section break requires a paragraph.");
    const index = metadata.sections.findIndex(
        (section) => section.rootNodeId === node.parentNodeId,
      ),
      section = metadata.sections[index];
    if (!section) fail("A section break requires a paragraph directly in its section.");
    ensureWordMainStoryContainer(metadata, edit, parse);
    const main = word.stories.find((story) => story.kind === "main")!;
    if (section.rootNodeId === main.rootNodeId) {
      const children = [...parse().peekNode(main.rootNodeId)!.childNodeIds],
        wrapper = edit.insertSource(main.rootNodeId, "<section></section>").rootNodeIds[0]!;
      edit.moveSource(children, wrapper);
      relocateWordSectionRoot(metadata, section.rootNodeId, wrapper);
    }
    const html = parse(),
      siblings = html.peekNode(main.rootNodeId)!.childNodeIds,
      ordinal = siblings.indexOf(section.rootNodeId);
    const rootId = edit.insertSource(main.rootNodeId, "<section></section>", siblings[ordinal + 1])
      .rootNodeIds[0]!;
    const children = html.peekNode(section.rootNodeId)!.childNodeIds,
      tail = children.slice(children.indexOf(operation.nodeId) + 1);
    if (tail.length) edit.moveSource(tail, rootId);
    else edit.insertSource(rootId, "<p></p>");
    const next = {
      ...section,
      sectionId: rootId,
      rootNodeId: rootId,
      breakType: operation.breakType,
    };
    delete next.pageNumberStart;
    metadata.sections.splice(index + 1, 0, next);
    return;
  }
  if (operation.op === "insertWordStory") {
    const section = metadata.sections.find((section) => section.rootNodeId === operation.nodeId);
    if (!section) fail("Header/footer creation requires a section.");
    const rootId = ensureWordMainStoryContainer(metadata, edit, parse);
    const storyRoot = edit.insertSource(
      rootId,
      '<div data-office-story="' + operation.kind + '">' + operation.html + "</div>",
    ).rootNodeIds[0]!;
    word.stories.push({ storyId: storyRoot, kind: operation.kind, rootNodeId: storyRoot });
    const normal = word.styles.find(
      (style) => style.kind === "paragraph" && style.styleId === "Normal",
    );
    if (normal) {
      const tree = parse(),
        visit = (id: string): void => {
          const paragraph = tree.peekNode(id)!;
          if (paragraph.tagName === "p" && !paragraph.attributes["data-word-style"])
            edit.attributes(id, {
              "data-word-style": normal.styleId,
              class: [paragraph.attributes.class, normal.cssClass].filter(Boolean).join(" "),
            });
          for (const child of paragraph.childNodeIds) visit(child);
        };
      visit(storyRoot);
    }
    const key = operation.kind === "header" ? "headers" : "footers";
    section[key] = { ...section[key], [operation.variant]: storyRoot };
    if (operation.variant === "first") section.differentFirstPage = true;
    if (operation.variant === "even") word.evenAndOddHeaders = true;
    return;
  }
  if (operation.op === "cloneWordStory") {
    const section = metadata.sections.find((section) => section.rootNodeId === operation.nodeId),
      story = word.stories.find((story) => story.storyId === operation.storyId);
    if (!section || !story || !["header", "footer"].includes(story.kind))
      fail("Copying a header/footer requires its source story and target section.");
    const source = parse(),
      article = source.nodeIds().find((id) => source.peekNode(id)!.parentNodeId === null)!,
      clone = edit.cloneSource(story.rootNodeId, article),
      rootId = clone.rootNodeIds[0]!;
    copyWordStorySemantics(word, clone.identities, edit, source, rootId);
    word.stories.push({ ...story, storyId: rootId, rootNodeId: rootId });
    const key = story.kind === "header" ? "headers" : "footers";
    section[key] = { ...section[key], [operation.variant]: rootId };
    if (operation.variant === "first") section.differentFirstPage = true;
    if (operation.variant === "even") word.evenAndOddHeaders = true;
    return;
  }
  if (operation.op === "setWordStoryContent") {
    if (
      !word.stories.some((story) => story.rootNodeId === operation.nodeId && story.kind !== "main")
    )
      fail("Auxiliary story editing requires an existing header, footer or comment story.");
    edit.replaceSourceContent(operation.nodeId, operation.html);
    return;
  }
  if (operation.op === "pasteWordBlockContent") {
    pasteWordBlockContent(metadata, edit, parse, styles, operation);
    return;
  }
  if (operation.op === "pasteWordInlineContent") {
    pasteWordInlineContent(metadata, edit, parse, operation);
    return;
  }
  if (operation.op === "deleteWordTable") {
    if (node.tagName !== "table") fail("Table deletion requires a table.");
    const rows = wordTableRows(parse(), operation.nodeId);
    if (word.trackChanges && rows.length)
      trackWordRowDeletion(word, edit, parse, operation.nodeId, 0, rows.length, limits);
    else edit.removeSource(operation.nodeId);
    return;
  }
  if (operation.op === "insertWordTable") {
    if (!/^p$|^h[1-6]$/.test(node.tagName)) fail("Insert a table beside an existing paragraph.");
    checkLimit(operation.rows * operation.columns * 3, limits.maxNodes);
    const width =
      operation.width ??
      wordContainerWidth(metadata, parse(), operation.nodeId, styles.resolver(), limits);
    const tableStyle = word.styles.find(
      (style) => style.kind === "table" && style.styleId === "TableGrid" && style.components,
    );
    const escapeStyle = (value: string) =>
      value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
    const source =
      "<table" +
      (tableStyle
        ? ' data-word-style="' +
          escapeStyle(tableStyle.styleId) +
          '" class="' +
          escapeStyle(tableStyle.cssClass) +
          '"'
        : "") +
      ' style="border-collapse:collapse;table-layout:fixed;width:' +
      width +
      'pt">' +
      Array.from(
        { length: operation.rows },
        () =>
          "<tr>" +
          Array.from(
            { length: operation.columns },
            () =>
              "<td" +
              (tableStyle
                ? ""
                : ' style="border:0.5pt solid #000000;padding:4pt;vertical-align:top"') +
              "><p></p></td>",
          ).join("") +
          "</tr>",
      ).join("") +
      "</table>";

    const parent = edit.sourceParent(operation.nodeId),
      siblings = parse().peekNode(parent)!.childNodeIds,
      index = siblings.indexOf(operation.nodeId);
    const reference = operation.position === "before" ? operation.nodeId : siblings[index + 1];
    const tableId = edit.insertSource(parent, source, reference).rootNodeIds[0]!;
    word.tables.push({
      nodeId: tableId,
      grid: Array.from({ length: operation.columns }, () => width / operation.columns),
    });
    if (word.trackChanges) {
      const author = word.trackingAuthor ?? "",
        date = new Date().toISOString(),
        groupId = tableId + "-table-insert";
      for (const id of wordTableRows(parse(), tableId))
        word.revisions.push({
          nodeId: id,
          sourceId: id + "-row-insert",
          revisionId: id + "-row-insert",
          groupId,
          kind: "rowInsert",
          author,
          date,
        });
    }
    return;
  }
  if (operation.op === "editWordTable" || operation.op === "formatWordTable") {
    if (node.tagName !== "table") fail("Word table commands require a table.");
    let info = word.tables.find((table) => table.nodeId === operation.nodeId);
    const source = parse(),
      rowIds = node.childNodeIds.flatMap((id) => {
        const row = source.peekNode(id)!;
        return row.tagName === "tr" ? [id] : row.childNodeIds;
      });
    if (!info) {
      const occupied: number[][] = Array.from({ length: rowIds.length }, () => []);
      let columns = 0;
      for (const [row, id] of rowIds.entries())
        for (const cellId of source.peekNode(id)!.childNodeIds) {
          const cell = source.peekNode(cellId)!;
          let column = 0;
          while (occupied[row]![column]) column++;
          const width = Number(cell.attributes.colspan ?? 1),
            height = Number(cell.attributes.rowspan ?? 1) || rowIds.length - row;
          checkLimit((column + width) * rowIds.length, limits.maxNodes);
          if (row + height > rowIds.length) fail("A merged cell exceeds the table.");
          for (let r = row; r < row + height; r++)
            for (let c = column; c < column + width; c++) occupied[r]![c] = 1;
          columns = Math.max(columns, column + width);
        }
      if (!columns) fail("An editable table requires a nonempty grid.");
      const resolver = styles.resolver(),
        project = (id: string): SourceStyleNode => {
          const node = source.peekNode(id)!;
          return {
            tag: node.tagName,
            attrs: node.attributes,
            parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
          };
        },
        css = resolver.get(project(operation.nodeId)),
        available = wordContainerWidth(metadata, source, operation.nodeId, resolver, limits),
        width =
          css.width && css.width !== "auto"
            ? length(css.width, { font: 12, percent: available })
            : available;
      info = {
        nodeId: operation.nodeId,
        grid: Array.from({ length: columns }, () => width / columns),
      };
      word.tables.push(info);
    }
    if (operation.op === "editWordTable") {
      if (operation.change.kind === "mergeCells" || operation.change.kind === "splitCell") {
        const change = operation.change,
          projection = wordTableCells(source, operation.nodeId, info.grid.length, limits),
          refine =
            change.kind === "splitCell" &&
            (change.rows !== undefined || change.columns !== undefined),
          selected = projection.cells.filter((cell) =>
            change.kind === "mergeCells"
              ? cell.row < change.row + change.rows &&
                cell.row + cell.rows > change.row &&
                cell.column < change.column + change.columns &&
                cell.column + cell.columns > change.column
              : cell.row <= change.row &&
                cell.row + cell.rows > change.row &&
                cell.column <= change.column &&
                cell.column + cell.columns > change.column,
          ),
          cells = new Set((refine ? projection.cells : selected).map((cell) => cell.id)),
          rows = new Set(refine ? projection.rows : []);
        if (!refine)
          for (const cell of selected)
            for (let row = cell.row; row < cell.row + cell.rows; row++)
              rows.add(projection.rows[row]!);
        const inserted = new Map(
            word.revisions
              .filter((revision) => revision.kind === "rowInsert")
              .map((revision) => [revision.nodeId, revision]),
          ),
          group = inserted.get(projection.rows[0]!)?.groupId,
          newTable =
            !!group &&
            projection.rows.every(
              (id) =>
                inserted.get(id)?.groupId === group &&
                inserted.get(id)?.author === (word.trackingAuthor ?? ""),
            );
        if (
          word.revisions.some(
            (revision) =>
              (revision.nodeId === operation.nodeId && revision.property === "tableGrid") ||
              (cells.has(revision.nodeId) &&
                (revision.kind.startsWith("cell") || revision.property === "cell")) ||
              (rows.has(revision.nodeId) &&
                revision.kind.startsWith("row") &&
                !(newTable && revision.kind === "rowInsert")),
          )
        )
          fail(
            "Review the existing table grid, cell formatting or structural revisions before merging or splitting these cells.",
          );
      }
      if (word.trackChanges && operation.change.kind === "resize" && operation.change.grid) {
        const previous = word.revisions.find(
          (item) =>
            item.nodeId === operation.nodeId &&
            item.kind === "format" &&
            item.property === "tableGrid",
        );
        if (previous) {
          previous.author = word.trackingAuthor ?? "";
          previous.date = new Date().toISOString();
        } else {
          const id = operation.nodeId + "-format-grid";
          word.revisions.push({
            nodeId: operation.nodeId,
            sourceId: id,
            revisionId: id,
            kind: "format",
            property: "tableGrid",
            before: { grid: [...info.grid], inheritNativeProperties: true },
            author: word.trackingAuthor ?? "",
            date: new Date().toISOString(),
          });
        }
        recordWordFormat(word, source, operation.nodeId, "table");
        for (const id of rowIds.flatMap((id) => source.peekNode(id)!.childNodeIds))
          if (
            ["td", "th"].includes(source.peekNode(id)!.tagName) &&
            !source.peekNode(id)!.attributes["data-word-grid-gap"]
          )
            recordWordFormat(word, source, id, "cell");
      }
      if (word.trackChanges && operation.change.kind === "resize" && operation.change.grid)
        for (const id of rowIds) recordWordFormat(word, source, id, "row");
      if (word.trackChanges && operation.change.kind === "resize" && operation.change.rowHeights)
        for (const id of rowIds) recordWordFormat(word, source, id, "row");
      if (
        word.trackChanges &&
        operation.change.kind === "resize" &&
        operation.change.rowIndex !== undefined
      ) {
        const id = rowIds[operation.change.rowIndex];
        if (!id) fail("The resized row exceeds the table.");
        recordWordFormat(word, source, id, "row");
      }
      if (word.trackChanges && operation.change.kind === "deleteRows") {
        trackWordRowDeletion(
          word,
          edit,
          parse,
          operation.nodeId,
          operation.change.at,
          operation.change.count,
          limits,
        );
        return;
      }
      if (word.trackChanges && operation.change.kind === "deleteColumns") {
        trackWordColumnDeletion(
          word,
          edit,
          parse,
          operation.nodeId,
          operation.change.at,
          operation.change.count,
          limits,
        );
        const grid = word.revisions.find(
          (revision) => revision.nodeId === operation.nodeId && revision.property === "tableGrid",
        );
        if (
          grid?.groupId &&
          grid.before?.grid?.length === info.grid.length &&
          !word.revisions.some(
            (revision) =>
              revision.groupId === grid.groupId &&
              revision.kind === "cellInsert" &&
              parse().hasNode(revision.nodeId),
          )
        )
          applyWordRevision(
            word,
            edit,
            parse,
            {
              op: "reviewWordRevision",
              nodeId: grid.nodeId,
              revisionId: grid.revisionId ?? grid.nodeId,
              decision: "reject",
            },
            limits,
            metadata.sections,
            styles,
          );
        return;
      }
      const insertion =
        word.trackChanges && operation.change.kind === "insertColumns"
          ? prepareWordColumnInsertion(word, source, operation.nodeId, limits)
          : undefined;
      const previousCells = insertion?.cells;
      const firstCell = rowIds
        .flatMap((id) => source.peekNode(id)!.childNodeIds)
        .map((id) => source.peekNode(id)!)
        .find(
          (cell) => ["td", "th"].includes(cell.tagName) && !cell.attributes["data-word-grid-gap"],
        );
      const escape = (value: string): string =>
        value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
      const cellAttributes = firstCell
        ? ["class", "style"]
            .filter((key) => firstCell.attributes[key])
            .map((key) => " " + key + '="' + escape(firstCell.attributes[key]!) + '"')
            .join("")
        : "";
      const splitContinuations =
          operation.change.kind === "splitCell" ? info.continuationFormats : undefined,
        previousGrid = [...info.grid];
      moveWordContinuationFormats(info, operation.change);
      const previousWidth = info.grid.reduce((a, b) => a + b, 0);
      const result = editSourceTable(
        {
          edit,
          parse,
          limits,
          wordOmissions: true,
          newCellHtml: "<td" + cellAttributes + "><p></p></td>",
        },
        operation.nodeId,
        info.grid,
        { width: previousWidth, height: Math.max(1, rowIds.length) * 20 },
        operation.change,
      );
      if (previousCells) {
        const old = new Set(previousCells.map((cell) => cell.id)),
          added = result.cells.filter((cell) => !old.has(cell.id)),
          author = word.trackingAuthor ?? "",
          date = new Date().toISOString(),
          groupId = insertion!.groupId;
        if (!added.length)
          throw new DocumentError(
            "UnsupportedFeature",
            "An insertion wholly inside merged or omitted regions requires explicit native column boundaries.",
          );
        for (const cell of added)
          word.revisions.push({
            nodeId: cell.id,
            sourceId: cell.id + "-cell-insert",
            revisionId: cell.id + "-cell-insert",
            groupId,
            kind: "cellInsert",
            author,
            date,
          });
      }
      if (word.trackChanges && operation.change.kind === "insertRows") {
        const added = wordTableRows(parse(), operation.nodeId).filter((id) => !rowIds.includes(id)),
          groupId = added[0] + "-row-insert";
        for (const id of added)
          word.revisions.push({
            nodeId: id,
            sourceId: id + "-row-insert",
            revisionId: id + "-row-insert",
            groupId,
            kind: "rowInsert",
            author: word.trackingAuthor ?? "",
            date: new Date().toISOString(),
          });
      }
      const preserveWidth =
        operation.change.kind === "insertColumns" || operation.change.kind === "deleteColumns";
      info.grid = preserveWidth
        ? result.grid.map((width) => (width * previousWidth) / result.width)
        : result.grid;
      if (
        operation.change.kind === "insertColumns" ||
        operation.change.kind === "deleteColumns" ||
        (operation.change.kind === "resize" && operation.change.grid) ||
        (operation.change.kind === "splitCell" && info.grid.length !== previousGrid.length)
      ) {
        const current = parse(),
          rows = wordTableRows(current, operation.nodeId);
        for (const cell of result.cells) {
          const side = current.peekNode(cell.id)!.attributes["data-word-grid-gap"];
          if (side !== "before" && side !== "after") continue;
          const value = String(
              Math.round(
                info.grid
                  .slice(cell.column, cell.column + cell.columns)
                  .reduce((a, b) => a + b, 0) * 20,
              ),
            ),
            prefix = "data-word-" + side + "-width-";
          for (let row = cell.row; row < cell.row + cell.rows; row++)
            edit.attributes(rows[row]!, {
              [prefix + "present"]: "true",
              [prefix + "type"]: "dxa",
              [prefix + "value"]: value,
            });
        }
      }
      if (splitContinuations)
        restoreWordSplitContinuations(
          word,
          info,
          source,
          rowIds,
          previousGrid,
          splitContinuations,
          result.cells,
          edit,
          parse(),
          limits,
          styles,
        );
      if (operation.change.kind === "resize" && operation.change.grid)
        edit.sourceStyle(operation.nodeId, "width", result.width + "pt");
      if (
        preserveWidth ||
        (operation.change.kind === "resize" && operation.change.grid) ||
        operation.change.kind === "splitCell"
      )
        for (const cell of result.cells)
          edit.sourceStyle(
            cell.id,
            "width",
            info.grid.slice(cell.column, cell.column + cell.columns).reduce((a, b) => a + b, 0) +
              "pt",
          );
      return;
    }
    if (
      operation.styleId !== undefined ||
      operation.lookMask !== undefined ||
      operation.rowBandSize !== undefined ||
      operation.columnBandSize !== undefined ||
      Object.keys(operation.properties ?? {}).length
    )
      recordWordFormat(word, parse(), operation.nodeId, "table");
    if (
      operation.repeatHeaderRows !== undefined ||
      operation.allowRowSplit !== undefined ||
      operation.rowHeightRule !== undefined
    )
      for (const id of rowIds) recordWordFormat(word, parse(), id, "row");
    if (operation.rowCellSpacing !== undefined) {
      const selected =
        operation.rowIndex === undefined
          ? rowIds
          : [rowIds[operation.rowIndex] ?? fail("The spacing row exceeds the table.")];
      for (const id of selected) {
        recordWordFormat(word, parse(), id, "row");
        for (const [key, value] of Object.entries({
          "--word-row-cell-spacing":
            operation.rowCellSpacing === null ? undefined : operation.rowCellSpacing + "pt",
          "--word-row-cell-spacing-type": operation.rowCellSpacing === null ? undefined : "dxa",
          "--word-row-cell-spacing-value":
            operation.rowCellSpacing === null
              ? undefined
              : String(Math.round(operation.rowCellSpacing * 20)),
        }))
          edit.sourceStyle(id, key, value);
      }
    }
    if (operation.styleId !== undefined) {
      const style =
        operation.styleId === null
          ? undefined
          : word.styles.find(
              (style) => style.styleId === operation.styleId && style.kind === "table",
            );
      if (operation.styleId !== null && !style) fail("The native table style is missing.");
      const old = word.styles.find((style) => style.styleId === node.attributes["data-word-style"]);
      const classes = (node.attributes.class ?? "")
        .split(/\s+/)
        .filter((value) => value && value !== old?.cssClass);
      if (style) classes.push(style.cssClass);
      edit.attributes(operation.nodeId, {
        "data-word-style": operation.styleId ?? undefined,
        class: classes.join(" "),
      });
    }
    for (const key of ["lookMask", "rowBandSize", "columnBandSize"] as const) {
      const value = operation[key];
      if (value === null) delete info[key];
      else if (value !== undefined) info[key] = value;
    }
    for (const [property, value] of Object.entries(operation.properties ?? {}))
      edit.sourceStyle(operation.nodeId, property, value);
    if (operation.repeatHeaderRows !== undefined) {
      if (operation.repeatHeaderRows > rowIds.length) fail("Repeated headers exceed the table.");
      for (const [index, id] of rowIds.entries())
        edit.attributes(id, {
          "data-word-repeat-header": index < operation.repeatHeaderRows ? "true" : "false",
        });
    }
    if (operation.allowRowSplit !== undefined)
      for (const id of rowIds)
        edit.sourceStyle(id, "break-inside", operation.allowRowSplit ? "auto" : "avoid");
    if (operation.rowHeightRule !== undefined)
      for (const id of rowIds)
        edit.sourceStyle(id, "--word-row-height-rule", operation.rowHeightRule);
    return;
  }
  if (operation.op === "formatWordCell") {
    if (node.attributes["data-word-grid-gap"])
      fail("An omitted table region has no cell appearance to format.");
    recordWordFormat(word, parse(), operation.nodeId, "cell");
    if (!["td", "th"].includes(node.tagName)) fail("Cell formatting requires a table cell.");
    for (const [property, value] of Object.entries(operation.properties))
      edit.sourceStyle(operation.nodeId, property, value);
    if (operation.textDirection !== undefined)
      edit.attributes(operation.nodeId, { "data-word-text-direction": operation.textDirection });
    return;
  }
  if (operation.op === "applyWordCharacterStyle") {
    const source = parse(),
      target = source.getNode(operation.nodeId)!,
      style = word.styles.find(
        (style) => style.styleId === operation.styleId && style.kind === "character",
      );
    if (!style) fail("The character style is missing.");
    if (target.text.slice(operation.from, operation.to) !== operation.expectedText)
      fail("The character-style range is stale.");
    const range = word.trackChanges
        ? prepareWordRunFormat(word, edit, parse, styles, operation)
        : { ...operation, preferred: new Set<string>() },
      tree = parse(),
      base = tree.textOffset(range.nodeId)!,
      covered = [...range.preferred]
        .map((id) => ({
          id,
          from: tree.textOffset(id)! - base,
          to: tree.textOffset(id)! - base + tree.getNode(id)!.text.length,
        }))
        .sort((a, b) => a.from - b.from),
      gaps: { from: number; to: number }[] = [];
    let at = range.from;
    for (const area of covered) {
      if (area.from > at) gaps.push({ from: at, to: area.from });
      at = Math.max(at, area.to);
    }
    if (at < range.to) gaps.push({ from: at, to: range.to });
    let groupId: string | undefined;
    for (const area of covered) {
      const revision = word.revisions.find(
          (item) => item.nodeId === area.id && item.kind === "format" && item.property === "run",
        )!,
        attributes = (revision.before!.attributes ??= {});
      for (const key of ["class", "data-word-style", "lang", "data-word-current-format"] as const)
        if (attributes[key] === undefined)
          attributes[key] = tree.peekNode(area.id)!.attributes[key] ?? null;
      groupId ??= revision.revisionId ?? revision.nodeId;
      revision.groupId = groupId;
      revision.author = word.trackingAuthor ?? "";
      revision.date = new Date().toISOString();
      const css = styles.resolver().specifiedClass(style.cssClass);
      if (Object.keys(css).length)
        edit.formatSourceText(
          area.id,
          { from: 0, to: area.to - area.from, expectedText: tree.getNode(area.id)!.text },
          css,
          undefined,
          range.preferred,
          false,
          true,
        );
      edit.attributes(area.id, {
        "data-word-style": style.styleId,
        class: style.cssClass,
        "data-word-run-overlay": "true",
        "data-word-current-format": "true",
      });
    }
    for (const gap of gaps)
      for (const id of edit.wrapSourceText(range.nodeId, gap.from, gap.to, "span")) {
        if (word.trackChanges) {
          const current = parse();
          let parent = current.peekNode(id)!.parentNodeId,
            previous: string | undefined;
          while (parent) {
            const owner = current.peekNode(parent)!;
            if (/^p$|^h[1-6]$/.test(owner.tagName)) break;
            if (owner.attributes["data-word-style"]) {
              previous = owner.attributes["data-word-style"];
              break;
            }
            parent = owner.parentNodeId;
          }
          const identity = id + "-format";
          groupId ??= identity;
          word.revisions.push({
            nodeId: id,
            sourceId: identity,
            revisionId: identity,
            groupId,
            kind: "format",
            property: "run",
            author: word.trackingAuthor ?? "",
            date: new Date().toISOString(),
            before: {
              inheritNativeProperties: true,
              attributes: {
                "data-word-style": previous ?? null,
                class: previous
                  ? word.styles.find((style) => style.styleId === previous)!.cssClass
                  : null,
                style: null,
                "data-word-current-format": null,
              },
            },
          });
        }
        edit.attributes(id, {
          "data-word-style": style.styleId,
          class: style.cssClass,
          "data-word-run-overlay": "true",
          "data-word-current-format": "true",
        });
      }
    return;
  }
  if (operation.op === "formatWordText") {
    const properties = { ...operation.properties };
    if (properties["font-weight"] !== undefined && properties["--word-bold-cs"] === undefined)
      properties["--word-bold-cs"] =
        properties["font-weight"] === "bold" || Number(properties["font-weight"]) >= 600
          ? "1"
          : "0";
    if (properties["font-style"] !== undefined && properties["--word-italic-cs"] === undefined)
      properties["--word-italic-cs"] = properties["font-style"] === "italic" ? "1" : "0";
    const changed: string[] = [],
      force =
        word.nativeContent?.fragments.some(
          (fragment) =>
            fragment.kind === "property" &&
            fragment.propertyName ===
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main rPr",
        ) ?? false;
    if (!word.trackChanges) {
      edit.formatSourceText(
        operation.nodeId,
        operation,
        properties,
        (change) => changed.push(change.nodeId),
        undefined,
        force,
        true,
      );
      for (const id of changed) edit.attributes(id, { "data-word-current-format": "true" });
      return;
    }
    const range = prepareWordRunFormat(word, edit, parse, styles, operation),
      source = parse();
    let groupId: string | undefined;
    edit.formatSourceText(
      range.nodeId,
      range,
      properties,
      (change) => {
        changed.push(change.nodeId);
        const character = source.hasNode(change.nodeId)
            ? wordRunCharacterStyle(word, source, change.nodeId)
            : wordRunCharacterStyleAt(word, source, range.nodeId, change.from),
          previous = word.revisions.find(
            (item) =>
              item.nodeId === change.nodeId && item.kind === "format" && item.property === "run",
          ),
          id = previous?.revisionId ?? change.nodeId + "-format";
        groupId ??= id;
        if (previous) {
          previous.author = word.trackingAuthor ?? "";
          previous.date = new Date().toISOString();
          previous.groupId = groupId;
        } else
          word.revisions.push({
            nodeId: change.nodeId,
            sourceId: id,
            revisionId: id,
            groupId,
            kind: "format",
            property: "run",
            author: word.trackingAuthor ?? "",
            date: new Date().toISOString(),
            before: {
              inheritNativeProperties: true,
              ...(character ? { characterStyleId: character } : {}),
              attributes: {
                style: change.previousStyle ?? null,
                "data-word-current-format":
                  source.peekNode(change.nodeId)?.attributes["data-word-current-format"] ?? null,
              },
            },
          });
      },
      range.preferred,
      force,
      true,
    );
    for (const id of changed) edit.attributes(id, { "data-word-current-format": "true" });
    return;
  }
  if (operation.op === "setWordNumbering") {
    if (node.tagName !== "article" || node.parentNodeId !== null)
      fail("Shared Word definitions require the complete document root.");
    if (operation.op === "setWordNumbering") {
      if (operation.copyFromNumberingId) {
        const previous = word.numbering.find(
          (item) => item.numberingId === operation.copyFromNumberingId,
        );
        if (!previous) fail("The numbering appearance source is missing.");
        for (const level of operation.levels) {
          const source =
            previous.levels.find((item) => item.level === level.level) ?? previous.levels.at(-1)!;
          if (source.cssClass !== level.cssClass)
            styles.cloneClass(source.cssClass, level.cssClass);
        }
      }
      for (const item of operation.levelProperties ?? []) {
        const level = operation.levels.find((level) => level.level === item.level);
        if (!level) fail("Numbering appearance requires an existing level.");
        styles.setClass(level.cssClass, item.properties);
      }
      const index = word.numbering.findIndex((n) => n.numberingId === operation.numberingId),
        definition = { numberingId: operation.numberingId, levels: operation.levels };
      if (index < 0) word.numbering.push(definition);
      else word.numbering[index] = definition;
      return;
    }
  }
  if (!/^p$|^h[1-6]$|^li$/.test(node.tagName))
    fail("Paragraph commands require a paragraph or list item.");
  recordWordFormat(word, parse(), operation.nodeId, "paragraph");
  if (operation.op === "setWordList") {
    recordWordFormat(word, parse(), operation.nodeId, "paragraph");
    const definition = word.numbering.find((n) => n.numberingId === operation.numberingId);
    if (
      operation.numberingId !== null &&
      !definition?.levels.some((l) => l.level === (operation.level ?? 0))
    )
      fail("A list requires an existing numbering level.");
    edit.attributes(operation.nodeId, {
      "data-word-numbering": operation.numberingId ?? undefined,
      "data-word-numbering-off": operation.numberingId === null ? "true" : undefined,
      "data-word-level": operation.numberingId === null ? undefined : String(operation.level ?? 0),
    });
    edit.sourceStyle(
      operation.nodeId,
      "margin-left",
      operation.numberingId === null ? "0pt" : ((operation.level ?? 0) + 1) * 24 + "pt",
    );
    edit.sourceStyle(
      operation.nodeId,
      "text-indent",
      operation.numberingId === null ? "0pt" : "-18pt",
    );
    return;
  }
  if (operation.clearDirectFormatting)
    edit.attributes(operation.nodeId, { style: undefined, "data-word-outline-level": undefined });
  if (operation.styleId !== undefined) {
    const definition =
      operation.styleId === null
        ? undefined
        : word.styles.find((s) => s.styleId === operation.styleId && s.kind === "paragraph");
    if (operation.styleId !== null && !definition)
      fail("A paragraph style is missing or has the wrong kind.");
    const previous = word.styles.find((s) => s.styleId === node.attributes["data-word-style"]);
    const ancestry = (style: typeof definition): typeof word.styles => {
      const result: typeof word.styles = [],
        seen = new Set<string>();
      for (
        let current = style;
        current && !seen.has(current.styleId);
        current = current.basedOn
          ? word.styles.find((s) => s.styleId === current!.basedOn)
          : undefined
      ) {
        result.unshift(current);
        seen.add(current.styleId);
      }
      return result;
    };
    const oldClasses = new Set(ancestry(previous).map((style) => style.cssClass));
    const classes = (node.attributes.class ?? "")
      .split(/\s+/)
      .filter((c) => c && !oldClasses.has(c));
    if (definition) classes.push(definition.cssClass);
    edit.attributes(operation.nodeId, {
      "data-word-style": operation.styleId ?? undefined,
      class: classes.join(" "),
    });
    const previousRun = previous?.runCssClass,
      nextRun = definition?.runCssClass;
    const tree = parse(),
      visit = (id: string): void => {
        const child = tree.peekNode(id)!;
        if (child.attributes["data-word-run"]) {
          const runClasses = (child.attributes.class ?? "")
            .split(/\s+/)
            .filter((c) => c && c !== previousRun);
          if (nextRun) runClasses.push(nextRun);
          edit.attributes(id, { class: [...new Set(runClasses)].join(" ") });
        }
        for (const next of child.childNodeIds) visit(next);
      };
    visit(operation.nodeId);
    let outline = operation.clearDirectFormatting
      ? definition?.outlineLevel
      : node.attributes["data-word-outline-level"] !== undefined
        ? Number(node.attributes["data-word-outline-level"])
        : definition?.outlineLevel;
    for (let current = definition; outline === undefined && current?.basedOn; ) {
      current = word.styles.find((style) => style.styleId === current!.basedOn);
      outline = current?.outlineLevel;
    }
    outline ??= definition?.styleId.match(/^Heading([1-9])$/)
      ? Number(definition.styleId.at(-1)) - 1
      : undefined;
    edit.renameSource(
      operation.nodeId,
      outline !== undefined && outline < 6 ? "h" + (outline + 1) : "p",
    );
  }
  if (operation.outlineLevel !== undefined) {
    edit.attributes(operation.nodeId, {
      "data-word-outline-level":
        operation.outlineLevel === null ? undefined : String(operation.outlineLevel),
    });
    if (operation.outlineLevel !== null)
      edit.renameSource(
        operation.nodeId,
        operation.outlineLevel < 6 ? "h" + (operation.outlineLevel + 1) : "p",
      );
  }
  for (const [property, value] of Object.entries(operation.properties ?? {})) {
    if (
      !/^(?:margin-(?:top|right|bottom|left)|text-(?:align|indent)|line-height|break-(?:before|after|inside)|widows|orphans|background-color|border(?:-(?:top|right|bottom|left)(?:-(?:style|color|width))?)?|--word-[a-z-]+)$/.test(
        property,
      )
    )
      fail("A paragraph property is outside the Word format contract.");
    edit.sourceStyle(operation.nodeId, property, value);
  }
}
