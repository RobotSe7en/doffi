import { DocumentError } from "../errors.js";
import type { WordDocumentMetadata } from "../contracts/document.js";
import type { WordContentSemantics } from "../contracts/word-content.js";
import type { SourceStyleEdit } from "./source-styles.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import { createOpaqueId } from "../identifiers.js";
import { remapWordFieldBookmarks } from "./word-field-bookmarks.js";
import { ensureWordMainStoryContainer } from "./word-story-container.js";
import { attachWordContentComments } from "./word-content-comments.js";
import type { WordContentCommentContext } from "./word-content-comments.js";
import { attachWordContentNative } from "./word-content-native.js";
const invalid = (message: string): never => {
  throw new DocumentError("InvalidOperation", message);
};
export function attachWordContent(
  metadata: WordDocumentMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  sourceIds: ReadonlyMap<string, string>,
  semantics: WordContentSemantics,
  styles?: SourceStyleEdit,
  insertedRootIds?: readonly string[],
  comments?: WordContentCommentContext,
): void {
  const word = metadata.word!,
    ids = new Map(sourceIds),
    storyIds = new Map<string, string>(),
    storyRoots: string[] = [];
  if (semantics.stories?.length) {
    const article = ensureWordMainStoryContainer(metadata, edit, parse);
    for (const story of semantics.stories) {
      if (storyIds.has(story.storyId)) invalid("Copied story identities must be unique.");
      const content = edit.insertSource(article, story.html, undefined, true),
        root =
          content.identities.get(story.rootNodeId) ??
          invalid("A copied story is missing its declared root.");
      if (
        !root ||
        content.rootNodeIds.length !== 1 ||
        content.rootNodeIds[0] !== root ||
        parse().peekNode(root)!.attributes["data-office-story"] !== story.kind
      )
        invalid("A copied story must contain exactly its declared source root.");
      for (const [before, after] of content.identities) {
        if (ids.has(before)) invalid("Clipboard source identities must be unique across stories.");
        ids.set(before, after);
      }
      const storyId = "word-content-story-" + createOpaqueId();
      storyIds.set(story.storyId, storyId);
      storyRoots.push(root);
      word.stories.push({ storyId, rootNodeId: root, kind: story.kind });
    }
  }
  const map = (id: string): string =>
    ids.get(id) ??
    invalid("A clipboard semantic reference must belong to the supplied inline source.");
  const usedNames = new Set(word.bookmarks.map((bookmark) => bookmark.name)),
    names = new Map<string, string>(),
    newName = (): string => {
      let name: string;
      do {
        name = "Copy_" + createOpaqueId().replaceAll("-", "");
      } while (usedNames.has(name));
      usedNames.add(name);
      return name;
    };
  for (const bookmark of semantics.bookmarks ?? []) {
    if (names.has(bookmark.name)) invalid("Copied bookmarks require unique source names.");
    const name =
      usedNames.has(bookmark.name) || !/^[_\p{L}][_\p{L}\p{N}]{0,39}$/u.test(bookmark.name)
        ? newName()
        : bookmark.name;
    usedNames.add(name);
    names.set(bookmark.name, name);
    word.bookmarks.push({
      ...bookmark,
      bookmarkId: map(bookmark.startNodeId),
      name,
      startNodeId: map(bookmark.startNodeId),
      endNodeId: map(bookmark.endNodeId),
    });
  }
  const referenceName = (name: string): string => {
    let result = names.get(name);
    if (!result) {
      result = newName();
      names.set(name, result);
    }
    return result;
  };
  const inserted = parse();
  const insertedIds = new Set(ids.values());
  const pending = [...(insertedRootIds ?? []), ...storyRoots];
  while (pending.length) {
    const id = pending.pop()!;
    insertedIds.add(id);
    pending.push(...inserted.peekNode(id)!.childNodeIds);
  }
  for (const after of insertedIds) {
    const node = inserted.peekNode(after)!;
    for (const key of ["data-word-field-end", "data-word-field-result"]) {
      const id = node.attributes[key];
      if (id) edit.attributes(after, { [key]: map(id) });
    }
  }
  const classes = new Map<string, string>();
  if (semantics.extraClasses) {
    const targetStyles =
      styles ?? invalid("Copied metadata classes require their bounded style adapter.");
    const prefix = "word-content-" + createOpaqueId();
    targetStyles.replaceClasses(
      Object.entries(semantics.extraClasses).map(([name, properties], index) => {
        const mapped = prefix + "-" + index;
        classes.set(name, mapped);
        return { name: mapped, properties };
      }),
    );
    for (const id of insertedIds) {
      const node = inserted.peekNode(id)!,
        names = node.attributes.class?.split(/\s+/);
      if (names?.some((name) => classes.has(name)))
        edit.attributes(id, { class: names.map((name) => classes.get(name) ?? name).join(" ") });
    }
  }
  const className = (name: string): string =>
    classes.get(name) ?? invalid("Copied metadata formatting requires its declared class.");
  const numberingIds = new Map<string, string>();
  for (const definition of semantics.numbering ?? []) {
    if (numberingIds.has(definition.numberingId))
      invalid("Copied numbering identities must be unique.");
    const numberingId = "word-content-numbering-" + createOpaqueId();
    numberingIds.set(definition.numberingId, numberingId);
    word.numbering.push({
      numberingId,
      levels: definition.levels.map((level) => ({ ...level, cssClass: className(level.cssClass) })),
    });
  }
  const numberingId = (id: string): string =>
    numberingIds.get(id) ?? invalid("Copied numbering must belong to the supplied content graph.");
  for (const id of insertedIds) {
    const sourceId = inserted.peekNode(id)!.attributes["data-word-numbering"];
    if (sourceId) edit.attributes(id, { "data-word-numbering": numberingId(sourceId) });
  }
  word.tabs.push(
    ...(semantics.tabs ?? []).map((tab) => ({
      ...tab,
      paragraphNodeId: map(tab.paragraphNodeId),
      stops: tab.stops.map((stop) => ({ ...stop })),
    })),
  );
  if (semantics.conditionalFormats?.length)
    (word.conditionalFormats ??= []).push(
      ...semantics.conditionalFormats.map((item) => ({
        nodeId: map(item.nodeId),
        format: { ...item.format },
      })),
    );
  word.tables.push(
    ...(semantics.tables ?? []).map((table) => {
      const copy = JSON.parse(JSON.stringify(table)) as typeof table;
      copy.nodeId = map(copy.nodeId);
      for (const context of copy.styleContexts ?? []) {
        const id = map(context.nodeId),
          node = inserted.peekNode(id)!;
        edit.attributes(id, {
          class: (node.attributes.class ?? "")
            .split(/\s+/)
            .filter((name) => name !== context.cssClass && name !== classes.get(context.cssClass))
            .join(" "),
        });
      }
      delete copy.styleContexts;
      for (const entry of copy.continuationFormats ?? []) {
        if (entry.cellCssClass) entry.cellCssClass = className(entry.cellCssClass);
        for (const paragraph of entry.emptyParagraphs ?? []) {
          if (paragraph.numbering)
            paragraph.numbering.numberingId = numberingId(paragraph.numbering.numberingId);
          if (paragraph.paragraphCssClass)
            paragraph.paragraphCssClass = className(paragraph.paragraphCssClass);
          if (paragraph.runCssClass) paragraph.runCssClass = className(paragraph.runCssClass);
          for (const run of paragraph.runs)
            if (run.cssClass) run.cssClass = className(run.cssClass);
        }
      }
      return copy;
    }),
  );
  word.fields.push(
    ...(semantics.fields ?? []).map((field) => {
      const copy = {
        ...field,
        nodeId: map(field.nodeId),
        ...(field.endNodeId ? { endNodeId: map(field.endNodeId) } : {}),
        ...(field.resultNodeIds ? { resultNodeIds: field.resultNodeIds.map(map) } : {}),
        ...(field.instructionFields
          ? {
              instructionFields: field.instructionFields.map((child) => ({
                ...child,
                fieldNodeId: map(child.fieldNodeId),
              })),
            }
          : {}),
      };
      remapWordFieldBookmarks(copy, referenceName);
      return copy;
    }),
  );
  word.hyperlinks.push(
    ...(semantics.hyperlinks ?? []).map((link) => ({
      ...link,
      nodeId: map(link.nodeId),
      ...(link.endNodeId ? { endNodeId: map(link.endNodeId) } : {}),
      ...(link.bookmark && !link.target ? { bookmark: referenceName(link.bookmark) } : {}),
    })),
  );
  if (semantics.equations?.length)
    (word.equations ??= []).push(
      ...semantics.equations.map((equation) => ({
        ...(JSON.parse(JSON.stringify(equation)) as typeof equation),
        nodeId: map(equation.nodeId),
      })),
    );
  if (semantics.controls?.length)
    (word.controls ??= []).push(
      ...semantics.controls.map((control) => {
        const copy = { ...control, nodeId: map(control.nodeId) };
        delete copy.sourceId;
        return copy;
      }),
    );
  const owners = new Set<string>(),
    noteIds = new Map<string, string>(),
    referenced = new Set<string>();
  for (const note of semantics.notes ?? []) {
    if (noteIds.has(note.noteId) || (note.type ?? "normal") !== "normal")
      invalid("Copied note definitions must be unique normal notes.");
    const storyId =
      storyIds.get(note.storyId) ?? invalid("A copied note requires its owned story.");
    if (semantics.stories?.find((s) => s.storyId === note.storyId)?.kind !== note.kind)
      invalid("A copied note story must match its reference kind.");
    const noteId = "word-content-note-" + createOpaqueId();
    noteIds.set(note.noteId, noteId);
    owners.add(note.storyId);
    (word.notes ??= []).push({ ...note, noteId, storyId });
  }
  for (const reference of semantics.noteReferences ?? []) {
    if (referenced.has(reference.noteId))
      invalid("Each copied note must have one independent reference.");
    referenced.add(reference.noteId);
    (word.noteReferences ??= []).push({
      ...reference,
      nodeId: map(reference.nodeId),
      noteId:
        noteIds.get(reference.noteId) ??
        invalid("A copied reference requires its selected note definition."),
    });
  }
  if (noteIds.size !== referenced.size || [...noteIds.keys()].some((id) => !referenced.has(id)))
    invalid("A clipboard note must belong to a selected reference.");
  if (semantics.drawings?.length) {
    const selected = new Set(semantics.drawings.map((d) => d.nodeId));
    (word.drawings ??= []).push(
      ...semantics.drawings.map((drawing) => {
        const copy = JSON.parse(JSON.stringify(drawing)) as typeof drawing;
        copy.nodeId = map(copy.nodeId);
        if ("textStoryId" in copy && copy.textStoryId) {
          owners.add(copy.textStoryId);
          copy.textStoryId =
            storyIds.get(copy.textStoryId) ??
            invalid("Copied drawing text requires its owned story.");
        }
        if (copy.kind === "connector")
          for (const side of ["start", "end"] as const) {
            const key = side === "start" ? "startNodeId" : "endNodeId",
              site = side === "start" ? "startSite" : "endSite",
              id = copy.connector[key];
            if (id && selected.has(id)) copy.connector[key] = map(id);
            else {
              delete copy.connector[key];
              delete copy.connector[site];
            }
          }
        if (copy.kind === "chart") {
          const series = new Map(
            copy.chart.series.map((s) => [s.seriesId, "word-content-series-" + createOpaqueId()]),
          );
          for (const item of copy.chart.series) item.seriesId = series.get(item.seriesId)!;
          const native = copy.chart.native;
          if (native) {
            const prefix = "word/clipboard-" + createOpaqueId() + "/",
              parts = new Map(native.parts.map((p) => [p.sourcePart, prefix + p.sourcePart]));
            native.chartPart = parts.get(native.chartPart)!;
            if (native.workbookPart) native.workbookPart = parts.get(native.workbookPart)!;
            for (const part of native.parts) part.sourcePart = parts.get(part.sourcePart)!;
            for (const relation of native.relationships) {
              relation.sourcePart =
                parts.get(relation.sourcePart) ??
                invalid("A copied chart relationship requires its source part.");
              if (!relation.external)
                relation.target =
                  parts.get(relation.target) ??
                  invalid("A copied chart relationship requires its target part.");
            }
            for (const binding of native.bindings)
              binding.seriesId =
                series.get(binding.seriesId) ??
                invalid("A copied chart binding requires its series.");
            for (const item of native.baseline?.series ?? [])
              item.seriesId =
                series.get(item.seriesId) ??
                invalid("A copied chart baseline requires its series.");
          }
        }
        return copy;
      }),
    );
  }
  attachWordContentNative(word, edit, ids, semantics.nativeContent, numberingIds);
  attachWordContentComments(word, semantics, ids, storyIds, owners, parse(), comments);
  if (storyIds.size !== owners.size || [...storyIds.keys()].some((id) => !owners.has(id)))
    invalid("Clipboard stories must belong to selected objects or note references.");
}
