import { WordContentSemanticsSchema } from "@doffi/core";
import type { ContentDocumentModel, WordContentSemantics, SourceStyleResolver } from "@doffi/core";
import { clipboardNodeText } from "./word-clipboard-stories.js";
/** Only complete supported structures enter the semantic graph. Partial fields
 * and controls keep their selected text/format without copying outside text. */
export function copyWordClipboardSemantics(
  holder: HTMLElement,
  model: ContentDocumentModel,
  completeNodes: ReadonlySet<string> | undefined,
  css: SourceStyleResolver,
  roots: readonly HTMLElement[] = [holder],
): { semantics?: WordContentSemantics; attributes: ReadonlyMap<string, Record<string, string>> } {
  const word = model.snapshot().metadata.word!,
    nodes = new Map(
      roots
        .flatMap((root) => [...root.querySelectorAll<HTMLElement>("[data-node-id]")])
        .map((node) => [node.dataset.nodeId!, node]),
    ),
    complete = (id: string): boolean => {
      const node = nodes.get(id),
        source = model.getNode(id);
      return (
        !!node &&
        !!source &&
        (completeNodes ? completeNodes.has(id) : clipboardNodeText(node) === source.text)
      );
    },
    attributes = new Map<string, Record<string, string>>();
  const drawings = (word.drawings ?? []).filter((d) => d.kind !== "group" && complete(d.nodeId));
  const groups = (word.drawings ?? []).filter((d) => d.kind === "group" && nodes.has(d.nodeId));
  for (let changed = true; changed; ) {
    changed = false;
    for (const group of groups)
      if (
        !drawings.includes(group) &&
        drawings.some((d) => nodes.get(group.nodeId)!.contains(nodes.get(d.nodeId)!))
      ) {
        drawings.push(group);
        changed = true;
      }
  }
  const drawingIds = new Map(drawings.map((d) => [d.nodeId, d.kind]));
  const fields = word.fields
      .filter(
        (field) =>
          complete(field.nodeId) &&
          (!field.endNodeId || complete(field.endNodeId)) &&
          (field.instructionFields ?? []).every((child) => nodes.has(child.fieldNodeId)),
      )
      .map((field) => ({
        ...field,
        ...(field.resultNodeIds
          ? { resultNodeIds: field.resultNodeIds.filter((id) => nodes.has(id)) }
          : {}),
      })),
    equations = word.equations?.filter((item) => complete(item.nodeId)) ?? [],
    controls =
      word.controls?.filter(
        (item) =>
          complete(item.nodeId) &&
          !word.nativeContent?.fragments.some((fragment) => fragment.nodeId === item.nodeId),
      ) ?? [],
    hyperlinks = word.hyperlinks.filter(
      (link) => nodes.has(link.nodeId) && (!link.endNodeId || nodes.has(link.endNodeId)),
    ),
    bookmarks = word.bookmarks.filter(
      (bookmark) => nodes.has(bookmark.startNodeId) && nodes.has(bookmark.endNodeId),
    );
  const bookmarkIds = new Set(
    bookmarks.flatMap((bookmark) => [bookmark.startNodeId, bookmark.endNodeId]),
  );
  const controlIds = new Set(controls.map((item) => item.nodeId)),
    equationIds = new Set(equations.map((item) => item.nodeId)),
    linkIds = new Set(
      hyperlinks.flatMap((link) => [link.nodeId, ...(link.endNodeId ? [link.endNodeId] : [])]),
    );
  let selectedFields = fields;
  for (let changed = true; changed; ) {
    const ids = new Set(selectedFields.map((field) => field.nodeId)),
      next = selectedFields.filter(
        (field) => !field.instructionFields?.some((child) => !ids.has(child.fieldNodeId)),
      );
    changed = next.length !== selectedFields.length;
    selectedFields = next;
  }
  const fieldIds = new Set(selectedFields.map((field) => field.nodeId)),
    fieldMarkers = new Set(
      selectedFields.flatMap((field) => [
        field.nodeId,
        ...(field.endNodeId ? [field.endNodeId] : []),
      ]),
    );
  for (const [id, node] of nodes) {
    const kind = drawingIds.get(id);
    if (kind) node.setAttribute("data-word-drawing", kind);
    else node.removeAttribute("data-word-drawing");
    if (node.hasAttribute("data-word-field") && !fieldMarkers.has(id))
      node.removeAttribute("data-word-field");
    for (const key of ["data-word-field-end", "data-word-field-result"])
      if (node.hasAttribute(key) && !fieldIds.has(node.getAttribute(key)!))
        node.removeAttribute(key);
    if (node.hasAttribute("data-word-content-control") && !controlIds.has(id))
      for (const attribute of [...node.attributes])
        if (
          attribute.name === "data-word-content-control" ||
          attribute.name.startsWith("data-word-control-")
        )
          node.removeAttribute(attribute.name);
    if (node.hasAttribute("data-word-equation") && !equationIds.has(id))
      node.removeAttribute("data-word-equation");
    if (node.hasAttribute("data-word-link") && !linkIds.has(id))
      node.removeAttribute("data-word-link");
    if (!bookmarkIds.has(id)) node.removeAttribute("data-word-bookmark");
    const own = Object.fromEntries(
      [...node.attributes]
        .filter((attribute) =>
          [
            "data-word-field",
            "data-word-field-end",
            "data-word-field-result",
            "data-word-numbering",
            "data-word-level",
            "data-word-numbering-off",
            "data-word-drawing",
          ].includes(attribute.name),
        )
        .map((attribute) => [attribute.name, attribute.value]),
    );
    if (Object.keys(own).length) attributes.set(id, own);
  }
  const tables = word.tables
      .filter((table) => complete(table.nodeId))
      .map((table) => {
        const copy = JSON.parse(JSON.stringify(table)) as NonNullable<
          WordContentSemantics["tables"]
        >[number];
        delete copy.styleContexts;
        return copy;
      }),
    tabs = word.tabs.filter((tab) => complete(tab.paragraphNodeId)),
    conditionalFormats = word.conditionalFormats?.filter((item) => complete(item.nodeId)),
    classes = new Set<string>(),
    numberingIds = new Set(
      [...nodes.values()].flatMap((node) =>
        node.dataset.wordNumbering ? [node.dataset.wordNumbering] : [],
      ),
    );
  for (const table of tables)
    for (const entry of table.continuationFormats ?? []) {
      if (entry.cellCssClass) classes.add(entry.cellCssClass);
      for (const paragraph of entry.emptyParagraphs ?? []) {
        if (paragraph.numbering) numberingIds.add(paragraph.numbering.numberingId);
        if (paragraph.paragraphCssClass) classes.add(paragraph.paragraphCssClass);
        if (paragraph.runCssClass) classes.add(paragraph.runCssClass);
        for (const run of paragraph.runs) if (run.cssClass) classes.add(run.cssClass);
      }
    }
  const numbering = word.numbering.filter((definition) => numberingIds.has(definition.numberingId));
  for (const definition of numbering)
    for (const level of definition.levels) classes.add(level.cssClass);
  const semantics = WordContentSemanticsSchema.parse({
    fields: selectedFields,
    equations,
    controls,
    hyperlinks,
    bookmarks,
    ...(drawings.length ? { drawings } : {}),
    ...(numbering.length ? { numbering } : {}),
    ...(tables.length ? { tables } : {}),
    ...(tabs.length ? { tabs } : {}),
    ...(conditionalFormats?.length ? { conditionalFormats } : {}),
    ...(classes.size
      ? {
          extraClasses: Object.fromEntries(
            [...classes].map((name) => [name, css.specifiedClass(name)]),
          ),
        }
      : {}),
  });
  return {
    ...(Object.values(semantics).some((items) => Array.isArray(items) && items.length > 0)
      ? { semantics }
      : {}),
    attributes,
  };
}
