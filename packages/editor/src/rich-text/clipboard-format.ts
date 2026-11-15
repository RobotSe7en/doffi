import {
  createOpaqueId,
  DEFAULT_DOCUMENT_READ_LIMITS,
  DocumentError,
  SourceStyleResolver,
  WordStyleSchema,
  WordContentSemanticsSchema,
} from "@doffi/core";
import type {
  ContentDocumentModel,
  DocumentOperation,
  WordMetadata,
  WordContentSemantics,
  TransactionPrepareOptions,
} from "@doffi/core";
import { pasteClipboardImages } from "./clipboard-images.js";
import { pasteClipboardResources, mergeClipboardResources } from "./clipboard-resources.js";
import {
  WORD_CLIPBOARD_BYTES,
  WORD_CLIPBOARD_FORMAT_ATTRIBUTE as marker,
  WORD_CLIPBOARD_NODE_ATTRIBUTE as identity,
  wordClipboardStyleDefinition as readDefinition,
} from "@doffi/render";
import type {
  WordClipboardFormat as Format,
  WordClipboardStyleDefinition as Definition,
} from "@doffi/render";
type Style = WordMetadata["styles"][number];
const styleClasses = (style: Style): string[] =>
  [
    style.cssClass,
    style.runCssClass,
    ...Object.values(style.components ?? {}),
    ...(style.conditions ?? []).flatMap((condition) => [
      condition.cssClass,
      ...Object.values(condition.components ?? {}),
    ]),
  ].filter((name): name is string => !!name);
const properties = (value: unknown): value is Record<string, string> =>
  !!value &&
  typeof value === "object" &&
  !Array.isArray(value) &&
  Object.values(value).every((item) => typeof item === "string");
const canonical = (value: unknown): string =>
  JSON.stringify(value, (_key, item: unknown) =>
    item && typeof item === "object" && !Array.isArray(item)
      ? Object.fromEntries(Object.entries(item).sort(([a], [b]) => a.localeCompare(b)))
      : item,
  );
/** Import only through existing typed style operations. The complete style and
 * content batch is validated before either reaches the authoritative session. */
export function pasteClipboardFormat(
  root: ParentNode,
  model: ContentDocumentModel,
): {
  operations: DocumentOperation[];
  semantics?: WordContentSemantics;
  resources?: TransactionPrepareOptions;
} {
  const wrappers = root.querySelectorAll<HTMLElement>("[" + marker + "]");
  if (!wrappers.length) return { operations: [] };
  if (wrappers.length !== 1) throw new DocumentError("InvalidOperation", "剪贴板格式包必须唯一。");
  let value: Format;
  try {
    const encoded = wrappers[0]!.getAttribute(marker)!;
    if (encoded.length > WORD_CLIPBOARD_BYTES) throw new Error();
    value = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        Uint8Array.from(atob(encoded), (character) => character.charCodeAt(0)),
      ),
    ) as Format;
  } catch {
    throw new DocumentError("InvalidOperation", "剪贴板格式包无效。");
  }
  if (
    value?.format !== "doffi-word-clipboard" ||
    ![1, 2].includes(value.version) ||
    !Array.isArray(value.styles) ||
    !value.nodes ||
    typeof value.nodes !== "object" ||
    Array.isArray(value.nodes) ||
    value.styles.length > DEFAULT_DOCUMENT_READ_LIMITS.maxNodes
  )
    throw new DocumentError("InvalidOperation", "剪贴板格式包版本或结构无效。");
  const seenIds = new Set<string>();
  value.styles = value.styles.map((item) => {
    if (
      !item ||
      typeof item !== "object" ||
      !properties(item.p) ||
      !properties(item.r) ||
      (item.table &&
        (!properties(item.table.table) ||
          !properties(item.table.cell) ||
          !properties(item.table.row))) ||
      (item.conditions &&
        (!Array.isArray(item.conditions) ||
          item.conditions.some(
            (condition) =>
              !condition ||
              typeof condition.region !== "string" ||
              !["p", "r", "table", "cell", "row"].every((scope) =>
                properties(condition[scope as keyof typeof condition]),
              ),
          )))
    )
      throw new DocumentError("InvalidOperation", "剪贴板样式属性无效。");
    const style = WordStyleSchema.parse(item.style);
    if (seenIds.has(style.styleId))
      throw new DocumentError("InvalidOperation", "剪贴板样式身份重复。");
    seenIds.add(style.styleId);
    return {
      style,
      p: item.p,
      r: item.r,
      ...(item.table ? { table: item.table, conditions: item.conditions ?? [] } : {}),
    };
  });
  if (value.version === 2) value.semantics = WordContentSemanticsSchema.parse(value.semantics);
  const owned = (value.semantics?.stories ?? []).map((story) => ({
      story,
      holder: new DOMParser().parseFromString(story.html, "text/html").body,
    })),
    roots = [root, ...owned.map((item) => item.holder)];

  const rootId =
    model.snapshot().metadata.word?.stories[0]?.rootNodeId ??
    new DOMParser()
      .parseFromString(model.snapshot().html, "text/html")
      .querySelector<HTMLElement>('[data-office-kind="document"]')!.dataset.nodeId!;
  let documentRoot = model.getNode(rootId)!;
  while (documentRoot.parentNodeId) documentRoot = model.getNode(documentRoot.parentNodeId)!;
  const snapshot = model.snapshot(),
    targetCss = new SourceStyleResolver(
      snapshot.styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    existing = new Map(snapshot.metadata.word?.styles.map((style) => [style.styleId, style])),
    reused = new Set(
      value.styles
        .filter(
          (item) =>
            existing.has(item.style?.styleId) &&
            canonical(item) ===
              canonical(readDefinition(existing.get(item.style.styleId)!, targetCss)),
        )
        .map((item) => item.style.styleId),
    );
  for (let changed = true; changed; ) {
    changed = false;
    for (const item of value.styles)
      if (
        reused.has(item.style.styleId) &&
        [item.style.basedOn, item.style.linkedStyleId, item.style.nextStyleId].some(
          (id) => id && !reused.has(id),
        )
      ) {
        reused.delete(item.style.styleId);
        changed = true;
      }
  }
  const namespace = "clipboard-" + createOpaqueId(),
    ids = new Map<string, string>(),
    classes = new Map<string, string>(),
    definitions: Definition[] = value.styles.flatMap((item, index) => {
      const style = WordStyleSchema.parse(item.style);
      if (
        !["paragraph", "character", "table"].includes(style.kind) ||
        ids.has(style.styleId) ||
        !item.p ||
        !item.r
      )
        throw new DocumentError("InvalidOperation", "剪贴板样式定义无效。");
      ids.set(style.styleId, namespace + "-" + index);
      for (const [nameIndex, name] of styleClasses(style).entries())
        classes.set(name, namespace + "-class-" + index + "-" + nameIndex);
      if (reused.has(style.styleId)) {
        ids.set(style.styleId, style.styleId);
        for (const name of styleClasses(style)) classes.set(name, name);
        return [];
      }
      return [{ ...item, style }];
    }),
    mapped = (definition: Style, links: boolean): Style => {
      const {
          basedOn,
          nextStyleId,
          linkedStyleId,
          components,
          paragraphProperties,
          runProperties,
          ...rest
        } = definition,
        style: Style = {
          ...rest,
          styleId: ids.get(definition.styleId)!,
          cssClass: classes.get(definition.cssClass)!,
          default: false,
          ...(definition.runCssClass ? { runCssClass: classes.get(definition.runCssClass)! } : {}),
        };
      if (links)
        for (const [key, id] of Object.entries({ basedOn, nextStyleId, linkedStyleId }))
          if (id) {
            const mapped = ids.get(id);
            if (!mapped) throw new DocumentError("InvalidOperation", "剪贴板样式依赖缺失。");
            (style as unknown as Record<string, unknown>)[key] = mapped;
          }
      if (components)
        style.components = Object.fromEntries(
          Object.entries(components).map(([scope, name]) => [scope, classes.get(name!)!]),
        );
      if (definition.conditions)
        style.conditions = definition.conditions.map((condition) => ({
          ...condition,
          cssClass: classes.get(condition.cssClass)!,
          ...(condition.components
            ? {
                components: Object.fromEntries(
                  Object.entries(condition.components).map(([scope, name]) => [
                    scope,
                    classes.get(name!)!,
                  ]),
                ),
              }
            : {}),
        }));
      return style;
    },
    operations: DocumentOperation[] = definitions.map((item) => ({
      op: "setWordStyle",
      nodeId: documentRoot.nodeId,
      style: mapped(item.style, false),
    })),
    ordered: Definition[] = [],
    active = new Set<string>(),
    done = new Set<string>(),
    byId = new Map(definitions.map((item) => [item.style.styleId, item]));
  const order = (item: Definition): void => {
    const id = item.style.styleId;
    if (done.has(id)) return;
    if (active.has(id) || active.size >= DEFAULT_DOCUMENT_READ_LIMITS.maxDepth)
      throw new DocumentError("InvalidOperation", "剪贴板样式继承无效。");
    active.add(id);
    if (item.style.basedOn && !reused.has(item.style.basedOn)) {
      const parent = byId.get(item.style.basedOn);
      if (!parent) throw new DocumentError("InvalidOperation", "剪贴板样式依赖缺失。");
      order(parent);
    }
    active.delete(id);
    done.add(id);
    ordered.push(item);
  };
  for (const item of definitions) order(item);
  for (const item of ordered) {
    const definition = mapped(item.style, true);
    operations.push({
      op: "setWordStyle",
      nodeId: documentRoot.nodeId,
      style: definition,
      paragraphProperties: item.p,
      runProperties: item.r,
      ...(item.table
        ? {
            tableProperties: item.table.table,
            cellProperties: item.table.cell,
            rowProperties: item.table.row,
          }
        : {}),
    });
    for (const condition of item.conditions ?? []) {
      const { region, p, r, table, cell, row } = condition;
      operations.push({
        op: "setWordStyle",
        nodeId: documentRoot.nodeId,
        style: definition,
        tableCondition: {
          region,
          paragraphProperties: p,
          runProperties: r,
          tableProperties: table,
          cellProperties: cell,
          rowProperties: row,
        },
      });
    }
  }
  for (const table of value.semantics?.tables ?? [])
    for (const entry of table.continuationFormats ?? [])
      for (const paragraph of entry.emptyParagraphs ?? []) {
        if (paragraph.styleId) {
          const id = ids.get(paragraph.styleId);
          if (!id) throw new DocumentError("InvalidOperation", "剪贴板续接段落样式缺失。");
          paragraph.styleId = id;
        }
        for (const run of paragraph.runs)
          if (run.styleId) {
            const id = ids.get(run.styleId);
            if (!id) throw new DocumentError("InvalidOperation", "剪贴板续接文字样式缺失。");
            run.styleId = id;
          }
      }

  for (const element of roots.flatMap((root) => [
    ...root.querySelectorAll<HTMLElement>("[" + identity + "]"),
  ])) {
    const key = element.getAttribute(identity)!,
      attributes = Object.hasOwn(value.nodes, key) ? value.nodes[key] : undefined;
    if (
      !attributes ||
      ![attributes.style, attributes.class, attributes.styleId].every(
        (value) => value === null || typeof value === "string",
      )
    )
      throw new DocumentError("InvalidOperation", "剪贴板节点格式无效。");
    if (attributes.style === null) element.removeAttribute("style");
    else element.setAttribute("style", attributes.style);
    if (attributes.class === null) element.removeAttribute("class");
    else
      element.setAttribute(
        "class",
        attributes.class
          .split(/\s+/)
          .map((name) => classes.get(name) ?? name)
          .join(" "),
      );
    if (attributes.styleId === null) element.removeAttribute("data-word-style");
    else {
      const id = ids.get(attributes.styleId);
      if (!id) throw new DocumentError("InvalidOperation", "剪贴板节点样式缺失。");
      element.setAttribute("data-word-style", id);
    }
    if (value.version === 2) {
      if (typeof attributes.sourceId !== "string")
        throw new DocumentError("InvalidOperation", "剪贴板语义节点身份无效。");
      element.setAttribute("data-node-id", attributes.sourceId);
      for (const [key, value] of Object.entries(attributes.attributes ?? {})) {
        if (
          ![
            "data-word-field",
            "data-word-field-end",
            "data-word-field-result",
            "data-word-numbering",
            "data-word-level",
            "data-word-numbering-off",
            "data-word-drawing",
          ].includes(key) ||
          typeof value !== "string"
        )
          throw new DocumentError("InvalidOperation", "剪贴板语义节点属性无效。");
        element.setAttribute(key, value);
      }
    }
    element.removeAttribute(identity);
  }
  wrappers[0]!.removeAttribute(marker);
  if (
    value.version === 2 &&
    (!value.semantics || typeof value.semantics !== "object" || Array.isArray(value.semantics))
  )
    throw new DocumentError("InvalidOperation", "剪贴板语义包无效。");
  const resources = mergeClipboardResources(
    pasteClipboardResources(value),
    ...owned.map((item) => pasteClipboardImages(item.holder, model)),
  );
  for (const item of owned) item.story.html = item.holder.innerHTML;
  return {
    operations,
    ...(value.version === 2 ? { semantics: value.semantics! } : {}),
    ...(resources ? { resources } : {}),
  };
}
