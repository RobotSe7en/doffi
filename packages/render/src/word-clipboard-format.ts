import { copyWordClipboardStories } from "./word-clipboard-stories.js";
import { copyWordClipboardImages } from "./word-clipboard-images.js";
import { copyWordClipboardSemantics } from "./word-clipboard-semantics.js";
import type { WordContentSemantics } from "@doffi/core";
import {
  DEFAULT_DOCUMENT_READ_LIMITS,
  getWordTextDefaults,
  SourceStyleResolver,
  WordStyleSchema,
  DocumentError,
} from "@doffi/core";
import type {
  ContentDocumentModel,
  SourceStyleNode,
  WordMetadata,
  DeepReadonly,
} from "@doffi/core";
import {
  WORD_CLIPBOARD_BYTES,
  WORD_CLIPBOARD_FORMAT_ATTRIBUTE as marker,
  WORD_CLIPBOARD_NODE_ATTRIBUTE as identity,
} from "./word-clipboard-constants.js";
type Style = WordMetadata["styles"][number];
export interface WordClipboardTableStyleProperties {
  p: Record<string, string>;
  r: Record<string, string>;
  table: Record<string, string>;
  cell: Record<string, string>;
  row: Record<string, string>;
}
export interface WordClipboardStyleDefinition {
  style: Style;
  p: Record<string, string>;
  r: Record<string, string>;
  table?: Pick<WordClipboardTableStyleProperties, "table" | "cell" | "row">;
  conditions?: ({
    region: NonNullable<Style["conditions"]>[number]["region"];
  } & WordClipboardTableStyleProperties)[];
}
export interface WordClipboardFormat {
  format: "doffi-word-clipboard";
  version: 1 | 2;
  semantics?: WordContentSemantics;
  styles: WordClipboardStyleDefinition[];
  resources?: { path: string; mimeType: string; data: string }[];
  nodes: Record<
    string,
    {
      style: string | null;
      class: string | null;
      styleId: string | null;
      sourceId?: string;
      attributes?: Record<string, string>;
    }
  >;
}
const runKey = (key: string): boolean =>
  /^(?:font-|color$|background-color$|text-(?:decoration|transform)|vertical-align$|letter-spacing$|word-spacing$|--word-(?:font|run|kerning|language|bold|italic))/.test(
    key,
  );
const bools = ["font-weight", "font-style", "font-variant", "text-transform"];
const declarations = (properties: Readonly<Record<string, string>>): string =>
  Object.entries(properties)
    .map(([key, value]) => key + ":" + value)
    .join(";");
export function wordClipboardStyleDefinition(
  style: DeepReadonly<Style>,
  css: SourceStyleResolver,
): WordClipboardStyleDefinition {
  const own = {
      ...css.specifiedClass(style.cssClass),
      ...(style.runCssClass ? css.specifiedClass(style.runCssClass) : {}),
    },
    p = style.components?.p
      ? css.specifiedClass(style.components.p)
      : Object.fromEntries(
          Object.entries(own).filter(
            ([key]) => style.paragraphProperties?.includes(key) ?? !runKey(key),
          ),
        ),
    r = {
      ...(style.components?.r
        ? css.specifiedClass(style.components.r)
        : Object.fromEntries(
            Object.entries(own).filter(
              ([key]) => style.runProperties?.includes(key) ?? runKey(key),
            ),
          )),
    },
    resolved = style.runCssClass
      ? css.classStyle(style.runCssClass, css.classStyle(style.cssClass))
      : css.classStyle(style.cssClass);
  if (style.kind !== "table")
    for (const key of bools)
      if (r[key] !== undefined && resolved[key] !== undefined) r[key] = resolved[key]!;
  const table =
    style.kind === "table"
      ? {
          table: style.components?.table
            ? css.specifiedClass(style.components.table)
            : !style.components
              ? css.specifiedClass(style.cssClass)
              : {},
          cell: style.components?.cell ? css.specifiedClass(style.components.cell) : {},
          row: style.components?.row ? css.specifiedClass(style.components.row) : {},
        }
      : undefined;
  const conditions =
    style.kind === "table"
      ? (style.conditions ?? []).map((condition) => ({
          region: condition.region,
          ...(Object.fromEntries(
            (["p", "r", "table", "cell", "row"] as const).map((scope) => {
              const name = condition.components?.[scope];
              return [
                scope,
                name
                  ? css.specifiedClass(name)
                  : !condition.components && scope === "cell"
                    ? css.specifiedClass(condition.cssClass)
                    : {},
              ];
            }),
          ) as unknown as WordClipboardTableStyleProperties),
        }))
      : undefined;
  return {
    style: WordStyleSchema.parse(style),
    p: { ...p },
    r,
    ...(table ? { table, conditions: conditions ?? [] } : {}),
  };
}
const encode = (value: unknown): string => {
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  let text = "";
  for (let offset = 0; offset < bytes.length; offset += 32768)
    text += String.fromCharCode(...bytes.subarray(offset, offset + 32768));
  return btoa(text);
};

/** The standard HTML has authored font values; SDK readers also recover the
 * named style definitions. No unselected source text or runtime fonts enter it. */
export function copyWordClipboardFormat(
  holder: HTMLElement,
  model: ContentDocumentModel,
  contextId: string,
  completeNodes?: ReadonlySet<string>,
): void {
  const snapshot = model.snapshot(),
    word = snapshot.metadata.word;
  if (!word || holder.querySelector("[data-word-native-object]")) return;
  for (const table of [...holder.querySelectorAll<HTMLElement>("table")])
    if (completeNodes && !completeNodes.has(table.dataset.nodeId!))
      for (const element of [table, ...table.querySelectorAll("thead,tbody,tfoot,tr,td,th")]
        .filter((element) => element.closest("table") === table)
        .reverse())
        element.replaceWith(...element.childNodes);
  const block = [...holder.children].some((element) =>
    element.matches("p,h1,h2,h3,h4,h5,h6,div,section,article,ul,ol,table"),
  );
  const css = new SourceStyleResolver(
    snapshot.styles.map((sheet) => sheet.css),
    DEFAULT_DOCUMENT_READ_LIMITS,
    new Set(
      word.revisions
        .filter((item) => ["insert", "delete", "moveFrom", "moveTo"].includes(item.kind))
        .map((item) => item.nodeId),
    ),
  );
  const graph = copyWordClipboardStories(holder, model, completeNodes);
  for (const story of graph.stories) {
    copyWordClipboardImages(story.holder, model);
    const root = story.holder.firstElementChild!;
    root.setAttribute(
      "style",
      declarations(getWordTextDefaults(model, story.rootNodeId)) +
        ";" +
        (root.getAttribute("style") ?? ""),
    );
  }
  const semantic = copyWordClipboardSemantics(holder, model, graph.completeNodes, css, graph.roots);
  for (const drawing of semantic.semantics?.drawings ?? [])
    if (drawing.kind === "chart" && drawing.chart.native) {
      const native = drawing.chart.native,
        book = native.workbookPart;
      // The current vectors own the copied data. Rebuild a bounded embedded book
      // on output instead of carrying unrelated source workbook cells.
      if (book) {
        native.parts = native.parts.filter((p) => p.sourcePart !== book);
        native.relationships = native.relationships.filter(
          (r) => r.sourcePart !== book && (r.external || r.target !== book),
        );
        delete native.workbookPart;
      }
      for (const binding of native.bindings) {
        delete binding.nameFormula;
        delete binding.categoryFormula;
        delete binding.valueFormula;
        delete binding.xFormula;
      }
    }
  // The clipboard carries the selected content and its formatting. Receiver
  // revision ownership is created by the receiving transaction.
  for (const revision of graph.roots.flatMap((root) => [...root.querySelectorAll("ins,del")])) {
    const span = holder.ownerDocument.createElement("span");
    for (const attribute of revision.attributes) span.setAttribute(attribute.name, attribute.value);
    span.append(...revision.childNodes);
    revision.replaceWith(span);
  }
  for (const marker of graph.roots.flatMap((root) => [
    ...root.querySelectorAll("[data-word-move-range]"),
  ]))
    marker.remove();
  if (!block)
    for (const mark of holder.querySelectorAll("[data-word-paragraph-mark]")) mark.remove();
  const cache = new Map<string, SourceStyleNode>();
  const project = (id: string): SourceStyleNode => {
    let value = cache.get(id);
    if (!value) {
      const node = model.getNode(id)!;
      value = {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
      cache.set(id, value);
    }
    return value;
  };
  const elements = graph.roots.flatMap((root) => [
      ...root.querySelectorAll<HTMLElement>("[data-node-id]"),
    ]),
    classes = new Set(elements.flatMap((element) => [...element.classList])),
    used = new Set<string>(
      elements.flatMap((element) => (element.dataset.wordStyle ? [element.dataset.wordStyle] : [])),
    ),
    byId = new Map(word.styles.map((style) => [style.styleId, style]));
  for (const style of word.styles)
    if (classes.has(style.cssClass) || (style.runCssClass && classes.has(style.runCssClass)))
      used.add(style.styleId);
  for (const table of semantic.semantics?.tables ?? [])
    for (const entry of table.continuationFormats ?? [])
      for (const paragraph of entry.emptyParagraphs ?? []) {
        if (paragraph.styleId) used.add(paragraph.styleId);
        for (const run of paragraph.runs) if (run.styleId) used.add(run.styleId);
      }
  const pending = [...used];
  for (let index = 0; index < pending.length; index++) {
    const style = byId.get(pending[index]!);
    if (!style || !["paragraph", "character", "table"].includes(style.kind)) return;
    for (const id of [style.basedOn, style.nextStyleId, style.linkedStyleId])
      if (id && !used.has(id)) {
        used.add(id);
        pending.push(id);
      }
  }
  const definitions: WordClipboardStyleDefinition[] = pending.map((id) =>
    wordClipboardStyleDefinition(byId.get(id)!, css),
  );
  const knownClasses = new Set(
    definitions.flatMap((item) =>
      [
        item.style.cssClass,
        item.style.runCssClass,
        ...Object.values(item.style.components ?? {}),
      ].filter((name): name is string => !!name),
    ),
  );
  const tableIds = new Set(semantic.semantics?.tables?.map((table) => table.nodeId)),
    contextClasses = new Set(
      word.tables
        .filter((table) => tableIds.has(table.nodeId))
        .flatMap((table) => table.styleContexts?.map((context) => context.cssClass) ?? []),
    );
  if (word.defaultRunCssClass) knownClasses.add(word.defaultRunCssClass);
  const excludedClasses = new Set([...knownClasses, ...contextClasses]),
    defaults = word.defaultRunCssClass ? css.specifiedClass(word.defaultRunCssClass) : {},
    nodes: WordClipboardFormat["nodes"] = {};
  for (const [index, element] of elements.entries()) {
    const id = element.dataset.nodeId!,
      key = String(index),
      sourceNode = project(id),
      authored = [
        declarations(css.specifiedOutsideClasses(sourceNode, excludedClasses)),
        element.getAttribute("style") ?? "",
      ]
        .filter(Boolean)
        .join(";");
    nodes[key] = {
      ...(semantic.semantics
        ? {
            sourceId: id,
            ...(semantic.attributes.has(id) ? { attributes: semantic.attributes.get(id)! } : {}),
          }
        : {}),
      style: authored || null,
      class:
        [...element.classList]
          .filter((name) => knownClasses.has(name) && name !== word.defaultRunCssClass)
          .join(" ") || null,
      styleId: element.getAttribute("data-word-style"),
    };
    element.setAttribute(identity, key);
    const source = css.get(project(id));
    element.setAttribute(
      "style",
      [
        element.getAttribute("style") ?? "",
        declarations(
          Object.fromEntries(Object.entries(source).filter(([property]) => runKey(property))),
        ),
      ]
        .filter(Boolean)
        .join(";"),
    );
    if (
      word.defaultRunCssClass &&
      element.classList.contains(word.defaultRunCssClass) &&
      !element.hasAttribute("data-word-drawing") &&
      !element.closest('[data-word-drawing="group"]') &&
      !element.hasAttribute("data-word-paragraph-mark")
    ) {
      const parent = holder.ownerDocument.createElement("span");
      parent.dataset.wordRunOverlay = "true";
      element.dataset.wordRunOverlay = "true";
      parent.setAttribute(
        "style",
        declarations(
          Object.fromEntries(
            Object.entries(defaults).filter(
              ([property]) =>
                property === "background-color" ||
                property.startsWith("text-decoration") ||
                property === "vertical-align" ||
                property === "--word-run-shading",
            ),
          ),
        ),
      );
      element.before(parent);
      parent.append(element);
    }
  }
  const wrapper = holder.ownerDocument.createElement(block ? "div" : "span");
  if (!block) wrapper.dataset.wordRunOverlay = "true";
  wrapper.setAttribute("style", declarations(getWordTextDefaults(model, contextId)));
  wrapper.append(...holder.childNodes);
  holder.append(wrapper);
  const format: WordClipboardFormat = {
    format: "doffi-word-clipboard",
    version: semantic.semantics ? 2 : 1,
    ...(semantic.semantics ? { semantics: semantic.semantics } : {}),
    styles: definitions,
    nodes,
  };
  if (graph.stories.length) {
    const semantics = (format.semantics ??= {});
    format.version = 2;
    semantics.stories = graph.stories.map(({ holder, ...story }) => ({
      ...story,
      html: holder.innerHTML,
    }));
  }
  const paths = new Set(
    (semantic.semantics?.drawings ?? []).flatMap((d) =>
      d.kind === "chart" ? (d.chart.native?.parts.map((p) => p.assetPath) ?? []) : [],
    ),
  );
  let files: ReturnType<ContentDocumentModel["serialize"]> | undefined,
    total = 0;
  for (const path of paths) {
    const asset = snapshot.manifest.assets.find((a) => a.path === path),
      bytes = model.getAsset
        ? model.getAsset(path, WORD_CLIPBOARD_BYTES)
        : (files ??= model.serialize()).get(path);
    if (!asset || !bytes)
      throw new DocumentError("AssetUnavailable", "复制图表需要当前文档已授权的依赖资源。");
    total += bytes.byteLength;
    if (total > WORD_CLIPBOARD_BYTES)
      throw new DocumentError("ResourceLimitExceeded", "图表依赖超过剪贴板限制。");
    let binary = "";
    for (let at = 0; at < bytes.length; at += 32768)
      binary += String.fromCharCode(...bytes.subarray(at, at + 32768));
    (format.resources ??= []).push({ path, mimeType: asset.mimeType, data: btoa(binary) });
  }
  wrapper.setAttribute(marker, encode(format));
  if (holder.innerHTML.length > WORD_CLIPBOARD_BYTES)
    throw new DocumentError("ResourceLimitExceeded", "复制格式超过剪贴板大小限制。");
}
