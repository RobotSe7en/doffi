import { registerFloatingStoryPaint } from "./word-floating-placement.js";
import { sectionFlow } from "./section-flow.js";
import { sameColumnGeometry } from "./column-regions.js";
import { cloneViewTree } from "./view-clone.js";
import { wordPunctuationLayout } from "./word-punctuation.js";
import type { WordRunWidthLayout } from "./word-run-width.js";
import type { WordTabLayout } from "./word-tabs.js";
import { DocumentError } from "@doffi/core";
import type {
  DeepReadonly,
  ContentDocumentMetadata,
  DocumentSection,
  PageSetup,
} from "@doffi/core";
import type { LayoutDiagnostic, RenderedPage, RenderLimits } from "./contracts.js";
import { markViewNode, viewNodeRole } from "./view-nodes.js";
import { repeatTableHeaders } from "./table-headers.js";
import { pageFlow } from "./flow.js";
import { columnFlow } from "./column-flow.js";
import { composeWordPageBorders } from "./word-page-borders.js";
import type { PageFlow } from "./flow.js";
import { ownViewStyle } from "./view-nodes.js";
import type { PreviewSource } from "./source-index.js";
import { applyWordGrid, wordPageSetup, facingWordPageSetup } from "./word-grid.js";
import type { WordNoteLayout } from "./word-notes.js";

const px = (value: number): number => (value * 96) / 72;
const pt = (value: number): number => (value * 72) / 96;
const gutter = 18;
export function pageLabel(value: number, format = "decimal"): string {
  if (format === "upperLetter" || format === "lowerLetter") {
    let label = "";
    for (let n = value; n > 0; n = Math.floor((n - 1) / 26))
      label = String.fromCharCode(65 + ((n - 1) % 26)) + label;
    return format === "lowerLetter" ? label.toLowerCase() : label;
  }
  if (format === "upperRoman" || format === "lowerRoman") {
    if (value > 1_000_000)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Roman page numbering exceeds the rendering budget.",
      );
    let label = "",
      n = value;
    for (const [amount, letters] of [
      [1000, "M"],
      [900, "CM"],
      [500, "D"],
      [400, "CD"],
      [100, "C"],
      [90, "XC"],
      [50, "L"],
      [40, "XL"],
      [10, "X"],
      [9, "IX"],
      [5, "V"],
      [4, "IV"],
      [1, "I"],
    ] as const)
      while (n >= amount) {
        label += letters;
        n -= amount;
      }
    return format === "lowerRoman" ? label.toLowerCase() : label;
  }
  return String(value);
}
interface SectionView {
  readonly definition: DeepReadonly<DocumentSection>;
  readonly element: HTMLElement;
}
interface Group {
  startIndex: number;
  readonly columns: string;
  readonly setup: DeepReadonly<PageSetup>;
  readonly box: HTMLElement;
  readonly flow: HTMLElement;
  readonly sections: SectionView[];
  readonly paginator: PageFlow;
}
export interface Pagination {
  measure(): readonly RenderedPage[];
  storyDiagnostics(): readonly LayoutDiagnostic[];
  contentBottom(index: number): number;
  nodePages(pages: readonly RenderedPage[]): Readonly<Record<string, readonly string[]>>;
  readonly diagnostics: readonly LayoutDiagnostic[];
  dispose(): void;
}

/** Browser-independent measured pages over a single, reversible source DOM. */
export function paginate(
  document: Document,
  source: PreviewSource,
  metadata: DeepReadonly<ContentDocumentMetadata>,
  limits: Readonly<RenderLimits>,
  notes?: WordNoteLayout,
  runWidths?: WordRunWidthLayout,
  tabs?: WordTabLayout,
): Pagination {
  if (metadata.presentation !== undefined)
    throw new DocumentError(
      "UnsupportedFeature",
      "Presentation surfaces do not use Word pagination.",
    );
  const root = source.root;
  const punctuation = wordPunctuationLayout(
    document,
    metadata.word,
    new Set(
      metadata.sections
        .filter(
          (section) =>
            section.grid?.type === "snapToChars" || section.grid?.type === "linesAndChars",
        )
        .map((section) => section.rootNodeId),
    ),
    limits,
    metadata.sections.every(section => (section.columns?.count ?? 1) === 1),
  );
  const diagnostics: LayoutDiagnostic[] = [],
    groups: Group[] = [];
  const sourceElements = new Map(
    [root, ...root.querySelectorAll<HTMLElement>("[data-node-id]")].map((node) => [
      node.dataset.nodeId!,
      node,
    ]),
  );
  if (document.querySelectorAll("*").length > limits.maxNodes)
    throw new DocumentError("ResourceLimitExceeded", "Document exceeds the render node budget.");
  const mainSections = metadata.sections
    .map((definition) => ({ definition, element: sourceElements.get(definition.rootNodeId) }))
    .filter((item): item is SectionView => Boolean(item.element && root.contains(item.element)))
    .sort((a, b) =>
      a.element === b.element ? 0 : a.element.compareDocumentPosition(b.element) & 4 ? -1 : 1,
    );
  if (!mainSections.length)
    throw new DocumentError("InvalidDocument", "The main story has no page section.");
  if (
    mainSections.some(({ element }, index) =>
      element === root
        ? mainSections.length !== 1
        : element.parentElement !== root ||
          mainSections.some(
            (other, ordinal) =>
              ordinal !== index &&
              (element.contains(other.element) || other.element.contains(element)),
          ),
    )
  )
    throw new DocumentError(
      "UnsupportedFeature",
      "Pagination requires non-overlapping section roots directly in the main story.",
    );
  if (mainSections[0]!.element !== root) {
    // Core fragment operations may place valid story content between section
    // roots. Keep its exact source order and inherit the preceding page setup
    // (or the first setup for leading content), without moving it into a source
    // section or inventing persisted metadata.
    const wrappers = mainSections.map((section) => {
      const wrapper = markViewNode(document.createElement("div"), "unwrap");
      wrapper.style.display = "flow-root";
      section.element.style.width = "auto";
      section.element.style.padding = "0";
      section.element.style.margin = "0";
      return { definition: section.definition, element: wrapper };
    });
    const ordinals = new Map(
      mainSections.map((section, ordinal) => [section.element as Node, ordinal]),
    );
    let index = 0;
    for (const node of [...root.childNodes]) {
      const ordinal = ordinals.get(node);
      if (ordinal !== undefined) index = ordinal;
      wrappers[index]!.element.append(node);
    }
    root.append(...wrappers.map((section) => section.element));
    mainSections.splice(0, mainSections.length, ...wrappers);
  }
  const grids: NonNullable<ReturnType<typeof applyWordGrid>>[] = [];
  const variableSections = new Set<SectionView>();
  notes?.attachSections(mainSections);
  for (let start = 0; start < mainSections.length; ) {
    let end = start + 1;
    const setup = JSON.stringify(
      wordPageSetup(mainSections[start]!.definition, metadata.pageSetup),
    );
    while (
      end < mainSections.length &&
      JSON.stringify(wordPageSetup(mainSections[end]!.definition, metadata.pageSetup)) === setup
    )
      end++;
    const run = mainSections.slice(start, end);
    if (
      run.length > 1 &&
      run.some(
        (section) =>
          (section.definition.columns?.count ?? 1) > 1 || section.definition.columns?.widths,
      )
    )
      for (const section of run) variableSections.add(section);
    start = end;
  }
  const maximumWidth = Math.max(
    ...mainSections.map((item) => (item.definition.pageSetup ?? metadata.pageSetup).width),
  );
  root.style.width = maximumWidth + "pt";
  root.style.padding = "0";
  root.style.margin = "0 auto";
  root.style.background = "transparent";
  root.style.position = "relative";
  const article = root.closest<HTMLElement>('[data-office-kind="document"]')!;
  article.style.width = maximumWidth + "pt";
  article.style.padding = "0";
  article.style.background = "transparent";
  article.style.position = "relative";
  article.style.zIndex = "1";
  const layoutStyle = document.createElement("style");
  layoutStyle.textContent =
    ".sdk-page-group{position:relative;box-sizing:border-box;margin:0 auto 18pt}.sdk-page-group:last-child{margin-bottom:0}.sdk-page-flow{position:relative;display:flow-root}.sdk-page-flow .sdk-section{display:flow-root}.sdk-page-layer{position:absolute;pointer-events:none;z-index:0}.sdk-paper{position:absolute;box-sizing:border-box;background:white;box-shadow:0 1px 5px #0003}.sdk-paper[data-sdk-gridlines='true']::after{content:'';position:absolute;inset:0;pointer-events:none;background-image:linear-gradient(to right,rgba(70,114,196,.12) 1px,transparent 1px),linear-gradient(to bottom,rgba(70,114,196,.12) 1px,transparent 1px);background-size:18pt 18pt;z-index:2}.sdk-line-numbers{position:absolute;left:0;top:0;bottom:0;pointer-events:none;z-index:2;font:9pt system-ui,sans-serif;color:#5c687a}.sdk-line-number{position:absolute;text-align:right;line-height:1}.sdk-page-story{position:absolute;box-sizing:border-box;overflow:hidden;pointer-events:auto}.sdk-page-story p{min-height:0}.sdk-page-story .sdk-auxiliary{display:block;border:0;padding:0}.sdk-story-label,.sdk-auxiliary{display:none}.sdk-page-story{display:block}.sdk-page-number{position:absolute;right:0;bottom:-14pt;color:#68717b;font:9px system-ui}[data-word-break]{border:0!important;margin:0!important;height:0;line-height:0;display:block}";
  document.head.append(layoutStyle);
  for (const section of mainSections) {
    const previousSection = mainSections[mainSections.indexOf(section) - 1];
    const setup = wordPageSetup(section.definition, metadata.pageSetup),
      { width, height, margins } = setup,
      changedPageGeometry =
        previousSection &&
        JSON.stringify(setup) !==
          JSON.stringify(wordPageSetup(previousSection.definition, metadata.pageSetup));
    if (
      section.definition.breakType === "nextColumn" &&
      previousSection &&
      (changedPageGeometry ||
        ((previousSection.definition.columns?.count ?? 1) > 1 &&
          !sameColumnGeometry(
            previousSection.definition.columns,
            section.definition.columns,
            width - margins.left - margins.right,
          )))
    )
      diagnostics.push({
        code: "SectionGeometryChange",
        nodeId: section.definition.rootNodeId,
        message: changedPageGeometry
          ? "A next-column section with different page geometry starts on a new page."
          : "A next-column section with different column geometry starts on a new page.",
      });
    const grid = applyWordGrid(
      document,
      section.element,
      section.definition,
      limits,
      metadata.word,
      (message) => {
        if (
          !diagnostics.some(
            (item) => item.nodeId === section.definition.rootNodeId && item.message === message,
          )
        )
          diagnostics.push({
            code: "NativeFormatting",
            nodeId: section.definition.rootNodeId,
            message,
          });
      },
    );
    if (grid) grids.push(grid);
    if (
      width - margins.left - margins.right <= 0 ||
      height - margins.top - margins.bottom <= 0 ||
      width > 20_000 ||
      height > 20_000
    )
      throw new DocumentError(
        "UnsupportedFeature",
        "Page dimensions must provide a positive printable area within the render size limit.",
      );
    let group = groups.at(-1);
    const columns = variableSections.has(section) ? "sections" : "regular";
    const samePaperContinuousMargins =
      Boolean(group) &&
      section.definition.breakType === "continuous" &&
      group!.columns === columns &&
      group!.setup.width === setup.width &&
      group!.setup.height === setup.height &&
      group!.setup.margins.top === setup.margins.top &&
      group!.setup.margins.bottom === setup.margins.bottom &&
      (group!.setup.margins.left !== setup.margins.left ||
        group!.setup.margins.right !== setup.margins.right);
    if (samePaperContinuousMargins && group) {
      ownViewStyle(section.element, "margin-left", "margin-right");
      section.element.style.marginLeft = setup.margins.left - group.setup.margins.left + "pt";
      section.element.style.marginRight = setup.margins.right - group.setup.margins.right + "pt";
    }
    if (
      !group ||
      (!samePaperContinuousMargins && JSON.stringify(group.setup) !== JSON.stringify(setup)) ||
      group.columns !== columns
    ) {
      if (changedPageGeometry && section.definition.breakType === "continuous")
        diagnostics.push({
          code: "SectionGeometryChange",
          nodeId: section.definition.rootNodeId,
          message: "A continuous section with different page geometry starts on a new page.",
        });
      const box = markViewNode(document.createElement("div"), "unwrap"),
        flow = markViewNode(document.createElement("div"), "unwrap");
      box.className = "sdk-page-group";
      flow.className = "sdk-page-flow";
      box.style.cssText = `width:${width}pt;padding:${margins.top}pt 0 ${margins.bottom}pt;overflow:clip`;
      flow.style.cssText = `width:${width - margins.left - margins.right}pt;margin-left:${margins.left}pt`;
      box.append(flow);
      const definition = section.definition.columns;
      const constraints = notes
        ? {
            reserved: (index: number) => notes.reserved(group!.startIndex + index),
            minimumPage: (node: Node) => {
              const minimum = notes.minimumPage(node);
              return minimum === undefined ? undefined : Math.max(0, minimum - group!.startIndex);
            },
          }
        : undefined;
      const balanceEnd = (section: SectionView) =>
        !metadata.word?.noColumnBalance &&
        mainSections.indexOf(section) < mainSections.length - 1 &&
        mainSections[mainSections.indexOf(section) + 1]?.definition.breakType !== "nextColumn";
      group = {
        startIndex: 0,
        setup,
        columns,
        box,
        flow,
        sections: [],
        paginator:
          columns === "sections"
            ? sectionFlow(
                document,
                flow,
                px(height - margins.top - margins.bottom),
                px(height + gutter),
                px(width - margins.left - margins.right),
                limits,
                source,
                () => group!.sections,
                balanceEnd,
                constraints,
                metadata.word?.mirrorMargins
                  ? (index) =>
                      (group!.startIndex + index) % 2 === 1
                        ? px(setup.margins.right - setup.margins.left)
                        : 0
                  : undefined,
                () => group!.startIndex,
              )
            : (definition && (definition.count > 1 || definition.widths)) ||
                metadata.word?.mirrorMargins
              ? columnFlow(
                  document,
                  flow,
                  px(height - margins.top - margins.bottom),
                  px(height + gutter),
                  px(width - margins.left - margins.right),
                  definition?.count ?? 1,
                  px(definition?.gap ?? 0),
                  definition?.separator ?? false,
                  limits,
                  source,
                  constraints,
                  metadata.word?.mirrorMargins
                    ? (index) =>
                        (group!.startIndex + index) % 2 === 1
                          ? px(setup.margins.right - setup.margins.left)
                          : 0
                    : undefined,
                  definition?.widths?.map(px),
                  definition?.spaces?.map(px),
                  {
                    firstPageIndex: () => group!.startIndex,
                    balanceEnd: () => balanceEnd(group!.sections.at(-1)!),
                  },
                )
              : pageFlow(
                  document,
                  flow,
                  px(height - margins.top - margins.bottom),
                  px(height + gutter),
                  limits,
                  1,
                  constraints,
                ),
      };
      groups.push(group);
      if (section.element === root) {
        flow.append(...root.childNodes);
        root.append(box);
      } else root.insertBefore(box, section.element);
    }
    group.sections.push(section);
    if (section.element !== root) {
      const node = section.element;
      node.style.width = "auto";
      node.style.padding = "0";
      node.style.margin = "0";
      node.style.boxSizing = "border-box";
      node.style.breakBefore = "auto";
      group.flow.append(node);
    }
  }
  const layer = markViewNode(document.createElement("div"));
  layer.className = "sdk-page-layer";
  article.before(layer);
  const front = markViewNode(document.createElement("div"));
  front.className = "sdk-page-border-layer";
  front.style.cssText = "position:absolute;pointer-events:none;z-index:3";
  article.after(front);
  const storyTemplates = new Map(
    (metadata.word?.stories ?? []).map((story) => [
      story.rootNodeId,
      document.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(story.rootNodeId) + '"]'),
    ]),
  );
  const repeated = repeatTableHeaders(document, root, source, limits);
  let pages: readonly RenderedPage[] = [],
    disposed = false;
  const sectionStarts = (group: Group): number[] =>
    group.sections.map(({ element }) => (element === root ? 0 : group.paginator.pageAt(element)));
  function measure(): readonly RenderedPage[] {
    if (disposed) return pages;
    source.clearInstances();
    repeated.clear();
    for (const group of groups) group.paginator.reset();
    runWidths?.reset();
    punctuation.reset();
    for (const grid of grids) grid.measure();
    for (const group of groups) {
      punctuation.measure(group.flow);
      runWidths?.measure(group.flow);
      tabs?.measure(group.flow);
    }
    let physical = 0;
    for (const group of groups) {
      group.startIndex = physical;
      for (const section of group.sections) {
        const kind = section.definition.breakType,
          first = section === mainSections[0];
        const parity =
          !first && (kind === "oddPage" || kind === "evenPage")
            ? ((kind === "oddPage" ? 0 : 1) + physical) % 2
            : undefined;
        group.paginator.layout(
          section.element === root ? group.flow : section.element,
          !first &&
            kind !== "continuous" &&
            (kind !== "nextColumn" || group.columns !== "sections"),
          parity,
        );
      }
      physical += group.paginator.count;
      if (physical > limits.maxPages)
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Document exceeds the render page budget.",
        );
      group.paginator.project?.();
    }
    source.reindex();
    const result: RenderedPage[] = [],
      surface = root.getBoundingClientRect();
    let nextNumber = 1;
    const startsBySection = new Map<string, number>(),
      previousRefs: { headers: Record<string, string>; footers: Record<string, string> } = {
        headers: {},
        footers: {},
      };
    const refs = new Map<
      string,
      { headers: Record<string, string>; footers: Record<string, string> }
    >();
    for (const section of mainSections) {
      Object.assign(previousRefs.headers, section.definition.headers);
      Object.assign(previousRefs.footers, section.definition.footers);
      refs.set(section.definition.sectionId, {
        headers: { ...previousRefs.headers },
        footers: { ...previousRefs.footers },
      });
    }
    layer.replaceChildren();
    front.replaceChildren();
    layer.style.left = surface.left - document.body.getBoundingClientRect().left + "px";
    layer.style.top = surface.top - document.body.getBoundingClientRect().top + "px";
    layer.style.width = maximumWidth + "pt";
    front.style.left = layer.style.left;
    front.style.top = layer.style.top;
    front.style.width = layer.style.width;
    let renderNodes = document.querySelectorAll("*").length;
    for (const group of groups) {
      const { setup } = group,
        starts = sectionStarts(group),
        pitch = px(setup.height + gutter);
      const count =
        group === groups.at(-1)
          ? Math.max(group.paginator.count, (notes?.requiredPages ?? 0) - group.startIndex)
          : group.paginator.count;
      group.box.style.minHeight = count * setup.height + (count - 1) * gutter + "pt";
      const box = group.box.getBoundingClientRect();
      for (let index = 0; index < count; index++) {
        if (result.length >= limits.maxPages)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Document exceeds the render page budget.",
          );
        let owner = 0;
        for (let candidate = 1; candidate < starts.length; candidate++)
          if (starts[candidate]! <= index) owner = candidate;
        let definition = group.sections[owner]!.definition;
        const blank = group.paginator.blanks.has(index)
          ? group.sections.find((_section, ordinal) => index === starts[ordinal]! - 1)
          : undefined;
        if (blank)
          definition = mainSections[Math.max(0, mainSections.indexOf(blank) - 1)]!.definition;
        const first = !blank && !startsBySection.has(definition.sectionId);
        if (first) {
          startsBySection.set(definition.sectionId, result.length);
          if (definition.pageNumberStart !== undefined) nextNumber = definition.pageNumberStart;
        }
        const number = nextNumber++,
          page: RenderedPage = Object.freeze({
            pageId: "page-" + (result.length + 1),
            index: result.length,
            sectionId: definition.sectionId,
            number,
            label: pageLabel(number, definition.pageNumberFormat),
            rect: Object.freeze({
              x: pt(box.left - surface.left),
              y: pt(box.top - surface.top + index * pitch),
              width: setup.width,
              height: setup.height,
            }),
            setup: facingWordPageSetup(
              setup,
              result.length,
              metadata.word?.mirrorMargins ?? false,
            ) as PageSetup,
            blank: Boolean(blank),
          });
        result.push(page);
        const paper = document.createElement("div");
        paper.className = "sdk-paper";
        paper.dataset.pageId = page.pageId;
        paper.setAttribute("aria-hidden", "true");
        const bgColor = metadata.word?.pageBackground?.color
          ? "#" + metadata.word.pageBackground.color
          : "white";
        paper.style.cssText = `left:${page.rect.x}pt;top:${page.rect.y}pt;width:${setup.width}pt;height:${setup.height}pt;background:${bgColor}`;
        if (document.documentElement.dataset.wordGridlines === "true")
          paper.dataset.sdkGridlines = "true";
        layer.append(paper);
        let storyPaint: HTMLElement | undefined;
        const paint = (node: HTMLElement, story: HTMLElement): void => {
          if (!storyPaint) {
            if (++renderNodes > limits.maxNodes)
              throw new DocumentError(
                "ResourceLimitExceeded",
                "Repeated stories exceed the render node budget.",
              );
            storyPaint = markViewNode(document.createElement("div"));
            storyPaint.className = "sdk-page-story-paint";
            storyPaint.style.cssText =
              "position:absolute;inset:0;overflow:hidden;pointer-events:none";
            paper.append(storyPaint);
          }
          registerFloatingStoryPaint(node, story);
          storyPaint.append(node);
        };
        for (const kind of ["headers", "footers"] as const) {
          if (blank) continue;
          const variant =
            first && definition.differentFirstPage
              ? "first"
              : metadata.word?.evenAndOddHeaders && number % 2 === 0
                ? "even"
                : "default";
          const reference = refs.get(definition.sectionId)?.[kind][variant];
          const id = metadata.word?.stories.find(
            (story) => story.storyId === reference,
          )?.rootNodeId;
          const template = id ? storyTemplates.get(id) : null;
          if (!template) continue;
          renderNodes += template.querySelectorAll("*").length + 1;
          if (renderNodes > limits.maxNodes)
            throw new DocumentError(
              "ResourceLimitExceeded",
              "Repeated stories exceed the render node budget.",
            );
          const story = cloneViewTree(template, limits.maxNodes);
          story.classList.remove("sdk-auxiliary");
          story.classList.add("sdk-page-story");
          story.dataset.sdkAuxiliary = "true";
          if (reference) story.dataset.sdkPageStoryId = reference;
          story.dataset.sdkPageStoryKind = kind === "headers" ? "header" : "footer";
          const distance =
            kind === "headers"
              ? (definition.headerDistance ?? 36)
              : (definition.footerDistance ?? 36);
          const available = Math.max(
            0,
            (kind === "headers" ? setup.margins.top : setup.margins.bottom) - distance,
          );
          story.style.cssText += `;left:${page.setup.margins.left}pt;width:${setup.width - setup.margins.left - setup.margins.right}pt;max-height:${available}pt;${kind === "headers" ? "top:" : "bottom:"}${distance}pt`;
          paper.append(story);
          story.style.zIndex = "1";
          punctuation.measure(story, false);
          runWidths?.measure(story);
          tabs?.measure(story);
          source.registerInstances(story);
          for (const object of [...story.querySelectorAll<HTMLElement>("[data-sdk-word-object]")]) {
            const definition = metadata.word?.drawings?.find(
              (drawing) => drawing.nodeId === object.dataset.sdkWordObject,
            );
            if (
              definition?.layout.mode !== "floating" ||
              object.previousElementSibling?.getAttribute("data-sdk-word-anchor") !==
                definition.nodeId
            )
              continue;
            const paragraph = object.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6");
            if (!paragraph?.dataset.nodeId) continue;
            object.dataset.sdkWordStoryAnchor = paragraph.dataset.nodeId;
            object.style.zIndex = definition.layout.behindText ? "0" : "2";
            paint(object, story);
          }
          for (const table of story.querySelectorAll<HTMLElement>(".sdk-word-floating-table"))
            if (viewNodeRole(table) === "unwrap") paint(table, story);
        }
        renderNodes += composeWordPageBorders(
          document,
          article,
          paper,
          front,
          page,
          definition.pageBorders,
          first,
          metadata.word,
        );
        if (renderNodes > limits.maxNodes)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Page borders exceed the rendering node budget.",
          );
        const label = document.createElement("span");
        label.className = "sdk-page-number";
        label.textContent = String(page.index + 1);
        paper.append(label);
        layer.append(paper);
      }
    }
    // Paint page line numbers (w:lnNumType) and section vertical alignment (w:vAlign).
    let runningLineNumber = 1;
    let lastSectionId = "";
    for (let pageIdx = 0; pageIdx < result.length; pageIdx++) {
      const page = result[pageIdx]!;
      if (page.blank) continue;
      const sectionDef = mainSections.find((s) => s.definition.sectionId === page.sectionId)?.definition;
      if (!sectionDef) continue;
      if (sectionDef.verticalAlign && sectionDef.verticalAlign !== "top" && result.length === 1) {
        const group = groups[0];
        if (group) {
          const printablePt = page.setup.height - page.setup.margins.top - page.setup.margins.bottom;
          const usedPt = Math.max(0, group.paginator.bottom(0) * 0.75);
          const sparePt = Math.max(0, printablePt - usedPt);
          if (sparePt > 1) {
            ownViewStyle(group.flow, "padding-top");
            group.flow.style.paddingTop =
              sectionDef.verticalAlign === "center"
                ? sparePt / 2 + "pt"
                : sectionDef.verticalAlign === "bottom"
                  ? sparePt + "pt"
                  : "0pt";
          }
        }
      }
      if (sectionDef.lineNumbers) {
        const ln = sectionDef.lineNumbers;
        const restart = ln.restart ?? "newPage";
        const start = ln.start ?? 1;
        const countBy = Math.max(1, ln.countBy ?? 1);
        const distance = ln.distance ?? 18;
        if (restart === "newPage" || (restart === "newSection" && page.sectionId !== lastSectionId) || !lastSectionId) {
          runningLineNumber = start;
        }
        const paper = layer.querySelector<HTMLElement>(`.sdk-paper[data-page-id="${CSS.escape(page.pageId)}"]`);
        if (paper) {
          const topPx = surface.top + px(page.rect.y + page.setup.margins.top) - 1;
          const bottomPx = surface.top + px(page.rect.y + page.setup.height - page.setup.margins.bottom) + 1;
          const lineBands: { top: number; bottom: number }[] = [];
          const range = document.createRange();
          const walker = document.createTreeWalker(root, 4);
          for (let n = walker.nextNode(); n; n = walker.nextNode()) {
            const el = n.parentElement;
            if (
              !n.textContent?.trim() ||
              !el ||
              el.closest(
                ".sdk-auxiliary,.sdk-number-marker,.sdk-word-ruby-rt,.sdk-word-enclosed,.sdk-word-drop-cap,[data-sdk-word-drop-cap=\"drop\"],[data-sdk-word-drop-cap=\"margin\"],table",
              )
            )
              continue;
            const paraBlock = el.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6");
            if (
              paraBlock &&
              document.defaultView!.getComputedStyle(paraBlock).getPropertyValue("--word-paragraph-native-suppressLineNumbers").trim() === "1"
            )
              continue;
            range.selectNodeContents(n);
            for (const rect of range.getClientRects()) {
              if (rect.height < 2 || rect.width < 1 || rect.top < topPx || rect.bottom > bottomPx) continue;
              const match = lineBands.find((band) => {
                const overlap = Math.min(band.bottom, rect.bottom) - Math.max(band.top, rect.top);
                const minH = Math.min(band.bottom - band.top, rect.bottom - rect.top);
                return overlap > minH * 0.35;
              });
              if (match) {
                match.top = Math.min(match.top, rect.top);
                match.bottom = Math.max(match.bottom, rect.bottom);
              } else {
                lineBands.push({ top: rect.top, bottom: rect.bottom });
              }
            }
          }
          lineBands.sort((a, b) => a.top - b.top);
          if (lineBands.length) {
            const container = markViewNode(document.createElement("div"));
            container.className = "sdk-line-numbers";
            const rightEdge = Math.max(8, page.setup.margins.left - distance);
            container.style.width = rightEdge + "pt";
            const paperTop = surface.top + px(page.rect.y);
            for (const band of lineBands) {
              const num = runningLineNumber++;
              if (num % countBy === 0) {
                const span = document.createElement("span");
                span.className = "sdk-line-number";
                span.textContent = String(num);
                const mid = (band.top + band.bottom) / 2;
                span.style.top = pt(mid - paperTop - 5) + "pt";
                span.style.right = "0pt";
                container.append(span);
              }
            }
            paper.append(container);
          }
        }
      }
      lastSectionId = page.sectionId;
    }
    pages = Object.freeze(result);
    repeated.project(pages);
    // Clip source painting to printable areas; borders and merged-cell backgrounds
    // cannot bleed into margins or the inter-page gutter. Stories are siblings.
    for (const group of groups) {
      const contours = Array.from({ length: group.paginator.count }, (_, index) =>
        group.paginator.blanks.has(index)
          ? ""
          : `M ${-px(group.setup.margins.left)} ${px(index * (group.setup.height + gutter))} h ${px(group.setup.width)} v ${px(group.setup.height - group.setup.margins.top - group.setup.margins.bottom) - (notes?.reserved(group.startIndex + index) ?? 0)} h ${-px(group.setup.width)} Z`,
      ).join(" ");
      ownViewStyle(group.flow, "clip-path");
      group.flow.style.clipPath = `path("${contours}")`;
    }
    return pages;
  }
  return {
    measure,
    diagnostics,
    storyDiagnostics() {
      const overflow = new Set<string>();
      for (const story of layer.querySelectorAll<HTMLElement>(".sdk-page-story")) {
        const id = story.dataset.nodeId;
        if (id && story.scrollHeight > px(parseFloat(story.style.maxHeight)) + 1) overflow.add(id);
      }
      return [...overflow].map((nodeId) => ({
        code: "ContentOverflow",
        nodeId,
        message: "A header or footer exceeds the available margin area and is clipped.",
      }));
    },
    contentBottom(index) {
      const group = groups.find(
        (group) => index >= group.startIndex && index < group.startIndex + group.paginator.count,
      );
      return group
        ? group.setup.margins.top + group.paginator.bottom(index - group.startIndex) * 0.75
        : (pages[index]?.setup.margins.top ?? 0);
    },
    nodePages(current) {
      const surface = root.getBoundingClientRect(),
        result: Record<string, readonly string[]> = Object.create(null) as Record<
          string,
          readonly string[]
        >;
      const drawings = new Set(metadata.word?.drawings?.map((drawing) => drawing.nodeId));
      const bands = (paper: boolean) => {
        const intervals = current.flatMap((page, index) => {
          if (page.blank) return [];
          const { y, height } = page.rect, margins = page.setup.margins;
          return [{
            index,
            top: surface.top + px(y + (paper ? 0 : margins.top)) + 0.1,
            bottom: surface.top + px(y + height - (paper ? 0 : margins.bottom)) - 0.1,
          }];
        }).sort((a, b) => a.top - b.top);
        let bottom = -Infinity;
        return intervals.map(interval => ({ ...interval, furthestBottom: bottom = Math.max(bottom, interval.bottom) }));
      };
      const body = bands(false), paper = drawings.size ? bands(true) : body;
      for (const [id, rectangles] of source.elementRectangleEntries(sourceElements.keys())) {
        const intervals = drawings.has(id) ? paper : body, matches = new Set<number>();
        for (const { top, bottom } of rectangles) {
          let low = 0, high = intervals.length;
          while (low < high) {
            const mid = (low + high) >>> 1;
            if (intervals[mid]!.furthestBottom <= top) low = mid + 1;
            else high = mid;
          }
          for (let i = low; i < intervals.length && intervals[i]!.top < bottom; i++) {
            const interval = intervals[i]!;
            if (interval.bottom > top) matches.add(interval.index);
          }
        }
        result[id] = Object.freeze([...matches].sort((a, b) => a - b).map(index => current[index]!.pageId));
      }
      return Object.freeze(result);
    },
    dispose() {
      disposed = true;
      for (const grid of grids) grid.dispose();
      repeated.dispose();
      layer.remove();
      front.remove();
      layoutStyle.remove();
    },
  };
}
