import { wordCenterWrap } from "./word-center-wrap.js";
import { composeWordRowExceptions } from "./word-row-exceptions.js";
import { composeWordContextualSpacing } from "./word-contextual-spacing.js";
import { composeWordFloatingTables } from "./word-floating-tables.js";
import { wordFloatingPlacement } from "./word-floating-placement.js";
import { composeWordTextRuns } from "./word-text-runs.js";
import { composeWordLanguages } from "./word-languages.js";
import { composeWordTypography } from "./word-typography.js";
import { wordPunctuationLayout } from "./word-punctuation.js";
import { wordRunWidthLayout } from "./word-run-width.js";
import { wordTabLayout } from "./word-tabs.js";
import { composeWordNumbering } from "./word-numbering.js";
import type { WordTabLayout } from "./word-tabs.js";
import { retainFlowSelection } from "./flow-selection.js";
import { composeWordTableLayout } from "./word-table-layout.js";
import { composeWordContinuationCells } from "./word-continuation-cells.js";
import { composeWordTableBidi } from "./word-table-bidi.js";
import { composeWordTableSpacing } from "./word-table-spacing.js";
import { composeWordRowSpacing } from "./word-row-spacing.js";
import { composeWordParagraphMarks } from "./word-paragraph-marks.js";
import { composeWordNative } from "./word-native.js";
import { composeWordControls } from "./word-controls.js";
import { DocumentError, openDocument } from "@doffi/core";
import { loadWordWebFonts, resolvePreviewFonts } from "./fonts.js";
import type { FontResolution, WordWebFont } from "./fonts.js";
import type { DeepReadonly, ContentDocumentModel, WordMetadata } from "@doffi/core";
import { PreviewSource } from "./source-index.js";
import { renderLimits } from "./contracts.js";
import type { DocumentRenderReport, RenderLimits } from "./contracts.js";
import { paginate } from "./pagination.js";
import type { Pagination } from "./pagination.js";
import { createPrintDocument } from "./print.js";
import { markViewNode, ownViewStyle } from "./view-nodes.js";
import { layoutFields } from "./fields.js";
import { composeWordDrawings } from "./word-drawings.js";
import { composeWordNotes } from "./word-notes.js";
import { composeWordMath } from "./word-math.js";
import type { LayoutFields } from "./fields.js";
import { SourceProjection } from "./projection.js";
import { copyPreviewSelection } from "./clipboard.js";
import { mountPresentationPreview } from "./presentation/preview.js";
import { mountWorkbookPreview } from "./workbook/preview.js";
import type {
  WorkbookPreview,
  WorkbookPreviewInput,
  WorkbookPreviewOptions,
} from "./workbook/preview-contracts.js";

export interface DocumentPreview {
  readonly element: HTMLIFrameElement;
  readonly ready: Promise<void>;
  fonts(): readonly FontResolution[];
  layout(): DocumentRenderReport;
  update(model: ContentDocumentModel): void;
  settled(): Promise<DocumentRenderReport>;
  printDocument(): Document;
  dispose(): void;
}
export interface DocumentPreviewOptions {
  readonly title?: string;
  readonly layout?: "pages" | "continuous";
  readonly limits?: Partial<RenderLimits>;
  readonly wordWebFonts?: readonly WordWebFont[];
  /** Read-only editing projection of a registered master/layout/notes-master root. */
  readonly presentationSurfaceId?: string;
}

export function packagePath(reference: string, entrypoint: string): string {
  const parts = entrypoint.split("/").slice(0, -1);
  for (const token of reference.split("/")) {
    const value = decodeURIComponent(token);
    if (value === "..") parts.pop();
    else if (value !== ".") parts.push(value);
  }
  return parts.join("/");
}
export function decorate(
  document: Document,
  word: DeepReadonly<WordMetadata> | undefined,
  scope: ParentNode = document,
): void {
  if (!word) return;
  for (const object of scope.querySelectorAll<HTMLElement>("[data-word-native-object]")) {
    object.contentEditable = "false";
    object.setAttribute("aria-label", object.title || "保留的原生内容");
  }
  const byId = new Map(
    [...scope.querySelectorAll<HTMLElement>("[data-node-id]")].map((node) => [
      node.dataset.nodeId!,
      node,
    ]),
  );
  if ((scope as Node).nodeType === 1) {
    const root = scope as HTMLElement;
    if (root.dataset.nodeId) byId.set(root.dataset.nodeId, root);
  }
  for (const cell of scope.querySelectorAll<HTMLElement>("td,th")) {
    const direction =
      cell.dataset.wordTextDirection ||
      document
        .defaultView!.getComputedStyle(cell)
        .getPropertyValue("--word-cell-text-direction")
        .trim();
    if (direction && direction !== "lrTb") {
      cell.style.writingMode = "vertical-rl";
      if (direction === "btLr") cell.style.transform = "rotate(180deg)";
    }
  }
  for (const run of scope.querySelectorAll<HTMLElement>("[data-word-run]")) {
    const computed = document.defaultView!.getComputedStyle(run);
    const threshold = parseFloat(computed.getPropertyValue("--word-kerning-threshold"));
    if (Number.isFinite(threshold))
      run.style.fontKerning =
        (parseFloat(computed.fontSize) * 72) / 96 >= threshold ? "normal" : "none";
  }
  for (const paragraph of scope.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")) {
    const computed = document.defaultView!.getComputedStyle(paragraph);
    if (computed.getPropertyValue("--word-line-rule").trim() === "atLeast") {
      const minimum = parseFloat(computed.lineHeight) || 0;
      for (const run of [paragraph, ...paragraph.querySelectorAll<HTMLElement>("span,strong,em")]) {
        const size = parseFloat(document.defaultView!.getComputedStyle(run).fontSize) * 1.2;
        if (size > minimum) run.style.lineHeight = size + "px";
      }
    }
    if (computed.getPropertyValue("--word-distributed").trim() === "1") {
      paragraph.style.textAlignLast = "justify";
      paragraph.style.textJustify = "inter-character";
    }
    paragraph.style.hyphens =
      computed.getPropertyValue("--word-suppress-auto-hyphens").trim() === "1"
        ? "manual"
        : word.automaticHyphenation
          ? "auto"
          : "manual";
  }
  for (const paragraph of scope.querySelectorAll<HTMLElement>('[data-word-line-rule="auto"]')) {
    const typed = paragraph as HTMLElement & {
      computedStyleMap?: () => {
        get: (name: string) => { unit?: string; value?: number } | undefined;
      };
    };
    const line = typed.computedStyleMap?.().get("line-height") as
      | { unit?: string; value?: number }
      | undefined;
    let multiplier =
      line?.unit === "number" ? (line.value ?? 0) : Number(paragraph.style.lineHeight);
    const computed = document.defaultView!.getComputedStyle(paragraph);
    if ((!Number.isFinite(multiplier) || multiplier <= 0) && !line) {
      // Firefox/WebKit expose resolved pixels rather than Typed OM numbers.
      // A temporary font-size probe distinguishes a unitless inherited rule
      // from an authored absolute line height, without reimplementing cascade.
      const size = parseFloat(computed.fontSize),
        height = parseFloat(computed.lineHeight);
      if (Number.isFinite(size) && size > 0 && Number.isFinite(height)) {
        const value = paragraph.style.getPropertyValue("font-size"),
          priority = paragraph.style.getPropertyPriority("font-size");
        paragraph.style.setProperty("font-size", size * 2 + "px", "important");
        const doubled = parseFloat(document.defaultView!.getComputedStyle(paragraph).lineHeight);
        if (value) paragraph.style.setProperty("font-size", value, priority);
        else paragraph.style.removeProperty("font-size");
        if (Math.abs(doubled - height * 2) < 0.2) multiplier = height / size;
      }
    }
    if (!Number.isFinite(multiplier) || multiplier <= 0) continue;
    const probe = document.createElement("span");
    probe.style.cssText =
      "position:absolute;visibility:hidden;display:block;line-height:normal;white-space:pre";
    probe.style.fontFamily = computed.fontFamily;
    probe.style.fontSize = computed.fontSize;
    probe.style.fontWeight = computed.fontWeight;
    probe.style.fontStyle = computed.fontStyle;
    probe.textContent = "汉M";
    document.body.append(probe);
    const height = probe.getBoundingClientRect().height;
    probe.remove();
    paragraph.style.lineHeight = height * multiplier + "px";
  }
  for (const table of word.tables) {
    const element = byId.get(table.nodeId);
    if (!element) continue;
    const columns = markViewNode(document.createElement("colgroup"));
    for (const width of table.grid) {
      const column = document.createElement("col");
      column.style.width = width + "pt";
      columns.append(column);
    }
    element.prepend(columns);
  }
  for (const section of scope.querySelectorAll<HTMLElement>('[data-office-story="main"] > section'))
    section.classList.add("sdk-section");
  for (const story of word.stories) {
    const root = byId.get(story.rootNodeId);
    if (!root) continue;
    if (story.kind !== "main") {
      const heading = markViewNode(document.createElement("h2"));
      heading.className = "sdk-story-label";
      heading.textContent =
        story.kind === "header"
          ? "页眉"
          : story.kind === "footer"
            ? "页脚"
            : story.kind === "footnote"
              ? "脚注"
              : story.kind === "endnote"
                ? "尾注"
                : story.kind === "textBox"
                  ? "对象文字"
                  : "批注";
      root.before(heading);
      root.classList.add("sdk-auxiliary");
    }
  }
  for (const field of word.fields) {
    for (const nested of field.instructionFields ?? []) {
      const definition = word.fields.find((item) => item.nodeId === nested.fieldNodeId);
      for (const id of definition?.resultNodeIds ?? []) {
        const node = byId.get(id);
        if (node) {
          node.hidden = true;
          node.style.display = "none";
        }
      }
    }
  }
}

function prepareDocument(input: ContentDocumentModel): {
  model: ContentDocumentModel;
  document: Document;
  urls: string[];
  nativeUrls: Map<string, string>;
} {
  const model = openDocument(input.serialize(), { kind: "document" });
  const snapshot = model.snapshot();
  if (snapshot.manifest.kind !== "document" || snapshot.metadata.presentation !== undefined)
    throw new DocumentError("UnsupportedFormat", "This renderer supports the document flow view.");
  const files = model.serialize();
  const parsed = new DOMParser().parseFromString(snapshot.html, "text/html");
  for (const link of parsed.querySelectorAll("link")) link.remove();
  const policy = parsed.createElement("meta");
  policy.httpEquiv = "Content-Security-Policy";
  policy.content =
    "default-src 'none'; img-src blob:; style-src 'unsafe-inline'; font-src 'none'; base-uri 'none'; form-action 'none'";
  parsed.head.prepend(policy);
  const style = parsed.createElement("style");
  style.setAttribute("data-sdk-source-styles", "");
  // outerHTML leaves raw style text intact. Escape the slash in HTML end tags
  // using CSS syntax, preserving the string's value and existing escapes.
  style.textContent = snapshot.styles
    .map((item) => item.css)
    .join("\n")
    .replace(/<\/style/gi, (terminator) => "<\\" + terminator.slice(1));
  parsed.head.append(style);
  const frameStyle = parsed.createElement("style");
  const setup = snapshot.metadata.pageSetup;
  frameStyle.textContent = `html{background:#eceef1}body{margin:0;padding:24px;font-family:'Noto Sans CJK SC',serif;font-size:12pt}article{box-sizing:border-box;margin:0 auto;width:${setup.width}pt;background:white;padding:${setup.margins.top}pt ${setup.margins.right}pt ${setup.margins.bottom}pt ${setup.margins.left}pt}p,h1,h2,h3,h4,h5,h6{margin:0;font-size:inherit;font-weight:inherit;line-height:normal}td,th{padding:0;font-weight:inherit;text-align:inherit;vertical-align:top}p{min-height:1em}.sdk-section{box-sizing:border-box}.sdk-story-label{font:600 12px sans-serif!important;color:#58606b!important;margin:24px 0 8px!important;padding:12px 0 0!important;border-top:1px solid #ccc!important}.sdk-auxiliary{border-left:2px solid #ddd;padding-left:12px}[data-word-field]{background:#f5f4ee}ins{color:#006400;text-decoration:underline}del{color:#a00000;text-decoration:line-through}[data-sdk-word-text-review="insert"]{color:#006400!important;text-decoration-line:var(--sdk-word-authored-insert-decoration,underline)!important}[data-sdk-word-text-review="delete"]{color:#a00000!important;text-decoration-line:var(--sdk-word-authored-delete-decoration,line-through)!important}[data-word-break="page"]{display:block;border-top:1px dashed #bbb;margin:18px 0}.sdk-number-marker{pointer-events:none}[data-word-native-object]{display:inline-block;min-width:6em;min-height:1.5em;border:1px dashed #8290a5;background:#f1f4f8;color:transparent;position:relative;vertical-align:middle}[data-word-native-object]::after{content:attr(title);position:absolute;inset:0;padding:2px 5px;color:#526175;font:10pt sans-serif;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}[data-sdk-native-preview]::after{display:none!important}[data-word-native-boundary]{display:none!important}`;
  parsed.head.insertBefore(frameStyle, style);
  const byId = new Map(
    [...parsed.querySelectorAll<HTMLElement>("[data-node-id]")].map((node) => [
      node.dataset.nodeId!,
      node,
    ]),
  );
  let maximumWidth = setup.width;
  const sectionNodes: HTMLElement[] = [];
  for (const section of snapshot.metadata.sections) {
    const effectiveSetup = section.pageSetup ?? setup;
    const node = byId.get(section.rootNodeId);
    if (!node) continue;
    maximumWidth = Math.max(maximumWidth, effectiveSetup.width);
    const { width, margins } = effectiveSetup;
    node.classList.add("sdk-section");
    node.style.boxSizing = "border-box";
    node.style.width = width + "pt";
    node.style.margin = "0 auto";
    node.style.padding = `${margins.top}pt ${margins.right}pt ${margins.bottom}pt ${margins.left}pt`;
    sectionNodes.push(node);
  }
  const article = parsed.querySelector("article");
  if (article && sectionNodes.length && !sectionNodes.includes(article)) {
    article.style.width = maximumWidth + "pt";
    article.style.padding = "0";
    const auxiliary = parsed.createElement("div");
    auxiliary.style.padding = `0 ${setup.margins.right}pt ${setup.margins.bottom}pt ${setup.margins.left}pt`;
    const auxiliaryRoots = new Set(
      snapshot.metadata.word?.stories
        .filter((story) => story.kind !== "main")
        .map((story) => story.rootNodeId) ?? [],
    );
    for (const node of [...article.children])
      if (auxiliaryRoots.has(node.getAttribute("data-node-id") ?? "")) auxiliary.append(node);
    if (auxiliary.children.length) article.append(auxiliary);
  }
  const urls: string[] = [],
    nativeUrls = new Map<string, string>();
  const imageUrl = (path: string): string => {
    const known = nativeUrls.get(path);
    if (known) return known;
    const asset = snapshot.manifest.assets.find((item) => item.path === path)!,
      bytes = files.get(path)!,
      buffer = new ArrayBuffer(bytes.length);
    new Uint8Array(buffer).set(bytes);
    const url = URL.createObjectURL(new Blob([buffer], { type: asset.mimeType }));
    urls.push(url);
    nativeUrls.set(path, url);
    return url;
  };
  for (const fragment of snapshot.metadata.word?.nativeContent?.fragments ?? [])
    if (fragment.preview) imageUrl(fragment.preview.assetPath);
  for (const image of parsed.querySelectorAll("img")) {
    const path = packagePath(image.getAttribute("src")!, snapshot.manifest.entrypoint);
    image.src = imageUrl(path);
  }
  return { model, document: parsed, urls, nativeUrls };
}

/** Mount a validated, scriptless read-only view. The iframe DOM is never serialized. */
export function mountDocumentPreview(
  container: HTMLElement,
  input: ContentDocumentModel,
  options?: DocumentPreviewOptions,
): DocumentPreview;
export function mountDocumentPreview(
  container: HTMLElement,
  input: WorkbookPreviewInput,
  options?: WorkbookPreviewOptions,
): WorkbookPreview;
export function mountDocumentPreview(
  container: HTMLElement,
  input: ContentDocumentModel | WorkbookPreviewInput,
  options?: DocumentPreviewOptions | WorkbookPreviewOptions,
): DocumentPreview | WorkbookPreview;
export function mountDocumentPreview(
  container: HTMLElement,
  input: ContentDocumentModel | WorkbookPreviewInput,
  options: DocumentPreviewOptions | WorkbookPreviewOptions = {},
): DocumentPreview | WorkbookPreview {
  return input.kind === "workbook"
    ? mountWorkbookPreview(container, input, options as WorkbookPreviewOptions)
    : mountPreview(container, input, options as DocumentPreviewOptions);
}

/** Internal renderer connection; source positions never become a public document model. */
export function mountPreview(
  container: HTMLElement,
  input: ContentDocumentModel,
  options: DocumentPreviewOptions = {},
): DocumentPreview & { source(): PreviewSource | undefined } {
  if (input.snapshot().manifest.kind === "presentation")
    return mountPresentationPreview(container, input, options);
  if (options.presentationSurfaceId)
    throw new DocumentError(
      "InvalidOperation",
      "Definition surfaces are available only for presentations.",
    );
  const limits = renderLimits(options.limits),
    hostDocument = container.ownerDocument;
  let prepared = prepareDocument(input),
    model = prepared.model;
  const iframe = hostDocument.createElement("iframe");
  iframe.title = options.title ?? "文档只读预览";
  iframe.setAttribute("sandbox", "allow-same-origin allow-scripts");
  iframe.style.cssText = "width:100%;height:480px;border:0;display:block";
  let fontReport: readonly FontResolution[] = [],
    source: PreviewSource | undefined,
    pagination: Pagination | undefined;
  let baseFonts: readonly FontResolution[] = [],
    fields: LayoutFields | undefined;
  let projection: SourceProjection | undefined;
  let drawings: ReturnType<typeof composeWordDrawings> | undefined;
  let floatingTables: ReturnType<typeof composeWordFloatingTables> | undefined;
  let centeredWrap: ReturnType<typeof wordCenterWrap> | undefined;
  let continuationCells: (() => void) | undefined;
  let notes: ReturnType<typeof composeWordNotes> | undefined;
  let disposed = false,
    loaded = false,
    generation = 0,
    scheduled: number | undefined,
    observer: ResizeObserver | undefined;
  let report: DocumentRenderReport = Object.freeze({
    documentId: model.snapshot().manifest.documentId,
    revision: model.snapshot().manifest.revision,
    layout: options.layout ?? "pages",
    pages: [],
    nodePages: {},
    fields: [],
    diagnostics: [],
    fonts: [],
  });
  let continuousPunctuation: ReturnType<typeof wordPunctuationLayout> | undefined;
  let runWidths: ReturnType<typeof wordRunWidthLayout> | undefined;
  let tabs: WordTabLayout | undefined;
  let hostedFonts: ReadonlyMap<string, string> | undefined;
  let resolveReady: () => void = () => {},
    rejectReady: (reason: unknown) => void = () => {};
  const ready = new Promise<void>((resolve, reject) => {
    resolveReady = resolve;
    rejectReady = reject;
  });
  function measure(): void {
    if (disposed || !source) return;
    // A display:none host has no usable geometry. Keep the last accepted
    // layout instead of removing page gaps and measuring zero-sized lines.
    if (!iframe.clientWidth || !iframe.clientHeight) return;
    // IME owns temporary DOM until compositionend; splitting its text node would
    // invalidate the operating system's composition target.
    if (composing) return;
    const restoreSelection = retainFlowSelection(source.root, projection?.ignored);
    runWidths?.reset();
    composeWordContextualSpacing(iframe.contentDocument!, model);
    composeWordTypography(iframe.contentDocument!, model, limits);
    centeredWrap?.reset();
    drawings?.reset();
    if (continuousPunctuation) {
      continuousPunctuation.reset();
      continuousPunctuation.measure(source.root, false);
      runWidths?.measure(source.root);
      tabs?.measure(source.root);
    }
    if (centeredWrap?.active || continuousPunctuation) source.reindex();
    const placeFloats = (
      pages: DocumentRenderReport["pages"],
      root: HTMLElement,
      storiesOnly = false,
    ): void => {
      const placement = wordFloatingPlacement(limits);
      const finish = drawings?.measure(pages, root, placement, storiesOnly);
      floatingTables?.measure(pages, root, placement, storiesOnly);
      placement.place();
      finish?.();
    };
    let pages: DocumentRenderReport["pages"] = [],
      nodePages: DocumentRenderReport["nodePages"] = {},
      results: DocumentRenderReport["fields"] = [],
      stable = !pagination;
    for (let pass = 0; pagination && pass < limits.maxLayoutPasses; pass++) {
      pages = pagination.measure();
      pages = notes?.measure(pages) ?? pages;
      placeFloats(pages, source.root);
      centeredWrap?.measure(pages, source.root);
      if (centeredWrap?.changed) source.reindex();
      nodePages = pagination.nodePages(pages);
      results = fields?.evaluate(pages, nodePages) ?? [];
      if (
        !fields?.changed &&
        !notes?.changed &&
        !drawings?.changed &&
        !floatingTables?.changed &&
        !centeredWrap?.changed
      ) {
        stable = true;
        break;
      }
      const derivedFonts = resolvePreviewFonts(
        iframe.contentDocument!,
        ".sdk-field-display [data-word-script]",
        hostedFonts,
      );
      fontReport = Object.freeze([...baseFonts, ...derivedFonts]);
    }
    if (!pagination && (centeredWrap?.active || drawings?.active || floatingTables?.active)) {
      const origin = source.root.getBoundingClientRect(),
        snapshot = model.snapshot();
      const canvases = (snapshot.metadata.sections ?? []).flatMap((section, index) => {
        const node = source!.root.querySelector<HTMLElement>(
          '[data-node-id="' + CSS.escape(section.rootNodeId) + '"]',
        );
        if (!node) return [];
        const bounds = node.getBoundingClientRect(),
          setup = section.pageSetup ?? snapshot.metadata.pageSetup;
        if (!setup) return [];
        return [
          {
            pageId: "continuous-" + section.sectionId,
            index,
            sectionId: section.sectionId,
            number: index + 1,
            label: String(index + 1),
            blank: false,
            setup,
            rect: {
              x: (bounds.left - origin.left) * 0.75,
              y: (bounds.top - origin.top) * 0.75,
              width: setup.width,
              height: Math.max(setup.height, bounds.height * 0.75),
            },
          },
        ];
      });
      stable = false;
      for (let pass = 0; pass < limits.maxLayoutPasses; pass++) {
        source.clearInstances();
        placeFloats(canvases, source.root);
        centeredWrap?.measure(canvases, source.root);
        if (centeredWrap?.changed) source.reindex();
        if (!drawings?.changed && !floatingTables?.changed && !centeredWrap?.changed) {
          stable = true;
          break;
        }
      }
    }
    if (!stable)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Document layout did not converge within the render pass budget.",
      );
    fields?.repeat(pages);
    if (tabs && fields)
      for (const story of iframe.contentDocument!.querySelectorAll(".sdk-page-story,.sdk-repeat-header"))
        tabs.measure(story);
    if (iframe.contentDocument!.querySelector(".sdk-page-story-paint")) {
      for (let pass = 0; pass < limits.maxLayoutPasses; pass++) {
        placeFloats(pages, source.root, true);
        centeredWrap?.measure(pages, source.root, true);
        if (!centeredWrap?.storiesChanged) break;
        if (pass + 1 === limits.maxLayoutPasses)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Repeated story layout did not converge within the render pass budget.",
          );
      }
    }
    continuationCells?.();
    restoreSelection();
    source.setPages(pages);
    report = Object.freeze({
      documentId: model.snapshot().manifest.documentId,
      revision: model.snapshot().manifest.revision,
      layout: options.layout ?? "pages",
      pages,
      nodePages,
      fields: results,
      diagnostics: Object.freeze([
        ...(pagination?.diagnostics ?? []),
        ...(pagination?.storyDiagnostics() ?? []),
        ...(fields?.diagnostics ?? []),
        ...(drawings?.diagnostics ?? []),
        ...(floatingTables?.diagnostics ?? []),
      ]),
      fonts: fontReport,
    });
    const body = iframe.contentDocument!.body;
    const height =
      Math.ceil(Math.max(200, body.scrollHeight, body.getBoundingClientRect().height)) + "px";
    if (iframe.style.height !== height) iframe.style.height = height;
  }
  function render(): void {
    const frame = iframe.contentDocument!;
    runWidths?.reset();
    notes?.dispose();
    notes = undefined;
    pagination?.dispose();
    pagination = undefined;
    fields = undefined;
    source = undefined;
    const snapshot = model.snapshot();
    runWidths = snapshot.html.includes("--word-run-native-w") ||
      snapshot.styles.some(({ css }) => css.includes("--word-run-native-w"))
      ? wordRunWidthLayout(frame, limits)
      : undefined;
    composeWordTextRuns(frame, model, limits);
    composeWordParagraphMarks(frame, model);
    composeWordLanguages(frame, model, limits);
    composeWordContextualSpacing(frame, model);
    const fontSelector = "[data-node-id],[data-sdk-word-script-view]";
    const hostedResolution = hostedFonts?.size ? resolvePreviewFonts(frame, fontSelector, hostedFonts) : undefined;
    composeWordTypography(frame, model, limits);
    const sourceFonts = hostedResolution ?? resolvePreviewFonts(frame, fontSelector);
    source = new PreviewSource(frame, model);
    // Empty paragraphs reserve a line in the main story and repeated headers
    // and footers. A blank line in a bottom-anchored footer raises the preceding
    // text. The derived break carries no source newline or identity.
    for (const paragraph of [
      ...source.root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6"),
      ...frame.querySelectorAll<HTMLElement>(
        ':is([data-office-story="header"],[data-office-story="footer"]) :is(p,h1,h2,h3,h4,h5,h6)',
      ),
    ]) {
      if (paragraph.textContent || paragraph.querySelector("br,img,[data-word-tab]")) continue;
      const cell = paragraph.parentElement;
      if (
        cell?.matches("td,th") &&
        cell.lastElementChild === paragraph &&
        cell.style.getPropertyValue("--word-cell-hide-mark").trim() === "1" &&
        !paragraph.querySelector(
          "[data-word-drawing],[data-word-equation],[data-word-field],[data-word-note-reference],[data-word-native-object]:not([data-word-native-boundary])",
        )
      ) {
        // An empty final paragraph represents Word's end-of-cell marker.
        // Keep its source identity and formatting available for input.
        ownViewStyle(paragraph, "min-height", "line-height");
        paragraph.style.minHeight = "0";
        paragraph.style.lineHeight = "0";
      } else paragraph.append(markViewNode(frame.createElement("br")));
    }
    const word = snapshot.metadata.word;
    const numberingFonts = composeWordNumbering(frame, word, hostedFonts);
    decorate(frame, word);
    tabs = word?.tabs.length ? wordTabLayout(frame, word) : undefined;
    tabs?.measure(frame);
    composeWordTableSpacing(frame, model);
    const rowSpacing = composeWordRowSpacing(frame, model, limits);
    composeWordRowExceptions(frame, model, limits);
    composeWordTableBidi(frame, model);
    composeWordTableLayout(frame, limits, snapshot.metadata.word);
    continuationCells = composeWordContinuationCells(frame, model, limits, rowSpacing);
    floatingTables = composeWordFloatingTables(frame, model, limits);
    composeWordNative(frame, model, prepared.nativeUrls, limits);
    composeWordControls(frame, model);
    composeWordMath(frame, model, limits);
    if ((options.layout ?? "pages") === "pages")
      fields = layoutFields(frame, source, snapshot.metadata.word);
    drawings = composeWordDrawings(frame, model, limits, source);
    centeredWrap = wordCenterWrap(frame, model, limits);
    notes = composeWordNotes(frame, model, source, limits, (index) =>
      pagination?.contentBottom(index),
    );
    if ((options.layout ?? "pages") === "pages") {
      pagination = paginate(frame, source, snapshot.metadata, limits, notes, runWidths, tabs);
      continuousPunctuation = undefined;
    } else
      continuousPunctuation = wordPunctuationLayout(
        frame,
        snapshot.metadata.word,
        new Set(),
        limits,
      );
    baseFonts = Object.freeze([
      ...sourceFonts,
      ...numberingFonts,
      ...resolvePreviewFonts(
        frame,
        ".sdk-word-equation-view,.sdk-word-equation-view *,[data-sdk-word-chart] text,[data-sdk-word-chart] tspan",
        hostedFonts,
      ),
    ]);
    fontReport = baseFonts;
    report = Object.freeze({
      documentId: snapshot.manifest.documentId,
      revision: snapshot.manifest.revision,
      layout: options.layout ?? "pages",
      pages: [],
      nodePages: {},
      fields: [],
      diagnostics: [],
      fonts: fontReport,
    });
    measure();
    projection = new SourceProjection();
    const raw = new DOMParser()
      .parseFromString(snapshot.html, "text/html")
      .querySelector<HTMLElement>(
        '[data-node-id="' + CSS.escape(source.root.dataset.nodeId!) + '"]',
      )!;
    projection.capture(source.root, raw);
    const substitutions = fontReport.filter((item) => item.substituted),
      missing = [...new Set(fontReport.flatMap((item) => [...item.missingCharacters]))];
    if (
      substitutions.length ||
      missing.length ||
      fontReport.some((item) => !item.glyphAuditComplete)
    ) {
      const note = frame.createElement("p");
      note.setAttribute("role", "note");
      note.style.cssText = "font:12px sans-serif;color:#694800;margin:16px auto;max-width:600px";
      note.textContent =
        (substitutions.length
          ? "字体替代：" +
            substitutions
              .map((item) => item.requestedFamily + " → " + item.selectedFamily)
              .join("；") +
            "。"
          : "") +
        (missing.length
          ? "缺少字形：" +
            missing
              .map(
                (character) =>
                  "U+" +
                  character.codePointAt(0)!.toString(16).toUpperCase() +
                  "（" +
                  character +
                  "）",
              )
              .join("、") +
            "。"
          : "") +
        (fontReport.some((item) => !item.glyphAuditComplete) ? "部分字形未核验。" : "") +
        "断行和版式可能与原字体不同。";
      frame.body.append(note);
    }
    for (const image of frame.images) image.addEventListener("load", measure, { once: true });
    observer?.disconnect();
    observer = new ResizeObserver(() => {
      if (scheduled !== undefined || disposed) return;
      scheduled = hostDocument.defaultView!.requestAnimationFrame(() => {
        scheduled = undefined;
        measure();
      });
    });
    observer.observe(frame.body);
    observer.observe(iframe);
  }
  let composing = false;
  const load = (): void => {
    const frame = iframe.contentDocument;
    if (disposed || !frame) return;
    void frame.fonts.ready
      .then(async () => {
        if (options.wordWebFonts?.length && !disposed)
          hostedFonts = await loadWordWebFonts(frame, options.wordWebFonts);
        if (!disposed) {
          loaded = true;
          iframe.contentDocument!.addEventListener(
            "compositionstart",
            () => {
              composing = true;
            },
            true,
          );
          iframe.contentDocument!.addEventListener(
            "compositionend",
            () => {
              composing = false;
            },
            true,
          );
          iframe.contentDocument!.addEventListener("copy", (event) => {
            if (source && projection) copyPreviewSelection(event, source.root, projection, model);
          });
          render();
          resolveReady();
        }
      })
      .catch(rejectReady);
  };
  iframe.addEventListener("load", load, { once: true });
  iframe.srcdoc = "<!doctype html>" + prepared.document.documentElement.outerHTML;
  container.append(iframe);
  return {
    element: iframe,
    ready,
    fonts: () => fontReport,
    source: () => source,
    layout: () => report,
    update(next) {
      if (disposed) throw new DocumentError("InvalidOperation", "The renderer has been disposed.");
      if (!loaded || !iframe.contentDocument)
        throw new DocumentError(
          "InvalidOperation",
          "Wait for renderer.ready before updating the view.",
        );
      const incoming = prepareDocument(next),
        previous = prepared;
      generation++;
      observer?.disconnect();
      notes?.dispose();
      notes = undefined;
      pagination?.dispose();
      pagination = undefined;
      prepared = incoming;
      model = incoming.model;
      const frame = iframe.contentDocument;
      frame.head.replaceWith(frame.importNode(incoming.document.head, true));
      frame.body.replaceWith(frame.importNode(incoming.document.body, true));
      for (const url of previous.urls) URL.revokeObjectURL(url);
      render();
    },
    async settled() {
      const expected = generation;
      const current = (): void => {
        if (disposed || generation !== expected)
          throw new DocumentError(
            "InvalidOperation",
            "The pending render was disposed or replaced.",
          );
      };
      await ready;
      current();
      const frame = iframe.contentDocument!;
      await frame.fonts.ready;
      current();
      await Promise.all(
        [...frame.images].map(async (image) => {
          if (image.complete && image.naturalWidth) return;
          await image.decode();
        }),
      );
      current();
      measure();
      return report;
    },
    printDocument() {
      if (disposed || !loaded || !source)
        throw new DocumentError("InvalidOperation", "The renderer is not ready for printing.");
      if (report.layout !== "pages")
        throw new DocumentError("InvalidOperation", "PDF printing requires pages layout.");
      measure();
      return createPrintDocument(iframe.contentDocument!, source.root, report.pages, limits);
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      generation++;
      observer?.disconnect();
      notes?.dispose();
      notes = undefined;
      pagination?.dispose();
      source = undefined;
      projection = undefined;
      if (scheduled !== undefined) hostDocument.defaultView!.cancelAnimationFrame(scheduled);
      iframe.removeEventListener("load", load);
      iframe.remove();
      iframe.srcdoc = "";
      for (const url of prepared.urls) URL.revokeObjectURL(url);
      resolveReady();
    },
  };
}
