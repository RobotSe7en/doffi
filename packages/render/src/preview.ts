import { DocumentError, openDocument } from "@doffi/core";
import { resolvePreviewFonts } from "./fonts.js";
import type { FontResolution } from "./fonts.js";
import type { DeepReadonly, ContentDocumentModel, WordMetadata } from "@doffi/core";
import { PreviewSource } from "./source-index.js";
import { renderLimits } from "./contracts.js";
import type { DocumentRenderReport, RenderLimits } from "./contracts.js";
import { paginate } from "./pagination.js";
import type { Pagination } from "./pagination.js";
import { createPrintDocument } from "./print.js";
import { markViewNode } from "./view-nodes.js";
import { layoutFields } from "./fields.js";
import type { LayoutFields } from "./fields.js";
import { SourceProjection } from "./projection.js";
import { copyPreviewSelection } from "./clipboard.js";
import { mountPresentationPreview } from "./presentation/preview.js";
import { mountWorkbookPreview } from "./workbook/preview.js";
import type { WorkbookPreview, WorkbookPreviewInput, WorkbookPreviewOptions } from "./workbook/preview-contracts.js";

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
  /** Read-only editing projection of a registered master/layout/notes-master root. */
  readonly presentationSurfaceId?:string;
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
function numeral(value: number, format: string): string {
  if (format === "none") return "";
  if (format === "upperLetter" || format === "lowerLetter") {
    let result = "";
    for (let n = value; n > 0; n = Math.floor((n - 1) / 26))
      result = String.fromCharCode(65 + ((n - 1) % 26)) + result;
    return format === "lowerLetter" ? result.toLowerCase() : result;
  }
  if (format === "upperRoman" || format === "lowerRoman") {
    if (value > 1_000_000)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "A derived Roman numbering label exceeds the preview budget.",
      );
    let result = "",
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
    ] as const) {
      while (n >= amount) {
        result += letters;
        n -= amount;
      }
    }
    return format === "lowerRoman" ? result.toLowerCase() : result;
  }
  return String(value);
}
export function decorate(
  document: Document,
  word: DeepReadonly<WordMetadata> | undefined,
  scope: ParentNode = document,
): void {
  if (!word) return;
  const byId = new Map(
    [...scope.querySelectorAll<HTMLElement>("[data-node-id]")].map(
      (node) => [node.dataset.nodeId!, node],
    ),
  );
  if ((scope as Node).nodeType === 1) {
    const root = scope as HTMLElement;
    if (root.dataset.nodeId) byId.set(root.dataset.nodeId, root);
  }
  for (const run of scope.querySelectorAll<HTMLElement>("[data-word-run]")) {
    const computed = document.defaultView!.getComputedStyle(run);
    const threshold = parseFloat(
      computed.getPropertyValue("--word-kerning-threshold"),
    );
    if (Number.isFinite(threshold))
      run.style.fontKerning =
        (parseFloat(computed.fontSize) * 72) / 96 >= threshold
          ? "normal"
          : "none";
  }
  for (const paragraph of scope.querySelectorAll<HTMLElement>(
    "p,h1,h2,h3,h4,h5,h6",
  )) {
    const computed = document.defaultView!.getComputedStyle(paragraph);
    paragraph.style.hyphens =
      computed.getPropertyValue("--word-suppress-auto-hyphens").trim() === "1"
        ? "manual"
        : word.automaticHyphenation
          ? "auto"
          : "manual";
  }
  for (const paragraph of scope.querySelectorAll<HTMLElement>(
    '[data-word-line-rule="auto"]',
  )) {
    const typed = paragraph as HTMLElement & {
      computedStyleMap?: () => {
        get: (name: string) => { unit?: string; value?: number } | undefined;
      };
    };
    const line = typed.computedStyleMap?.().get("line-height") as
      | { unit?: string; value?: number }
      | undefined;
    let multiplier =
      line?.unit === "number"
        ? (line.value ?? 0)
        : Number(paragraph.style.lineHeight);
    const computed = document.defaultView!.getComputedStyle(paragraph);
    if ((!Number.isFinite(multiplier) || multiplier <= 0) && !line) {
      // Firefox/WebKit expose resolved pixels rather than Typed OM numbers.
      // A temporary font-size probe distinguishes a unitless inherited rule
      // from an authored absolute line height, without reimplementing cascade.
      const size = parseFloat(computed.fontSize), height = parseFloat(computed.lineHeight);
      if (Number.isFinite(size) && size > 0 && Number.isFinite(height)) {
        const value = paragraph.style.getPropertyValue("font-size"), priority = paragraph.style.getPropertyPriority("font-size");
        paragraph.style.setProperty("font-size", size * 2 + "px", "important");
        const doubled = parseFloat(document.defaultView!.getComputedStyle(paragraph).lineHeight);
        if (value) paragraph.style.setProperty("font-size", value, priority); else paragraph.style.removeProperty("font-size");
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
  for (const section of scope.querySelectorAll<HTMLElement>(
    '[data-office-story="main"] > section',
  ))
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
            : "批注";
      root.before(heading);
      root.classList.add("sdk-auxiliary");
    }
    const counters = new Map<string, number[]>();
    for (const paragraph of root.querySelectorAll<HTMLElement>(
      "[data-word-numbering]",
    )) {
      const definition = word.numbering.find(
        (item) => item.numberingId === paragraph.dataset.wordNumbering,
      );
      const level = Number(paragraph.dataset.wordLevel);
      const rule = definition?.levels.find((item) => item.level === level);
      if (!definition || !rule) continue;
      const values = counters.get(definition.numberingId) ?? [];
      values[level] = (values[level] ?? rule.start - 1) + 1;
      for (const deeper of definition.levels)
        if (
          deeper.level > level &&
          deeper.restart !== 0 &&
          (deeper.restart === undefined || deeper.restart - 1 >= level)
        )
          values[deeper.level] = deeper.start - 1;
      counters.set(definition.numberingId, values);
      const marker = markViewNode(document.createElement("span"));
      marker.className = "sdk-number-marker " + rule.cssClass;
      marker.setAttribute("aria-hidden", "true");
      marker.textContent = rule.text.replace(
        /%([1-9])/g,
        (_match, token: string) => {
          const index = Number(token) - 1;
          const format = definition.levels.find((item) => item.level === index);
          return numeral(
            values[index] ?? format?.start ?? 1,
            format?.format ?? "decimal",
          );
        },
      );
      if (rule.suffix === "space") marker.textContent += " ";
      const computed = document.defaultView!.getComputedStyle(paragraph);
      const hanging = Math.abs(parseFloat(computed.textIndent) || 24);
      paragraph.style.position = "relative";
      paragraph.style.textIndent = "0";
      marker.style.cssText +=
        ";position:absolute;left:-" +
        hanging +
        "px;width:" +
        hanging +
        "px;margin:0;padding:0;text-indent:0;white-space:pre";
      paragraph.prepend(marker);
    }
  }
  for (const field of word.fields) {
    for (const nested of field.instructionFields ?? []) {
      const definition = word.fields.find(
        (item) => item.nodeId === nested.fieldNodeId,
      );
      for (const id of definition?.resultNodeIds ?? []) {
        const node = byId.get(id);
        if (node) {
          node.hidden = true;
          node.style.display = "none";
        }
      }
    }
  }
  for (const definition of word.tabs) {
    const paragraph = byId.get(definition.paragraphNodeId);
    if (!paragraph) continue;
    // Word tab stops are relative to the text margin, before paragraph indent.
    const indent = parseFloat(
      document.defaultView!.getComputedStyle(paragraph).marginLeft,
    ) || 0;
    for (const tab of paragraph.querySelectorAll<HTMLElement>(
      "[data-word-tab]",
    )) {
      if (tab.closest("[hidden]")) continue;
      const start = paragraph.getBoundingClientRect();
      const box = tab.getBoundingClientRect();
      const cursor = ((box.left - start.left + indent) * 72) / 96;
      const stop = [...definition.stops]
        .filter((item) => item.alignment !== "clear" && item.position > cursor)
        .sort((a, b) => a.position - b.position)[0];
      const target =
        stop?.position ??
        (Math.floor(cursor / definition.defaultInterval) + 1) *
          definition.defaultInterval;
      let followingWidth = 0,
        decimalWidth = 0;
      const range = document.createRange();
      range.setStartAfter(tab);
      range.setEnd(paragraph, paragraph.childNodes.length);
      const following = range.cloneContents().textContent?.split("\t")[0] ?? "";
      const canvas = document.createElement("canvas");
      const context = canvas.getContext("2d")!;
      context.font = document.defaultView!.getComputedStyle(tab).font;
      followingWidth = (context.measureText(following).width * 72) / 96;
      decimalWidth =
        (context.measureText(following.split(/[.,]/)[0] ?? "").width * 72) / 96;
      const adjustment =
        stop?.alignment === "right"
          ? followingWidth
          : stop?.alignment === "center"
            ? followingWidth / 2
            : stop?.alignment === "decimal"
              ? decimalWidth
              : 0;
      const width = Math.max(0, target - cursor - adjustment);
      tab.textContent = "";
      tab.style.display = "inline-block";
      tab.style.width = width + "pt";
      tab.style.whiteSpace = "nowrap";
      tab.style.overflow = "hidden";
      tab.style.verticalAlign = "bottom";
      const leader = stop?.leader;
      if (leader && leader !== "none")
        tab.textContent = (
          leader === "dot"
            ? "."
            : leader === "hyphen"
              ? "-"
              : leader === "middleDot"
                ? "·"
                : "_"
        ).repeat(Math.min(2000, Math.ceil(width)));
    }
  }
}

function prepareDocument(input: ContentDocumentModel): { model: ContentDocumentModel; document: Document; urls: string[] } {
  const model = openDocument(input.serialize(), { kind: "document" });
  const snapshot = model.snapshot();
  if (snapshot.manifest.kind !== "document" || snapshot.metadata.presentation !== undefined) throw new DocumentError("UnsupportedFormat", "This renderer supports the document flow view.");
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
  style.textContent = snapshot.styles.map((item) => item.css).join("\n")
    .replace(/<\/style/gi, (terminator) => "<\\" + terminator.slice(1));
  parsed.head.append(style);
  const frameStyle = parsed.createElement("style");
  const setup = snapshot.metadata.pageSetup;
  frameStyle.textContent = `html{background:#eceef1}body{margin:0;padding:24px;font-family:'Noto Sans CJK SC',serif;font-size:12pt}article{box-sizing:border-box;margin:0 auto;width:${setup.width}pt;background:white;padding:${setup.margins.top}pt ${setup.margins.right}pt ${setup.margins.bottom}pt ${setup.margins.left}pt}p,h1,h2,h3,h4,h5,h6{margin:0;font-weight:inherit;line-height:normal}td,th{padding:0;font-weight:inherit;text-align:inherit;vertical-align:top}p{min-height:1em}.sdk-section{box-sizing:border-box}.sdk-story-label{font:600 12px sans-serif!important;color:#58606b!important;margin:24px 0 8px!important;padding:12px 0 0!important;border-top:1px solid #ccc!important}.sdk-auxiliary{border-left:2px solid #ddd;padding-left:12px}[data-word-field]{background:#f5f4ee}ins{color:#006400;text-decoration:underline}del{color:#a00000;text-decoration:line-through}[data-word-break="page"]{display:block;border-top:1px dashed #bbb;margin:18px 0}.sdk-number-marker{pointer-events:none}`;
  parsed.head.insertBefore(frameStyle, style);
  const byId = new Map(
    [...parsed.querySelectorAll<HTMLElement>("[data-node-id]")].map(
      (node) => [node.dataset.nodeId!, node],
    ),
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
    const auxiliaryRoots = new Set(snapshot.metadata.word?.stories.filter((story) => story.kind !== "main").map((story) => story.rootNodeId) ?? []);
    for (const node of [...article.children])
      if (auxiliaryRoots.has(node.getAttribute("data-node-id") ?? ""))
        auxiliary.append(node);
    if (auxiliary.children.length) article.append(auxiliary);
  }
  const urls: string[] = [];
  for (const image of parsed.querySelectorAll("img")) {
    const path = packagePath(
      image.getAttribute("src")!,
      snapshot.manifest.entrypoint,
    );
    const asset = snapshot.manifest.assets.find((item) => item.path === path)!;
    const bytes = files.get(path)!;
    const buffer = new ArrayBuffer(bytes.length);
    new Uint8Array(buffer).set(bytes);
    const url = URL.createObjectURL(
      new Blob([buffer], { type: asset.mimeType }),
    );
    urls.push(url);
    image.src = url;
  }
  return { model, document: parsed, urls };
}

/** Mount a validated, scriptless read-only view. The iframe DOM is never serialized. */
export function mountDocumentPreview(
  container: HTMLElement,
  input: ContentDocumentModel,
  options?: DocumentPreviewOptions,
): DocumentPreview;
export function mountDocumentPreview(container: HTMLElement, input: WorkbookPreviewInput, options?: WorkbookPreviewOptions): WorkbookPreview;
export function mountDocumentPreview(container: HTMLElement, input: ContentDocumentModel | WorkbookPreviewInput, options?: DocumentPreviewOptions | WorkbookPreviewOptions): DocumentPreview | WorkbookPreview;
export function mountDocumentPreview(container: HTMLElement, input: ContentDocumentModel | WorkbookPreviewInput, options: DocumentPreviewOptions | WorkbookPreviewOptions = {}): DocumentPreview | WorkbookPreview {
  return input.kind === "workbook" ? mountWorkbookPreview(container, input, options as WorkbookPreviewOptions) : mountPreview(container, input, options as DocumentPreviewOptions);
}

/** Internal renderer connection; source positions never become a public document model. */
export function mountPreview(
  container: HTMLElement,
  input: ContentDocumentModel,
  options: DocumentPreviewOptions = {},
): DocumentPreview & { source(): PreviewSource | undefined } {
  if (input.snapshot().manifest.kind === "presentation") return mountPresentationPreview(container, input, options);
  if(options.presentationSurfaceId)throw new DocumentError("InvalidOperation","Definition surfaces are available only for presentations.");
  const limits = renderLimits(options.limits), hostDocument = container.ownerDocument;
  let prepared = prepareDocument(input), model = prepared.model;
  const iframe = hostDocument.createElement("iframe");
  iframe.title = options.title ?? "文档只读预览";
  iframe.setAttribute("sandbox", "allow-same-origin allow-scripts");
  iframe.style.cssText = "width:100%;height:480px;border:0;display:block";
  let fontReport: readonly FontResolution[] = [], source: PreviewSource | undefined, pagination: Pagination | undefined;
  let baseFonts: readonly FontResolution[] = [], fields: LayoutFields | undefined;
  let projection: SourceProjection | undefined;
  let disposed = false, loaded = false, generation = 0, scheduled: number | undefined, observer: ResizeObserver | undefined;
  let report: DocumentRenderReport = Object.freeze({ documentId: model.snapshot().manifest.documentId, revision: model.snapshot().manifest.revision,
    layout: options.layout ?? "pages", pages: [], nodePages: {}, fields: [], diagnostics: [], fonts: [] });
  let resolveReady: () => void = () => {}, rejectReady: (reason: unknown) => void = () => {};
  const ready = new Promise<void>((resolve, reject) => { resolveReady = resolve; rejectReady = reject; });
  function measure(): void {
    if (disposed || !source) return;
    // A display:none host has no usable geometry. Keep the last accepted
    // layout instead of removing page gaps and measuring zero-sized lines.
    if (!iframe.clientWidth || !iframe.clientHeight) return;
    // IME owns temporary DOM until compositionend; splitting its text node would
    // invalidate the operating system's composition target.
    if (composing) return;
    let pages: DocumentRenderReport["pages"] = [], nodePages: DocumentRenderReport["nodePages"] = {}, results: DocumentRenderReport["fields"] = [], stable = !pagination;
    for (let pass = 0; pagination && pass < limits.maxLayoutPasses; pass++) {
      pages = pagination.measure(); nodePages = pagination.nodePages(pages); results = fields?.evaluate(pages, nodePages) ?? [];
      if (!fields?.changed) { stable = true; break; }
      const derivedFonts = resolvePreviewFonts(iframe.contentDocument!, ".sdk-field-display [data-word-script]");
      fontReport = Object.freeze([...baseFonts, ...derivedFonts]);
    }
    if (!stable) throw new DocumentError("ResourceLimitExceeded", "Layout fields did not converge within the render pass budget.");
    fields?.repeat(pages);
    source.setPages(pages);
    report = Object.freeze({ documentId: model.snapshot().manifest.documentId, revision: model.snapshot().manifest.revision,
      layout: options.layout ?? "pages", pages, nodePages, fields: results, diagnostics: Object.freeze([...(pagination?.diagnostics ?? []), ...(fields?.diagnostics ?? [])]), fonts: fontReport });
    const body = iframe.contentDocument!.body;
    const height = Math.ceil(Math.max(200, body.scrollHeight, body.getBoundingClientRect().height)) + "px";
    if (iframe.style.height !== height) iframe.style.height = height;
  }
  function render(): void {
    const frame = iframe.contentDocument!;
    pagination?.dispose(); pagination = undefined; fields = undefined; source = undefined;
    const snapshot = model.snapshot();
    const sourceFonts = resolvePreviewFonts(frame);
    source = new PreviewSource(frame, model);
    // Empty native runs still need a real line box in preview and editing.
    // A derived break provides it without adding a source newline/identity.
    for (const paragraph of source.root.querySelectorAll("p,h1,h2,h3,h4,h5,h6")) {
      if (!paragraph.textContent && !paragraph.querySelector("br,img,[data-word-tab]")) paragraph.append(markViewNode(frame.createElement("br")));
    }
    decorate(frame, snapshot.metadata.word);
    if ((options.layout ?? "pages") === "pages") {
      fields = layoutFields(frame, source, snapshot.metadata.word);
      pagination = paginate(frame, source, snapshot.metadata, limits);
    }
    baseFonts = Object.freeze([...sourceFonts, ...resolvePreviewFonts(frame, ".sdk-number-marker")]); fontReport = baseFonts;
    report = Object.freeze({ documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision,
      layout: options.layout ?? "pages", pages: [], nodePages: {}, fields: [], diagnostics: [], fonts: fontReport });
    measure();
    projection = new SourceProjection();
    const raw = new DOMParser().parseFromString(snapshot.html, "text/html").querySelector<HTMLElement>('[data-node-id="' + CSS.escape(source.root.dataset.nodeId!) + '"]')!;
    projection.capture(source.root, raw);
    const substitutions = fontReport.filter((item) => item.substituted), missing = [...new Set(fontReport.flatMap((item) => [...item.missingCharacters]))];
    if (substitutions.length || missing.length || fontReport.some((item) => !item.glyphAuditComplete)) {
      const note = frame.createElement("p"); note.setAttribute("role", "note");
      note.style.cssText = "font:12px sans-serif;color:#694800;margin:16px auto;max-width:600px";
      note.textContent = (substitutions.length ? "字体替代：" + substitutions.map((item) => item.requestedFamily + " → " + item.selectedFamily).join("；") + "。" : "")
        + (missing.length ? "缺少字形：" + missing.map((character) => "U+" + character.codePointAt(0)!.toString(16).toUpperCase() + "（" + character + "）").join("、") + "。" : "")
        + (fontReport.some((item) => !item.glyphAuditComplete) ? "部分字形未核验。" : "") + "断行和版式可能与原字体不同。";
      frame.body.append(note);
    }
    for (const image of frame.images) image.addEventListener("load", measure, { once: true });
    observer?.disconnect(); observer = new ResizeObserver(() => {
      if (scheduled !== undefined || disposed) return;
      scheduled = hostDocument.defaultView!.requestAnimationFrame(() => { scheduled = undefined; measure(); });
    }); observer.observe(frame.body); observer.observe(iframe);
  }
  let composing = false;
  const load = (): void => {
    if (disposed || !iframe.contentDocument) return;
    void iframe.contentDocument.fonts.ready.then(() => { if (!disposed) {
      loaded = true;
      iframe.contentDocument!.addEventListener("compositionstart", () => { composing = true; }, true);
      iframe.contentDocument!.addEventListener("compositionend", () => { composing = false; }, true);
      iframe.contentDocument!.addEventListener("copy", (event) => { if (source && projection) copyPreviewSelection(event, source.root, projection); });
      render(); resolveReady();
    } }).catch(rejectReady);
  };
  iframe.addEventListener("load", load, { once: true });
  iframe.srcdoc = "<!doctype html>" + prepared.document.documentElement.outerHTML; container.append(iframe);
  return {
    element: iframe, ready, fonts: () => fontReport, source: () => source, layout: () => report,
    update(next) {
      if (disposed) throw new DocumentError("InvalidOperation", "The renderer has been disposed.");
      if (!loaded || !iframe.contentDocument) throw new DocumentError("InvalidOperation", "Wait for renderer.ready before updating the view.");
      const incoming = prepareDocument(next), previous = prepared;
      generation++;
      observer?.disconnect(); pagination?.dispose(); pagination = undefined;
      prepared = incoming; model = incoming.model;
      const frame = iframe.contentDocument;
      frame.head.replaceWith(frame.importNode(incoming.document.head, true)); frame.body.replaceWith(frame.importNode(incoming.document.body, true));
      for (const url of previous.urls) URL.revokeObjectURL(url);
      render();
    },
    async settled() {
      const expected = generation;
      const current = (): void => { if (disposed || generation !== expected) throw new DocumentError("InvalidOperation", "The pending render was disposed or replaced."); };
      await ready; current();
      const frame = iframe.contentDocument!;
      await frame.fonts.ready; current();
      await Promise.all([...frame.images].map(async (image) => { if (image.complete && image.naturalWidth) return; await image.decode(); }));
      current(); measure(); return report;
    },
    printDocument() {
      if (disposed || !loaded || !source) throw new DocumentError("InvalidOperation", "The renderer is not ready for printing.");
      if (report.layout !== "pages") throw new DocumentError("InvalidOperation", "PDF printing requires pages layout.");
      measure(); return createPrintDocument(iframe.contentDocument!, source.root, report.pages, limits);
    },
    dispose() {
      if (disposed) return; disposed = true; generation++; observer?.disconnect(); pagination?.dispose(); source = undefined; projection = undefined;
      if (scheduled !== undefined) hostDocument.defaultView!.cancelAnimationFrame(scheduled);
      iframe.removeEventListener("load", load); iframe.remove(); iframe.srcdoc = "";
      for (const url of prepared.urls) URL.revokeObjectURL(url); resolveReady();
    },
  };
}
