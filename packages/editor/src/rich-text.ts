import { createWordFontControls } from "./word-font-controls.js";
import { createWordStyleControls } from "./word-style-controls.js";
import { trackedWordInput, trackedSelection } from "./rich-text/tracked-input.js";
import { createWordReviewControls } from "./word-review-controls.js";
import type { WordRevisionDisplay } from "./word-review-controls.js";
import { createWordChartControls } from "./word-chart-controls.js";
import { createWordControlControls } from "./word-control-controls.js";
import { createWordPageControls } from "./word-page-controls.js";
import { createWordListControls } from "./word-list-controls.js";
import { wordStoryPreview } from "./word-story-preview.js";
import { createWordTableControls } from "./word-table-controls.js";
import { createWordDrawingControls } from "./word-drawing-controls.js";
import { createWordReferenceControls } from "./word-reference-controls.js";
import { createWordLinkControls } from "./word-link-controls.js";
import { createWordMathControls } from "./word-math-controls.js";
import { createWordViewport, mountWordRulers } from "./word-ruler.js";
import { bindWordLineSelection } from "./line-selection.js";
import { bindWordSelectionScroll } from "./word-selection-scroll.js";
import { showWordWhitespace } from "./word-marks.js";
import { captureZoomAnchor, restoreZoomAnchor } from "./viewport-zoom.js";
import type { ZoomAnchor } from "./viewport-zoom.js";
import { executeClipboard } from "./clipboard-command.js";
import {
  copyClipboardImages,
  pasteClipboardImages,
  WORD_CLIPBOARD_BYTES,
} from "./rich-text/clipboard-images.js";
import { pasteClipboardFormat } from "./rich-text/clipboard-format.js";
import { copyWordClipboardFormat as copyClipboardFormat } from "@doffi/render";
import { readTextFormat } from "./text-format-state.js";
import type { TextFormatState } from "./text-format-state.js";
import {
  findWordText,
  wordTextReplacementOperations,
  getWordOutline,
  getWordTextStatistics,
  getWordTextDefaults,
  createOpaqueId,
  WordStyleSchema,
  DocumentError,
} from "@doffi/core";
import type {
  ContentDocumentModel,
  DocumentEditingSession,
  DocumentSelection,
  ReconciledTextEdit,
  TransactionPrepareOptions,
  DocumentOperation,
} from "@doffi/core";
import { mountDocumentPreview } from "./preview.js";
import { wordReviewPreview } from "./word-review-preview.js";
import { projectWordColumnReview } from "./word-column-review.js";
import { bindWordTableResize } from "./word-table-resize.js";
import { createWordWatermarkControls } from "./word-watermark-controls.js";
import {
  PreviewSource,
  sourceSelectionText,
  stripClipboardIdentities,
  DEFAULT_RENDER_LIMITS,
} from "@doffi/render";
import type { DocumentPreviewOptions } from "./preview.js";
import type { DocumentRenderReport } from "@doffi/render";
import {
  blockSelector,
  bookmark,
  caretAt,
  inputProofs,
  inputStep,
  offsetAt,
  restoreBookmark,
  sourceText,
} from "./rich-text/dom.js";
import type { InputStep, SelectionBookmark } from "./rich-text/dom.js";
import { SourceProjection } from "./rich-text/projection.js";
import {
  deleteNativeRange,
  nativeDeletionRange,
  styleNativeRange,
} from "./rich-text/operations.js";
import { workspaceControls } from "./workspace.js";

export interface DocumentTextEditorOptions extends DocumentPreviewOptions {
  readonly appearance?: "default" | "workspace";
  readonly storyId?: string;
  readonly onFormatStateChange?: (state: TextFormatState) => void;
  readonly onError?: (error: Error) => void;
}
export type DocumentTextCommand =
  | { readonly type: "toggle"; readonly property: "bold" | "italic" | "underline" }
  | {
      readonly type: "format";
      readonly property:
        | "font-weight"
        | "font-style"
        | "text-decoration"
        | "color"
        | "font-size"
        | "font-family"
        | "background-color"
        | "vertical-align"
        | "letter-spacing"
        | "text-decoration-line";
      readonly value: string;
    }
  | { readonly type: "align"; readonly value: "left" | "center" | "right" | "justify" }
  | {
      readonly type: "word";
      readonly action: "paragraph";
      readonly properties: Readonly<Record<string, string>>;
    }
  | {
      readonly type: "word";
      readonly action: "paragraphMark";
      readonly properties: Readonly<Record<string, string>>;
    }
  | { readonly type: "word"; readonly action: "style"; readonly styleId: string }
  | {
      readonly type: "word";
      readonly action: "list";
      readonly kind: "bullet" | "numbering" | "none";
      readonly level?: number;
    }
  | { readonly type: "word"; readonly action: "clearFormat" }
  | { readonly type: "word"; readonly action: "indent"; readonly direction: -1 | 1 }
  | { readonly type: "word"; readonly action: "showMarks"; readonly visible: boolean }
  | {
      readonly type: "word";
      readonly action: "paintFormat";
      readonly properties: Readonly<Record<string, string>>;
    }
  | {
      readonly type: "word";
      readonly action: "defineStyle";
      readonly styleId: string;
      readonly properties: Readonly<Record<string, string>>;
    }
  | {
      readonly type: "word";
      readonly action:
        | "find"
        | "replace"
        | "outline"
        | "wordCount"
        | "insertTable"
        | "tableSettings"
        | "listSettings"
        | "pageSettings"
        | "mainStory"
        | "sectionBreak"
        | "pageBreak"
        | "pageNumber"
        | "insertImage"
        | "insertShape"
        | "insertTextBox"
        | "insertConnector"
        | "drawingSettings"
        | "insertFootnote"
        | "insertEndnote"
        | "referenceSettings"
        | "insertLink"
        | "bookmarks"
        | "crossReference"
        | "caption"
        | "insertEquation"
        | "equationSettings"
        | "insertControl"
        | "controlSettings"
        | "insertChart"
        | "chartSettings"
        | "review"
        | "styleSettings"
        | "fontSettings"
        | "watermark";
    };
export interface DocumentTextEditor {
  readonly element: HTMLElement;
  readonly ready: Promise<void>;
  settled(): Promise<DocumentRenderReport>;
  layout(): DocumentRenderReport;
  focus(): void;
  execute(command: DocumentTextCommand): void;
  formatState(): TextFormatState;
  setZoom(zoom: number): void;
  setRulers(visible: boolean): void;
  editStory(storyId?: string): void;
  selection(): DocumentSelection | undefined;
  clipboard(action: "copy" | "cut" | "paste"): Promise<void>;
  /** Core owns persistent state and history. Disposal releases only this view. */
  dispose(): void;
}

/** Native input over the same scriptless HTML/CSS renderer used by preview. */
export function mountDocumentTextEditor(
  container: HTMLElement,
  session: DocumentEditingSession,
  options: DocumentTextEditorOptions = {},
): DocumentTextEditor {
  const wordMetadata = () => {
    const metadata = session.snapshot().metadata;
    if (metadata.presentation !== undefined)
      throw new DocumentError("UnsupportedFeature", "Presentation input is not implemented yet.");
    return metadata;
  };
  wordMetadata();
  const document = container.ownerDocument,
    element = document.createElement("div"),
    shadow = element.attachShadow({ mode: "open" });
  const style = document.createElement("style");
  style.textContent =
    ":host{display:block;font:14px system-ui}button,select{font:inherit;padding:6px 10px}.toolbar{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}.stage{max-height:70vh;overflow:auto;border:1px solid #ccd3dc}.error{color:#a00000}";
  if (options.appearance === "workspace")
    style.textContent +=
      workspaceControls +
      ".toolbar{display:var(--doffi-text-controls,flex);margin:0;padding:7px 16px;background:#f8faff;border-bottom:1px solid #e0e6f0}.stage{border:0;min-height:420px;height:var(--doffi-canvas-height,calc(100dvh - 282px));max-height:var(--doffi-canvas-height,calc(100dvh - 282px));background:#eceef1}.toolbar button,.toolbar select{font-size:11px}";
  const toolbar = document.createElement("div");
  toolbar.className = "toolbar";
  toolbar.setAttribute("role", "toolbar");
  toolbar.setAttribute("aria-label", "正文格式");
  const stage = document.createElement("div");
  stage.className = "stage";
  const error = document.createElement("p");
  error.setAttribute("role", "alert");
  error.className = "error";
  error.hidden = true;
  const plane = document.createElement("div");
  plane.style.cssText = "position:relative;margin-inline:auto;min-width:0";
  stage.append(plane);
  const viewport = createWordViewport(stage);
  shadow.append(style, toolbar, viewport, error);
  container.append(element);
  const controls = new AbortController(),
    initial = session.snapshot();
  let viewControls = new AbortController();
  let disposed = false,
    committing = false,
    composing = false,
    compositionRevision: number | undefined;
  let root: HTMLElement | undefined,
    frame: Document | undefined,
    revision = session.snapshot().manifest.revision;
  let projection = new SourceProjection(),
    proofs: ReconciledTextEdit[] = [],
    pending: InputStep | undefined;
  let savedSelection: SelectionBookmark | undefined;
  let compositionInsertion: { from: number; to: number } | undefined;
  let wholeStoryReplacement = false;
  let pendingCut: { token: string; cutId: string; revision: number; html: string } | undefined;
  const typingStyles = new Map<string, string>();
  let typingAnchor: string | undefined;
  let activeStoryId = options.storyId;
  let revisionDisplay: WordRevisionDisplay = "all";
  let showMarks = false,
    clearMarks: () => void = () => {};
  let wordPanel: HTMLElement | undefined;
  const report = (reason: unknown): void => {
    if (disposed) return;
    const value = reason instanceof Error ? reason : new Error("正文输入失败。");
    error.textContent = value.message;
    error.hidden = false;
    options.onError?.(value);
  };
  const tableControls = createWordTableControls(
    shadow,
    viewport,
    session,
    () => {
      const range = selectedRange(),
        node = range?.startContainer;
      const element = node?.nodeType === 1 ? (node as Element) : node?.parentElement;
      const paragraphId = element?.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6")?.dataset.nodeId,
        tableId = element?.closest<HTMLElement>("table")?.dataset.nodeId,
        cellId = element?.closest<HTMLElement>("td,th")?.dataset.nodeId;
      return {
        ...(paragraphId ? { paragraphId } : {}),
        ...(tableId ? { tableId } : {}),
        ...(cellId ? { cellId } : {}),
      };
    },
    report,
  );
  const currentSection = (): string => {
    const range = selectedRange(),
      metadata = wordMetadata();
    return (
      metadata.sections.find((section) => {
        const element =
          root?.querySelector('[data-node-id="' + CSS.escape(section.rootNodeId) + '"]') ??
          (root?.dataset.nodeId === section.rootNodeId ? root : undefined);
        return element && range && element.contains(range.startContainer);
      })?.rootNodeId ?? metadata.sections[0]!.rootNodeId
    );
  };
  const pageControls = createWordPageControls(
    shadow,
    viewport,
    session,
    currentSection,
    editStory,
    report,
  );
  const watermarkControls = createWordWatermarkControls(
    shadow,
    viewport,
    session,
    currentSection,
    report,
  );
  const fontControls = createWordFontControls(
    shadow,
    viewport,
    (mark) => {
      if (mark && root && frame) {
        const paragraph = selectedParagraphs(false)[0],
          marker = paragraph?.querySelector('[data-word-paragraph-mark="true"]');
        if (marker) {
          const range = frame.createRange();
          range.selectNodeContents(marker);
          range.collapse(true);
          return readTextFormat(root, range, new Map(), projection.ignored);
        }
      }
      return formatState();
    },
    (properties, mark) => {
      if (mark)
        wordCommand({
          type: "word",
          action: "paragraphMark",
          properties: Object.fromEntries(properties),
        });
      else formatAppearance(properties);
    },
    report,
  );
  const styleControls = createWordStyleControls(
    shadow,
    viewport,
    session,
    (style) => {
      const snapshot = session.snapshot(),
        range = selectedRange();
      if (style.kind === "paragraph") {
        wordCommand({ type: "word", action: "style", styleId: style.styleId });
        return;
      }
      let operations: import("@doffi/core").DocumentOperation[] = [];
      if (style.kind === "character" && range) {
        const anchor = new PreviewSource(frame!, viewModel()).textAnchor(range, session);
        if (anchor?.kind === "text" || anchor?.kind === "text-set")
          operations = (anchor.kind === "text" ? [anchor] : anchor.ranges).map((item) => ({
            op: "applyWordCharacterStyle",
            nodeId: item.nodeId,
            from: item.from,
            to: item.to,
            expectedText: item.quote,
            styleId: style.styleId,
          }));
      }
      if (style.kind === "table" && range) {
        const node =
            range.startContainer.nodeType === 1
              ? (range.startContainer as Element)
              : range.startContainer.parentElement,
          table = node?.closest<HTMLElement>("table");
        if (table?.dataset.nodeId)
          operations = [
            { op: "formatWordTable", nodeId: table.dataset.nodeId, styleId: style.styleId },
          ];
      }
      if (!operations.length)
        throw new DocumentError("InvalidOperation", "请先选择相应文字或表格。");
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations,
        }),
      );
    },
    report,
  );
  const listControls = createWordListControls(
    shadow,
    viewport,
    session,
    () => selectedParagraphs().map((node) => node.dataset.nodeId!),
    report,
  );
  const drawingControls = createWordDrawingControls(
    shadow,
    viewport,
    session,
    () => {
      const range = selectedRange(),
        paragraph = selectedParagraphs()[0];
      if (!range || !paragraph) return;
      return {
        paragraphId: paragraph.dataset.nodeId!,
        offset: offsetAt(paragraph, range.startContainer, range.startOffset, projection.ignored),
      };
    },
    editStory,
    report,
  );
  const referenceControls = createWordReferenceControls(
    shadow,
    viewport,
    session,
    () => {
      const range = selectedRange(),
        paragraph = selectedParagraphs()[0];
      if (!range || !paragraph) return;
      return {
        paragraphId: paragraph.dataset.nodeId!,
        offset: offsetAt(paragraph, range.startContainer, range.startOffset, projection.ignored),
      };
    },
    editStory,
    report,
  );
  const chartControls = createWordChartControls(
    shadow,
    viewport,
    session,
    () => {
      const range = selectedRange(),
        paragraph = selectedParagraphs()[0];
      if (!range || !paragraph) return;
      return {
        paragraphId: paragraph.dataset.nodeId!,
        offset: offsetAt(paragraph, range.startContainer, range.startOffset, projection.ignored),
      };
    },
    report,
  );
  const controlControls = createWordControlControls(
    shadow,
    viewport,
    session,
    () => {
      const range = selectedRange(),
        paragraph = selectedParagraphs()[0];
      if (!range || !paragraph) return;
      return {
        paragraphId: paragraph.dataset.nodeId!,
        offset: offsetAt(paragraph, range.startContainer, range.startOffset, projection.ignored),
      };
    },
    report,
  );
  const reviewControls = createWordReviewControls(
    shadow,
    viewport,
    session,
    report,
    () => revisionDisplay,
    (mode) => {
      revisionDisplay = mode;
      refresh(savedSelection);
    },
    (nodeId) => {
      const node = root?.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(nodeId) + '"]');
      if (node) node.scrollIntoView({ block: "center" });
    },
  );
  const mathControls = createWordMathControls(
    shadow,
    viewport,
    session,
    () => {
      const range = selectedRange(),
        paragraph = selectedParagraphs()[0];
      if (!range || !paragraph) return;
      return {
        paragraphId: paragraph.dataset.nodeId!,
        offset: offsetAt(paragraph, range.startContainer, range.startOffset, projection.ignored),
      };
    },
    report,
  );
  const linkControls = createWordLinkControls(
    shadow,
    viewport,
    session,
    () => {
      const range = selectedRange(),
        paragraph = selectedParagraphs()[0];
      const last = selectedParagraphs().at(-1);
      if (!range || !paragraph || !last || !root) return;
      const from = paragraph.contains(range.startContainer)
          ? offsetAt(paragraph, range.startContainer, range.startOffset, projection.ignored)
          : 0,
        to = last.contains(range.endContainer)
          ? offsetAt(last, range.endContainer, range.endOffset, projection.ignored)
          : session.getNode(last.dataset.nodeId!)!.text.length,
        begin = offsetAt(root, range.startContainer, range.startOffset, projection.ignored),
        end = offsetAt(root, range.endContainer, range.endOffset, projection.ignored);
      return {
        paragraphId: paragraph.dataset.nodeId!,
        endNodeId: last.dataset.nodeId!,
        from,
        to,
        text: session.getNode(root.dataset.nodeId!)!.text.slice(begin, end),
      };
    },
    (nodeId) => {
      const node = root?.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(nodeId) + '"]');
      if (node) {
        node.scrollIntoView({ block: "center" });
        const range = frame!.createRange();
        range.selectNode(node);
        range.collapse(true);
        root!.focus();
        frame!.getSelection()!.removeAllRanges();
        frame!.getSelection()!.addRange(range);
      }
    },
    report,
  );
  let cachedView:
    | {
        snapshot: ReturnType<typeof session.snapshot>;
        storyId: string | undefined;
        display: WordRevisionDisplay;
        model: ContentDocumentModel;
      }
    | undefined;
  const viewModel = () => {
    const snapshot = session.snapshot();
    if (
      activeStoryId &&
      !snapshot.metadata.word?.stories.some((story) => story.storyId === activeStoryId)
    )
      activeStoryId = undefined;
    if (
      cachedView?.snapshot === snapshot &&
      cachedView.storyId === activeStoryId &&
      cachedView.display === revisionDisplay
    )
      return cachedView.model;
    const source = wordStoryPreview(session, activeStoryId),
      model = revisionDisplay === "original" ? wordReviewPreview(source) : source;
    cachedView = { snapshot, storyId: activeStoryId, display: revisionDisplay, model };
    return model;
  };
  function editStory(storyId?: string): void {
    if (composing) throw new DocumentError("InvalidOperation", "请先完成当前输入法组合。");
    activeStoryId = storyId;
    savedSelection = undefined;
    refresh();
  }
  const sourceDocument = (): Document =>
    new DOMParser().parseFromString(session.snapshot().html, "text/html");
  const rootElement = (source: Document): HTMLElement => {
    const main = session
      .snapshot()
      .metadata.word?.stories.find((story) =>
        activeStoryId ? story.storyId === activeStoryId : story.kind === "main",
      );
    const result = main
      ? source.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(main.rootNodeId) + '"]')
      : source.querySelector<HTMLElement>('body > [data-office-kind="document"]');
    if (!result) throw new DocumentError("InvalidDocument", "正文内容域不存在。");
    return result;
  };
  const selectedRange = (): Range | undefined => {
    const selection = frame?.getSelection();
    if (!selection?.rangeCount || !root) return;
    const range = selection.getRangeAt(0);
    return root.contains(range.startContainer) && root.contains(range.endContainer)
      ? range.cloneRange()
      : undefined;
  };
  function formatState(): TextFormatState {
    return !root || composing || disposed
      ? { enabled: false }
      : {
          ...readTextFormat(root, selectedRange(), typingStyles, projection.ignored),
          editingText: true,
          storyKind:
            session
              .snapshot()
              .metadata.word?.stories.find((story) =>
                activeStoryId ? story.storyId === activeStoryId : story.kind === "main",
              )?.kind ?? "main",
        };
  }
  let lastFormatState = "";
  function notifyFormat(): void {
    if (disposed) return;
    const state = formatState(),
      key = JSON.stringify(state);
    if (key === lastFormatState) return;
    lastFormatState = key;
    try {
      options.onFormatStateChange?.(state);
    } catch {
      /* Host UI only. */
    }
  }
  function toggle(property: "bold" | "italic" | "underline"): void {
    const state = formatState(),
      active = state[property] === true;
    if (property === "underline") {
      format(
        "text-decoration",
        [!active ? "underline" : "", state.strikethrough ? "line-through" : ""]
          .filter(Boolean)
          .join(" ") || "none",
      );
      notifyFormat();
      return;
    }
    format(
      property === "bold"
        ? "font-weight"
        : property === "italic"
          ? "font-style"
          : "text-decoration",
      property === "bold"
        ? active
          ? "400"
          : "700"
        : property === "italic"
          ? active
            ? "normal"
            : "italic"
          : active
            ? "none"
            : "underline",
    );
    notifyFormat();
  }
  const select = (node: Node, offset: number): void => {
    const range = frame!.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    const selection = frame!.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  };
  function refresh(selection = savedSelection): void {
    if (!root || !frame || disposed) return;
    const renderedModel = viewModel();
    const source = rootElement(sourceDocument());
    viewControls.abort();
    viewControls = new AbortController();
    try {
      preview.update(renderedModel);
      root = rootElement(frame);
      projection = new SourceProjection();
      activate(source);
      revision = session.snapshot().manifest.revision;
      proofs = [];
      pending = undefined;
      typingStyles.clear();
      typingAnchor = undefined;
      compositionInsertion = undefined;
      wholeStoryReplacement = false;
    } catch (reason) {
      root.contentEditable = "false";
      report(reason);
      return;
    }
    root.focus();
    restoreBookmark(root, selection, projection.ignored);
    savedSelection = bookmark(root, projection.ignored);
    notifyFormat();
  }
  function normalizedRoot(): HTMLElement {
    const value = projection.serialize(root!),
      source = sourceDocument(),
      original = rootElement(source);
    const sections = wordMetadata()
      .sections.map((section) =>
        source.querySelector<HTMLElement>(
          '[data-node-id="' + CSS.escape(section.rootNodeId) + '"]',
        ),
      )
      .filter((section): section is HTMLElement =>
        Boolean(section && section !== original && original.contains(section)),
      )
      .sort((a, b) => (a.compareDocumentPosition(b) & 4 ? -1 : 1));
    const survivingSections = sections.some((section) =>
      value.querySelector('[data-node-id="' + CSS.escape(section.dataset.nodeId!) + '"]'),
    );
    const survivingOutside = [...original.querySelectorAll<HTMLElement>("[data-node-id]")].some(
      (node) =>
        !sections.some((section) => section === node || section.contains(node)) &&
        value.querySelector('[data-node-id="' + CSS.escape(node.dataset.nodeId!) + '"]'),
    );
    for (const [index, section] of sections.entries()) {
      if (value.querySelector('[data-node-id="' + CSS.escape(section.dataset.nodeId!) + '"]'))
        continue;
      let reference: Element | null = null;
      const following =
        section.parentElement === original
          ? [...original.children].slice([...original.children].indexOf(section) + 1)
          : sections.slice(index + 1);
      for (const node of following) {
        const id = node.getAttribute("data-node-id");
        if (id) reference = value.querySelector('[data-node-id="' + CSS.escape(id) + '"]');
        if (reference) break;
      }
      while (reference?.parentElement && reference.parentElement !== value)
        reference = reference.parentElement;
      value.insertBefore(section.cloneNode(false), reference);
    }
    let target: Element = value;
    if (wholeStoryReplacement && sections.length && !survivingSections && !survivingOutside) {
      target = value.querySelector(
        '[data-node-id="' + CSS.escape(sections[0]!.dataset.nodeId!) + '"]',
      )!;
      for (const child of [...value.childNodes]) {
        if (
          child === target ||
          (child.nodeType === 1 &&
            sections.some(
              (section) =>
                (child as Element).querySelector(
                  '[data-node-id="' + CSS.escape(section.dataset.nodeId!) + '"]',
                ) || (child as Element).getAttribute("data-node-id") === section.dataset.nodeId,
            ))
        )
          continue;
        if (child.nodeType !== 3 || child.textContent!.trim()) target.append(child);
      }
    }
    if (
      [...target.childNodes].some((child) => child.nodeType === 3 && child.textContent!.trim()) &&
      !target.querySelector(blockSelector)
    ) {
      const paragraph = document.createElement("p");
      paragraph.append(...target.childNodes);
      target.append(paragraph);
    }
    const paragraph = original.querySelector("p"),
      run = paragraph?.querySelector('[data-word-run="true"]');
    for (const node of value.querySelectorAll("p:not([data-node-id])")) {
      if (paragraph && !node.hasAttribute("class") && !node.hasAttribute("style"))
        for (const name of ["class", "style", "data-word-style", "data-word-line-rule"]) {
          const attribute = paragraph.getAttribute(name);
          if (attribute) node.setAttribute(name, attribute);
        }
      if (run && !node.querySelector("[data-word-run]")) {
        const wrapper = document.createElement("span");
        for (const name of ["data-word-run", "class", "style"]) {
          const attribute = run.getAttribute(name);
          if (attribute) wrapper.setAttribute(name, attribute);
        }
        wrapper.append(...node.childNodes);
        node.append(wrapper);
      }
    }
    return value;
  }
  function flush(
    resources?: TransactionPrepareOptions,
    prefix: readonly DocumentOperation[] = [],
  ): void {
    if (!root || disposed || composing) return;
    let selection = bookmark(root, projection.ignored);
    try {
      const baseRevision = compositionRevision ?? revision;
      if (baseRevision !== session.snapshot().manifest.revision)
        throw new DocumentError("RevisionConflict", "输入期间文档已变化，请基于当前修订重新输入。");
      const normalized = normalizedRoot(),
        values = new Map<string, string>();
      for (const edit of proofs) {
        const previous = values.get(edit.nodeId) ?? session.getNode(edit.nodeId)!.text;
        values.set(edit.nodeId, previous.slice(0, edit.from) + edit.text + previous.slice(edit.to));
      }
      const changedByNormalization = new Set(
        [...values]
          .filter(([id, value]) => {
            const node =
              id === normalized.dataset.nodeId
                ? normalized
                : normalized.querySelector('[data-node-id="' + CSS.escape(id) + '"]');
            return node && sourceText(node, () => false) !== value;
          })
          .map(([id]) => id),
      );
      const tracked = trackedWordInput(
        session,
        rootElement(sourceDocument()),
        normalized,
        typingStyles,
        proofs.filter((edit) => !changedByNormalization.has(edit.nodeId)),
      );
      const candidate = session.prepare(
        {
          protocolVersion: 1,
          documentId: session.snapshot().manifest.documentId,
          baseRevision,
          operations: [
            ...prefix,
            ...(tracked ?? [
              {
                op: "reconcileFragment",
                nodeId: root.dataset.nodeId!,
                html: normalized.innerHTML,
                textEdits: proofs.filter((edit) => !changedByNormalization.has(edit.nodeId)),
              },
            ]),
          ],
        },
        resources,
      );
      if (tracked) selection = trackedSelection(selection, session, candidate.model);
      committing = true;
      try {
        session.commit(candidate);
      } catch (reason) {
        session.discard(candidate);
        throw reason;
      } finally {
        committing = false;
      }
      error.hidden = true;
      compositionRevision = undefined;
      refresh(selection);
    } catch (reason) {
      compositionRevision = undefined;
      proofs = [];
      refresh(selection);
      report(reason);
    }
  }
  function mutate(
    range: Range,
    action: () => void,
    split = false,
    resources?: TransactionPrepareOptions,
    prefix: readonly DocumentOperation[] = [],
  ): void {
    if (!root || composing || disposed) return;
    const step = inputStep(root, range, projection.ignored);
    wholeStoryReplacement =
      step.from === 0 && step.to > 0 && step.to === sourceText(root, projection.ignored).length;
    const parent =
      range.startContainer.nodeType === 1
        ? (range.startContainer as Element)
        : range.startContainer.parentElement;
    const block = parent?.closest(blockSelector);
    const moved =
      split && block && !block.matches("td,th,li")
        ? {
            ...step,
            to: Math.max(
              step.to,
              offsetAt(root, block, block.childNodes.length, projection.ignored),
            ),
          }
        : undefined;
    try {
      action();
      const actual = inputProofs(root, step, projection.ignored),
        known = new Set(actual.map((edit) => edit.nodeId));
      proofs.push(
        ...actual,
        ...(moved
          ? inputProofs(root, moved, projection.ignored).filter((edit) => !known.has(edit.nodeId))
          : []),
      );
      flush(resources, prefix);
    } catch (reason) {
      refresh();
      report(reason);
    }
  }
  function insert(range: Range, content: Node): void {
    deleteRange(range);
    range.insertNode(content);
    select(
      content.parentNode!,
      [...content.parentNode!.childNodes].indexOf(content as ChildNode) + 1,
    );
  }
  function deleteRange(range: Range): void {
    deleteNativeRange(range);
  }
  function deletionRange(range: Range, kind: string): Range {
    return root ? nativeDeletionRange(root, range, kind, projection.ignored) : range;
  }
  function lineBreak(range: Range): void {
    insert(range, frame!.createElement("br"));
  }
  function paragraphBreak(range: Range): void {
    deleteRange(range);
    const parent =
      range.startContainer.nodeType === 1
        ? (range.startContainer as Element)
        : range.startContainer.parentElement;
    const block = parent?.closest<HTMLElement>(blockSelector);
    if (!block || block === root || block.matches("td,th,li")) {
      lineBreak(range);
      return;
    }
    if (block.dataset.wordNumbering && !sourceText(block, projection.ignored).trim()) {
      block.removeAttribute("data-word-numbering");
      block.removeAttribute("data-word-level");
      projection.setStyles(
        block,
        new Map([
          ["margin-left", "0pt"],
          ["text-indent", "0pt"],
        ]),
      );
      return;
    }
    const tail = frame!.createRange();
    tail.setStart(range.startContainer, range.startOffset);
    tail.setEnd(block, block.childNodes.length);
    const authored = projection.serialize(block),
      word = session.snapshot().metadata.word,
      style = word?.styles.find(
        (style) => style.styleId === authored.getAttribute("data-word-style"),
      ),
      following = word?.styles.find(
        (item) => item.styleId === style?.nextStyleId && item.kind === "paragraph",
      ),
      next = frame!.createElement(
        following
          ? following.outlineLevel !== undefined && following.outlineLevel < 6
            ? "h" + (following.outlineLevel + 1)
            : "p"
          : block.localName,
      );
    for (const attribute of authored.attributes)
      if (!["id", "data-node-id"].includes(attribute.name))
        next.setAttribute(attribute.name, attribute.value);
    const content = projection.serializeRange(tail);
    // A range beginning at the end of inserted text still clones its empty
    // revision ancestor. Keep its font shell, but never create an unowned
    // revision in the next paragraph. Real content and zero-width native
    // structures retain their source ownership and review boundary.
    const protectedContent =
      "img,br,[data-word-field],[data-word-field-end],[data-word-content-control],[data-word-native-object],[data-word-drawing],[data-word-equation],[data-word-note-reference],[data-word-bookmark],[data-word-move-range]";
    for (const revision of content.querySelectorAll<HTMLElement>("ins,del")) {
      const id = revision.dataset.nodeId;
      if (
        !id ||
        !session.getNode(id)?.text ||
        sourceText(revision, () => false) ||
        revision.matches(protectedContent) ||
        revision.querySelector(protectedContent)
      )
        continue;
      const shell = frame!.createElement("span");
      for (const attribute of revision.attributes)
        if (!["data-node-id", "id"].includes(attribute.name))
          shell.setAttribute(attribute.name, attribute.value);
      shell.dataset.wordRunOverlay = "true";
      shell.append(...revision.childNodes);
      revision.replaceWith(shell);
    }
    tail.deleteContents();
    next.append(content);
    if (following) {
      next.dataset.wordStyle = following.styleId;
      next.className = [...authored.classList]
        .filter((name) => name !== style?.cssClass)
        .concat(following.cssClass)
        .join(" ");
      next.removeAttribute("data-word-outline-level");
      for (const run of next.querySelectorAll<HTMLElement>("[data-word-run]")) {
        if (style?.runCssClass) run.classList.remove(style.runCssClass);
        if (following.runCssClass) run.classList.add(following.runCssClass);
      }
    }
    block.after(next);
    select(next, 0);
  }
  function format(property: string, value: string): void {
    const properties = new Map([[property, value]]);
    if (property === "font-family")
      for (const role of ["ascii", "hAnsi", "eastAsia", "cs"])
        properties.set("--word-font-" + role, value);
    formatAppearance(properties);
  }
  function formatAppearance(properties: ReadonlyMap<string, string>): void {
    if (!root || composing) return;
    root.focus();
    if (!selectedRange()) restoreBookmark(root, savedSelection, projection.ignored);
    const range = selectedRange();
    if (!range) return;
    if (range.collapsed) {
      for (const [name, value] of properties) typingStyles.set(name, value);
      typingAnchor = JSON.stringify(bookmark(root, projection.ignored));
      notifyFormat();
      return;
    }
    const selection = bookmark(root, projection.ignored),
      snapshot = session.snapshot();
    if (snapshot.metadata.word) {
      const anchor = new PreviewSource(frame!, viewModel()).textAnchor(range, session);
      if (!anchor || (anchor.kind !== "text" && anchor.kind !== "text-set"))
        throw new DocumentError("InvalidOperation", "请选择可格式化的正文。");
      const ranges = anchor.kind === "text" ? [anchor] : anchor.ranges,
        candidate = session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations: ranges.map((item) => ({
            op: "formatWordText",
            nodeId: item.nodeId,
            from: item.from,
            to: item.to,
            expectedText: item.quote,
            properties: Object.fromEntries(properties),
          })),
        });
      committing = true;
      try {
        session.commit(candidate);
      } catch (error) {
        session.discard(candidate);
        throw error;
      } finally {
        committing = false;
      }
      refresh(selection);
      return;
    }
    mutate(range, () => {
      styleRange(range, properties);
      if (root) restoreBookmark(root, selection, projection.ignored);
    });
  }
  function styleRange(range: Range, properties: ReadonlyMap<string, string>): void {
    if (root) styleNativeRange(root, range, properties, projection.ignored, projection.setStyles);
  }
  function button(label: string, action: () => void): void {
    const control = document.createElement("button");
    control.type = "button";
    control.textContent = label;
    control.addEventListener("mousedown", (event) => event.preventDefault(), {
      signal: controls.signal,
    });
    control.addEventListener(
      "click",
      () => {
        try {
          action();
        } catch (reason) {
          report(reason);
        }
      },
      { signal: controls.signal },
    );
    toolbar.append(control);
  }
  button("加粗", () => toggle("bold"));
  button("常规字重", () => format("font-weight", "400"));
  button("斜体", () => toggle("italic"));
  button("下划线", () => toggle("underline"));
  button("深蓝色", () => format("color", "#1f4b99"));
  const size = document.createElement("select");
  size.setAttribute("aria-label", "字号");
  for (const value of [9, 10, 11, 12, 14, 16, 18, 24]) {
    const option = document.createElement("option");
    option.value = value + "pt";
    option.textContent = value + " 磅";
    size.append(option);
  }
  size.value = "12pt";
  size.addEventListener("change", () => format("font-size", size.value), {
    signal: controls.signal,
  });
  toolbar.append(size);
  function align(alignment: "left" | "center" | "right" | "justify"): void {
    if (!root || composing) return;
    root.focus();
    if (!selectedRange()) restoreBookmark(root, savedSelection, projection.ignored);
    const range = selectedRange();
    if (!range) return;
    const nodes = [...root.querySelectorAll<HTMLElement>(blockSelector)].filter(
      (node) => range.intersectsNode(node) && node.dataset.nodeId,
    );
    if (!nodes.length) return;
    savedSelection = bookmark(root, projection.ignored);
    const candidate = session.prepare({
      protocolVersion: 1,
      documentId: session.snapshot().manifest.documentId,
      baseRevision: revision,
      operations: nodes.map((node) => ({
        op: "setWordParagraph" as const,
        nodeId: node.dataset.nodeId!,
        properties: { "text-align": alignment, "--word-distributed": "0" },
      })),
    });
    try {
      session.commit(candidate);
    } catch (reason) {
      session.discard(candidate);
      throw reason;
    }
  }
  function openWordPanel(action: "find" | "replace" | "outline" | "wordCount"): void {
    wordPanel?.remove();
    const panel = document.createElement("section");
    wordPanel = panel;
    panel.setAttribute("role", "region");
    panel.setAttribute(
      "aria-label",
      action === "outline" ? "文档大纲" : action === "wordCount" ? "字数统计" : "查找和替换",
    );
    panel.style.cssText =
      "padding:12px;background:white;border:1px solid #dbe3ef;max-height:240px;overflow:auto;display:flex;flex-wrap:wrap;gap:8px";
    const addButton = (label: string, run: () => void): HTMLButtonElement => {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = label;
      button.addEventListener(
        "click",
        () => {
          try {
            run();
          } catch (error) {
            report(error);
          }
        },
        { signal: controls.signal },
      );
      panel.append(button);
      return button;
    };
    addButton("关闭", () => {
      panel.remove();
      wordPanel = undefined;
    });
    const focus = (nodeId: string, from?: number, to?: number, endNodeId = nodeId): void => {
      const node = root?.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(nodeId) + '"]');
      if (!node || !root || !frame) return;
      root.focus();
      node.scrollIntoView({ block: "center" });
      const range = frame.createRange();
      if (from === undefined) range.selectNodeContents(node);
      else {
        range.setStart(...caretAt(node, from, projection.ignored, "forward"));
        range.setEnd(
          ...caretAt(
            root.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(endNodeId) + '"]') ??
              node,
            to ?? from,
            projection.ignored,
          ),
        );
      }
      const selected = frame.getSelection()!;
      selected.removeAllRanges();
      selected.addRange(range);
      savedSelection = bookmark(root, projection.ignored);
      notifyFormat();
    };
    if (action === "outline")
      for (const heading of getWordOutline(session, { revisionView: revisionDisplay })) {
        const button = addButton(heading.text || "空标题", () => focus(heading.nodeId));
        button.style.marginLeft = (heading.level - 1) * 12 + "px";
      }
    else if (action === "wordCount") {
      const stats = getWordTextStatistics(session, { revisionView: revisionDisplay }),
        text = document.createElement("p");
      text.textContent = `字数 ${stats.words} · 字符数 ${stats.characters} · 不计空格 ${stats.charactersWithoutSpaces} · 段落 ${stats.paragraphs}`;
      panel.append(text);
    } else {
      const query = document.createElement("input"),
        replacement = document.createElement("input"),
        matchCase = document.createElement("input"),
        wholeWord = document.createElement("input"),
        results = document.createElement("div"),
        summary = document.createElement("span");
      query.setAttribute("aria-label", "查找内容");
      query.placeholder = "查找内容";
      replacement.setAttribute("aria-label", "替换为");
      replacement.placeholder = "替换为";
      matchCase.type = wholeWord.type = "checkbox";
      panel.append(query);
      if (action === "replace") panel.append(replacement);
      for (const [input, label] of [
        [matchCase, "区分大小写"],
        [wholeWord, "全字匹配"],
      ] as const) {
        const wrapper = document.createElement("label");
        wrapper.append(input, label);
        panel.append(wrapper);
      }
      let matches: ReturnType<typeof findWordText> = [],
        current = 0;
      const search = (): void => {
        matches = findWordText(session, query.value, {
          matchCase: matchCase.checked,
          wholeWord: wholeWord.checked,
          revisionView: revisionDisplay,
        });
        current = Math.min(current, Math.max(0, matches.length - 1));
        summary.textContent = matches.length + " 处匹配";
        results.replaceChildren();
        for (const [index, match] of matches.entries()) {
          const button = document.createElement("button");
          button.type = "button";
          button.textContent = index + 1 + " · " + match.expectedText;
          button.addEventListener(
            "click",
            () => {
              current = index;
              focus(
                match.nodeId,
                match.from,
                match.ranges?.at(-1)?.to ?? match.to,
                match.ranges?.at(-1)?.nodeId,
              );
            },
            { signal: controls.signal },
          );
          results.append(button);
        }
      };
      const next = (direction: number): void => {
        search();
        if (!matches.length) return;
        current = (current + direction + matches.length) % matches.length;
        const match = matches[current]!;
        focus(
          match.nodeId,
          match.from,
          match.ranges?.at(-1)?.to ?? match.to,
          match.ranges?.at(-1)?.nodeId,
        );
      };
      query.addEventListener("input", search, { signal: controls.signal });
      matchCase.addEventListener("change", search, { signal: controls.signal });
      wholeWord.addEventListener("change", search, { signal: controls.signal });
      query.addEventListener(
        "keydown",
        (event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            next(event.shiftKey ? -1 : 1);
          }
        },
        { signal: controls.signal },
      );
      addButton("上一处", () => next(-1));
      addButton("下一处", () => next(1));
      const replace = (all: boolean): void => {
        search();
        const selected = all ? matches : matches.slice(current, current + 1);
        if (!selected.length) return;
        const snapshot = session.snapshot();
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: snapshot.manifest.documentId,
            baseRevision: snapshot.manifest.revision,
            operations: [...wordTextReplacementOperations(session, selected, replacement.value)],
          }),
        );
        search();
      };
      if (action === "replace") {
        addButton("替换当前", () => replace(false));
        addButton("全部替换", () => replace(true));
      }
      panel.append(summary, results);
      queueMicrotask(() => query.focus());
    }
    shadow.insertBefore(panel, viewport);
  }
  function selectedParagraphs(focus = true): HTMLElement[] {
    if (!root) return [];
    if (focus) {
      root.focus();
      if (!selectedRange()) restoreBookmark(root, savedSelection, projection.ignored);
    }
    const range = selectedRange();
    if (!range) return [];
    if (range.collapsed) {
      const element =
        range.startContainer.nodeType === 1
          ? (range.startContainer as Element)
          : range.startContainer.parentElement;
      const paragraph = element?.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li");
      return paragraph?.dataset.nodeId ? [paragraph] : [];
    }
    return [...root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li")].filter(
      (node) => node.dataset.nodeId && range.intersectsNode(node),
    );
  }
  function wordCommand(command: Extract<DocumentTextCommand, { type: "word" }>): void {
    if (!root || composing) return;
    if (command.action === "paragraphMark") {
      const paragraphs = selectedParagraphs(),
        snapshot = session.snapshot();
      if (!paragraphs.length) throw new DocumentError("InvalidOperation", "请先选择正文段落。");
      savedSelection = bookmark(root, projection.ignored);
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations: paragraphs.map((paragraph) => ({
            op: "formatWordParagraphMark",
            nodeId: paragraph.dataset.nodeId!,
            properties: { ...command.properties },
          })),
        }),
      );
      return;
    }
    if (command.action === "review") {
      reviewControls.open();
      return;
    }
    if (command.action === "listSettings") {
      listControls.open();
      return;
    }
    if (
      command.action === "insertImage" ||
      command.action === "insertShape" ||
      command.action === "insertTextBox" ||
      command.action === "insertConnector" ||
      command.action === "drawingSettings"
    ) {
      drawingControls.open(command.action);
      return;
    }
    if (
      command.action === "insertFootnote" ||
      command.action === "insertEndnote" ||
      command.action === "referenceSettings"
    ) {
      referenceControls.open(
        command.action === "insertFootnote"
          ? "footnote"
          : command.action === "insertEndnote"
            ? "endnote"
            : "settings",
      );
      return;
    }
    if (
      command.action === "insertLink" ||
      command.action === "bookmarks" ||
      command.action === "crossReference" ||
      command.action === "caption"
    ) {
      linkControls.open(
        command.action === "insertLink"
          ? "link"
          : command.action === "bookmarks"
            ? "bookmark"
            : command.action === "crossReference"
              ? "reference"
              : "caption",
      );
      return;
    }
    if (command.action === "insertChart" || command.action === "chartSettings") {
      chartControls.open(command.action === "insertChart");
      return;
    }
    if (command.action === "insertControl" || command.action === "controlSettings") {
      controlControls.open(command.action === "insertControl");
      return;
    }
    if (command.action === "insertEquation" || command.action === "equationSettings") {
      mathControls.open(command.action === "insertEquation");
      return;
    }
    if (command.action === "fontSettings") {
      fontControls.open();
      return;
    }
    if (command.action === "styleSettings") {
      styleControls.open();
      return;
    }
    if (command.action === "pageSettings") {
      pageControls.open();
      return;
    }
    if (command.action === "watermark") {
      watermarkControls.open();
      return;
    }
    if (command.action === "mainStory") {
      editStory();
      return;
    }
    if (command.action === "pageBreak") {
      const range = selectedRange();
      if (!range) return;
      mutate(range, () => {
        const breaker = frame!.createElement("br");
        breaker.dataset.wordBreak = "page";
        insert(range, breaker);
      });
      return;
    }
    if (command.action === "sectionBreak" || command.action === "pageNumber") {
      const paragraph = selectedParagraphs()[0];
      if (!paragraph) return;
      const snapshot = session.snapshot();
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations: [
            command.action === "sectionBreak"
              ? { op: "splitWordSection", nodeId: paragraph.dataset.nodeId!, breakType: "nextPage" }
              : {
                  op: "insertWordField",
                  nodeId: paragraph.dataset.nodeId!,
                  instruction: " PAGE ",
                  cachedText: "1",
                },
          ],
        }),
      );
      return;
    }
    if (command.action === "insertTable" || command.action === "tableSettings") {
      tableControls.open(command.action === "insertTable");
      return;
    }
    if (["find", "replace", "outline", "wordCount"].includes(command.action)) {
      openWordPanel(command.action as "find" | "replace" | "outline" | "wordCount");
      return;
    }
    if (command.action === "showMarks") {
      showMarks = command.visible;
      updateMarks();
      return;
    }
    if (command.action === "clearFormat") {
      const paragraph = selectedParagraphs()[0],
        range = selectedRange();
      if (!paragraph || !range) return;
      savedSelection = bookmark(root, projection.ignored);
      if (range.collapsed) {
        typingStyles.clear();
        for (const [property, value] of Object.entries(
          getWordTextDefaults(session, paragraph.dataset.nodeId!),
        ))
          typingStyles.set(property, value);
        typingAnchor = JSON.stringify(savedSelection);
        notifyFormat();
        return;
      }
      const anchor = new PreviewSource(frame!, viewModel()).textAnchor(range, session);
      if (!anchor || !(anchor.kind === "text" || anchor.kind === "text-set")) return;
      const snapshot = session.snapshot(),
        ranges = anchor.kind === "text" ? [anchor] : anchor.ranges;
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations: ranges.map((range) => ({
            op: "resetWordText",
            nodeId: range.nodeId,
            from: range.from,
            to: range.to,
            expectedText: range.quote,
            resetParagraph:
              range.from === 0 && range.to === session.getNode(range.nodeId)!.text.length,
          })),
        }),
      );
      return;
    }
    const paragraphs = selectedParagraphs();
    if (!paragraphs.length) return;
    savedSelection = bookmark(root, projection.ignored);
    const snapshot = session.snapshot(),
      operations: import("@doffi/core").DocumentOperation[] = [];
    const documentRoot = new DOMParser()
      .parseFromString(snapshot.html, "text/html")
      .querySelector<HTMLElement>('[data-office-kind="document"]')!.dataset.nodeId!;
    if (
      command.action === "find" ||
      command.action === "replace" ||
      command.action === "outline" ||
      command.action === "wordCount"
    )
      return;
    if (command.action === "defineStyle") {
      const current = snapshot.metadata.word?.styles.find(
        (style) => style.styleId === command.styleId,
      );
      const definition = current
        ? WordStyleSchema.parse(current)
        : {
            styleId: command.styleId,
            kind: "paragraph" as const,
            cssClass: "word-style-" + createOpaqueId().replaceAll("-", ""),
          };
      operations.push({
        op: "setWordStyle",
        nodeId: documentRoot,
        style: definition,
        runProperties: { ...command.properties },
      });
    } else if (command.action === "style") {
      const known = snapshot.metadata.word?.styles.find(
        (style) => style.styleId === command.styleId,
      );
      if (!known) {
        const heading = /^Heading([1-9])$/.exec(command.styleId),
          cssClass = "word-user-" + command.styleId.toLowerCase();
        operations.push({
          op: "setWordStyle",
          nodeId: documentRoot,
          style: {
            styleId: command.styleId,
            kind: "paragraph",
            cssClass,
            ...(heading ? { outlineLevel: Number(heading[1]) - 1 } : {}),
          },
          paragraphProperties: {
            "margin-top": heading ? "12pt" : "0pt",
            "margin-bottom": heading ? "6pt" : "0pt",
            ...(heading ? { "break-after": "avoid" } : {}),
          },
          runProperties: {
            "font-size": heading ? 20 - Number(heading[1]) * 2 + "pt" : "12pt",
            "font-weight": heading ? "700" : "400",
          },
        });
      }
      for (const paragraph of paragraphs)
        operations.push({
          op: "setWordParagraph",
          nodeId: paragraph.dataset.nodeId!,
          styleId: command.styleId,
          clearDirectFormatting: true,
        });
    } else if (command.action === "paragraph") {
      for (const paragraph of paragraphs)
        operations.push({
          op: "setWordParagraph",
          nodeId: paragraph.dataset.nodeId!,
          properties: { ...command.properties },
        });
    } else if (command.action === "list") {
      const numberingId = command.kind === "none" ? null : "word-user-" + command.kind;
      if (
        numberingId &&
        !snapshot.metadata.word?.numbering.some((n) => n.numberingId === numberingId)
      ) {
        operations.push({
          op: "setWordNumbering",
          nodeId: documentRoot,
          numberingId,
          levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            start: 1,
            format: command.kind === "bullet" ? ("bullet" as const) : ("decimal" as const),
            text:
              command.kind === "bullet"
                ? ["•", "◦", "▪"][level % 3]!
                : Array.from({ length: level + 1 }, (_, n) => "%" + (n + 1)).join(".") + ".",
            suffix: "tab" as const,
            cssClass: "word-list-" + command.kind + "-" + level,
          })),
        });
      }
      for (const paragraph of paragraphs)
        operations.push({
          op: "setWordList",
          nodeId: paragraph.dataset.nodeId!,
          numberingId,
          level: command.level ?? 0,
        });
    } else if (command.action === "indent") {
      for (const paragraph of paragraphs)
        operations.push({
          op: "setWordParagraph",
          nodeId: paragraph.dataset.nodeId!,
          properties: {
            "margin-left":
              Math.max(
                0,
                parseFloat(frame!.defaultView!.getComputedStyle(paragraph).marginLeft) * 0.75 +
                  command.direction * 24,
              ) + "pt",
          },
        });
    } else if (command.action === "paintFormat") {
      formatAppearance(
        new Map(
          Object.entries({
            ...command.properties,
            ...(command.properties["font-family"]
              ? Object.fromEntries(
                  ["ascii", "hAnsi", "eastAsia", "cs"].map((role) => [
                    "--word-font-" + role,
                    command.properties["font-family"]!,
                  ]),
                )
              : {}),
          }),
        ),
      );
      return;
    } else throw new DocumentError("InvalidOperation", "未知 Word 格式命令。");
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations,
      }),
    );
  }
  for (const [label, alignment] of [
    ["左对齐", "left"],
    ["居中", "center"],
    ["右对齐", "right"],
  ] as const)
    button(label, () => align(alignment));
  button("撤销输入", () => {
    session.undo();
  });
  button("重做输入", () => {
    session.redo();
  });

  function clipboardRange(range: Range): {
    start: HTMLElement;
    end: HTMLElement;
    from: number;
    to: number;
    expectedText: string;
  } {
    if (!root) throw new DocumentError("InvalidOperation", "剪贴板选区缺少源故事。");
    const selector = "p,h1,h2,h3,h4,h5,h6";
    const owner = (node: Node) =>
      (node.nodeType === 1 ? (node as Element) : node.parentElement)?.closest<HTMLElement>(
        selector,
      );
    const paragraphs = [...root.querySelectorAll<HTMLElement>(selector)].filter((p) =>
      range.intersectsNode(p),
    );
    const start = owner(range.startContainer) ?? paragraphs[0],
      end = owner(range.endContainer) ?? paragraphs.at(-1);
    if (!start || !end)
      throw new DocumentError("InvalidOperation", "剪贴板操作需要准确的段落选区。");
    const from = start.contains(range.startContainer)
        ? offsetAt(start, range.startContainer, range.startOffset, projection.ignored)
        : 0,
      to = end.contains(range.endContainer)
        ? offsetAt(end, range.endContainer, range.endOffset, projection.ignored)
        : sourceText(end, projection.ignored).length;
    return {
      start,
      end,
      from,
      to,
      expectedText: session
        .getNode(root.dataset.nodeId!)!
        .text.slice(
          offsetAt(root, start, 0, projection.ignored) + from,
          offsetAt(root, end, 0, projection.ignored) + to,
        ),
    };
  }

  function paste(event: ClipboardEvent): void {
    event.preventDefault();
    if (!event.clipboardData || composing || !root) return;
    const raw = event.clipboardData.getData("text/html"),
      image = Array.from(event.clipboardData?.files ?? []).find((file) =>
        /^image\/(?:png|jpeg|webp)$/.test(file.type),
      );
    if (image && !raw) {
      event.preventDefault();
      void drawingControls.insertImage(image).catch(report);
      return;
    }
    try {
      const text = event.clipboardData.getData("text/plain");
      if (raw.length > WORD_CLIPBOARD_BYTES || text.length > WORD_CLIPBOARD_BYTES)
        throw new DocumentError("ResourceLimitExceeded", "粘贴内容超过 64 MiB。");
      const template = document.createElement("template");
      if (raw) template.innerHTML = raw;
      else
        text
          .replace(/\r\n?/g, "\n")
          .split("\n")
          .forEach((line, index) => {
            if (index) template.content.append(document.createElement("br"));
            template.content.append(document.createTextNode(line));
          });
      const format = pasteClipboardFormat(template.content, session),
        prefix = format.operations,
        resources = pasteClipboardImages(template.content, session);
      const cutId =
        pendingCut &&
        pendingCut.revision === session.snapshot().manifest.revision &&
        pendingCut.html === raw &&
        template.content
          .querySelector("[data-doffi-word-cut]")
          ?.getAttribute("data-doffi-word-cut") === pendingCut.token
          ? pendingCut.cutId
          : undefined;
      for (const node of template.content.querySelectorAll("[data-doffi-word-cut]"))
        node.removeAttribute("data-doffi-word-cut");
      for (const node of template.content.querySelectorAll("[data-node-id],[id]")) {
        if (!format.semantics) node.removeAttribute("data-node-id");
        node.removeAttribute("id");
      }
      const block = [...template.content.children].some((node) =>
        /^(?:p|h[1-6]|div|section|article|table|ul|ol)$/.test(node.localName),
      );
      if (!format.semantics) {
        const validation = session.prepare(
          {
            protocolVersion: 1,
            documentId: session.snapshot().manifest.documentId,
            baseRevision: revision,
            operations: [
              ...prefix,
              {
                op: "insertFragment",
                nodeId: root.dataset.nodeId!,
                position: "end",
                html: block ? template.innerHTML : "<p>" + template.innerHTML + "</p>",
              },
            ],
          },
          resources,
        );
        session.discard(validation);
      }
      const range = selectedRange();
      if (!range) return;
      const receiving = session.snapshot().metadata.word ? clipboardRange(range) : undefined;
      const start = receiving?.start,
        end = receiving?.end;
      if (
        format.semantics ||
        cutId ||
        (session.snapshot().metadata.word && start && end && start !== end) ||
        (block &&
          start &&
          end &&
          session.snapshot().metadata.word &&
          !template.content.querySelector("table,section,article"))
      ) {
        if (!start || !end) throw new DocumentError("InvalidOperation", "粘贴需要准确的段落选区。");
        const from = receiving!.from,
          to = receiving!.to,
          nodeId = start.dataset.nodeId!,
          snapshot = session.snapshot(),
          word = snapshot.metadata.word!,
          author = word.trackingAuthor ?? "",
          fold =
            start === end &&
            word.revisions.some((revision) => {
              if (revision.kind !== "insert" || revision.author !== author) return false;
              const node = start.querySelector<HTMLElement>(
                '[data-node-id="' + CSS.escape(revision.nodeId) + '"]',
              );
              if (!node) return false;
              const begin = offsetAt(start, node, 0, projection.ignored);
              return from >= begin && to <= begin + session.getNode(revision.nodeId)!.text.length;
            }),
          caret =
            (word.trackChanges && !fold ? to : from) +
            sourceText(template.content, () => false).length;
        const candidate = session.prepare(
          {
            protocolVersion: 1,
            documentId: snapshot.manifest.documentId,
            baseRevision: snapshot.manifest.revision,
            operations: [
              ...prefix,
              {
                op: block ? "pasteWordBlockContent" : "pasteWordInlineContent",
                nodeId,
                ...(start !== end ? { endNodeId: end.dataset.nodeId! } : {}),
                from,
                to,
                expectedText: receiving!.expectedText,
                html: template.innerHTML,
                semantics: format.semantics,
                ...(cutId ? { cutId } : {}),
              },
            ],
          },
          resources,
        );
        let landing = { nodeId, offset: caret, path: [] as number[] };
        if (block) {
          const beforeParent = session.getNode(session.getNode(nodeId)!.parentNodeId!)!,
            next = beforeParent.childNodeIds[beforeParent.childNodeIds.indexOf(nodeId) + 1],
            parent = candidate.model.getNode(beforeParent.nodeId)!,
            siblings = parent.childNodeIds.slice(parent.childNodeIds.indexOf(nodeId) + 1),
            until = next ? siblings.slice(0, siblings.indexOf(next)) : siblings,
            lastBlock = until.at(-2);
          if (lastBlock) {
            const walk = (id: string): string[] => {
                const node = candidate.model.getNode(id)!;
                return /^p$|^h[1-6]$/.test(node.tagName) ? [id] : node.childNodeIds.flatMap(walk);
              },
              last = walk(lastBlock).at(-1);
            if (last)
              landing = {
                nodeId: last,
                offset: candidate.model.getNode(last)!.text.length,
                path: [],
              };
          }
          const baseIds = new Set(
            [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(
              (match) => match[1]!,
            ),
          );
          const tails = new Set(candidate.mapping.textTransfers?.map((move) => move.targetNodeId));
          const walkInserted = (id: string): string[] => {
            const node = candidate.model.getNode(id)!;
            return /^p$|^h[1-6]$/.test(node.tagName)
              ? !baseIds.has(id) && !tails.has(id)
                ? [id]
                : []
              : node.childNodeIds.flatMap(walkInserted);
          };
          const last = walkInserted(root.dataset.nodeId!).at(-1);
          if (last)
            landing = {
              nodeId: last,
              offset: candidate.model.getNode(last)!.text.length,
              path: [],
            };
        } else if (start !== end) {
          const retained =
            candidate.model.getNode(nodeId) ?? candidate.model.getNode(end.dataset.nodeId!)!;
          landing = {
            nodeId: retained.nodeId,
            offset:
              retained.nodeId === nodeId
                ? retained.text.length
                : from + sourceText(template.content, () => false).length,
            path: [],
          };
        }
        committing = true;
        try {
          session.commit(candidate);
          pendingCut = undefined;
        } catch (error) {
          session.discard(candidate);
          throw error;
        } finally {
          committing = false;
        }
        refresh({ anchor: landing, focus: landing });
        return;
      }
      if (session.snapshot().metadata.word?.trackChanges && !block) {
        const start = (
            range.startContainer.nodeType === 1
              ? (range.startContainer as Element)
              : range.startContainer.parentElement
          )?.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6"),
          end = (
            range.endContainer.nodeType === 1
              ? (range.endContainer as Element)
              : range.endContainer.parentElement
          )?.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6");
        if (!start || start !== end)
          throw new DocumentError("UnsupportedFeature", "记录修订的行内粘贴需要限定在同一段落。");
        const from = offsetAt(start, range.startContainer, range.startOffset, projection.ignored),
          to = offsetAt(start, range.endContainer, range.endOffset, projection.ignored),
          nodeId = start.dataset.nodeId!,
          snapshot = session.snapshot(),
          word = snapshot.metadata.word!,
          author = word.trackingAuthor ?? "",
          fold = word.revisions.some((revision) => {
            if (revision.kind !== "insert" || revision.author !== author) return false;
            const node = start.querySelector<HTMLElement>(
              '[data-node-id="' + CSS.escape(revision.nodeId) + '"]',
            );
            if (!node) return false;
            const begin = offsetAt(start, node, 0, projection.ignored);
            return from >= begin && to <= begin + session.getNode(revision.nodeId)!.text.length;
          }),
          caret = (fold ? from : to) + sourceText(template.content, () => false).length,
          selection = {
            anchor: { nodeId, offset: caret, path: [] },
            focus: { nodeId, offset: caret, path: [] },
          };
        const candidate = session.prepare(
          {
            protocolVersion: 1,
            documentId: snapshot.manifest.documentId,
            baseRevision: snapshot.manifest.revision,
            operations: [
              ...prefix,
              {
                op: "trackWordInlineContent",
                nodeId,
                from,
                to,
                expectedText: session.getNode(nodeId)!.text.slice(from, to),
                html: template.innerHTML,
              },
            ],
          },
          resources,
        );
        committing = true;
        try {
          session.commit(candidate);
        } catch (error) {
          session.discard(candidate);
          throw error;
        } finally {
          committing = false;
        }
        refresh(selection);
        return;
      }
      mutate(
        range,
        () => {
          if (block) {
            paragraphBreak(range);
            const current = selectedRange()!,
              parent =
                current.startContainer.nodeType === 1
                  ? (current.startContainer as Element)
                  : current.startContainer.parentElement!;
            const paragraph = parent.closest(blockSelector);
            if (!paragraph)
              throw new DocumentError("UnsupportedFeature", "当前位置不支持区块粘贴。");
            const last = template.content.lastChild!;
            paragraph.before(template.content);
            select(last, last.childNodes.length);
          } else {
            const last = template.content.lastChild;
            deleteRange(range);
            range.insertNode(template.content);
            if (last)
              select(
                last.parentNode!,
                [...last.parentNode!.childNodes].indexOf(last as ChildNode) + 1,
              );
          }
        },
        false,
        resources,
        prefix,
      );
    } catch (reason) {
      report(reason);
    }
  }
  function copy(event: ClipboardEvent, cut: boolean): void {
    if (!root || !frame || !event.clipboardData) return;
    event.preventDefault();
    if (composing) return;
    const range = selectedRange();
    if (!range || range.collapsed) return;
    const holder = frame.createElement("div");
    let fragment: Node = projection.serializeRange(range),
      ancestor = range.commonAncestorContainer;
    if (ancestor.nodeType !== 1) ancestor = ancestor.parentNode!;
    while (ancestor !== root && root.contains(ancestor)) {
      // A selection within a paragraph is inline. Only explicitly selected
      // blocks or table structure carry block wrappers to the clipboard.
      if ((ancestor as Element).matches(blockSelector + ",section,article,div,ul,ol")) break;
      const wrapper = projection.shell(ancestor as Element);
      if (wrapper) {
        wrapper.appendChild(fragment);
        fragment = wrapper;
      }
      if ((ancestor as Element).matches("table")) break;
      ancestor = ancestor.parentNode!;
    }
    holder.append(fragment);
    const source = holder;
    try {
      copyClipboardImages(source, viewModel());
      copyClipboardFormat(
        source,
        viewModel(),
        (ancestor as HTMLElement).dataset.nodeId ??
          session.snapshot().metadata.word?.stories[0]?.rootNodeId ??
          root.dataset.nodeId!,
        projection.completeClipboardNodes(range),
      );
      if (source.innerHTML.length > WORD_CLIPBOARD_BYTES)
        throw new DocumentError("ResourceLimitExceeded", "复制内容超过 64 MiB。");
    } catch (reason) {
      report(reason);
      return;
    }
    stripClipboardIdentities(source);
    pendingCut = undefined;
    if (cut && session.snapshot().metadata.word) {
      try {
        const { start, end, from, to, expectedText } = clipboardRange(range);
        const nodeId = start.dataset.nodeId!,
          endNodeId = end.dataset.nodeId!,
          snapshot = session.snapshot(),
          cutId = "word-cut-" + createOpaqueId(),
          token = createOpaqueId();
        const candidate = session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations: [
            {
              op: "cutWordContent",
              nodeId,
              ...(nodeId !== endNodeId ? { endNodeId } : {}),
              from,
              to,
              expectedText,
              cutId,
            },
          ],
        });
        const retained = candidate.model.getNode(nodeId) ?? candidate.model.getNode(endNodeId)!;
        const landing = {
          nodeId: retained.nodeId,
          offset: Math.min(from, retained.text.length),
          path: [] as number[],
        };
        committing = true;
        try {
          session.commit(candidate);
        } catch (reason) {
          session.discard(candidate);
          throw reason;
        } finally {
          committing = false;
        }
        const marker = source.firstElementChild ?? source.appendChild(frame.createElement("span"));
        marker.setAttribute("data-doffi-word-cut", token);
        pendingCut = {
          token,
          cutId,
          revision: session.snapshot().manifest.revision,
          html: source.innerHTML,
        };
        refresh({ anchor: landing, focus: landing });
      } catch (reason) {
        report(reason);
        return;
      }
    } else if (cut)
      mutate(range, () => {
        deleteRange(range);
        select(range.startContainer, range.startOffset);
      });
    event.clipboardData.setData("text/html", source.innerHTML);
    event.clipboardData.setData("text/plain", sourceSelectionText(source));
  }
  function changeListLevel(direction: -1 | 1): void {
    const snapshot = session.snapshot(),
      definitions = new Map<string, import("@doffi/core").WordMetadata["numbering"][number]>(),
      operations: import("@doffi/core").DocumentOperation[] = [];
    for (const paragraph of selectedParagraphs()) {
      const numberingId = paragraph.dataset.wordNumbering;
      if (!numberingId) continue;
      const current = Number(paragraph.dataset.wordLevel ?? 0),
        level = Math.min(8, current + direction);
      if (level === current) continue;
      if (level < 0) {
        operations.push({
          op: "setWordList",
          nodeId: paragraph.dataset.nodeId!,
          numberingId: null,
        });
        continue;
      }
      const source = snapshot.metadata.word!.numbering.find(
        (item) => item.numberingId === numberingId,
      )!;
      let definition = definitions.get(numberingId);
      if (
        !source.levels.some((item) => item.level === level) &&
        !definition?.levels.some((item) => item.level === level)
      ) {
        definition ??= JSON.parse(
          JSON.stringify(source),
        ) as import("@doffi/core").WordMetadata["numbering"][number];
        const base = source.levels.find((item) => item.level === current) ?? source.levels[0]!;
        definition.levels.push({
          ...base,
          level,
          cssClass: "number-level-" + createOpaqueId(),
          text:
            base.format === "bullet"
              ? base.text
              : Array.from({ length: level + 1 }, (_, n) => "%" + (n + 1)).join(".") + ".",
          start: 1,
        });
        definitions.set(numberingId, definition);
      }
      const style = frame!.defaultView!.getComputedStyle(paragraph);
      operations.push(
        { op: "setWordList", nodeId: paragraph.dataset.nodeId!, numberingId, level },
        {
          op: "setWordParagraph",
          nodeId: paragraph.dataset.nodeId!,
          properties: {
            "margin-left":
              Math.max(
                0,
                Math.round(parseFloat(style.marginLeft) * 0.75 * 100) / 100 + direction * 24,
              ) + "pt",
            "text-indent": Math.round(parseFloat(style.textIndent) * 0.75 * 100) / 100 + "pt",
          },
        },
      );
    }
    if (!operations.length) return;
    const article = new DOMParser()
      .parseFromString(snapshot.html, "text/html")
      .querySelector<HTMLElement>('[data-office-kind="document"]')!;
    savedSelection = root ? bookmark(root, projection.ignored) : undefined;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations: [
          ...[...definitions.values()].map((definition) => ({
            op: "setWordNumbering" as const,
            nodeId: article.dataset.nodeId!,
            numberingId: definition.numberingId,
            copyFromNumberingId: definition.numberingId,
            levels: definition.levels,
          })),
          ...operations,
        ],
      }),
    );
  }
  function keydown(event: KeyboardEvent): void {
    if (!root || composing) return;
    const key = event.key.toLowerCase();
    if (
      (event.ctrlKey || event.metaKey) &&
      ["z", "y", "b", "i", "u", "e", "l", "r"].includes(key)
    ) {
      event.preventDefault();
      try {
        if (key === "z" || key === "y") {
          if (event.shiftKey || key === "y") session.redo();
          else session.undo();
        } else if (["e", "l", "r"].includes(key))
          align(key === "e" ? "center" : key === "l" ? "left" : "right");
        else toggle(key === "b" ? "bold" : key === "i" ? "italic" : "underline");
      } catch (reason) {
        report(reason);
      }
      return;
    }
    const range = selectedRange();
    if (!range) return;
    if (event.key === "Enter") {
      event.preventDefault();
      mutate(
        range,
        () => (event.shiftKey ? lineBreak(range) : paragraphBreak(range)),
        !event.shiftKey,
      );
    }
    if (event.key === "Tab") {
      const parent =
        range.startContainer.nodeType === 1
          ? (range.startContainer as Element)
          : range.startContainer.parentElement;
      const cell = parent?.closest("td,th"),
        cells = [...root.querySelectorAll("td,th")].filter(
          (node) => !projection.ignored(node) && !node.hasAttribute("data-word-grid-gap"),
        );
      if (cell && cells.includes(cell)) {
        event.preventDefault();
        const next = cells[cells.indexOf(cell) + (event.shiftKey ? -1 : 1)];
        if (next) select(...caretAt(next, 0, projection.ignored));
      } else if (parent?.closest("[data-word-numbering]")) {
        event.preventDefault();
        try {
          changeListLevel(event.shiftKey ? -1 : 1);
        } catch (error) {
          report(error);
        }
      }
    }
  }
  function beforeinput(event: InputEvent): void {
    if (!root) return;
    let range = selectedRange();
    const target = event.getTargetRanges?.()[0];
    if (target) {
      const native = frame!.createRange();
      native.setStart(target.startContainer, target.startOffset);
      native.setEnd(target.endContainer, target.endOffset);
      // Browsers can canonicalize a run-start caret onto the preceding run's
      // end. Equal text positions retain the user's explicit source context.
      if (
        !(
          event.inputType.startsWith("insert") &&
          range?.collapsed &&
          native.collapsed &&
          offsetAt(root, range.startContainer, range.startOffset, projection.ignored) ===
            offsetAt(root, native.startContainer, native.startOffset, projection.ignored)
        )
      )
        range = native;
    }
    if (!range || !root.contains(range.commonAncestorContainer)) {
      event.preventDefault();
      return;
    }
    if (composing || event.isComposing || event.inputType === "insertCompositionText") {
      pending = inputStep(root, range, projection.ignored);
      wholeStoryReplacement ||=
        pending.from === 0 &&
        pending.to > 0 &&
        pending.to === sourceText(root, projection.ignored).length;
      return;
    }
    event.preventDefault();
    if (event.inputType === "historyUndo" || event.inputType === "historyRedo") {
      try {
        if (event.inputType === "historyUndo") session.undo();
        else session.redo();
      } catch (reason) {
        report(reason);
      }
      return;
    }
    if (["insertText", "insertReplacementText"].includes(event.inputType))
      mutate(range, () => {
        let content: Node = frame!.createTextNode(event.data ?? "");
        if (typingStyles.size) {
          const span = frame!.createElement("span");
          for (const [property, value] of typingStyles) span.style.setProperty(property, value);
          span.append(content);
          content = span;
        }
        insert(range!, content);
      });
    else if (event.inputType.startsWith("delete")) {
      range = deletionRange(range, event.inputType);
      mutate(range, () => {
        deleteRange(range!);
        select(range!.startContainer, range!.startOffset);
      });
    } else if (["insertParagraph", "insertLineBreak"].includes(event.inputType))
      mutate(
        range,
        () => (event.inputType === "insertParagraph" ? paragraphBreak(range!) : lineBreak(range!)),
        event.inputType === "insertParagraph",
      );
    else
      report(new DocumentError("UnsupportedFeature", "当前不支持此输入操作：" + event.inputType));
  }
  function updateMarks(): void {
    clearMarks();
    clearMarks = () => {};
    if (!root) return;
    projection.setViewAttribute(root, "data-doffi-show-marks", showMarks ? "" : null);
    if (showMarks)
      clearMarks = showWordWhitespace(
        root,
        projection,
        options.limits?.maxNodes ?? DEFAULT_RENDER_LIMITS.maxNodes,
      );
  }
  function activate(source: HTMLElement): void {
    if (!root || !frame) return;
    const reviewStyle = frame.createElement("style");
    reviewStyle.textContent =
      '[data-sdk-word-revision="insert"]{background:#dff7e5;text-decoration:underline}[data-sdk-word-revision="delete"]{background:#ffe3e3;text-decoration:line-through}[data-sdk-word-revision="format"]{outline:1px dotted #827600}[data-sdk-word-text-review="insert"]{background:#dff7e5!important}[data-sdk-word-text-review="delete"]{background:#ffe3e3!important}' +
      (revisionDisplay === "final"
        ? '[data-sdk-word-deleted]{display:none!important}[data-sdk-word-revision="insert"]{background:transparent;text-decoration:none}[data-sdk-word-text-review="insert"]{color:var(--sdk-word-authored-color)!important;background:var(--sdk-word-authored-background)!important;text-decoration-line:var(--sdk-word-authored-decoration)!important}'
        : revisionDisplay === "original"
          ? '[data-sdk-word-inserted]{display:none!important}[data-sdk-word-revision="delete"]{background:transparent;text-decoration:none}[data-sdk-word-text-review="delete"]{color:var(--sdk-word-authored-color)!important;background:var(--sdk-word-authored-background)!important;text-decoration-line:var(--sdk-word-authored-decoration)!important}'
          : "");
    frame.head.append(reviewStyle);
    viewControls.signal.addEventListener("abort", () => reviewStyle.remove(), { once: true });
    for (const item of session.snapshot().metadata.word?.revisions ?? []) {
      const node = root.querySelector<HTMLElement>(
        '[data-node-id="' + CSS.escape(item.nodeId) + '"]',
      );
      if (node) {
        projection.setViewAttribute(
          node,
          "data-sdk-word-revision",
          ["moveTo", "insert", "rowInsert", "cellInsert"].includes(item.kind)
            ? "insert"
            : ["moveFrom", "delete", "rowDelete", "cellDelete"].includes(item.kind)
              ? "delete"
              : item.kind,
        );
        if (["moveTo", "insert", "rowInsert", "cellInsert"].includes(item.kind))
          projection.setViewAttribute(node, "data-sdk-word-inserted", "true");
        if (["delete", "moveFrom", "rowDelete", "cellDelete"].includes(item.kind)) {
          projection.setViewAttribute(node, "data-sdk-word-deleted", "true");
          projection.setViewAttribute(node, "contenteditable", "false");
        }
      }
    }
    projectWordColumnReview(
      root,
      session.snapshot().metadata.word,
      revisionDisplay,
      projection,
      options.limits?.maxNodes ?? DEFAULT_RENDER_LIMITS.maxNodes,
    );
    drawingControls.bind(root, projection, viewControls.signal);
    bindWordTableResize(root, session, report, viewControls.signal);
    chartControls.bind(root, viewControls.signal);
    controlControls.bind(root, viewControls.signal);
    mathControls.bind(root, viewControls.signal);
    bindWordSelectionScroll(stage, preview.element, root, {
      signal: viewControls.signal,
      enabled: () => !composing && !disposed,
      ignored: projection.ignored,
      changed() {
        if (root) savedSelection = bookmark(root, projection.ignored);
        notifyFormat();
      },
    });
    root.contentEditable = "true";
    root.setAttribute("data-office-editor", "");
    root.setAttribute("role", "textbox");
    root.setAttribute("aria-label", "文档正文");
    root.setAttribute("aria-multiline", "true");
    root.spellcheck = false;
    for (const marker of root.querySelectorAll<HTMLElement>(".sdk-number-marker,[data-word-tab]"))
      marker.contentEditable = "false";
    // Noneditable marker text corrupts Chromium's native composition offsets:
    // later preedit updates can target the number instead of the new paragraph.
    // Keep the same visible label without adding text to the input DOM.
    for (const marker of root.querySelectorAll<HTMLElement>(".sdk-number-marker")) {
      if (!projection.ignored(marker)) continue;
      marker.setAttribute("data-sdk-number-text", marker.textContent ?? "");
      marker.replaceChildren();
    }
    projection.capture(root, source);
    updateMarks();
    bindWordLineSelection(frame, root, {
      signal: viewControls.signal,
      enabled: () => !disposed && !composing,
      ignored: projection.ignored,
      layout: () => preview.layout(),
      setup: () => wordMetadata().pageSetup,
      changed() {
        if (root) savedSelection = bookmark(root, projection.ignored);
        notifyFormat();
      },
    });
    if (!frame.querySelector("style[data-sdk-input]")) {
      const viewStyle = frame.createElement("style");
      viewStyle.setAttribute("data-sdk-input", "");
      viewStyle.textContent =
        '[data-office-editor][data-doffi-show-marks] :is(p,h1,h2,h3,h4,h5,h6)::after{content:"¶";color:var(--sdk-word-mark-color,#8a9ab0);font-family:var(--sdk-word-mark-font-family,inherit);font-size:var(--sdk-word-mark-font-size,inherit);font-weight:var(--sdk-word-mark-font-weight,normal);font-style:var(--sdk-word-mark-font-style,normal);text-decoration-line:var(--sdk-word-mark-text-decoration-line,none);text-decoration-style:var(--sdk-word-mark-text-decoration-style,solid);position:absolute;pointer-events:none}[data-office-editor] .sdk-space-mark{position:relative}[data-office-editor][data-doffi-show-marks] .sdk-space-mark::after{content:attr(data-sdk-space-mark);position:absolute;left:0;top:0;color:#8a9ab0;pointer-events:none}[data-office-editor][data-doffi-show-marks] [data-word-tab]::after{content:"→";color:#8a9ab0;position:absolute;left:0;pointer-events:none}[data-office-editor]:focus{outline:none}[data-office-editor]{min-height:100px}[data-office-editor] [data-word-tab]{user-select:all;position:relative}[data-office-editor] .sdk-number-marker::before{content:attr(data-sdk-number-text)}';
      frame.head.append(viewStyle);
    }
    if (!root.querySelector(blockSelector)) {
      const first = wordMetadata().sections[0]!.rootNodeId;
      const target =
        first === root.dataset.nodeId
          ? (root.querySelector(".sdk-page-flow") ?? root)
          : (root.querySelector('[data-node-id="' + CSS.escape(first) + '"]') ?? root);
      const paragraph = frame.createElement("p");
      paragraph.append(...target.childNodes);
      target.append(paragraph);
      if (!paragraph.childNodes.length) {
        const filler = frame.createElement("br");
        projection.markDerived(filler);
        paragraph.append(filler);
      }
    }
    root.addEventListener("beforeinput", beforeinput as EventListener, {
      signal: viewControls.signal,
    });
    frame.addEventListener(
      "dblclick",
      (event) => {
        if (composing) return;
        const target = event.target as Element | null,
          storyId = target?.closest<HTMLElement>("[data-sdk-word-note-story]")?.dataset
            .sdkWordNoteStory;
        if (
          storyId &&
          session
            .snapshot()
            .metadata.word?.stories.some(
              (story) => story.storyId === storyId && ["footnote", "endnote"].includes(story.kind),
            )
        ) {
          event.preventDefault();
          editStory(storyId);
        }
      },
      { signal: viewControls.signal },
    );
    root.addEventListener(
      "input",
      () => {
        if (!root || disposed) return;
        if (pending) {
          const edits = inputProofs(root, pending, projection.ignored);
          proofs.push(...edits);
          const rootEdit = edits.find((edit) => edit.nodeId === root!.dataset.nodeId);
          if (rootEdit)
            compositionInsertion = {
              from: rootEdit.from,
              to: rootEdit.from + rootEdit.text.length,
            };
        }
        pending = undefined;
        if (!composing) flush();
      },
      { signal: viewControls.signal },
    );
    root.addEventListener("keydown", keydown, { signal: viewControls.signal });
    root.addEventListener("paste", paste as EventListener, { signal: viewControls.signal });
    root.addEventListener(
      "copy",
      ((event: ClipboardEvent) => copy(event, false)) as EventListener,
      { signal: viewControls.signal },
    );
    root.addEventListener("cut", ((event: ClipboardEvent) => copy(event, true)) as EventListener, {
      signal: viewControls.signal,
    });
    root.addEventListener(
      "drop",
      (event) => {
        event.preventDefault();
        const image = Array.from(event.dataTransfer?.files ?? []).find((file) =>
          /^image\/(?:png|jpeg|webp)$/.test(file.type),
        );
        if (image) {
          void drawingControls.insertImage(image).catch(report);
          return;
        }
        report(new DocumentError("UnsupportedFeature", "请通过粘贴或宿主资源接口插入内容。"));
      },
      { signal: viewControls.signal },
    );
    root.addEventListener(
      "compositionstart",
      () => {
        composing = true;
        compositionRevision = revision;
        notifyFormat();
      },
      { signal: viewControls.signal },
    );
    root.addEventListener(
      "compositionend",
      () => {
        queueMicrotask(() => {
          if (disposed || !root || !frame) return;
          composing = false;
          if (
            compositionInsertion &&
            typingStyles.size &&
            revision === session.snapshot().manifest.revision
          ) {
            const selected = bookmark(root, projection.ignored),
              range = frame.createRange();
            range.setStart(...caretAt(root, compositionInsertion.from, projection.ignored));
            range.setEnd(...caretAt(root, compositionInsertion.to, projection.ignored));
            styleRange(range, typingStyles);
            restoreBookmark(root, selected, projection.ignored);
          }
          flush();
        });
      },
      { signal: viewControls.signal },
    );
  }
  const unsubscribe = session.subscribe((event) => {
    if (event.type !== "error" && !committing && !disposed && !composing) refresh();
  });
  const preview = mountDocumentPreview(plane, viewModel(), { ...options, title: "正文编辑" });
  let zoom = 1;
  const rulers = mountWordRulers(
    stage,
    preview.element,
    () => root,
    () => preview.layout(),
    () => wordMetadata().pageSetup,
    {
      pageId() {
        const range = selectedRange();
        if (!range || !root) return;
        const rect = range.getBoundingClientRect();
        if (!rect.height) return;
        const y = (rect.top + rect.height / 2 - root.getBoundingClientRect().top) * 0.75;
        return preview
          .layout()
          .pages.find((page) => y >= page.rect.y && y < page.rect.y + page.rect.height)?.pageId;
      },
      read() {
        const paragraph = selectedParagraphs(false)[0];
        if (!paragraph || !frame) return;
        const style = frame.defaultView!.getComputedStyle(paragraph),
          tabs = wordMetadata().word?.tabs.find(
            (tab) => tab.paragraphNodeId === paragraph.dataset.nodeId,
          );
        return {
          left: Math.round((parseFloat(style.marginLeft) * 0.75 || 0) * 100) / 100,
          right: Math.round((parseFloat(style.marginRight) * 0.75 || 0) * 100) / 100,
          first: Math.round((parseFloat(style.textIndent) * 0.75 || 0) * 100) / 100,
          tabs:
            tabs?.stops.filter((stop) => stop.alignment !== "clear").map((stop) => stop.position) ??
            [],
        };
      },
      commit(kind, value, previous) {
        const snapshot = session.snapshot(),
          operations: import("@doffi/core").DocumentOperation[] = [];
        for (const paragraph of selectedParagraphs()) {
          if (kind === "tab") {
            const tabs = snapshot.metadata.word?.tabs.find(
              (tab) => tab.paragraphNodeId === paragraph.dataset.nodeId,
            );
            operations.push({
              op: "setWordTabs",
              nodeId: paragraph.dataset.nodeId!,
              defaultInterval:
                tabs?.defaultInterval ?? snapshot.metadata.word?.tabs[0]?.defaultInterval ?? 36,
              stops: [
                ...(tabs?.stops ?? []).filter(
                  (stop) => stop.position !== previous && stop.position !== value,
                ),
                ...(value >= 0
                  ? [{ position: value, alignment: "left" as const, leader: "none" as const }]
                  : []),
              ].sort((a, b) => a.position - b.position),
            });
          } else
            operations.push({
              op: "setWordParagraph",
              nodeId: paragraph.dataset.nodeId!,
              properties: {
                [kind === "left"
                  ? "margin-left"
                  : kind === "right"
                    ? "margin-right"
                    : "text-indent"]: value + "pt",
              },
            });
        }
        if (operations.length)
          session.commit(
            session.prepare({
              protocolVersion: 1,
              documentId: snapshot.manifest.documentId,
              baseRevision: snapshot.manifest.revision,
              operations,
            }),
          );
      },
      commitMargin(side, value, sectionId, pageIndex) {
        const snapshot = session.snapshot(),
          metadata = wordMetadata(),
          section =
            metadata.sections.find((section) => section.sectionId === sectionId) ??
            metadata.sections.find((section) => section.rootNodeId === currentSection()) ??
            metadata.sections[0]!,
          setup = section.pageSetup ?? metadata.pageSetup,
          physical =
            preview.layout().pages.find((page) => page.index === pageIndex)?.setup ?? setup,
          sourceSide =
            metadata.word?.mirrorMargins &&
            pageIndex !== undefined &&
            pageIndex % 2 === 1 &&
            (side === "left" || side === "right")
              ? side === "left"
                ? "right"
                : "left"
              : side,
          delta = physical.margins[side] - setup.margins[sourceSide],
          next = {
            ...setup,
            margins: { ...setup.margins, [sourceSide]: Math.max(0, value - delta) },
          };
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: snapshot.manifest.documentId,
            baseRevision: snapshot.manifest.revision,
            operations: [
              { op: "setWordSection", nodeId: section.rootNodeId, changes: { pageSetup: next } },
            ],
          }),
        );
      },
      onError: report,
    },
  );
  function resizePlane(
    anchor: ZoomAnchor | undefined = captureZoomAnchor(stage, plane, zoom),
  ): void {
    if (disposed || !stage.clientWidth) return;
    const word = wordMetadata(),
      minimum =
        (Math.max(
          word.pageSetup.width,
          ...word.sections.map((section) => (section.pageSetup ?? word.pageSetup).width),
        ) *
          4) /
          3 +
        48;
    const viewportWidth = Math.max(stage.clientWidth, minimum) + "px";
    if (preview.element.style.width !== viewportWidth) preview.element.style.width = viewportWidth;
    preview.element.style.transformOrigin = "0 0";
    preview.element.style.transform = `scale(${zoom})`;
    const width = preview.element.offsetWidth * zoom + "px",
      height = preview.element.offsetHeight * zoom + "px";
    if (plane.style.width !== width) plane.style.width = width;
    if (plane.style.height !== height) plane.style.height = height;
    restoreZoomAnchor(stage, plane, zoom, anchor);
    rulers.refresh();
  }
  function setZoom(value: number): void {
    if (disposed || !Number.isFinite(value) || value < 0.5 || value > 2)
      throw new DocumentError("InvalidOperation", "缩放范围为 50% 至 200%。");
    const anchor = captureZoomAnchor(stage, plane, zoom);
    zoom = value;
    resizePlane(anchor);
  }
  const geometryObserver = new ResizeObserver(() => resizePlane());
  geometryObserver.observe(stage);
  geometryObserver.observe(preview.element);
  const ready = preview.ready
    .then(() => {
      if (disposed) return;
      frame = preview.element.contentDocument!;
      root = rootElement(frame);
      const source = rootElement(new DOMParser().parseFromString(initial.html, "text/html"));
      activate(source);
      if (revision !== session.snapshot().manifest.revision) refresh();
      resizePlane();
      notifyFormat();
      frame.addEventListener(
        "selectionchange",
        () => {
          if (root) {
            const next = bookmark(root, projection.ignored);
            if (!composing && typingAnchor && next && JSON.stringify(next) !== typingAnchor) {
              typingStyles.clear();
              typingAnchor = undefined;
            }
            savedSelection = next ?? savedSelection;
            notifyFormat();
            rulers.refresh();
          }
        },
        { signal: controls.signal },
      );
    })
    .catch((reason: unknown) => {
      report(reason);
      throw reason;
    });
  return {
    element,
    ready,
    async settled() {
      await ready;
      const result = await preview.settled();
      resizePlane();
      return result;
    },
    formatState,
    setZoom,
    editStory,
    setRulers(visible) {
      const anchor = captureZoomAnchor(stage, plane, zoom);
      rulers.setVisible(visible);
      resizePlane(anchor);
    },
    selection() {
      const range = selectedRange();
      if (!frame || !range || range.collapsed || composing || disposed) return undefined;
      const anchor = new PreviewSource(frame, viewModel()).textAnchor(range, session);
      if (!anchor) return undefined;
      const manifest = session.snapshot().manifest;
      return { documentId: manifest.documentId, baseRevision: manifest.revision, anchor };
    },
    clipboard(action) {
      if (!root || disposed || composing)
        return Promise.reject(new DocumentError("InvalidOperation", "正文尚不可编辑。"));
      const expected = revision,
        selection = JSON.stringify(bookmark(root, projection.ignored));
      return executeClipboard(
        root,
        () => {
          root?.focus();
          if (root) restoreBookmark(root, savedSelection, projection.ignored);
        },
        action,
        () =>
          !disposed &&
          !composing &&
          revision === expected &&
          (!root || JSON.stringify(bookmark(root, projection.ignored)) === selection),
      );
    },
    layout: () => preview.layout(),
    focus: () => {
      root?.focus();
    },
    execute(command) {
      if (disposed || !root || composing)
        throw new DocumentError("InvalidOperation", "请等待正文输入就绪或完成输入法组合。");
      if (command?.type === "toggle" && ["bold", "italic", "underline"].includes(command.property))
        toggle(command.property);
      else if (
        command?.type === "format" &&
        [
          "font-weight",
          "font-style",
          "text-decoration",
          "color",
          "font-size",
          "font-family",
          "background-color",
          "vertical-align",
          "letter-spacing",
          "text-decoration-line",
        ].includes(command.property) &&
        typeof command.value === "string"
      )
        format(command.property, command.value);
      else if (
        command?.type === "align" &&
        ["left", "center", "right", "justify"].includes(command.value)
      )
        align(command.value);
      else if (command?.type === "word") wordCommand(command);
      else throw new DocumentError("InvalidOperation", "未知正文格式命令。");
      notifyFormat();
    },
    dispose: () => {
      if (disposed) return;
      disposed = true;
      geometryObserver.disconnect();
      rulers.dispose();
      unsubscribe();
      controls.abort();
      viewControls.abort();
      preview.dispose();
      clearMarks();
      reviewControls.dispose();
      chartControls.dispose();
      controlControls.dispose();
      mathControls.dispose();
      linkControls.dispose();
      referenceControls.dispose();
      drawingControls.dispose();
      listControls.dispose();
      styleControls.dispose();
      fontControls.dispose();
      watermarkControls.dispose();
      pageControls.dispose();
      tableControls.dispose();
      wordPanel?.remove();
      wordPanel = undefined;
      root = undefined;
      frame = undefined;
      proofs = [];
      pending = undefined;
      savedSelection = undefined;
      element.remove();
    },
  };
}
