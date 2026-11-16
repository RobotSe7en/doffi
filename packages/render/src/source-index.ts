import { DocumentError } from "@doffi/core";
import type { Anchor, DeepReadonly, ContentDocumentModel, Rectangle, TextRange } from "@doffi/core";
import type { RenderedPage } from "./contracts.js";
import { viewNodeRole } from "./view-nodes.js";
import { wordRowVisibleRect } from "./word-table-layout.js";

interface SourceElement {
  readonly element: HTMLElement;
  readonly from: number;
  readonly to: number;
}
interface Token {
  readonly node: Node;
  readonly text: string;
  readonly from: number;
  readonly to: number;
  readonly owner: HTMLElement;
}
const blocks = "p,h1,h2,h3,h4,h5,h6,li,td,th";
const hitTargets = blocks + ",img,table,[data-word-drawing]";

function boundary(text: string, offset: number): boolean {
  const before = text.charCodeAt(offset - 1),
    after = text.charCodeAt(offset),
    code = text.codePointAt(offset);
  return (
    !(before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff) &&
    !(
      offset > 0 &&
      code !== undefined &&
      ((code >= 0xfe00 && code <= 0xfe0f) || (code >= 0xe0100 && code <= 0xe01ef))
    )
  );
}
function intersects(a: Rectangle, b: Rectangle): boolean {
  return a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;
}

/** Capture source positions before the renderer inserts labels or replaces tab text. */
export class PreviewSource {
  readonly elements: Map<string, SourceElement> = new Map();
  readonly tokens: Token[] = [];
  readonly root: HTMLElement;
  private readonly document: Document;
  private pages: readonly RenderedPage[] = [];
  private fieldViews = new WeakMap<Node, HTMLElement>();
  private readonly fieldScopes = new Map<string, HTMLElement>();
  private readonly instances = new Map<string, HTMLElement[]>();
  private readonly mainOrder: string[];
  clearInstances(): void {
    this.instances.clear();
  }
  registerInstances(root: HTMLElement): void {
    for (const node of [root, ...root.querySelectorAll<HTMLElement>("[data-node-id]")]) {
      const id = node.dataset.nodeId;
      if (id && this.elements.has(id))
        this.instances.set(id, [...(this.instances.get(id) ?? []), node]);
    }
  }
  instanceElements(nodeId: string): readonly HTMLElement[] {
    return Object.freeze([...(this.instances.get(nodeId) ?? [])]);
  }
  elementRectangles(id: string): DOMRect[] {
    const entry = this.elements.get(id),
      copies = [...(this.instances.get(id) ?? [])].flatMap((node) =>
        [...node.getClientRects()].flatMap((rect) => this.instanceClip(node, rect)),
      );
    if (!entry) return copies;
    if (entry.element.hasAttribute("data-word-drawing")) {
      const frame = entry.element.closest<HTMLElement>("[data-sdk-word-object]") ?? entry.element;
      if (
        copies.length ||
        this.document.defaultView!.getComputedStyle(frame).visibility === "hidden"
      )
        return copies;
      return [...frame.getClientRects()];
    }
    if (!entry.element.matches(blocks + ",span,a,strong,em,ins,del"))
      return [...entry.element.getClientRects(), ...copies];
    let lo = 0,
      hi = this.tokens.length;
    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2);
      if (this.tokens[mid]!.to <= entry.from) lo = mid + 1;
      else hi = mid;
    }
    const rectangles: DOMRect[] = [];
    for (
      let index = lo;
      index < this.tokens.length && this.tokens[index]!.from < entry.to;
      index++
    ) {
      const token = this.tokens[index]!;
      if (this.visible(token)) rectangles.push(...this.tokenRectangles(token));
    }
    return [
      ...(rectangles.length
        ? rectangles
        : [...entry.element.getClientRects()].flatMap((rect) => {
            const visible = wordRowVisibleRect(entry.element, rect);
            return visible ? [visible] : [];
          })),
      ...copies,
    ];
  }
  setPages(pages: readonly RenderedPage[]): void {
    this.pages = pages;
  }
  setFieldView(nodeId: string, display: HTMLElement): void {
    this.fieldScopes.set(nodeId, display);
    const entry = this.elements.get(nodeId);
    if (!entry) return;
    for (const token of this.tokens)
      if (token.from >= entry.from && token.to <= entry.to && entry.element.contains(token.node))
        this.fieldViews.set(token.node, display);
  }

  constructor(document: Document, model: ContentDocumentModel) {
    this.document = document;
    this.model = model;
    const documentRoot = document.querySelector<HTMLElement>(
      'body > [data-office-kind="document"][data-node-id],body > [data-office-kind="presentation"][data-node-id]',
    );
    if (!documentRoot) throw new DocumentError("InvalidDocument", "预览文档缺少根容器。");
    this.index(documentRoot);
    const main = model.snapshot().metadata.word?.stories.find((story) => story.kind === "main");
    this.root = main ? this.elements.get(main.rootNodeId)!.element : documentRoot;
    this.mainOrder = [this.root, ...this.root.querySelectorAll<HTMLElement>("[data-node-id]")].map(
      (node) => node.dataset.nodeId!,
    );
  }
  private readonly model: ContentDocumentModel;
  /** Text splitting at measured line boundaries changes DOM positions, never source offsets. */
  reindex(): void {
    this.elements.clear();
    this.tokens.length = 0;
    this.fieldViews = new WeakMap();
    this.index(
      this.document.querySelector<HTMLElement>(
        'body > [data-office-kind="document"][data-node-id],body > [data-office-kind="presentation"][data-node-id]',
      )!,
    );
    for (const [id, display] of this.fieldScopes) this.setFieldView(id, display);
  }
  private index(documentRoot: HTMLElement): void {
    let offset = 0;
    const visit = (node: Node): void => {
      if (viewNodeRole(node) === "remove") return;
      const from = offset;
      const element = node.nodeType === 1 ? (node as HTMLElement) : undefined;
      const atomic = element?.matches(
        'br,img,[data-word-tab],[data-word-drawing]:not([data-word-drawing="group"]),[data-word-note-reference],[data-word-equation]',
      );
      const text =
        node.nodeType === 3
          ? (node.nodeValue ?? "")
          : atomic
            ? element!.tagName === "IMG"
              ? "\ufffc"
              : element!.tagName === "BR"
                ? "\n"
                : (this.model.getNode(element!.dataset.nodeId!)?.text ?? "\t")
            : undefined;
      if (text !== undefined) {
        offset += text.length;
        const parent = element ?? node.parentElement!;
        const owner = parent.closest<HTMLElement>(blocks);
        // An un-identified paragraph cannot occur in a validated package.
        const identified = owner?.hasAttribute("data-node-id")
          ? owner
          : parent.closest<HTMLElement>("[data-node-id]")!;
        this.tokens.push({ node, text, from, to: offset, owner: identified });
      } else {
        for (const child of node.childNodes) visit(child);
      }
      if (element?.dataset.nodeId)
        this.elements.set(element.dataset.nodeId, { element, from, to: offset });
    };
    visit(documentRoot);
  }

  /** Only the main content story participates in document selection. */
  contains(node: Node): boolean {
    return this.root === node || this.root.contains(node);
  }

  private tokenRange(token: Token, from = 0, to = token.text.length): Range {
    const range = this.document.createRange();
    const display = this.fieldViews.get(token.node);
    if (display) {
      range.selectNodeContents(display);
      return range;
    }
    if (token.node.nodeType === 3) {
      range.setStart(token.node, from);
      range.setEnd(token.node, to);
    } else range.selectNode(token.node);
    return range;
  }
  private visible(token: Token): boolean {
    const element =
      this.fieldViews.get(token.node) ??
      (token.node.nodeType === 1 ? (token.node as HTMLElement) : token.node.parentElement!);
    if (
      element.hasAttribute("data-word-drawing") &&
      element.dataset.nodeId &&
      this.instances.has(element.dataset.nodeId)
    )
      return token.node.isConnected && this.elementRectangles(element.dataset.nodeId).length > 0;
    const style = this.document.defaultView!.getComputedStyle(element);
    return (
      token.node.isConnected &&
      style.visibility === "visible" &&
      this.tokenRectangles(token).length > 0
    );
  }
  private tokenRectangles(token: Token, from = 0, to = token.text.length): DOMRect[] {
    const owner =
      token.node.nodeType === 1 ? (token.node as HTMLElement) : token.node.parentElement!;
    if (owner.hasAttribute("data-word-drawing") && owner.dataset.nodeId)
      return this.elementRectangles(owner.dataset.nodeId);
    return [...this.tokenRange(token, from, to).getClientRects()].flatMap((rect) => {
      const visible = wordRowVisibleRect(owner, rect);
      return visible ? [visible] : [];
    });
  }
  private instanceClip(node: Element, rect: DOMRect): DOMRect[] {
    if (this.document.defaultView!.getComputedStyle(node).visibility !== "visible") return [];
    let left = rect.left,
      right = rect.right,
      top = rect.top,
      bottom = rect.bottom;
    for (let parent: Element | null = node.parentElement; parent; parent = parent.parentElement) {
      const style = this.document.defaultView!.getComputedStyle(parent),
        x = /^(?:hidden|clip|scroll|auto)$/.test(style.overflowX),
        y = /^(?:hidden|clip|scroll|auto)$/.test(style.overflowY);
      if (x || y) {
        const box = parent.getBoundingClientRect();
        if (x) {
          left = Math.max(left, box.left);
          right = Math.min(right, box.right);
        }
        if (y) {
          top = Math.max(top, box.top);
          bottom = Math.min(bottom, box.bottom);
        }
      }
    }
    return right > left + 0.01 && bottom > top + 0.01
      ? [new DOMRect(left, top, right - left, bottom - top)]
      : [];
  }
  private instanceTextRectangles(nodeId: string, from: number, to: number): DOMRect[] {
    return (this.instances.get(nodeId) ?? []).flatMap((root) => {
      let offset = 0;
      const rects: DOMRect[] = [];
      const visit = (node: Node): void => {
        if (viewNodeRole(node) === "remove") return;
        const element = node.nodeType === 1 ? (node as HTMLElement) : undefined,
          atomic = element?.matches(
            "br,img,[data-word-tab],[data-word-drawing],[data-word-native-object]:not([data-word-native-boundary]),[data-word-note-reference],[data-word-note-number],[data-word-equation]",
          ),
          size = node.nodeType === 3 ? node.textContent!.length : atomic ? 1 : undefined;
        if (size !== undefined) {
          const start = Math.max(from, offset),
            end = Math.min(to, offset + size);
          if (start < end) {
            const range = this.document.createRange();
            if (node.nodeType === 3) {
              range.setStart(node, start - offset);
              range.setEnd(node, end - offset);
            } else range.selectNode(node);
            for (const rect of range.getClientRects())
              rects.push(...this.instanceClip(element ?? node.parentElement!, rect));
          }
          offset += size;
        } else for (const child of node.childNodes) visit(child);
      };
      visit(root);
      return rects;
    });
  }

  textAnchor(range: Range, model: ContentDocumentModel): Anchor | undefined {
    if (range.collapsed) return undefined;
    if (!this.contains(range.startContainer) || !this.contains(range.endContainer))
      throw new DocumentError("InvalidOperation", "文字选择必须位于同一正文内容域。");
    const ranges: TextRange[] = [];
    const ownerTexts = new Map<string, string>();
    for (const token of this.tokens) {
      const display = this.fieldViews.get(token.node);
      if (
        !this.contains(token.node) ||
        !this.visible(token) ||
        !range.intersectsNode(display ?? token.node)
      )
        continue;
      const from =
        !display && token.node.nodeType === 3 && range.startContainer === token.node
          ? range.startOffset
          : 0;
      const to =
        !display && token.node.nodeType === 3 && range.endContainer === token.node
          ? range.endOffset
          : token.text.length;
      if (from === to) continue;
      const nodeId = token.owner.dataset.nodeId!;
      const source = this.elements.get(nodeId)!;
      const start = token.from - source.from + from,
        end = token.from - source.from + to;
      let text = ownerTexts.get(nodeId);
      if (text === undefined) {
        text = model.getNode(nodeId)!.text;
        ownerTexts.set(nodeId, text);
      }
      if (!boundary(text, start) || !boundary(text, end))
        throw new DocumentError(
          "InvalidOperation",
          "选择不能拆开 UTF-16 代理对或字符的变体选择符。",
        );
      const previous = ranges.at(-1);
      if (previous?.nodeId === nodeId && previous.to === start) {
        previous.to = end;
        previous.quote = text.slice(previous.from, end);
      } else ranges.push({ nodeId, from: start, to: end, quote: text.slice(start, end) });
    }
    return ranges.length === 1
      ? { kind: "text", ...ranges[0]! }
      : ranges.length
        ? { kind: "text-set", ranges }
        : undefined;
  }

  orderedIds(ids: readonly string[]): string[] {
    const wanted = new Set(ids);
    return this.mainOrder.filter((id) => wanted.has(id));
  }
  point(clientX: number, clientY: number): { x: number; y: number } {
    const rect = this.root.getBoundingClientRect();
    return {
      x: (Math.max(0, Math.min(rect.width, clientX - rect.left)) * 72) / 96,
      y: (Math.max(0, Math.min(rect.height, clientY - rect.top)) * 72) / 96,
    };
  }
  region(rect: Rectangle, layoutKey: string): Anchor {
    const surface = this.root.getBoundingClientRect();
    const hits = [...this.root.querySelectorAll<HTMLElement>("[data-node-id]")].filter(
      (element) => {
        if (!element.matches(hitTargets)) return false;
        return this.elementRectangles(element.dataset.nodeId!).some((box) =>
          intersects(rect, {
            x: ((box.left - surface.left) * 72) / 96,
            y: ((box.top - surface.top) * 72) / 96,
            width: (box.width * 72) / 96,
            height: (box.height * 72) / 96,
          }),
        );
      },
    );
    const page = this.pages.find(
      (item) =>
        rect.x >= item.rect.x &&
        rect.y >= item.rect.y &&
        rect.x + rect.width <= item.rect.x + item.rect.width &&
        rect.y + rect.height <= item.rect.y + item.rect.height,
    );
    return {
      kind: "region",
      surfaceId: this.root.dataset.nodeId!,
      rect,
      ...(page ? { pageId: page.pageId } : {}),
      nodeIds: [
        ...new Set(
          hits
            .filter(
              (element) => !hits.some((child) => child !== element && element.contains(child)),
            )
            .map((element) => element.dataset.nodeId!),
        ),
      ],
      layoutKey,
    };
  }

  /** A PPT region is relative to one stable slide, independent of its order. */
  slideRegion(surfaceId: string, rect: Rectangle): Anchor {
    const surface = this.elements.get(surfaceId)?.element,
      p = this.model.snapshot().metadata.presentation;
    if (!surface || !p?.slides.some((s) => s.rootNodeId === surfaceId))
      throw new DocumentError("InvalidOperation", "A presentation region requires a source slide.");
    const box = surface.getBoundingClientRect(),
      width = box.width * 0.75,
      height = box.height * 0.75;
    if (
      !Object.values(rect).every(Number.isFinite) ||
      rect.x < 0 ||
      rect.y < 0 ||
      rect.width <= 0 ||
      rect.height <= 0 ||
      rect.x + rect.width > width + 0.01 ||
      rect.y + rect.height > height + 0.01
    )
      throw new DocumentError(
        "InvalidOperation",
        "The selected region must stay inside its slide.",
      );
    const nodeIds = [...surface.querySelectorAll<HTMLElement>("[data-node-id]")]
      .filter(
        (node) =>
          this.elements.get(node.dataset.nodeId!)?.element === node && node.matches(hitTargets),
      )
      .filter((node) =>
        [...node.getClientRects()].some((hit) =>
          intersects(rect, {
            x: (hit.left - box.left) * 0.75,
            y: (hit.top - box.top) * 0.75,
            width: hit.width * 0.75,
            height: hit.height * 0.75,
          }),
        ),
      )
      .map((node) => node.dataset.nodeId!);
    return {
      kind: "region",
      surfaceId,
      rect: { ...rect },
      nodeIds: [...new Set(nodeIds)],
      layoutKey: this.layoutKey(surfaceId),
    };
  }

  rectangles(anchor: DeepReadonly<Anchor>): DOMRect[] {
    if (anchor.kind === "cells")
      throw new DocumentError(
        "UnsupportedFeature",
        "Cell anchors require a workbook grid projection.",
      );
    if (anchor.kind === "region") {
      const surface = this.elements.get(anchor.surfaceId)?.element;
      if (!surface) return [];
      const box = surface.getBoundingClientRect();
      return [
        new DOMRect(
          box.left + (anchor.rect.x * 96) / 72,
          box.top + (anchor.rect.y * 96) / 72,
          (anchor.rect.width * 96) / 72,
          (anchor.rect.height * 96) / 72,
        ),
      ];
    }
    if (anchor.kind === "node" || anchor.kind === "node-set") {
      const ids = anchor.kind === "node" ? [anchor.nodeId] : anchor.nodeIds;
      return ids.flatMap((id) => this.elementRectangles(id));
    }
    const ranges = anchor.kind === "text" ? [anchor] : anchor.ranges;
    return ranges.flatMap((range) => {
      const source = this.elements.get(range.nodeId);
      if (!source) return [];
      const start = source.from + range.from,
        end = source.from + range.to;
      return [
        ...this.instanceTextRectangles(range.nodeId, range.from, range.to),
        ...this.tokens.flatMap((token) => {
          if (token.to <= start || token.from >= end || !this.visible(token)) return [];
          return this.tokenRectangles(
            token,
            Math.max(0, start - token.from),
            Math.min(token.text.length, end - token.from),
          );
        }),
      ];
    });
  }

  /** Source-relative geometry and resolved faces; scrolling and SDK zoom are excluded. */
  layoutKey(surfaceId?: string): string {
    if (surfaceId) {
      const element = this.elements.get(surfaceId)?.element;
      if (!element)
        throw new DocumentError("InvalidOperation", "The source layout surface is missing.");
      let hash = 2166136261;
      const add = (value: unknown): void => {
        const text = JSON.stringify(value);
        for (let i = 0; i < text.length; i++) hash = Math.imul(hash ^ text.charCodeAt(i), 16777619);
      };
      const rootStyle = this.document.defaultView!.getComputedStyle(element);
      add([rootStyle.width, rootStyle.height]);
      for (const node of [
        element,
        ...element.querySelectorAll<HTMLElement>(
          "[data-node-id],[data-office-script],.sdk-ppt-marker,svg,text,path",
        ),
      ]) {
        const style = this.document.defaultView!.getComputedStyle(node),
          connector = !!node.closest('[data-office-kind="connector"]');
        // Global DOMRects and connector CTMs lose low bits when a slide moves
        // far down the document in Firefox. Local CSS, offsets and native SVG
        // coordinates describe the same layout without depending on page order.
        const offsets =
          node === element
            ? []
            : [node.offsetLeft, node.offsetTop, node.offsetWidth, node.offsetHeight];
        const geometry = [...node.attributes]
          .filter(
            (a) =>
              /^data-office-(?:x|y|width|height|rotation|flip-.+|child-.+)$/.test(a.name) ||
              (!connector &&
                [
                  "d",
                  "x",
                  "y",
                  "width",
                  "height",
                  "viewBox",
                  "cx",
                  "cy",
                  "r",
                  "points",
                  "transform",
                ].includes(a.name)),
          )
          .map((a) => [a.name, a.value]);
        let glyphBox: readonly number[] = [];
        if (node.localName === "text" && "getBBox" in node) {
          const box = (node as unknown as SVGGraphicsElement).getBBox();
          glyphBox = [box.x, box.y, box.width, box.height];
        }
        add([
          node.getAttribute("data-node-id"),
          node.localName,
          offsets,
          geometry,
          glyphBox,
          style.left,
          style.top,
          style.width,
          style.height,
          style.transform,
          style.fontFamily,
          style.fontSize,
          style.fontWeight,
          style.fontStyle,
          style.lineHeight,
          style.letterSpacing,
        ]);
      }
      return "slide-v1-" + (hash >>> 0).toString(16);
    }
    const view = this.document.defaultView!;
    const surface = this.root.getBoundingClientRect();
    let hash = 2166136261;
    const add = (value: string): void => {
      for (let index = 0; index < value.length; index++)
        hash = Math.imul(hash ^ value.charCodeAt(index), 16777619);
    };
    const prefix = this.pages.length ? "pages-v1" : "continuous-v1";
    add(prefix);
    add(JSON.stringify(this.pages));
    for (const [id, { element }] of this.elements) {
      if (!this.contains(element)) continue;
      const style = view.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      add(
        JSON.stringify([
          id,
          rect.left - surface.left,
          rect.top - surface.top,
          rect.width,
          rect.height,
          style.fontFamily,
          style.fontSize,
          style.fontWeight,
          style.fontStyle,
          style.lineHeight,
          style.letterSpacing,
        ]),
      );
    }
    for (const token of this.tokens) {
      if (!this.contains(token.node) || !this.visible(token)) continue;
      for (const rect of this.tokenRange(token).getClientRects())
        add(
          JSON.stringify([
            rect.left - surface.left,
            rect.top - surface.top,
            rect.width,
            rect.height,
          ]),
        );
    }
    return prefix + "-" + (hash >>> 0).toString(16);
  }
}
