import type {
  ContentDocumentModel,
  DocumentEditingSession,
  DocumentOperation,
} from "@doffi/core";
import { PreviewSource } from "@doffi/render";
import { caretAt } from "./rich-text/dom.js";

export interface WordMultiTextRange {
  readonly nodeId: string;
  readonly from: number;
  readonly to: number;
  readonly expectedText: string;
}

/** Non-contiguous multi-range text selection (hold Cmd/Ctrl while selecting).
 * Visualizes ranges purely in a view overlay without ever mutating source HTML,
 * and compiles batch formatting/deletion/case ops into ONE atomic transaction. */
export class WordMultiSelection {
  private items: WordMultiTextRange[] = [];
  private overlay: HTMLDivElement | undefined;

  get size(): number {
    return this.items.length;
  }

  ranges(): readonly WordMultiTextRange[] {
    return this.items;
  }

  clear(): void {
    this.items = [];
    this.overlay?.remove();
    this.overlay = undefined;
  }

  addFromDomRange(
    range: Range,
    frame: Document,
    session: DocumentEditingSession,
    viewModel: () => ContentDocumentModel,
  ): boolean {
    if (range.collapsed) return false;
    const anchor = new PreviewSource(frame, viewModel()).textAnchor(range, session);
    if (!anchor || !(anchor.kind === "text" || anchor.kind === "text-set")) return false;
    const incoming = (anchor.kind === "text" ? [anchor] : anchor.ranges).filter(
      (item) => item.to > item.from,
    );
    if (!incoming.length) return false;
    for (const item of incoming) {
      const node = session.getNode(item.nodeId);
      if (!node) continue;
      const expectedText = node.text.slice(item.from, item.to);
      if (!expectedText) continue;
      // Avoid exact duplicates or overlapping slices within the same node
      const exists = this.items.some(
        (existing) =>
          existing.nodeId === item.nodeId &&
          existing.from < item.to &&
          existing.to > item.from,
      );
      if (!exists) {
        this.items.push({
          nodeId: item.nodeId,
          from: item.from,
          to: item.to,
          expectedText,
        });
      }
    }
    return true;
  }

  paint(frame: Document, root: HTMLElement, ignored: (node: Node) => boolean): void {
    this.overlay?.remove();
    this.overlay = undefined;
    if (!this.items.length) return;
    const layer = frame.createElement("div");
    layer.className = "sdk-multi-select-layer";
    layer.dataset.sdkMultiSelectCount = String(this.items.length);
    layer.style.cssText = "position:absolute;inset:0;pointer-events:none;z-index:999";
    const bodyRect = frame.body.getBoundingClientRect();
    const scrollX = frame.defaultView?.scrollX ?? 0;
    const scrollY = frame.defaultView?.scrollY ?? 0;
    for (const item of this.items) {
      const host = root.querySelector<HTMLElement>(
        `[data-node-id="${CSS.escape(item.nodeId)}"]`,
      );
      if (!host) continue;
      try {
        const [startNode, startOffset] = caretAt(host, item.from, ignored);
        const [endNode, endOffset] = caretAt(host, item.to, ignored);
        const domRange = frame.createRange();
        domRange.setStart(startNode, startOffset);
        domRange.setEnd(endNode, endOffset);
        for (const rect of domRange.getClientRects()) {
          if (rect.width < 0.5 || rect.height < 0.5) continue;
          const mark = frame.createElement("div");
          mark.className = "sdk-multi-select-rect";
          mark.style.cssText =
            `position:absolute;pointer-events:none;background:rgba(24,90,189,0.24);border-radius:1px;` +
            `left:${rect.left - bodyRect.left + scrollX}px;` +
            `top:${rect.top - bodyRect.top + scrollY}px;` +
            `width:${rect.width}px;height:${rect.height}px`;
          layer.append(mark);
        }
      } catch {
        // Ignore stale range during intermediate re-render.
      }
    }
    if (layer.childNodes.length) {
      frame.body.append(layer);
      this.overlay = layer;
    }
  }

  private verifiedRanges(session: DocumentEditingSession): WordMultiTextRange[] {
    const valid: WordMultiTextRange[] = [];
    for (const item of this.items) {
      const node = session.getNode(item.nodeId);
      if (!node || item.to > node.text.length) continue;
      const text = node.text.slice(item.from, item.to);
      if (!text || text !== item.expectedText) continue;
      valid.push(item);
    }
    // Sort descending within each paragraph so mutations never shift earlier offsets.
    return valid.sort((a, b) =>
      a.nodeId === b.nodeId ? b.from - a.from : a.nodeId.localeCompare(b.nodeId),
    );
  }

  batchFormatOps(
    properties: Readonly<Record<string, string>>,
    session: DocumentEditingSession,
  ): DocumentOperation[] {
    return this.verifiedRanges(session).map((item) => ({
      op: "formatWordText" as const,
      nodeId: item.nodeId,
      from: item.from,
      to: item.to,
      expectedText: item.expectedText,
      properties: { ...properties },
    }));
  }

  batchResetOps(session: DocumentEditingSession): DocumentOperation[] {
    return this.verifiedRanges(session).map((item) => ({
      op: "resetWordText" as const,
      nodeId: item.nodeId,
      from: item.from,
      to: item.to,
      expectedText: item.expectedText,
      resetParagraph:
        item.from === 0 && item.to === (session.getNode(item.nodeId)?.text.length ?? -1),
    }));
  }

  batchDeleteOps(session: DocumentEditingSession): DocumentOperation[] {
    return this.verifiedRanges(session).map((item) => ({
      op: "replaceText" as const,
      nodeId: item.nodeId,
      from: item.from,
      to: item.to,
      expectedText: item.expectedText,
      text: "",
    }));
  }

  batchCaseOps(
    transform: (text: string) => string,
    session: DocumentEditingSession,
  ): DocumentOperation[] {
    const ops: DocumentOperation[] = [];
    for (const item of this.verifiedRanges(session)) {
      const next = transform(item.expectedText);
      if (next !== item.expectedText) {
        ops.push({
          op: "replaceText",
          nodeId: item.nodeId,
          from: item.from,
          to: item.to,
          expectedText: item.expectedText,
          text: next,
        });
      }
    }
    return ops;
  }

  plainText(session: DocumentEditingSession): string {
    return this.items
      .map((item) => session.getNode(item.nodeId)?.text.slice(item.from, item.to) ?? "")
      .filter(Boolean)
      .join("\n");
  }
}
