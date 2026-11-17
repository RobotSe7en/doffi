import { DocumentError, wordDrawingTransforms } from "@doffi/core";
import type { ContentDocumentModel, DeepReadonly, WordDrawing } from "@doffi/core";
import type { LayoutDiagnostic, RenderedPage, RenderLimits } from "./contracts.js";
import type { PreviewSource } from "./source-index.js";
import { cloneViewTree } from "./view-clone.js";
import {
  hideFromViewSelection,
  markViewNode,
  measureViewContent,
  ownViewStyle,
  viewNodeRole,
} from "./view-nodes.js";
import { wordDrawingPosition, wordDrawingLeadingTextRect } from "./word-drawing-position.js";
import { avoidWordDrawingOverlap, wordWrapPoints } from "./word-wrap-geometry.js";
import type { WrapBounds } from "./word-wrap-geometry.js";
import { wordDrawingDimensions as wordDrawingSize } from "@doffi/core";

/** Floating paint lives outside the clipped text flow. The original atomic
 * source stays at its insertion point, alongside a reversible layout anchor. */
export function wordFloatingLayer(
  document: Document,
  model: ContentDocumentModel,
  source: PreviewSource,
  limits: Readonly<RenderLimits>,
  resize: (
    frame: HTMLElement,
    drawing: DeepReadonly<WordDrawing>,
    geometry: WordDrawing["geometry"],
  ) => void,
) {
  const entries: { drawing: DeepReadonly<WordDrawing>; frame: HTMLElement }[] = [];
  const sections = new Map(
    model.snapshot().metadata.sections?.map((section) => [section.sectionId, section]),
  );
  const containers = new Map<string, HTMLElement>(),
    paints = new Map<HTMLElement, HTMLElement>(),
    cellSpaces = new Map<HTMLElement, HTMLElement>(),
    leadingPages = new Map<HTMLElement, { delta: number; index: number; rect: DOMRect } | null>();
  let layer: HTMLElement | undefined;
  let changed = false;
  let diagnostics: readonly LayoutDiagnostic[] = [];
  return {
    get diagnostics(): readonly LayoutDiagnostic[] {
      return diagnostics;
    },
    get active(): boolean {
      return entries.length > 0;
    },
    get changed(): boolean {
      return changed;
    },
    reset(): void {
      for (const space of cellSpaces.values()) space.remove();
      cellSpaces.clear();
      leadingPages.clear();
    },
    add(drawing: DeepReadonly<WordDrawing>, frame: HTMLElement): void {
      entries.push({ drawing, frame });
    },
    measure(pages: readonly RenderedPage[], root: HTMLElement): void {
      changed = false;
      const overflow: LayoutDiagnostic[] = [];
      diagnostics = overflow;
      if (!entries.length || !pages.length) return;
      ownViewStyle(root, "position", "isolation");
      root.style.position = "relative";
      root.style.isolation = "isolate";
      if (!layer) {
        layer = markViewNode(document.createElement("div"));
        layer.className = "sdk-word-floating-layer";
        layer.style.cssText =
          "position:absolute;left:0;top:0;width:100%;height:0;pointer-events:none";
        root.prepend(layer);
      }
      const origin = root.getBoundingClientRect(),
        anchors = new Map<string, HTMLElement[]>();
      for (const anchor of document.querySelectorAll<HTMLElement>("[data-sdk-word-anchor]")) {
        if (viewNodeRole(anchor) !== "remove") continue;
        const id = anchor.dataset.sdkWordAnchor!,
          values = anchors.get(id) ?? [];
        values.push(anchor);
        anchors.set(id, values);
      }
      const pageAt = (y: number): RenderedPage | undefined => {
        const value = (y - origin.top) * 0.75;
        let low = 0,
          high = pages.length;
        while (low < high) {
          const middle = Math.floor((low + high) / 2);
          if (pages[middle]!.rect.y <= value + 0.075) low = middle + 1;
          else high = middle;
        }
        const page = pages[low - 1];
        return page && !page.blank && value < page.rect.y + page.rect.height ? page : undefined;
      };
      const ordered = [...entries].sort((a, b) =>
        a.frame === b.frame ? 0 : a.frame.compareDocumentPosition(b.frame) & 4 ? -1 : 1,
      );
      const placements: {
        drawing: DeepReadonly<WordDrawing>;
        frame: HTMLElement;
        anchor: HTMLElement;
        paragraph: HTMLElement;
        page: RenderedPage;
        position: ReturnType<typeof wordDrawingPosition>;
        geometry: WordDrawing["geometry"];
      }[] = [];
      const placed = new Map<string, (WrapBounds & { allowOverlap: boolean })[]>();
      const projections = new Map<string, ReturnType<typeof wordDrawingTransforms>>();
      const measuredGeometry = (
        drawing: DeepReadonly<WordDrawing>,
        page: RenderedPage,
      ): WordDrawing["geometry"] => {
        if (
          drawing.kind !== "connector" ||
          (!drawing.connector.startNodeId && !drawing.connector.endNodeId)
        )
          return wordDrawingSize(drawing, page, sections.get(page.sectionId));
        let transforms = projections.get(page.pageId);
        if (!transforms) {
          const word = model.snapshot().metadata.word!;
          transforms = wordDrawingTransforms(
            {
              ...word,
              drawings: word.drawings!.map((item) => ({
                ...item,
                geometry: wordDrawingSize(item, page, sections.get(page.sectionId)),
              })),
            },
            model,
          );
          projections.set(page.pageId, transforms);
        }
        return transforms.connectorGeometry(drawing.nodeId);
      };
      let work = 0;
      const reserve = () => {
        if (++work > limits.maxNodes * 32)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Floating object collisions exceed the render measurement budget.",
          );
      };
      for (const { drawing, frame } of ordered) {
        ownViewStyle(frame, "visibility", "pointer-events");
        frame.style.visibility = "hidden";
        frame.style.pointerEvents = "none";
        for (const anchor of anchors.get(drawing.nodeId) ?? []) {
          if (anchor.closest(".sdk-auxiliary") || anchor.getClientRects().length === 0) continue;
          const a = anchor.getBoundingClientRect(),
            paragraph = anchor.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6")!,
            rawPage = pageAt(a.top);
          // Freeze a leading anchor's page correction before exclusions are
          // inserted. It must not chase text pushed down by its own wrap box.
          if (!leadingPages.has(anchor)) {
            const first = paragraph ? wordDrawingLeadingTextRect(paragraph, anchor) : undefined,
              firstPage = first ? pageAt(first.top + Math.min(first.height / 2, 3)) : undefined;
            leadingPages.set(
              anchor,
              first && firstPage && (!rawPage || firstPage.index > rawPage.index)
                ? {
                    delta: firstPage.index - (rawPage?.index ?? 0),
                    index: firstPage.index,
                    rect: first,
                  }
                : null,
            );
          }
          const correction = leadingPages.get(anchor),
            page = correction
              ? pages[rawPage ? rawPage.index + correction.delta : correction.index]
              : rawPage;
          if (page && paragraph) {
            const geometry = measuredGeometry(drawing, page),
              position = wordDrawingPosition(
                { ...drawing, geometry },
                page,
                origin,
                paragraph,
                anchor,
                sections.get(page.sectionId),
                correction?.rect,
              ),
              cell =
                drawing.layout.layoutInCell !== false
                  ? paragraph.closest<HTMLElement>("td,th")
                  : null,
              scope = page.pageId + ":" + (cell?.dataset.nodeId ?? "main"),
              previous = placed.get(scope) ?? [],
              points = wordWrapPoints(
                [
                  { x: 0, y: 0 },
                  { x: 21600, y: 0 },
                  { x: 21600, y: 21600 },
                  { x: 0, y: 21600 },
                ],
                {
                  left: position.left,
                  top: position.top,
                  width: geometry.width / 0.75,
                  height: geometry.height / 0.75,
                },
                drawing.geometry,
              ),
              bounds = {
                left: Math.min(...points.map((p) => p.x)),
                right: Math.max(...points.map((p) => p.x)),
                top: Math.min(...points.map((p) => p.y)),
                bottom: Math.max(...points.map((p) => p.y)),
              },
              delta = avoidWordDrawingOverlap(
                bounds,
                drawing.layout.allowOverlap !== false,
                previous,
                reserve,
              );
            position.top += delta;
            position.offsetY += delta * 0.75;
            if (
              delta > 0.1 &&
              bounds.bottom + delta > origin.top + (page.rect.y + page.rect.height) / 0.75 + 0.1
            )
              overflow.push(
                Object.freeze({
                  code: "ContentOverflow",
                  nodeId: drawing.nodeId,
                  message: "Collision avoidance moves this drawing beyond its anchor page.",
                }),
              );
            previous.push({
              ...bounds,
              top: bounds.top + delta,
              bottom: bounds.bottom + delta,
              allowOverlap: drawing.layout.allowOverlap !== false,
            });
            placed.set(scope, previous);
            placements.push({
              drawing,
              frame,
              anchor,
              paragraph,
              page,
              position,
              geometry,
            });
          }
        }
      }
      const retained = new Set(placements.map((placement) => placement.anchor)),
        keptContainers = new Set<string>();
      for (const [anchor, paint] of paints)
        if (!retained.has(anchor)) {
          paint.remove();
          paints.delete(anchor);
        }
      let count = document.querySelectorAll("*").length;
      const cellBottoms = new Map<HTMLElement, number>();
      // Read all source positions before inserting paint. Drawing layers
      // cannot perturb the next object's source measurements.
      for (const { drawing, frame, anchor, paragraph, page, position, geometry } of placements) {
        const key = page.pageId + (drawing.layout.behindText ? "-back" : "-front");
        keptContainers.add(key);
        let container = containers.get(key);
        if (!container) {
          container = document.createElement("div");
          container.className = "sdk-word-floating-page";
          container.dataset.pageId = page.pageId;
          layer.append(container);
          containers.set(key, container);
          count++;
        }
        container.style.cssText = `position:absolute;left:${page.rect.x}pt;top:${page.rect.y}pt;width:${page.rect.width}pt;height:${page.rect.height}pt;overflow:hidden;pointer-events:none;z-index:${drawing.layout.behindText ? -1 : 1}`;
        let paint = paints.get(anchor);
        if (!paint) {
          paint = markViewNode(cloneViewTree(frame, limits.maxNodes));
          paints.set(anchor, paint);
          count += paint.querySelectorAll("*").length + 1;
          if (count > limits.maxNodes)
            throw new DocumentError(
              "ResourceLimitExceeded",
              "Floating Word objects exceed the render node budget.",
            );
        }
        paint.dataset.sdkWordSourceParagraph = paragraph.dataset.nodeId!;
        resize(paint, drawing, geometry);
        paint.dataset.sdkWordPositionX = String(position.offsetX);
        paint.dataset.sdkWordPositionY = String(position.offsetY);
        paint.dataset.sdkWordLayoutLeft = String(position.left - origin.left);
        paint.dataset.sdkWordLayoutTop = String(position.top - origin.top);
        const cell =
          drawing.layout.layoutInCell !== false ? paragraph.closest<HTMLElement>("td,th") : null;
        if (cell) paint.dataset.sdkWordWrapCell = cell.dataset.nodeId!;
        else delete paint.dataset.sdkWordWrapCell;
        paint.style.visibility = "visible";
        paint.style.pointerEvents = "auto";
        paint.style.position = "absolute";
        paint.style.left = position.left - origin.left - page.rect.x / 0.75 + "px";
        paint.style.top = position.top - origin.top - page.rect.y / 0.75 + "px";
        container.append(paint);
        if (cell && drawing.layout.wrap !== "none" && !drawing.layout.behindText) {
          const bottom =
            paint.getBoundingClientRect().bottom + (drawing.layout.distances?.bottom ?? 0) / 0.75;
          cellBottoms.set(cell, Math.max(cellBottoms.get(cell) ?? -Infinity, bottom));
        }
      }
      // Reserve only the space below real cell content. This participates in
      // row measurement without adding an authored node or changing row rules.
      for (const [cell, bottom] of cellBottoms) {
        let space = cellSpaces.get(cell);
        if (!space) {
          if (++count > limits.maxNodes)
            throw new DocumentError(
              "ResourceLimitExceeded",
              "Floating cell layout exceeds the render node budget.",
            );
          space = markViewNode(document.createElement("div"));
          measureViewContent(space);
          hideFromViewSelection(space);
          space.contentEditable = "false";
          space.className = "sdk-word-cell-float-space";
          space.style.cssText =
            "height:0;padding:0;margin:0;border:0;pointer-events:none;user-select:none";
          cell.append(space);
          cellSpaces.set(cell, space);
        }
        const height = Math.max(0, bottom - space.getBoundingClientRect().top);
        if (Math.abs((parseFloat(space.style.height) || 0) - height) > 0.1) {
          space.style.height = height + "px";
          changed = true;
        }
      }
      for (const [key, container] of containers)
        if (!keptContainers.has(key)) {
          container.remove();
          containers.delete(key);
        }
      source.registerInstances(layer);
    },
  };
}
