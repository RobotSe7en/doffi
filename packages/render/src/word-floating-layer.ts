import { DocumentError } from '@doffi/core';
import type { ContentDocumentModel, DeepReadonly, WordDrawing } from '@doffi/core';
import type { RenderedPage, RenderLimits } from './contracts.js';
import type { PreviewSource } from './source-index.js';
import { cloneViewTree } from './view-clone.js';
import { markViewNode, ownViewStyle, viewNodeRole } from './view-nodes.js';
import { wordDrawingPosition } from './word-drawing-position.js';

/** Floating paint lives outside the clipped text flow. The original atomic
 * source stays at its insertion point, alongside a reversible layout anchor. */
export function wordFloatingLayer(document: Document, model: ContentDocumentModel, source: PreviewSource, limits: Readonly<RenderLimits>) {
    const entries: { drawing: DeepReadonly<WordDrawing>; frame: HTMLElement }[] = [];
    const sections = new Map(model.snapshot().metadata.sections?.map(section => [section.sectionId, section]));
    const containers = new Map<string, HTMLElement>(), paints = new Map<HTMLElement, HTMLElement>();
    let layer: HTMLElement | undefined;
    return {
        add(drawing: DeepReadonly<WordDrawing>, frame: HTMLElement): void {
            entries.push({ drawing, frame });
        },
        measure(pages: readonly RenderedPage[], root: HTMLElement): void {
            if (!entries.length || !pages.length) return;
            ownViewStyle(root, 'position', 'isolation'); root.style.position = 'relative'; root.style.isolation = 'isolate';
            if (!layer) {
                layer = markViewNode(document.createElement('div')); layer.className = 'sdk-word-floating-layer';
                layer.style.cssText = 'position:absolute;left:0;top:0;width:100%;height:0;pointer-events:none'; root.prepend(layer);
            }
            const origin = root.getBoundingClientRect(), anchors = new Map<string, HTMLElement[]>();
            for (const anchor of document.querySelectorAll<HTMLElement>('[data-sdk-word-anchor]')) {
                if (viewNodeRole(anchor) !== 'remove') continue;
                const id = anchor.dataset.sdkWordAnchor!, values = anchors.get(id) ?? []; values.push(anchor); anchors.set(id, values);
            }
            const pageAt = (y: number): RenderedPage | undefined => {
                const value = (y - origin.top) * .75; let low = 0, high = pages.length;
                while (low < high) { const middle = Math.floor((low + high) / 2); if (pages[middle]!.rect.y <= value + .075) low = middle + 1; else high = middle; }
                const page = pages[low - 1]; return page && !page.blank && value < page.rect.y + page.rect.height ? page : undefined;
            };
            const ordered = [...entries].sort((a, b) => a.frame === b.frame ? 0 : a.frame.compareDocumentPosition(b.frame) & 4 ? -1 : 1);
            const placements: { drawing: DeepReadonly<WordDrawing>; frame: HTMLElement; anchor: HTMLElement; paragraph: HTMLElement; page: RenderedPage; position: ReturnType<typeof wordDrawingPosition> }[] = [];
            for (const { drawing, frame } of ordered) {
                ownViewStyle(frame, 'visibility', 'pointer-events'); frame.style.visibility = 'hidden'; frame.style.pointerEvents = 'none';
                for (const anchor of anchors.get(drawing.nodeId) ?? []) {
                    if (anchor.closest('.sdk-auxiliary') || anchor.getClientRects().length === 0) continue;
                    const a = anchor.getBoundingClientRect(), paragraph = anchor.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6')!, page = pageAt(a.top);
                    if (page && paragraph) placements.push({ drawing, frame, anchor, paragraph, page, position: wordDrawingPosition(drawing, page, origin, paragraph, anchor, sections.get(page.sectionId)) });
                }
            }
            const retained = new Set(placements.map(placement => placement.anchor)), keptContainers = new Set<string>();
            for (const [anchor, paint] of paints) if (!retained.has(anchor)) { paint.remove(); paints.delete(anchor); }
            let count = document.querySelectorAll('*').length;
            // Read all source positions before inserting paint. Drawing layers
            // cannot perturb the next object's source measurements.
            for (const { drawing, frame, anchor, paragraph, page, position } of placements) {
                const key = page.pageId + (drawing.layout.behindText ? '-back' : '-front');
                keptContainers.add(key);
                let container = containers.get(key);
                if (!container) {
                    container = document.createElement('div'); container.className = 'sdk-word-floating-page'; container.dataset.pageId = page.pageId;
                    layer.append(container); containers.set(key, container); count++;
                }
                container.style.cssText = `position:absolute;left:${page.rect.x}pt;top:${page.rect.y}pt;width:${page.rect.width}pt;height:${page.rect.height}pt;overflow:hidden;pointer-events:none;z-index:${drawing.layout.behindText ? -1 : 1}`;
                let paint = paints.get(anchor);
                if (!paint) {
                    paint = markViewNode(cloneViewTree(frame, limits.maxNodes)); paints.set(anchor, paint);
                    count += paint.querySelectorAll('*').length + 1;
                    if (count > limits.maxNodes) throw new DocumentError('ResourceLimitExceeded', 'Floating Word objects exceed the render node budget.');
                }
                paint.dataset.sdkWordSourceParagraph = paragraph.dataset.nodeId!;
                paint.dataset.sdkWordPositionX = String(position.offsetX); paint.dataset.sdkWordPositionY = String(position.offsetY);
                paint.style.visibility = 'visible'; paint.style.pointerEvents = 'auto'; paint.style.position = 'absolute';
                paint.style.left = position.left - origin.left - page.rect.x / .75 + 'px'; paint.style.top = position.top - origin.top - page.rect.y / .75 + 'px';
                container.append(paint);
            }
            for (const [key, container] of containers) if (!keptContainers.has(key)) { container.remove(); containers.delete(key); }
            source.registerInstances(layer);
        },
    };
}
