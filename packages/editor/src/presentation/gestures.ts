import { resolvePresentationGeometry } from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation, PresentationGeometry } from "@doffi/core";
import type { PreviewSource } from "@doffi/render";
export interface ObjectGestureHost {
    readonly session: DocumentEditingSession;
    readonly source: PreviewSource;
    selected(): readonly string[];
    select(ids: readonly string[]): void;
    current(): boolean;
    perform(operations: DocumentOperation[]): void;
    error(error: unknown): void;
}
export interface ObjectGestures {
    paint(): void;
    cancel(): void;
    dispose(): void;
}
type Kind = "move" | "resize" | "rotate";
interface Item {
    id: string;
    geometry: Readonly<PresentationGeometry>;
    source: Readonly<PresentationGeometry>;
    autoFitShape: boolean;
    coordinateScale: number;
    parent: DOMMatrix;
    inverse: DOMMatrix;
    start: DOMPoint;
}
interface Drag {
    kind: Kind;
    revision: number;
    items: Item[];
    startX: number;
    startY: number;
    changed: boolean;
    target: Element;
    pointerId: number;
    next: Map<string, PresentationGeometry>;
}
/** A short-lived SVG probe measures the renderer's actual parent coordinate system. */
function matrix(node: HTMLElement, coordinateScale = 1): DOMMatrix {
    const frame = node.ownerDocument, probe = frame.createElementNS("http://www.w3.org/2000/svg", "svg");
    // 750 pt is exactly 1000 CSS px: a 1 pt probe suffers layout-unit rounding.
    probe.setAttribute("viewBox", "0 0 750 750");
    probe.style.cssText = "position:absolute;left:0;top:0;width:750pt;height:750pt;visibility:hidden;pointer-events:none";
    node.append(probe);
    const result = probe.getScreenCTM();
    probe.remove();
    if (!result)
        throw new Error("对象坐标暂不可测量。");
    return result.scale(coordinateScale);
}
function transformed(geometry: Readonly<PresentationGeometry>, parent: DOMMatrix): DOMMatrix {
    return parent.translate(geometry.x + geometry.width / 2, geometry.y + geometry.height / 2).rotate(geometry.rotation).scale(geometry.flipH ? -1 : 1, geometry.flipV ? -1 : 1).translate(-geometry.width / 2, -geometry.height / 2);
}
function point(frame: Document, x: number, y: number, matrix: DOMMatrix): DOMPoint { return new frame.defaultView!.DOMPoint(x, y).matrixTransform(matrix); }
function equal(a: Readonly<PresentationGeometry>, b: Readonly<PresentationGeometry>): boolean {
    return Math.abs(a.x - b.x) < 1e-8 && Math.abs(a.y - b.y) < 1e-8 && Math.abs(a.width - b.width) < 1e-8 && Math.abs(a.height - b.height) < 1e-8 && Math.abs((a.rotation - b.rotation) % 360) < 1e-8 && a.flipH === b.flipH && a.flipV === b.flipV;
}
/** Gestures draw disposable outlines and commit one complete transaction at pointer-up. */
export function bindObjectGestures(frame: Document, host: ObjectGestureHost): ObjectGestures {
    const controller = new AbortController(), overlays: HTMLElement[] = [];
    let drag: Drag | undefined, disposed = false;
    const remove = (): void => { for (const node of overlays)
        node.remove(); overlays.length = 0; };
    const element = (id: string): HTMLElement | undefined => host.source.elements.get(id)?.element;
    const overlay = (css: string): HTMLDivElement => { const node = frame.createElement("div"); node.dataset.sdkPptSelection = ""; node.style.cssText = "all:initial!important;position:fixed!important;box-sizing:border-box!important;z-index:2147483646!important;" + css; frame.documentElement.append(node); overlays.push(node); return node; };
    const paint = (): void => {
        remove();
        if (disposed || !host.current())
            return;
        const ids = host.selected();
        for (const id of ids) {
            const node = element(id);
            if (!node)
                continue;
            const g = drag?.next.get(id), item = drag?.items.find(item => item.id === id);
            let box: {
                left: number;
                top: number;
                width: number;
                height: number;
            };
            if (g && item) {
                const m = transformed(g, item.parent), points = [[0, 0], [g.width, 0], [0, g.height], [g.width, g.height]].map(([x, y]) => point(frame, x!, y!, m));
                const left = Math.min(...points.map(p => p.x)), top = Math.min(...points.map(p => p.y));
                box = { left, top, width: Math.max(...points.map(p => p.x)) - left, height: Math.max(...points.map(p => p.y)) - top };
            }
            else
                box = node.getBoundingClientRect();
            overlay(`pointer-events:none!important;border:2px ${drag ? "dashed" : "solid"} #245aa5!important;background:${drag ? "rgba(36,90,165,.08)" : "transparent"}!important;left:${box.left}px!important;top:${box.top}px!important;width:${Math.max(box.width, 2)}px!important;height:${Math.max(box.height, 2)}px!important;`);
            if (ids.length === 1 && !drag) {
                const coordinateScale = Number(node.dataset.sdkPptCoordinateScale ?? 1);
                const original = resolvePresentationGeometry(host.session, id), style = frame.defaultView!.getComputedStyle(node), auto = node.getAttribute("data-sdk-text-auto-fit") === "shape", width = parseFloat(style.width) * .75 / coordinateScale, height = parseFloat(style.height) * .75 / coordinateScale;
                const g = auto ? { ...original, width: Math.abs(width - original.width) * coordinateScale > .02 ? width : original.width, height: Math.abs(height - original.height) * coordinateScale > .02 ? height : original.height } : original, m = matrix(node, coordinateScale);
                for (const [kind, x, y] of [["resize", g.width, g.height], ["rotate", g.width / 2, -24 / coordinateScale]] as const) {
                    const p = point(frame, x, y, m), handle = overlay(`pointer-events:auto!important;cursor:${kind === "resize" ? "nwse-resize" : "grab"}!important;left:${p.x - 5}px!important;top:${p.y - 5}px!important;width:10px!important;height:10px!important;background:white!important;border:2px solid #245aa5!important;border-radius:${kind === "rotate" ? "50%" : "0"}!important;`);
                    handle.dataset.sdkPptHandle = kind;
                    handle.dataset.sdkPptObject = id;
                    handle.setAttribute("role", "button");
                    handle.setAttribute("aria-label", kind === "resize" ? "调整对象大小" : "旋转对象");
                }
            }
        }
    };
    const cancel = (): void => {
        const previous = drag;
        drag = undefined;
        if (previous)
            try {
                previous.target.releasePointerCapture(previous.pointerId);
            }
            catch { /* Capture may already have ended. */ }
        paint();
    };
    const start = (event: PointerEvent): void => {
        if (disposed || !host.current() || event.button !== 0)
            return;
        const target = event.target as Element | null;
        if (!target?.closest || target.closest('[contenteditable="true"],input,textarea,select,button,summary'))
            return;
        const handle = target.closest<HTMLElement>("[data-sdk-ppt-handle]"), hit = handle ?? target.closest<HTMLElement>('[data-office-kind]');
        const id = handle?.dataset.sdkPptObject ?? hit?.dataset.nodeId, object = id && host.session.snapshot().metadata.presentation!.objects.find(o => o.nodeId === id);
        if (!object || !id || !hit || !handle && element(id) !== hit) {
            if (target.closest('[data-office-kind="slide"]'))
                host.select([]);
            return;
        }
        try {
            if (!handle) {
                const selected = host.selected();
                if (event.shiftKey) {
                    event.preventDefault();
                    frame.defaultView!.focus();
                    host.select(selected.includes(id) ? selected.filter(value => value !== id) : [...selected, id]);
                    return;
                }
                if (!selected.includes(id))
                    host.select([id]);
            }
            const items = host.selected().map(id => {
                const node = element(id)!, coordinateScale = Number(node.dataset.sdkPptCoordinateScale ?? 1), parent = matrix(node.parentElement!, coordinateScale), inverse = parent.inverse();
                const source = resolvePresentationGeometry(host.session, id), style = frame.defaultView!.getComputedStyle(node), autoFitShape = node.getAttribute("data-sdk-text-auto-fit") === "shape", width = parseFloat(style.width) * .75 / coordinateScale, height = parseFloat(style.height) * .75 / coordinateScale;
                const geometry = autoFitShape ? { ...source, width: Math.abs(width - source.width) * coordinateScale > .02 ? width : source.width, height: Math.abs(height - source.height) * coordinateScale > .02 ? height : source.height } : source;
                return { id, geometry, source, autoFitShape, coordinateScale, parent, inverse, start: point(frame, event.clientX, event.clientY, inverse) };
            });
            if (!items.length)
                return;
            const capture = element(id)!;
            drag = { kind: handle?.dataset.sdkPptHandle as Kind ?? "move", revision: host.session.snapshot().manifest.revision, items, startX: event.clientX, startY: event.clientY, changed: false, target: capture, pointerId: event.pointerId, next: new Map() };
            capture.setPointerCapture(event.pointerId);
            event.preventDefault();
            frame.defaultView!.focus();
            paint();
        }
        catch (error) {
            cancel();
            host.error(error);
        }
    };
    const move = (event: PointerEvent): void => {
        if (!drag || event.pointerId !== drag.pointerId)
            return;
        if (!host.current() || drag.revision !== host.session.snapshot().manifest.revision) {
            cancel();
            return;
        }
        if (Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) < 2 && !drag.changed)
            return;
        drag.changed = true;
        for (const item of drag.items) {
            const p = point(frame, event.clientX, event.clientY, item.inverse), g = item.geometry, dx = p.x - item.start.x, dy = p.y - item.start.y;
            let next: PresentationGeometry = { ...g };
            if (drag.kind === "move")
                next = { ...g, x: g.x + dx, y: g.y + dy };
            else if (drag.kind === "rotate") {
                const cx = g.x + g.width / 2, cy = g.y + g.height / 2, angle = (Math.atan2(p.y - cy, p.x - cx) - Math.atan2(item.start.y - cy, item.start.x - cx)) * 180 / Math.PI;
                next.rotation = event.shiftKey ? Math.round((g.rotation + angle) / 15) * 15 : g.rotation + angle;
            }
            else {
                const r = g.rotation * Math.PI / 180, u = { x: Math.cos(r) * (g.flipH ? -1 : 1), y: Math.sin(r) * (g.flipH ? -1 : 1) }, v = { x: -Math.sin(r) * (g.flipV ? -1 : 1), y: Math.cos(r) * (g.flipV ? -1 : 1) };
                let width = Math.max(1 / item.coordinateScale, g.width + dx * u.x + dy * u.y), height = Math.max(1 / item.coordinateScale, g.height + dx * v.x + dy * v.y);
                if (event.shiftKey && g.width && g.height) {
                    const scale = Math.max(width / g.width, height / g.height);
                    width = g.width * scale;
                    height = g.height * scale;
                }
                const ax = g.x + g.width / 2 - u.x * g.width / 2 - v.x * g.height / 2, ay = g.y + g.height / 2 - u.y * g.width / 2 - v.y * g.height / 2;
                next = { ...g, x: ax + u.x * width / 2 + v.x * height / 2 - width / 2, y: ay + u.y * width / 2 + v.y * height / 2 - height / 2, width, height };
            }
            drag.next.set(item.id, next);
        }
        event.preventDefault();
        paint();
    };
    const end = (event: PointerEvent): void => {
        if (!drag || event.pointerId !== drag.pointerId)
            return;
        const previous = drag, valid = host.current() && previous.revision === host.session.snapshot().manifest.revision;
        cancel();
        if (!valid || !previous.changed)
            return;
        const operations: DocumentOperation[] = previous.items.flatMap(item => {
            const geometry = previous.next.get(item.id);
            if (!geometry || equal(item.geometry, geometry))
                return [];
            const value = previous.kind === "resize" ? geometry : { ...item.source, x: geometry.x, y: geometry.y, rotation: geometry.rotation };
            const operations: DocumentOperation[] = [{ op: "setPresentationGeometry", nodeId: item.id, geometry: value }];
            if (previous.kind === "resize" && item.autoFitShape)
                operations.push({ op: "setPresentationTextProperties", nodeId: item.id, changes: { autoFit: "none" } });
            return operations;
        });
        if (operations.length)
            try {
                host.perform(operations);
            }
            catch (error) {
                host.error(error);
            }
    };
    frame.addEventListener("pointerdown", start, { signal: controller.signal });
    frame.addEventListener("pointermove", move, { signal: controller.signal });
    frame.addEventListener("pointerup", end, { signal: controller.signal });
    frame.addEventListener("pointercancel", cancel, { signal: controller.signal });
    frame.addEventListener("keydown", event => { if (event.key === "Escape")
        cancel(); }, { signal: controller.signal });
    frame.addEventListener("lostpointercapture", event => { if (drag?.pointerId === event.pointerId)
        cancel(); }, { signal: controller.signal });
    frame.defaultView!.addEventListener("blur", cancel, { signal: controller.signal });
    frame.defaultView!.addEventListener("resize", paint, { signal: controller.signal });
    frame.addEventListener("scroll", paint, { signal: controller.signal });
    return { paint, cancel, dispose() { if (disposed)
            return; disposed = true; drag = undefined; controller.abort(); remove(); } };
}
